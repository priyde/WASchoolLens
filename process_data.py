"""
WASchoolLens — OSPI Data Pipeline v2
=====================================
Converts raw OSPI Washington Report Card Excel files into clean JSON
for the WASchoolLens dashboard.

Changes in v2:
  - Joins on SchoolCode (integer) instead of school name strings — robust
    to name variations, double-spaces, and future renames
  - Filters districts by DistrictCode (integer) instead of DistrictName
  - WaKIDS (which lacks SchoolCode) joined via SchoolOrganizationId bridge
  - Feeder chains read from FeederFlow.xlsx — no longer hardcoded
  - All 4 districts processed in one run

Usage:
  1. Place OSPI Excel files and FeederFlow.xlsx in INPUT_DIR (default: same folder)
  2. Run:  python process_data.py
  3. JSON files appear in OUTPUT_DIR (default: ./data/)

Download data from: https://ospi.k12.wa.us/data-reporting/data-portal
"""

import json, os, re, sys
import numpy as np
import pandas as pd

# ── CONFIG ──────────────────────────────────────────────────────────────────

INPUT_DIR  = "."       # folder containing OSPI Excel files + FeederFlow.xlsx
OUTPUT_DIR = "data"    # folder where JSON files are written

SCHOOL_YEAR = "2024-25"

# Update file names each year when OSPI releases new data
FILES = {
    "enrollment":  "Enrollment_202425_20260610.xlsx",
    "assessment":  "Assessment_Data_202425_20260610.xlsx",
    "growth":      "StudentGrowth_202425_20260610.xlsx",
    "discipline":  "Discipline_202425_20260610.xlsx",
    "graduation":  "Graduation_202425_20260610.xlsx",
    "sqss":        "SQSS_202425_20260610.xlsx",
    "wakids":      "WaKidsKindergarten_202425_20260610.xlsx",
    "feeders":     "FeederFlow.xlsx",
    "help_assessment": "HELP_Assessment_Data.xlsx",  # contains SchoolTypeCodes tab
}

# slug → DistrictCode (OSPI integer identifier)
DISTRICTS = {
    "lwsd":  17414,   # Lake Washington School District
    "bsd":   17405,   # Bellevue School District
    "nsd":   17417,   # Northshore School District
    "isd":   17411,   # Issaquah School District
    "sps":   17001,   # Seattle Public Schools
    "shsd":  17412,   # Shoreline School District
    "esd":   31015,   # Edmonds School District
    "eps":   31002,   # Everett School District
    "snosd": 31201,   # Snohomish School District
    "rvsd":  17407,   # Riverview School District
    "rsd":   17403,   # Renton School District
    "misd":  17400,   # Mercer Island School District
    # Add more districts here — same Excel files, no other changes needed
}

DISTRICT_NAMES = {
    17414: "Lake Washington School District",
    17405: "Bellevue School District",
    17417: "Northshore School District",
    17411: "Issaquah School District",
    17001: "Seattle Public Schools",
    17412: "Shoreline School District",
    31015: "Edmonds School District",
    31002: "Everett School District",
    31201: "Snohomish School District",
    17407: "Riverview School District",
    17403: "Renton School District",
    17400: "Mercer Island School District",
}

# ── HELPERS ─────────────────────────────────────────────────────────────────

def safe_int(v):
    try:
        f = float(v)
        return int(f) if not np.isnan(f) else None
    except (TypeError, ValueError):
        return None

def parse_pct(x):
    """Parse percentage-like values to float 0-1."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, float)):
        v = float(x)
        # Values already in 0-1 range (e.g. GraduationRate 0.92)
        return round(v, 4) if 0 <= v <= 1 else round(v / 100, 4)
    s = str(x).strip().replace('%', '').replace('<', '').replace('>', '')
    try:
        return round(float(s) / 100, 4)
    except ValueError:
        return None

def safe_div(num, den):
    try:
        n, d = float(num), float(den)
        if d == 0 or np.isnan(n) or np.isnan(d):
            return None
        return round(n / d, 4)
    except (TypeError, ValueError):
        return None

GRADE_ORDER = [
    'Pre-Kindergarten', 'Transition to Kindergarten', 'Half-Day Kindergarten', 'Kindergarten',
    '1st Grade', '2nd Grade', '3rd Grade', '4th Grade', '5th Grade',
    '6th Grade', '7th Grade', '8th Grade',
    '9th Grade', '10th Grade', '11th Grade', '12th Grade',
]
ELEM_GRADES = set(GRADE_ORDER[:9])   # Pre-K through 5th
MID_GRADES  = set(GRADE_ORDER[9:12]) # 6th-8th
HIGH_GRADES = set(GRADE_ORDER[12:])  # 9th-12th

# Fallback keyword list, only used for a school if OSPI's CurrentSchoolType
# field is missing entirely for that school (should be rare).
ALT_KEYWORDS_FALLBACK = [
    'community school', 'renaissance', 'stella schola',
    'environmental', 'old redmond', 'skill center',
    'international community', 'international school',
    'discovery community', 'explorer community', 'contractual',
    'big picture', 'digital discovery', 'open doors', 'reengagement',
    'learning options', 'online', 'special services',
    'secondary academy', 'community center', 'echo glen',
    'gibson ek', 'innovation lab', 'tesla stem', 'nikola tesla',
]

def is_alt_by_name(name):
    n = name.lower()
    return any(kw in n for kw in ALT_KEYWORDS_FALLBACK)

def read_school_type_codes(help_assessment_path):
    """Read the SchoolTypeCodes tab from HELP_Assessment_Data.xlsx.
    Returns {code_str: description_str}"""
    df = pd.read_excel(help_assessment_path, sheet_name='SchoolTypeCodes')
    code_col = df.columns[0]   # 'School Type Code'
    desc_col  = df.columns[1]  # 'School Type'
    result = {}
    for _, row in df.iterrows():
        code = str(row[code_col]).strip()
        desc = str(row[desc_col]).strip()
        if code and code != 'nan':
            result[code] = desc
    return result


def build_school_type_lookup(enr_df, school_type_codes):
    """OSPI's CurrentSchoolType field marks each school as 'P' (Public School /
    regular attendance-area school) or another code. Anything other than 'P' is
    treated as a choice/alternative/special program.
    Returns {(dist_code, school_code): {'is_alt': bool, 'code': str, 'desc': str}}"""
    d = enr_df[
        enr_df['SchoolCode'].notna() &
        (enr_df['GradeLevel'] == 'All Grades')
    ]
    out = {}
    for _, r in d.iterrows():
        dc = safe_int(r.get('DistrictCode'))
        sc = safe_int(r.get('SchoolCode'))
        st = r.get('CurrentSchoolType')
        if dc and sc and isinstance(st, str) and st.strip():
            code = st.strip()
            out[(dc, sc)] = {
                'is_alt': code.upper() != 'P',
                'code':   code,
                'desc':   school_type_codes.get(code, ''),
            }
    return out

def classify_by_name(name):
    """Fallback single-band guess, used only if a school has no per-grade
    enrollment data at all (e.g. fully suppressed)."""
    n = name.lower()
    if any(x in n for x in ['high school', 'senior high', ' hs', 'stem high']):
        return 'High'
    if 'middle' in n:
        return 'Middle'
    return 'Elementary'

def bands_to_type(bands):
    """Convert a set of {'Elementary','Middle','High'} into (types list, display label)."""
    if bands == {'Elementary'}:
        return ['Elementary'], 'Elementary'
    if bands == {'Middle'}:
        return ['Middle'], 'Middle'
    if bands == {'High'}:
        return ['High'], 'High'
    if bands == {'Elementary', 'Middle'}:
        return ['Elementary', 'Middle'], 'K-8'
    if bands == {'Middle', 'High'}:
        return ['Middle', 'High'], '6-12'
    if bands >= {'Elementary', 'Middle', 'High'} or bands == {'Elementary', 'High'}:
        return ['Elementary', 'Middle', 'High'], 'K-12'
    return [], None  # empty / no grade data

def jclean(v):
    """Make value safe for JSON serialisation."""
    if v is None:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return None if np.isnan(v) else float(v)
    if isinstance(v, float) and np.isnan(v):
        return None
    return v

# ── REFERENCE TABLES ────────────────────────────────────────────────────────

def build_reference_tables(enr_df):
    """
    Build lookup tables from the enrollment DataFrame:
      school_names     : {dist_code → {school_code → school_name}}
      org_id_bridge    : {(dist_code, school_org_id) → school_code}
      dist_org_ids     : {dist_code → dist_org_id}   (for WaKIDS)
    """
    school_names  = {}
    org_id_bridge = {}
    dist_org_ids  = {}

    # District-level rows for dist_org_id lookup (SchoolCode is null on these)
    dist_rows = enr_df[enr_df['SchoolCode'].isna() & (enr_df['GradeLevel'] == 'All Grades')]
    for _, r in dist_rows.iterrows():
        dc  = safe_int(r.get('DistrictCode'))
        doi = safe_int(r.get('DistrictOrganizationId'))
        if dc and doi:
            dist_org_ids[dc] = doi

    # School-level rows
    school_rows = enr_df[
        enr_df['SchoolCode'].notna() &
        (enr_df['GradeLevel'] == 'All Grades')
    ]
    for _, r in school_rows.iterrows():
        dc  = safe_int(r.get('DistrictCode'))
        sc  = safe_int(r.get('SchoolCode'))
        soi = safe_int(r.get('SchoolOrganizationID'))
        nm  = str(r.get('SchoolName', '')).strip()
        if dc and sc and nm:
            school_names.setdefault(dc, {})[sc] = nm
        if dc and sc and soi:
            org_id_bridge[(dc, soi)] = sc

    return school_names, org_id_bridge, dist_org_ids

# ── EXTRACTORS ──────────────────────────────────────────────────────────────

def extract_grade_bands(enr_df, dist_code):
    """Determine which grade bands (Elementary/Middle/High) each school
    actually serves, based on per-grade enrollment counts.
    Returns {school_code: {'types': [...], 'type_label': str}}"""
    d = enr_df[
        (enr_df['DistrictCode'] == dist_code) &
        enr_df['SchoolCode'].notna() &
        (enr_df['GradeLevel'] != 'All Grades')
    ]
    grades_by_school = {}
    for _, r in d.iterrows():
        sc = safe_int(r.get('SchoolCode'))
        grade = r.get('GradeLevel')
        n = r.get('All Students')
        if not sc or grade not in GRADE_ORDER:
            continue
        if n is None or (isinstance(n, float) and np.isnan(n)) or n <= 0:
            continue
        grades_by_school.setdefault(sc, set()).add(grade)

    out = {}
    for sc, grades in grades_by_school.items():
        bands = set()
        if grades & ELEM_GRADES:
            bands.add('Elementary')
        if grades & MID_GRADES:
            bands.add('Middle')
        if grades & HIGH_GRADES:
            bands.add('High')
        types, label = bands_to_type(bands)
        if types:
            out[sc] = {'types': types, 'type_label': label}
    return out


def extract_enrollment(enr_df, dist_code):
    """Returns {school_code: {...enrollment metrics...}}"""
    d = enr_df[
        (enr_df['DistrictCode'] == dist_code) &
        enr_df['SchoolCode'].notna() &
        (enr_df['GradeLevel'] == 'All Grades')
    ]
    out = {}
    for _, r in d.iterrows():
        sc = safe_int(r.get('SchoolCode'))
        if not sc:
            continue
        total = safe_int(r.get('All Students'))
        if not total or total == 0:
            continue
        out[sc] = {
            'enrollment': total,
            'frpl_pct':     safe_div(r.get('Low-Income'), total),
            'asian_pct':    safe_div(r.get('Asian'), total),
            'white_pct':    safe_div(r.get('White'), total),
            'hispanic_pct': safe_div(r.get('Hispanic/Latino of any race(s)'), total),
            'black_pct':    safe_div(r.get('Black/African American'), total),
            'multirace_pct':safe_div(r.get('Two or More Races'), total),
            'ell_pct':      safe_div(r.get('English Language Learners'), total),
            'sped_pct':     safe_div(r.get('Students with Disabilities'), total),
            'hc_pct':       safe_div(r.get('Highly Capable'), total),
        }
    return out


def extract_assessment(asmnt_df, dist_code):
    """Returns {school_code: {ela_prof, math_prof, sci_prof}}"""
    d = asmnt_df[
        (asmnt_df['DistrictCode'] == dist_code) &
        (asmnt_df['StudentGroup'] == 'All Students') &
        (asmnt_df['GradeLevel'] == 'All Grades')
    ]
    col = 'Percent Consistent Grade Level Knowledge And Above'
    out = {}
    spec = [('ELA', 'SBAC', 'ela_prof'),
            ('Math', 'SBAC', 'math_prof'),
            ('Science', 'WCAS', 'sci_prof')]
    for subject, admin, key in spec:
        sub = d[(d['TestSubject'] == subject) & (d['TestAdministration'] == admin)]
        for _, r in sub.iterrows():
            sc = safe_int(r.get('SchoolCode'))
            if sc:
                out.setdefault(sc, {})[key] = parse_pct(r.get(col))
    return out


def extract_growth(growth_df, dist_code):
    """Returns {school_code: {ela_sgp, math_sgp}}"""
    d = growth_df[
        (growth_df['DistrictCode'] == dist_code) &
        (growth_df['StudentGroup'] == 'All Students') &
        (growth_df['GradeLevel'] == 'All Grades')
    ]
    out = {}
    spec = [('English Language Arts', 'ela_sgp'), ('Math', 'math_sgp')]
    for subject, key in spec:
        sub = d[d['Subject'] == subject]
        for _, r in sub.iterrows():
            sc = safe_int(r.get('SchoolCode'))
            v  = r.get('MedianSGP')
            if sc:
                out.setdefault(sc, {})[key] = (
                    None if (v is None or (isinstance(v, float) and np.isnan(v)))
                    else round(float(v), 1)
                )
    return out


def extract_discipline(disc_df, dist_code):
    """Returns {school_code: {disc_rate, disc_suppressed}}"""
    d = disc_df[
        (disc_df['DistrictCode'] == dist_code) &
        (disc_df['Student Group'] == 'All Students') &
        (disc_df['GradeLevel'] == 'All')
    ]
    out = {}
    for _, r in d.iterrows():
        sc  = safe_int(r.get('SchoolCode'))
        raw = r.get('DisciplineRate')
        if not sc:
            continue
        suppressed = isinstance(raw, str) and ('<' in raw or '>' in raw)
        out[sc] = {
            'disc_rate':       parse_pct(raw),
            'disc_suppressed': bool(suppressed),
        }
    return out


def extract_graduation(grad_df, dist_code):
    """Returns {school_code: {grad_4yr}}"""
    d = grad_df[
        (grad_df['DistrictCode'] == dist_code) &
        (grad_df['StudentGroupType'] == 'All') &
        (grad_df['StudentGroup'] == 'All Students') &
        (grad_df['Cohort'] == 'Four Year')
    ]
    out = {}
    for _, r in d.iterrows():
        sc = safe_int(r.get('SchoolCode'))
        v  = r.get('GraduationRate')
        if sc:
            out[sc] = {'grad_4yr': parse_pct(v)}
    return out


def extract_sqss(sqss_df, dist_code):
    """Returns {school_code: {attendance_pct, ninth_on_track_pct, dual_credit_pct}}"""
    d = sqss_df[
        (sqss_df['DistrictCode'] == dist_code) &
        (sqss_df['StudentGroup'] == 'All Students') &
        (sqss_df['GradeLevel'] == 'All Grades')
    ]
    keymap = {
        'Regular Attendance':   'attendance_pct',
        'Ninth Grade on Track': 'ninth_on_track_pct',
        'Dual Credit':          'dual_credit_pct',
    }
    out = {}
    for measure, key in keymap.items():
        sub = d[d['Measure'] == measure]
        for _, r in sub.iterrows():
            sc = safe_int(r.get('SchoolCode'))
            v  = r.get('Percent')
            if sc:
                out.setdefault(sc, {})[key] = (
                    None if (v is None or (isinstance(v, float) and np.isnan(v)))
                    else round(float(v), 4)
                )
    return out


def extract_wakids(wakids_df, dist_code, dist_org_ids, org_id_bridge):
    """Returns {school_code: {kg_ready_pct}} using SchoolOrganizationId bridge."""
    dist_org_id = dist_org_ids.get(dist_code)
    if not dist_org_id:
        print(f'  WARNING: no DistrictOrganizationId found for DistrictCode {dist_code}')
        return {}

    d = wakids_df[
        (wakids_df['DistrictOrganizationId'] == dist_org_id) &
        wakids_df['SchoolOrganizationId'].notna() &
        (wakids_df['StudentGroup'] == 'All Students') &
        (wakids_df['Measure'] == 'NumberofDomainsReadyforKindergarten') &
        (wakids_df['MeasureValue'].astype(str).str.strip() == '6')
    ]
    out = {}
    for _, r in d.iterrows():
        soi = safe_int(r.get('SchoolOrganizationId'))
        v   = r.get('Percent')
        sc  = org_id_bridge.get((dist_code, soi)) if soi else None
        if sc:
            out[sc] = {
                'kg_ready_pct': (
                    None if (v is None or (isinstance(v, float) and np.isnan(v)))
                    else round(float(v), 4)
                )
            }
    return out


def extract_feeders(feeders_path, slug, school_name_lookup):
    """
    Read feeder chains from FeederFlow.xlsx.
    school_name_lookup: {school_code → canonical OSPI school name}
    Returns list of chain dicts.
    """
    sheet = slug.upper()
    try:
        df = pd.read_excel(feeders_path, sheet_name=sheet)
    except Exception as e:
        print(f'  WARNING: could not read feeder sheet "{sheet}": {e}')
        return []

    chains = []
    for _, row in df.iterrows():
        elem_code = safe_int(row.get('ElemSchoolCode'))
        ms_code   = safe_int(row.get('MSSchoolCode'))
        hs_code   = safe_int(row.get('HSSchoolCode'))

        # Use canonical OSPI names from lookup, fall back to Excel names
        elem = school_name_lookup.get(elem_code, str(row.get('Elementary School', '')).strip())
        ms   = school_name_lookup.get(ms_code,   str(row.get('Middle School', '')).strip())
        hs   = school_name_lookup.get(hs_code,   str(row.get('High School', '')).strip())

        note_raw = row.get('Notes', '')
        note = str(note_raw).strip() if (note_raw and str(note_raw).strip() not in ('nan', '')) else None

        chains.append({k: jclean(v) for k, v in {
            'elem_code': elem_code,
            'ms_code':   ms_code,
            'hs_code':   hs_code,
            'elem':      elem,
            'ms':        ms,
            'hs':        hs,
            'note':      note,
        }.items()})

    return chains

# ── MAIN BUILD ──────────────────────────────────────────────────────────────

def build_district(slug, dist_code, dfs, school_names_all, org_id_bridge, dist_org_ids, alt_type_lookup):
    print(f'\nProcessing {DISTRICT_NAMES.get(dist_code, dist_code)} ({slug}) ...')

    # School name lookup for this district
    school_name_lookup = school_names_all.get(dist_code, {})

    # Extract all metrics (keyed by SchoolCode)
    enrollment  = extract_enrollment(dfs['enrollment'],  dist_code)
    grade_bands = extract_grade_bands(dfs['enrollment'], dist_code)
    assessment  = extract_assessment(dfs['assessment'],  dist_code)
    growth      = extract_growth(    dfs['growth'],      dist_code)
    discipline  = extract_discipline(dfs['discipline'],  dist_code)
    graduation  = extract_graduation(dfs['graduation'],  dist_code)
    sqss        = extract_sqss(      dfs['sqss'],        dist_code)
    wakids      = extract_wakids(    dfs['wakids'],      dist_code,
                                     dist_org_ids, org_id_bridge)

    # Build school records (enrollment defines the canonical school list)
    schools = []
    for sc, base in sorted(enrollment.items(), key=lambda x: school_name_lookup.get(x[0], '')):
        name = school_name_lookup.get(sc, f'Unknown ({sc})')
        bands = grade_bands.get(sc)
        if bands:
            types, type_label = bands['types'], bands['type_label']
        else:
            # No per-grade enrollment data (fully suppressed) - fall back
            # to a single-band guess from the school name.
            label = classify_by_name(name)
            types, type_label = [label], label
        # Prefer OSPI's own CurrentSchoolType field (P = neighborhood school,
        # anything else = choice/alternative/reengagement/special program).
        # Only fall back to name-keyword matching if that field is missing.
        alt_info = alt_type_lookup.get((dist_code, sc))
        if alt_info is not None:
            is_alt         = alt_info['is_alt']
            school_type_code = alt_info['code']
            school_type_desc = alt_info['desc']
        else:
            is_alt         = is_alt_by_name(name)
            school_type_code = None
            school_type_desc = None
        rec  = {
            'school_code':     sc,
            'name':            name,
            'types':           types,
            'type_label':      type_label,
            'is_alt':          is_alt,
            'school_type_code': school_type_code if is_alt else None,
            'school_type_desc': school_type_desc if is_alt else None,
        }
        rec.update(base)
        for src in (assessment, growth, discipline, graduation, sqss, wakids):
            rec.update(src.get(sc, {}))
        schools.append({k: jclean(v) for k, v in rec.items()})

    # Feeder chains
    feeders_path = os.path.join(INPUT_DIR, FILES['feeders'])
    chains = extract_feeders(feeders_path, slug, school_name_lookup)

    # Validate chains against school list
    school_codes = {s['school_code'] for s in schools}
    valid_chains, skipped = [], 0
    for ch in chains:
        if ch['elem_code'] and ch['elem_code'] not in school_codes:
            # Elementary not in OSPI data (e.g. Bear Creek) — keep but flag
            pass
        valid_chains.append(ch)

    payload = {
        'district':      DISTRICT_NAMES.get(dist_code, str(dist_code)),
        'district_code': dist_code,
        'slug':          slug,
        'school_year':   SCHOOL_YEAR,
        'generated_by':  'WASchoolLens pipeline v2 (process_data.py)',
        'source':        'OSPI Washington State Report Card — https://ospi.k12.wa.us/data-reporting/data-portal',
        'school_count':  len(schools),
        'schools':       schools,
        'feeder_chains': valid_chains,
    }

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, f'{slug}.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=1)

    print(f'  → {out_path}  ({len(schools)} schools, {len(valid_chains)} feeder chains)')
    return {'slug': slug, 'district': payload['district'],
            'district_code': dist_code, 'school_year': SCHOOL_YEAR,
            'school_count': len(schools)}


def main():
    print('WASchoolLens Pipeline v2')
    print('=' * 50)

    # Verify files
    missing = [k for k, f in FILES.items()
               if not os.path.exists(os.path.join(INPUT_DIR, f))]
    if missing:
        print(f'ERROR: missing files: {missing}')
        print('Check INPUT_DIR and file names in the FILES config.')
        return 1

    print('Reading Excel files (this may take ~30 seconds) ...')
    # Read the first sheet in each file by position, not by name. OSPI's exports
    # have used inconsistent sheet names across releases ('Data' vs 'Sheet1'),
    # so relying on position is more robust than hardcoding a name.
    def read_first_sheet(label, path):
        xl = pd.ExcelFile(path)
        actual_name = xl.sheet_names[0]
        print(f'  {label}: reading sheet "{actual_name}" from {os.path.basename(path)}')
        return pd.read_excel(xl, sheet_name=0)

    dfs = {
        'enrollment': read_first_sheet('enrollment', os.path.join(INPUT_DIR, FILES['enrollment'])),
        'assessment': read_first_sheet('assessment', os.path.join(INPUT_DIR, FILES['assessment'])),
        'growth':     read_first_sheet('growth',     os.path.join(INPUT_DIR, FILES['growth'])),
        'discipline': read_first_sheet('discipline', os.path.join(INPUT_DIR, FILES['discipline'])),
        'graduation': read_first_sheet('graduation', os.path.join(INPUT_DIR, FILES['graduation'])),
        'sqss':       read_first_sheet('sqss',       os.path.join(INPUT_DIR, FILES['sqss'])),
        'wakids':     read_first_sheet('wakids',     os.path.join(INPUT_DIR, FILES['wakids'])),
    }
    print('Files loaded.')

    # Build reference tables once
    school_names_all, org_id_bridge, dist_org_ids = build_reference_tables(dfs['enrollment'])
    school_type_codes = read_school_type_codes(
        os.path.join(INPUT_DIR, FILES['help_assessment'])
    )
    alt_type_lookup = build_school_type_lookup(dfs['enrollment'], school_type_codes)
    print(f'School type codes loaded: {len(school_type_codes)} codes, {len(alt_type_lookup)} school entries')

    # Process all districts
    index = []
    for slug, dist_code in DISTRICTS.items():
        index.append(build_district(
            slug, dist_code, dfs,
            school_names_all, org_id_bridge, dist_org_ids, alt_type_lookup
        ))

    # Write district index
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    idx_path = os.path.join(OUTPUT_DIR, 'districts.json')
    with open(idx_path, 'w', encoding='utf-8') as f:
        json.dump({'districts': index}, f, ensure_ascii=False, indent=1)

    print(f'\nWrote {idx_path}  ({len(index)} districts)')
    print('Done. ✓')
    return 0


if __name__ == '__main__':
    sys.exit(main())
