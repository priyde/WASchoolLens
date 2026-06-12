# WASchoolLens

**See every school clearly. Compare what matters.**

WASchoolLens is an independent dashboard for comparing Washington State public schools side by side: test scores, student growth, milestones, discipline, demographics, and a first-of-its-kind **K-12 Pathways** view that maps each child's journey from kindergarten through graduation.

All data comes directly from the [OSPI Washington State Report Card](https://ospi.k12.wa.us/data-reporting/data-portal) (published via [data.wa.gov](https://data.wa.gov)).

> WASchoolLens is not affiliated with OSPI or any school district.

---


## Districts included (2024-25)

| District | Slug | Schools | Feeder chains |
|---|---|---|---|
| Lake Washington School District | lwsd | 55 | 31 |
| Bellevue School District | bsd | 30 | 22 |
| Northshore School District | nsd | 36 | 25 |
| Issaquah School District | isd | 30 | 16 |

**151 schools across 4 districts**, with more planned.

---

## Dashboard tabs

| Tab | What it shows |
|---|---|
| Test scores | ELA / Math / Science proficiency (SBA & WCAS) |
| Growth (SGP) | Median Student Growth Percentile vs. statewide peers |
| Milestones | KG readiness, attendance, 9th on track, dual credit, 4-year graduation |
| Discipline | Exclusionary suspension / expulsion rates |
| Demographics | Enrollment, FRPL (low income), race/ethnicity, ELL, SpEd, HC |
| Score vs. FRPL | Scatter showing which schools out/under-perform their demographics |
| K-12 Pathways | Every feeder chain (Elem to Middle to High) as one comparable row |

Every metric column has a hover tooltip explaining what it means and citing its exact data source. An "Abbreviation guide" panel (collapsed by default) covers all terms used.

---

## Project structure

```
WASchoolLens/
├── index.html           the dashboard (vanilla JS, no frameworks, no build step)
├── contact.html         About / Contact / Feedback page
├── process_data.py      ETL pipeline: OSPI Excel files + FeederFlow.xlsx -> JSON
├── FeederFlow.xlsx       feeder chain data for all districts (one sheet per district)
├── data/
│   ├── districts.json   index of available districts
│   ├── lwsd.json
│   ├── bsd.json
│   ├── nsd.json
│   └── isd.json
└── README.md
```

## Data notes & methodology

- **Proficiency**: OSPI "Percent Consistent Grade Level Knowledge And Above" from the Smarter Balanced Assessment (ELA/Math, grades 3-8 and 10) and WCAS (Science, grades 5/8/11). WA-AIM alternate assessment results are excluded, matching OSPI's own school-level reporting.
- **SGP**: median Student Growth Percentile, grades 4-8, spring administrations. 1-33 = low growth, 34-66 = typical, 67-99 = high growth.
- **KG Ready%**: WaKIDS, percent of kindergartners ready in all 6 developmental domains. Reporting the language and literacy domains is optional for districts, which can affect this number.
- **Attendance%**: SQSS Regular Attendance, percent of students attending 90%+ of days (fewer than 2 absences per month).
- **Graduation**: Four-Year adjusted cohort rate.
- **Pipeline join key**: all data sources are joined on `SchoolCode` (an OSPI integer identifier), not school name strings, for robustness against naming variations and renames.
- **WaKIDS bridge**: the WaKIDS file lacks `SchoolCode`, so it is joined via a `SchoolOrganizationId` to `SchoolCode` bridge table built from the enrollment file.
- **Privacy suppression**: OSPI redacts small student groups (N<10) and applies top/bottom range capping. Suppressed values appear as dashes or as capped (less than / greater than) bounds in the dashboard.
- **Feeder chains**: sourced from official district feeder pattern documents and boundary maps. Geographic splits, where an elementary school's attendance boundary is divided between two middle or high schools, are shown as separate rows with explanatory notes.

---

## License

All rights reserved. This repository has no open source license, so no permission is granted to copy, modify, or redistribute this code or its content without permission from the author.

Source data is public information published by Washington State OSPI.

(c) 2026 Priyank A Deshmukh.
