# School Results Analyser

An Excel VBA toolkit for importing school result exports, organising them into consistent worksheets, and producing subject, cohort, student-progress, and correlation reports.

The project is intended for teachers or school staff working with Cockpit-style Excel result exports. It supports:

- SEC results using G1, G2, and G3 grade schemes
- IP results using grades from A+ to U
- Multiple levels, assessments, and years
- Form-teacher progress comparisons between two assessments
- At-risk and top-performing student lists
- Subject-score correlation matrices
- Dashboard navigation buttons

> **Important:** this repository contains exported VBA source files (`.bas` and `.cls`). It does not contain a ready-to-use Excel workbook. The files must be imported into a macro-enabled workbook before they can run.

## What the analyser does

The normal data flow is:

```text
Cockpit Excel exports
        |
        v
Normalised staging sheets (for example, S2_WA1_2025)
        |
        +--> SEC/IP subject analysis and charts
        +--> At-risk and top-student reports
        +--> Form-teacher progress reports
        +--> SEC subject correlation reports
        |
        v
Dashboard navigation
```

### Import and normalisation

The importer reads one file or every Excel file in a selected folder. It prefers a source worksheet named `RE_RES_078`; if that sheet does not exist, it uses the first visible worksheet.

It produces one row per student with a consistent structure:

```text
RegNo | Name | Class | Assessment | Year | subject scores/grades | other columns
```

Output is separated by level, assessment, and year. Example worksheet names are:

- `S2_WA1_2025` for a SEC level
- `Y2_WA1_2025` for an IP level
- `Formatted_S2_WA1_2025` if `Formatted` is configured as a prefix

When files are imported, only affected staging sheets are refreshed. Unrelated assessment sheets are left in place.

### Reports

The toolkit can generate:

- **SEC subject analysis:** class and cohort grade distributions, pass/fail/top-grade rates, mean grade, charts, and distribution-validity guidance.
- **IP subject analysis:** A+ to U distributions, pass/fail/top-grade rates, GPA, charts, and validity guidance.
- **At-risk reports:** students meeting the configured failed-subject threshold, grouped into `AtRisk_S1` to `AtRisk_S5`.
- **Top-student reports:** top-grade summaries in `TopQual_S1` to `TopQual_S5`.
- **Form-teacher progress:** student-level and subject-level changes between assessments, including automatic remarks and talking points.
- **SEC correlations:** subject-score correlation matrices for a selected year, with automatically generated insights.
- **Results-summary email:** a compact HTML summary opened as an Outlook draft for review; it is never sent automatically.
- **Navigation:** buttons on a `Dashboard` worksheet linking to generated reports, plus Home buttons on report sheets.

## Requirements

- Microsoft Excel with VBA support (desktop Excel, not Excel for the web)
- A macro-enabled workbook saved as `.xlsm`
- Macros enabled for that workbook
- Cockpit-style source files in `.xls`, `.xlsx`, or `.xlsm` format
- A `Settings` worksheet in the macro workbook
- A `Dashboard` worksheet if navigation buttons are required

No external VBA libraries or add-ins are required by the source code.

The optional email-draft feature requires desktop Microsoft Outlook with Windows COM automation available. It uses late binding, so no Outlook reference needs to be added in the VBA editor. Outlook automation is not expected to work in Excel for macOS or Excel for the web.

## Installation

1. Open Excel and create a blank workbook.
2. Save it as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
3. Create worksheets named `Settings` and `Dashboard`.
4. Open the VBA editor:
   - Windows: press `Alt+F11`.
   - macOS: use **Tools > Macro > Visual Basic Editor**, or the keyboard shortcut configured for your Excel version.
5. In the VBA editor, select your workbook and use **File > Import File** to import every source file in this repository:
   - All `.bas` files
   - `cSubjPair.cls`
6. Save and close the VBA editor.
7. Configure the `Settings` worksheet as described below.

Do not paste `cSubjPair.cls` into a standard module. It must be imported as a **class module**.

## Settings worksheet

The importer reads the following cells. Values marked “optional” use the shown default when blank or invalid.

| Cell/range | Purpose | Typical/default value |
|---|---|---|
| `B3` | Optional prefix for generated staging-sheet names | Blank |
| `B4` | Main header row in each Cockpit export | `10` |
| `B5` | Sub-header row | `11` |
| `B6` | Row containing the year in column A | `3` |
| `B7` | Row containing the assessment name in column A | `6` |
| `B8` | First student-data row | `12` |
| `B9` | Text of the GEP header used as an anchor | `GEP Indicator` |
| `B10:B30` | Footer prefixes that mark non-student rows to exclude | School-specific |
| `B13` | Legacy logging option read by the importer | `TRUE` or `FALSE` |
| `B14` | Extra subject names, separated by commas | e.g. `Maths 1, Maths 2` |
| `B15` | Include grade columns as well as scores | `TRUE` |
| `B16` | Width of generated subject columns; minimum 5 | `9` |
| `D2:E50` | Class-pattern to level-key mappings | e.g. `3A` -> `S3` |
| `H2:I50` | Grade-to-point mappings used by progress analysis | e.g. `A1` -> `1` |
| `L5` | Optional width factor for IP validity panels | Leave blank for default |
| `L6` | Minimum cohort size for SEC subject analysis | `10` |
| `L7` | Number of failed subjects used for the at-risk threshold | `3` |
| `L8` | Percentage used to determine a student's main FSBB group | `70` |
| `N2:O20` | Optional level mode table for top-student reports | e.g. `S4` -> `AUTO_FSBB` |
| `Q2:R30` | Optional email settings table | See “Email draft settings” below |
| `T2:U100` | Optional source-to-display subject-name mappings for email | e.g. `EL - O` -> `English Language` |

For the optional level mode table, supported values in column O are `AUTO_FSBB` (the default) and `LEGACY_NO_DOWNWARD`.

### Email draft settings

Enter setting names in column Q and their values in column R. The row order does not matter.

| Q (setting name) | R (example value) |
|---|---|
| `SchoolName` | `Example Secondary School` |
| `PreparedBy` | `Assessment Committee` |
| `EmbargoText` | `For Internal Use only. Embargoed until 2 pm.` |
| `EmailTo` | `principal@example.edu; vp@example.edu` |
| `EmailCC` | `hod@example.edu` |
| `EmailSubjectPrefix` | `Results Summary` |

Recipients are optional. Leaving them blank creates an unaddressed draft. Always check names and addresses in Outlook before sending.

The email removes track suffixes such as ` - O` and ` - G2` from displayed subject names. To replace abbreviations or codes with full names, add mappings in `T2:U100`. Column T must match the staging subject name before `(Grade)`; column U contains the name leaders should see in the email.

### Example class mappings

Enter one class pattern per row in column D and its output level key in column E. Adjust these examples to match the school's class names.

| D (class pattern) | E (level key) |
|---|---|
| `1A` | `S1` |
| `1B` | `S1` |
| `1F` | `Y1` |
| `2A` | `S2` |
| `2F` | `Y2` |
| `3A` | `S3` |
| `3F` | `Y3` |
| `4A` | `S4` |
| `4F` | `Y4` |

More specific patterns should be placed before broad patterns. Verify the generated sheet names after the first import.

### Example grade-to-point mappings

The progress report compares grades through the mapping in `H2:I50`. Add every grade used in the source data. A typical SEC mapping begins as follows:

| H (grade) | I (point) |
|---|---:|
| `A1` | 1 |
| `A2` | 2 |
| `B3` | 3 |
| `B4` | 4 |
| `C5` | 5 |
| `C6` | 6 |
| `D7` | 7 |
| `E8` | 8 |
| `F9` | 9 |

Add the school's G1, G2, and IP mappings as needed. Use the same direction consistently: lower points should represent stronger results if you want the generated progress interpretation to follow the SEC convention.

> Set `B15` to `TRUE` before importing if you intend to build grade distributions, at-risk lists, top-student lists, or grade-based progress reports. If it was previously `FALSE`, enable it and import the source files again.

## Expected source-file structure

By default, the importer expects:

- Year information in column A of row 3
- Assessment information in column A of row 6
- Main headers in row 10
- Sub-headers in row 11
- Student data beginning in row 12
- Main headers named `Reg#`, `Name`, `Class`, and `GEP Indicator`
- Subject columns arranged in score/grade pairs
- A grade sub-header containing `B/G` immediately after each subject's score column

Subject names such as `EL - O`, `HCL - G3`, `Maths - G2`, and names ending in `IP` are detected automatically. Unusual subject names can be listed in `Settings!B14`.

If the export layout differs, update `B4:B9` instead of editing the VBA code where possible.

## Recommended operating procedure

Run macros from Excel using **Developer > Macros** (or `Alt+F8` on Windows).

### 1. Import results

Choose one of:

- `ParseCockpitFiles_ToStaging` — select one or more source workbooks.
- `ParseCockpitFolder_ToStaging` — select a folder and import all Excel workbooks in it.

Inspect the generated staging sheets before continuing. Confirm that:

- The level, assessment, and year in each sheet name are correct.
- Every student has the correct registration number, name, and class.
- Subject score and grade columns have been detected correctly.
- No footer or summary rows were imported as students.

### 2. Generate subject analyses

Run the appropriate macro:

- `BuildAllSec_SubjectAnalysis` for SEC G1/G2/G3 results.
- `BuildAllIp_SubjectAnalysis` for IP results.

These macros scan eligible staging sheets and generate formatted subject-analysis worksheets with tables, charts, and interpretation panels.

### 3. Generate student reports (optional)

- Run `BuildSec_AtRiskSummary` to create at-risk lists.
- Run `BuildSec_TopQualityByLevel` to create top-student lists.
- Run `BuildFormTeacherProgressAllClasses_Prompt` to compare two assessments for all matching classes.
- Run `BuildFormTeacherProgress_Prompt` to build a progress report for one class.

For a standard WA1-to-WA2 comparison, `BuildFormTeacherProgress_WA1_WA2` is also available.

Progress output includes:

- `FT_Progress_<Class>_<From>_<To>` — one summary row per student
- `FT_SubjectDelta_<Class>_<From>_<To>` — one row per student and subject

Both assessments must already have been imported, and their student identities and subject headers must be consistent.

### 4. Generate correlations (optional)

Run `BuildSecCorrelation_PromptYear` and enter the assessment year.

The result is written to `SEC_Correl_<Year>`. Correlations use score columns only and require at least 30 valid scores per subject and at least 30 overlapping students for each subject pair.

### 5. Build Dashboard navigation

Run the relevant navigation macros after creating the reports:

- `BuildSubjectAnalysisNavigation`
- `BuildIpSubjectAnalysisNavigation`
- `BuildFormTeacherProgressNavigation`
- `BuildTopQualityNavigation`
- `BuildSec_AtRiskNavigation`

Some report-building macros create their navigation automatically. Re-running a navigation macro refreshes its section of the `Dashboard` sheet.

The current code places both IP navigation and at-risk navigation at `Dashboard!M3`, so these two menus should not be used together without changing one module's start-cell constant. Whichever menu is built last may overwrite cells used by the other.

### 6. Create a results-summary email draft (optional)

1. Confirm that the relevant SEC staging sheet contains grade columns and has been checked for accuracy.
2. Configure `Settings!Q:R` and optional subject-name mappings in `Settings!T:U`.
3. Run `DraftResultsSummaryEmail`.
4. Enter the staging-sheet name when prompted. If a valid staging sheet is active, its name is offered as the default.
5. Review the Outlook draft carefully before sending it.

The draft contains:

- An embargo banner and results title
- Candidate, subject, overall-pass, highlight, and warning KPIs
- Lists of subjects with 100% passes or pass rates below 90%
- Separate G3, G2, and G1 subject-performance tables
- Up to ten top students for each G3/G2/G1 group
- Data warnings for low N, absences, or unrecognised grades
- Prepared-by, timestamp, source-sheet, and methodology information

The email calculates its results directly from the selected current staging sheet. It does not combine earlier sittings. Absences are excluded from subject rates and are reported as data checks. The macro calls Outlook's `Display` method and contains no `Send` call.

## Important data rules

- Failing SEC G3 grades are `D7`, `E8`, and `F9`.
- Failing G2 grade is `6`.
- Failing G1 grade is `E`.
- Failing IP grades include `D+`, `D`, and `U`.
- `AB` is treated as a failure/zero in relevant student reports, but is excluded from the count of subjects attempted.
- Validity and pattern labels are indicators for professional review, not proof that an assessment is valid or invalid.
- Correlation does not establish causation.

Always review generated reports before using them for student, parent, or staffing decisions.

## Re-running and clearing data

Importing files refreshes each affected staging sheet once and then combines matching input files into it. This makes it possible to import several class files belonging to the same level, assessment, and year.

`ClearStaging` clears the contents of staging worksheets identified by the configured prefix or expected naming pattern. Use it carefully: it modifies worksheet data and does not provide an undo facility.

For safety, keep an untouched copy of the source exports and back up the `.xlsm` workbook before a major refresh.

## Troubleshooting

### “Settings are incomplete”

Confirm that a sheet named exactly `Settings` exists and that `B4`, `B5`, and `B8` contain valid row numbers. The first data row must be below the sub-header row.

### A file is skipped or no staging sheet is created

Check that:

- The source has a visible `RE_RES_078` sheet, or at least one other visible sheet.
- `Reg#`, `Name`, `Class`, and the configured GEP header occur on the main-header row.
- The year and assessment rows contain recognisable values in column A.
- At least one subject score/grade pair is present after the GEP column.
- Footer prefixes in `B10:B30` do not match real student rows.

To retain parser messages, create a worksheet named `Logs` before importing. The parser appends the timestamp, filename, status, and notes to columns A:D. In the current code, the logger writes whenever this sheet exists; the value read from `Settings!B13` does not control `AppendLog`.

### Subject analysis finds no eligible data

Confirm that grades were included during import (`Settings!B15 = TRUE`) and that staging headers end in `(Grade)`. Also check class-to-level mappings and grade spelling.

### Progress reports cannot find an assessment

Ensure both assessment sheets exist, use matching assessment names, and have headers named `RegNo`, `Name`, `Class`, `Assessment`, and `Year`. Student registration numbers and class names should remain consistent across assessments.

### Dashboard buttons are missing

Create a worksheet named exactly `Dashboard`, generate the corresponding reports first, and then run the navigation macro again.

### Macros do not appear or will not run

- Confirm that the workbook is saved as `.xlsm`.
- Enable macros or place the file in a trusted location according to the organisation's security policy.
- Confirm that all `.bas` files and `cSubjPair.cls` were imported into the same VBA project.
- In the VBA editor, use **Debug > Compile VBAProject** to locate missing modules or compile errors.

## Module guide

| File | Responsibility |
|---|---|
| `modcockpitparser.bas` | Imports and normalises Cockpit result files |
| `cSubjPair.cls` | Holds a subject's score and grade column references during import |
| `modSecGradeDistribution.bas` | SEC analysis, at-risk reports, and top-student reports |
| `modIpAnalysis.bas` | IP grade distributions, GPA, charts, and validity panels |
| `modValidityEngine.bas` | Shared SEC distribution-validity and interpretation logic |
| `modFormTeacherProgress.bas` | Assessment-to-assessment student progress reporting |
| `modeSecCorrelation.bas` | SEC subject-score correlation matrices |
| `modResultsSummaryEmail.bas` | Outlook HTML results-summary draft generator |
| `modSubjectNavigation.bas` | SEC report navigation and Home buttons |
| `modIpSubjectNavigation.bas` | IP report navigation and Home buttons |
| `modPhysicsMsgDebug.bas` | Special-purpose Sec 4 Physics MSG diagnostic log |

`modPhysicsMsgDebug.bas` is not part of the normal workflow. It is hard-coded for worksheet `S4_PRELIMINARYEXAM_2025` and grade column `Phy - O (Grade)`, and writes detailed calculations to `PhyLogs`.

## Privacy and responsible use

The workbook processes student names, registration numbers, classes, and results. Store it only in an approved location, restrict access appropriately, and follow the school's data-protection and retention policies. Avoid sharing live workbooks or screenshots containing identifiable student information.
