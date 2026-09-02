# Smart Schedule System

One `MAINPROJECTINFO.xlsm` (the MPI) per project folder. Every schedule in that
folder links three values to it and works out everything else locally.

```
<project folder>/
    MAINPROJECTINFO.xlsm                     <- the only file you edit per project
    ...-SC-M-00000009-... - Radiators Schedule.xlsx
    ...-SC-M-00000010-... - Radiant Panel Schedule.xlsx
    ...
```

## The rule

**Only `Metadata!B4:B6` link out to the MPI: Project Name, Project Number, Client.**

Everything else in a schedule is a local reference. Copy a schedule to a laptop
without the MPI and it still opens correctly; only those three go stale (they
keep their last cached value, so nothing shows an error).

That is why adding a revision line is all you ever do to change the date or
revision across a schedule: the title block, front cover and revision summary
all read the revision table below them.

## Install (once, into your MPI template)

1. Open `MAINPROJECTINFO.xlsm`.
2. `Alt+F11` to open the VBA editor.
3. Delete the old modules: `Module1`, `Module2`, `Module3`, `Module4`,
   `Module11`, `UserForm1`, and the code in `Sheet3`. (Right-click → Remove.)
4. `File → Import File...` and import all three files from `vba/`:
   `modUtil.bas`, `modSchedule.bas`, `modMain.bas`.
5. Back in Excel: `Alt+F8` → run **InstallTool**.
6. Save. This is now your company template.

## Per project

1. Copy the template MPI into the project folder alongside the schedules.
2. Fill in **Client**, **Project Name**, **Project Number** on the Setup sheet.
3. Press **Set up / repair schedules**.

That's it. Nothing else is typed twice.

## The three buttons

**Set up / repair schedules** — opens every workbook in the folder, points its
Metadata at this MPI, rebuilds the local references, clears leftover links from
whatever it was copied from, and saves. Safe to press again at any time, so it
is also the "I just added a schedule" button and the "someone broke a link"
button. It backs up every file it touches into a timestamped `_backup` folder
first (turn that off on the Setup sheet if you don't want it).

**Refresh schedule list** — opens every schedule read-only and writes what it
actually contains onto the ScheduleList sheet, then checks it. The `Checks`
column says `OK` or what is wrong. It flags:

- no Metadata sheet (never been set up)
- Project Name / Number / Client that don't match the MPI
- a revision that isn't resolving
- a file name that disagrees with the schedule name inside
- links to any workbook other than the MPI
- **a revision or date that differs from most of the other schedules**

That last one is the "23 schedules say P04 and one still says P03" check.

Because the files are opened without updating links, the list shows the cached
values, i.e. exactly what a recipient sees when they open the file offline.

**Add revision to ticked** — fill in the *New revision* block on the Setup
sheet, put an `x` in the `Add?` column on ScheduleList next to the schedules
being reissued, press the button. It appends one row to each revision table and
refuses to add a revision number that is already there. Deliberately opt-in per
file: blanket-revving schedules that didn't change is its own QA problem.

## Setup sheet options

| Cell | Meaning |
|---|---|
| `B7` Schedules folder | blank = the folder this file is saved in. Filled in automatically if the tool has to ask |
| `B8` Backup before changes | `Yes` / `No` |
| `B9` Refresh list on open | `No` by default. `Yes` opens every schedule at startup, which is slow with 20+ files |
| `F1:F...` Suitability Codes | the dropdown pushed into every schedule's revision table. Populated from what your schedules already use on the first run; edit it here and re-run setup |

## What a schedule must contain

The tool finds cells by their labels, never by fixed addresses, and skips (and
reports) anything it cannot positively identify. It never guesses.

- A sheet named **Revision Page** with, in column A: `Project Name`,
  `Project no.`, `Recipient`, `Document type`, `Revision`, `Date`,
  `Prepared by`, `Checked by`, `Approved by`, `Document no`,
  `Suitability Status`, `Suitability Description` — values in the column to the
  right. A table named **RevisionTable** with columns `Revision`, `Status`,
  `Date`, `Prepared by`, `Checked by`, `Approved by`, `Description`.
  **`A4` holds the schedule title as typed text** — this is the one place the
  schedule name is written.
- A sheet named **Front Cover** with `Intended for` and `Date` labels (values
  below them) and the title text starting `SCHEDULE OF...`.
- A sheet named **Metadata** — fully owned by the tool, it rewrites it.

## What gets written to Metadata

| Row | Header | Source |
|---|---|---|
| 2 | DocumentNumber | the file name, up to the last `-` |
| 3 | ScheduleName | `Revision Page!A4` |
| 4 | Project Name | **MPI** |
| 5 | Project Number | **MPI** |
| 6 | Client | **MPI** |
| 7-14 | DocumentType, Revision, Date, Prepared/Checked/Approved by, Suitability Status, Suitability Description | the local Revision Page title block |

Rows 7-14 exist so the MPI can read one predictable sheet per schedule instead
of hunting for labels in a formatted page. Rows 2-6 keep the header names the
old files already used, so nothing existing breaks.

## Revision ordering

The title block reads the latest revision out of the revision table, ranking by
family first and then by number:

| Family | Meaning | Priority |
|---|---|---|
| `P` | Preliminary | 1 |
| `C` | Construction | 2 |
| `AF` | As fitted | 3 |

So `AF01` beats `C09` beats `P12`. The list lives in one place, the
`REV_PREFIXES` constant at the top of `modSchedule.bas`. Add a family there and
re-run **Set up / repair schedules** to push the new formulas into every file.

## Running from Filery / SharePoint

When a workbook is opened straight from Filery, SharePoint or a browser, its
path is a URL (`https://...`) rather than a drive path, and Excel's file
functions cannot read a URL. That is what caused `Run-time error 52`.

The tool now:

1. uses the folder in `Setup!B7` if it is set and readable,
2. otherwise tries to resolve the URL to your local synced folder,
3. otherwise asks you to pick the folder once and writes it into `Setup!B7`.

**For the cleanest links, open `MAINPROJECTINFO.xlsm` from the synced folder in
File Explorer, not from the browser.** When both the MPI and the schedules are
opened from the same local folder, Excel stores the link as a plain file name,
which is what survives a Filery export and re-import. If the MPI is open from a
URL, the tool still works but warns you that the links will be written as full
URLs.

## Known limits

- `DocumentNumber` comes from `CELL("filename")`, so it is blank until the file
  has been saved once.
- The title cell is located within the first 60 rows / 10 columns of a sheet.
- Revision numbers above 999 would collide across families.
