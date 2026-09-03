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

## The four buttons

**Set up / repair schedules** — opens every workbook in the folder, points its
Metadata at this MPI, rebuilds the local references, clears leftover links from
whatever it was copied from, and saves. It is idempotent, so **adding a new
schedule to the folder needs nothing more than pressing this again**. Same for
a link someone has broken, or a change to the suitability codes or the revision
families. It backs up every file it touches into a timestamped `_backup` folder
first (turn that off on the Setup sheet if you don't want it).

**Refresh schedule list** — reads every schedule and writes what it actually
contains onto the ScheduleList sheet, then checks it. The `Checks` column says
`OK` or what is wrong. It flags:

- no Metadata sheet (never been set up)
- Project Name / Number / Client that don't match the MPI
- a blank revision or schedule name
- links to any workbook other than the MPI
- **a revision or date that differs from most of the other schedules**

That last one is the "23 schedules say P04 and one still says P03" check.

There is deliberately **no check on the file name**. Document numbers and
schedule names do not have to match what a file happens to be called, and a QA
tool that complains about its own conventions is noise.

Files are read without updating links and without recalculating, so the list
shows what is actually saved in each file, i.e. exactly what a recipient sees
when they open it offline.

**Add revision to ticked** — put an `x` in the `Add?` column and fill in the
`New Rev` / `New Status` / `New Date` / ... columns on that row. Anything left
blank falls back to the *New revision* block on the Setup sheet.

So both jobs are the same operation:

- **All 24 on one revision** — fill in the Setup block once, tick every row.
- **6 of 24, on different revisions** — type each one on its own row, tick
  those six, leave the rest alone.

It appends one row per revision table, refuses a revision number that is
already in that file, and clears the tick and the typed values on the rows that
succeeded so the list is ready for the next round.

The new line goes in **directly under the last revision**, not at the bottom of
the table. Revision tables are usually drawn with spare rows below the last
entry, and appending past them leaves a gap that makes the title block look
like it skipped a revision. The table only grows by a row when every row in it
is already used.

The `Add?` and `New Rev` headers carry hover notes explaining this on the
sheet, and the same is written under the buttons on the Setup sheet.

## Setup sheet

Three blocks, top to bottom: **Project**, **Options**, **New revision**, with
the suitability codes in column F and the buttons to the right.

**Anything shaded yellow is yours to fill in.** Grey italic cells are written
by the tool. Everything else is a label or a note.

`InstallTool` rebuilds this sheet each time it runs. It clears columns A, C and
H and rewrites them, because option rows move between versions and a label left
on an old row is worse than no label at all. It does **not** touch column B
(your values), column F (the suitability codes), or any other sheet in the
workbook, so a README sheet of your own is safe. Option cells holding a value
from an older layout, such as a date serial sitting where a Yes/No now lives,
are reset to the default rather than left to be read wrong.

| Cell | Meaning |
|---|---|
| `B7` Schedules folder | blank = the folder this file is saved in. Filled in automatically if the tool has to ask |
| `B8` Backup before changes | `Yes` / `No` |
| `B9` Refresh list on open | `No` by default. `Yes` reads every schedule at startup |
| `B10` Full refresh every time | `No` by default, which reopens only files whose modified date changed since the last refresh. `Yes` always reopens everything |
| `B11` Header/footer source | the workbook to copy headers and footers from. Blank = the button asks, then fills this in |
| `B12`, `B13` | when setup and the refresh last ran. Written by the tool, read only |
| `F1:F...` Suitability Codes | the dropdown pushed into every schedule's Metadata sheet and revision table. Seeded with the ISO 19650 codes below; edit here and re-run setup to push the change to all files |

## Headers and footers (security classification)

**Copy headers & footers** takes one workbook's headers and footers and puts
them on all the others. This is how `OFFICIAL`, `OFFICIAL-SENSITIVE`,
`CONFIDENTIAL` (or nothing at all) gets applied consistently: set one schedule
up by hand under Page Layout, then push it to the rest.

Sheets are matched in this order:

1. a source sheet with the **same name**
2. `Front Cover` and `Revision Page` to their namesakes
3. anything else to the source's schedule sheet

Schedule tabs are not always called `Schedule`, which is why the name match
comes first and the role fallback second.

**Orientation, paper size and margins are never touched.** Only header and
footer properties are written, so a landscape schedule stays landscape. Header
and footer text is positioned relative to whatever page a sheet is set to, so
the same banner comes out correct on portrait and landscape alike.

### The logo in the footer

A footer containing `&G` is showing an image, and that image lives inside each
file. VBA cannot move image data between workbooks, so:

- where the target **already has an image** (all files descended from the same
  template do), its **size is matched to the source**, so a logo scaled to 20%
  is 20% everywhere;
- where the target **has none**, the source's original file path is tried,
  which only works if that file is still on the PC, and the log says so if it
  is not.

The summary box flags when the source uses `&G` at all, so it is worth checking
one print preview after the first run.

The source workbook is remembered in `Setup!B11`, so re-running after a tweak
is one click.

## Suitability codes

Held in one place, column F of the Setup sheet, and pushed into each schedule's
Metadata sheet (`D2` down) by the repair. The revision table's Status column
picks from that local list, so it still works with the MPI nowhere in sight.

The format is `<code> - <description>`. The title block splits on the ` - ` to
fill Suitability Status and Suitability Description, so **no description may
contain a hyphen surrounded by spaces**.

| Code | Description |
|---|---|
| `S0` | Work in Progress |
| `S1` | Suitable for Coordination |
| `S2` | Suitable for Information |
| `S3` | Suitable for Review and Comment |
| `S4` | Suitable for Stage Approval |
| `S6` | Suitable for PIM Authorisation |
| `S7` | Suitable for AIM Authorisation |
| `A1`, `A2` | Authorised and Accepted |
| `B1`, `B2` | Published with Comments |

Two decisions worth knowing about:

- **`S4` replaces `S5` as the approval status.** ISO 19650 leaves `S5` to
  project guidance, which is exactly why it caused confusion here.
- **`S5` is not in the list.** Nothing useful can be said about it on a shared
  dropdown. If a project genuinely needs it, add a row to column F.

Existing revision lines are never rewritten. A row that says
`S5 - Suitable for Client Acceptance` keeps saying that, which is correct: the
revision table is a record of what was actually issued. The list only governs
what the dropdown offers from here on.

`A1`/`A2` and `B1`/`B2` share a description on purpose. The number is the
issue count, not a different status. Add `A3`, `B3` and so on as needed.

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

## Progress, speed and the log

Both long-running buttons show progress and a time estimate on the status bar
(bottom-left of the Excel window, and there is a note saying so under the
buttons on the Setup sheet), e.g. `Setting up schedules: 7 of 24 (29%) - about 1m 20s left -
Radiators Schedule.xlsx`. Every run writes a line per file to the **Log** sheet
and finishes with a summary box: how many succeeded, were skipped, or failed,
and how long it took.

The refresh is **incremental by default**. A file whose modified date has not
changed since the last refresh is not reopened, its row is reused, and only its
comparison against the Setup values is recalculated. So the first refresh costs
what it costs, and every one after it is close to instant unless something
actually changed. Set `Full refresh every time` to `Yes` on the Setup sheet to
force a complete re-read.

The refresh also reads with calculation set to manual. That is both faster and
more honest: it reports what is saved in the file rather than what Excel would
compute after opening it.

### Why there is no "read the properties without opening the file"

Custom document properties can be read from a closed file, but a `.xlsx` has no
macros, so nothing inside it can keep those properties up to date when someone
edits it by hand. The tool could stamp them, and they would then be right until
the first manual edit and quietly wrong afterwards. For a QA tool that is worse
than being slow, so the incremental refresh above is used instead: it is exact,
because a file that was not modified cannot have changed.

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

The repair rewrites these formulas **everywhere in the workbook**, not just the
seven cells on the Revision Page, so a copy of the same formula sitting on a
front cover or in a schedule header gets the new ranking too. It only touches
formulas that already rank `RevisionTable[Revision]` with `MAX()`, works out
what each one returns, and leaves anything it cannot identify alone.

## What the repair does and does not touch

It **does** rebuild every cell that should be derived: the Metadata sheet, the
linked cells on the Front Cover and Revision Page, the revision formulas, the
suitability dropdown, leftover links and dead defined names.

It **does not** copy layout or static text between workbooks. Two schedules
whose front covers were built differently stay different. If you want one
golden front cover pushed into all 24, that is a separate deliberate step and
worth asking for on its own, because it overwrites whatever is there now.

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
