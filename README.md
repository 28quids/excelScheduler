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

### Re-installing a newer version

**Remove `modUtil`, `modSchedule` and `modMain` before importing the new
ones.** Importing a `.bas` whose module already exists does not replace it, it
adds `modMain1` alongside `modMain`, and which copy runs is anyone's guess.

Import **all three** every time. They change together, and a mismatched set
fails in confusing ways rather than obviously.

The install box and cell `H22` on the Setup sheet both show the tool version.
If that number is not what you expect, the import did not take.

### Changing the version number

One place, the declarations at the top of `vba/modUtil.bas`:

```vba
Public Const TOOL_VERSION As String = "1.1"
```

That single constant feeds both the "Ready. Schedule tool version ..." box and
the note on the Setup sheet. Change it, re-import `modUtil`, run `InstallTool`.

The **Version** table on your own README sheet is separate; the tool never
touches that sheet, so keep it in step by hand.

## Per project

1. Copy the template MPI into the project folder alongside the schedules.
2. Fill in **Client**, **Project Name**, **Project Number** and **Security
   classification** on the Setup sheet.
3. Press **Refresh schedule list**, then **Set up / repair links** with nothing
   ticked, and say Yes to all.

That's it. Nothing else is typed twice.

## The ticks: which schedules a button works on

**Every button that changes a schedule works on the rows with an `x` in
column A (`Select`) of ScheduleList, and only those.** The PDF export too.

With nothing ticked, a button works on every workbook in the folder, but its
confirmation box says **ALL** in capitals and defaults to **No** first. Ticked
files are listed by name in the box, so what you are agreeing to is on screen.

Most of what these buttons do cannot be undone except from the backup, so
doing the one schedule that needs it, rather than all 26, is the safer habit,
and the tool makes it the easy one.

Ticks stay until you delete them. That is deliberate: **Add revision to
ticked** then **Export ticked to PDF** issues exactly the schedules you just
revised. Pressing Add revision twice does no harm, a revision already in a
table is refused.

A ticked row whose file has gone from the folder is left out and named in the
confirmation box. The reference schedule is never a target of the buttons
that copy from it, even when ticked.

## The buttons

Three groups on the Setup sheet, left to right by how often you need them.

| Every issue | Adding a schedule | One job at a time |
|---|---|---|
| Refresh schedule list | Set up new schedule(s) | Set up / repair links |
| Add revision to ticked | | Set classification |
| Export ticked to PDF | | Copy headers & footers |
| | | Copy cover & revision page |
| | | Title block |
| | | Tidy sheets |
| | | Rename files |

## Adding a new schedule

1. Put the file in the project folder.
2. Press **Refresh schedule list**. Every workbook in the folder is listed,
   including one with no Revision Page yet, which shows as
   `NOT A SCHEDULE YET`.
3. Put an `x` next to it.
4. Press **Set up new schedule(s)**.

That one button opens each ticked file once and does, in order:

1. Front Cover and Revision Page from the reference schedule (its own title
   and any revision history are kept, as with **Copy cover & revision page**)
2. links set up to this MPI
3. the title block and General Notes box, fitted to its columns, and added
   above the schedule if it has none (see [Title block](#title-block))
4. headers and footers from the reference, logo included
5. sheets put in order, Metadata hidden

then saves it once. A file that fails any step is closed **without saving**,
so a half-done schedule is never written. It only ever works on ticked rows:
with nothing ticked it explains the steps above and stops.

**Set up / repair links** — opens the ticked schedules, points their
Metadata at this MPI, rebuilds the local references, clears leftover links from
whatever they were copied from, and saves. It is idempotent: safe to run on a
schedule that is already right. Use it for a link someone has broken, or after
a change to the suitability codes, the project fields or the revision
families. A workbook with no Revision Page is skipped and pointed at **Set up
new schedule(s)**. It backs up every file it touches into a timestamped
`_backup` folder first (turn that off on the Setup sheet if you don't want it).

**Refresh schedule list** — reads every workbook in the folder and writes what
it actually contains onto the ScheduleList sheet, then checks it. Besides the
title block values it shows:

- **RevDescription**, the Description on the revision table row of the latest
  revision, read straight from `RevisionTable`
- **Classification**, the security marking the file really carries (see
  [Security classification](#security-classification))

The `Checks` column says `OK` or what is wrong. It flags:

- not a schedule yet (no Revision Page), or never set up (no Metadata sheet)
- Project Name / Number / Client that don't match the MPI
- a blank revision or schedule name
- a latest revision with no description, or that is not in `RevisionTable`
- links to any workbook other than the MPI
- a sheet with no security marking while the others have one
- **a classification that differs from the Setup sheet**, or from most of the
  other schedules when the Setup sheet does not name one
- **a revision or date that differs from most of the other schedules**

Schedules that are not set up are left out of the "most of the others"
comparisons, so one new file does not make every other row look wrong.

Each file name is a link: click it to open that schedule and check something.
Every button leaves a file you have open alone, and says so on the Log sheet,
because closing it underneath you would lose whatever you had not saved. The
refresh reads an open file as it stands and leaves it open.

That last one is the "23 schedules say P04 and one still says P03" check.

There is deliberately **no check on the file name**. Document numbers and
schedule names do not have to match what a file happens to be called, and a QA
tool that complains about its own conventions is noise.

Files are read without updating links and without recalculating, so the list
shows what is actually saved in each file, i.e. exactly what a recipient sees
when they open it offline.

**Add revision to ticked** — put an `x` in column A and fill in the
`New Rev` / `New Status` / `New Date` / ... columns on that row. Anything left
blank falls back to the *New revision* block on the Setup sheet.

So both jobs are the same operation:

- **All 24 on one revision** — fill in the Setup block once, tick every row.
- **6 of 24, on different revisions** — type each one on its own row, tick
  those six, leave the rest alone.

It appends one row per revision table, refuses a revision number that is
already in that file, and clears the typed values on the rows that succeeded.
The tick stays, so **Export ticked to PDF** straight after issues the same set.

The new line goes in **directly under the last revision**, not at the bottom of
the table. Revision tables are usually drawn with spare rows below the last
entry, and appending past them leaves a gap that makes the title block look
like it skipped a revision. The table only grows by a row when every row in it
is already used.

The `Select` and `New Rev` headers carry hover notes explaining this on the
sheet, and the same is written under the buttons on the Setup sheet.

**Export ticked to PDF** — writes a PDF of each ticked schedule. The only button
that never writes to a schedule at all. Full detail below.

## Setup sheet

Three blocks, top to bottom: **Project**, **Options**, **New revision**, with
the suitability codes in column F and the buttons to the right.

**Anything shaded yellow is yours to fill in.** Grey italic cells are written
by the tool. Everything else is a label or a note.

Options are **written** at fixed rows but **read** by their label in column A.
Every time an option has been added, the rows below it have shifted and the old
values ended up under new meanings. Reading by label means an older sheet still
reads correctly and `InstallTool` simply tidies the layout.

`InstallTool` rebuilds this sheet each time it runs. It clears the labels in
column A, columns C and H to K, and rewrites them, because option rows move
between versions and a label left on an old row is worse than no label at all.
Option values are lifted out **by their label** first and put back on whatever
row that label now sits, so adding an option never leaves an old value under a
new name.

It does **not** touch the project values in `B1:B4`, the New revision values,
the **PROJECT FIELDS** block (names or values: schedules link to those cells by
address), column F (the suitability codes), or any other sheet in the workbook,
so a README sheet of your own is safe. Versions before 1.2 cleared the project
field **names** on every install; if yours went missing, type them back in.

| Cell | Meaning |
|---|---|
| `B7` Schedules folder | blank = the folder this file is saved in. Filled in automatically if the tool has to ask |
| `B8` Backup before changes | `Yes` / `No` |
| `B9` Refresh list on open | `No` by default. `Yes` reads every schedule at startup |
| `B10` Full refresh every time | `No` by default, which reopens only files whose modified date changed since the last refresh. `Yes` always reopens everything |
| `B11` Reference schedule | the schedule that is set up correctly. Used by both copy buttons. Blank = the button asks, then fills this in |
| `B12` Header image | logo for the top-right of the header. Blank = the button asks; decline and headers stay text only |
| `B13` Header image scale % | that logo's size as a percentage of the image's own size. Default `20` |
| `B14` Security classification | what every schedule should be marked: `None`, `OFFICIAL`, `OFFICIAL-SENSITIVE`, `CONFIDENTIAL`, `COMMERCIAL IN CONFIDENCE`. ScheduleList flags any file that is not. Blank = compare with most of the others instead |
| `B15`, `B16` | when setup and the refresh last ran. Written by the tool, read only |
| `F1:F...` Suitability Codes | the dropdown pushed into every schedule's Metadata sheet and revision table. Seeded with the ISO 19650 codes below; edit here and re-run setup to push the change to all files |

## Security classification

The marking normally sits in the headers and footers, sometimes in a cell of
its own on the cover. The refresh reads it from every printed sheet of every
schedule and shows it in the **Classification** column:

- `OFFICIAL`, `CONFIDENTIAL` and so on, when every marking agrees
- `None` when there is no marking anywhere
- `CONFIDENTIAL + OFFICIAL` when sheets disagree, which is how a stray
  footer shows up

`Checks` also names any printed sheet that has no marking while the others do,
and any file that does not match **Security classification** on the Setup
sheet. Leave that blank and each file is compared with the majority instead.

What counts as a marking: `OFFICIAL`, `OFFICIAL-SENSITIVE` (or with a space),
`CONFIDENTIAL`, `COMMERCIAL IN CONFIDENCE`, `RESTRICTED`, `SECRET`,
`TOP SECRET`, in any case. In headers and footers it is found anywhere in the
text, formatting codes and all. In cells it only counts when it is all the cell
says, and only on the Front Cover, the Revision Page and the title block rows
of a schedule sheet, so a word in the body of a schedule is never mistaken for
one. The `Delref Classification` row is never read or changed.

### Set classification

Choose the marking in **Security classification** on the Setup sheet, tick the
schedules, press **Set classification**. For each one:

- every marking already in a header or footer is changed **where it stands**,
  keeping its font, size and colour
- every cell on the cover or revision page that says nothing but a marking is
  changed (or emptied, for `None`)
- a printed sheet with no marking at all gets one, bold, in the centre of its
  footer, and of its header if that is empty
- `None` takes markings out, and a header or footer section left with nothing
  but formatting codes is emptied

Nothing else in a header or footer is touched: not the logo, page numbers,
orientation or margins. Each file is then read back, and anything other than
the chosen marking on every printed sheet is a **failure**: the file is closed
without saving and the Log says what it found. A marking shown through a
formula is never overwritten, and is named.

A file that is already right is not saved, so its modified date does not move.

## Headers and footers

**Copy headers & footers** takes the reference schedule's headers and footers
and puts them on the ticked ones. Use it for the whole header and footer
design: the logo, page numbers, a document reference. For only the marking,
**Set classification** is the direct way.

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

### The logo in the header

Excel does not keep the original path of an embedded header image, so
`PageSetup.RightHeaderPicture.Filename` comes back empty and the image cannot
be read out of another workbook. It **can** be written, though, so the logo
comes from a file you pick rather than from the source workbook.

The first time you press the button it asks whether you want a logo, and if so
which file. The answer is remembered on the Setup sheet, so it only asks once.
The image goes in the **top right of the header**, scaled by
`Setup!B13` (default **20%** of the image's own size). Any text already in the
right-hand header section is kept, not thrown away.

Keep the image somewhere everyone can reach: a company network path, or in the
project folder itself. If it sits with the schedules only its file name is
stored, so the setting survives the folder moving between Filery and local.

### Every write is checked

Header and footer text is read back after it is written and compared with what
was asked for. Anything that did not take is reported as a **failure**, named
on the Log sheet with what is there versus what should be. A file only counts
as succeeded when the check passes.

An earlier version wrapped the writes in `Application.PrintCommunication =
False` for speed. That queues page-setup changes rather than applying them, and
they can be silently dropped. It has been removed: 24 files is a few seconds
either way, and a tool that reports work it did not do is worse than useless in
a QA workflow.

The reference schedule is remembered in `Setup!B11`, the logo in `Setup!B12`
and its scale in `Setup!B13`, so re-running after a tweak is one click. The
same `B11` reference is used by **Copy cover & revision page**.

## Tidying the sheets

**Tidy sheets** puts the ticked workbooks the same way round:

```
Front Cover | Revision Page | Schedule | (Metadata, hidden)
```

A workbook with **one** schedule sheet has it renamed to `Schedule`, which
catches the `Sheet1` that most of them arrive with. One with **several** keeps
its names, because they cannot all be called `Schedule` and those names are the
only thing telling the sheets apart. Either way the log names what changed.

Metadata is set to hidden, not very hidden, so it can still be unhidden by hand
when you want to look at it. The other sheets are made sure to be visible.

**Only names, order and visibility change. No cell on any sheet is touched.** A
workbook that is already tidy is closed without saving and logged as unchanged.

## Renaming files

**Rename files** does a find and replace across schedule file names, on the
ticked rows when any are ticked, otherwise on every row. The
obvious use is swapping the `PROJECTNUMBER` placeholder in
`Z9A...-PROJECTNUMBER-SC-M-...xlsx` for the real number across all 24 at once.

It takes **two presses**:

1. The first asks what to find and what to replace it with, then writes the
   proposed names into the **New FileName** column and stops. Nothing is
   renamed. Read them, edit any you want different, clear any you do not want.
2. The second does the renaming.

Typing a name straight into that column skips to step 2. Leave the extension
off and the current one is kept.

Every problem is found **before anything is renamed**, and any problem refuses
the whole batch: an illegal character, a file that is open, a changed file
type, or two rows that would both become the same name. A half-renamed folder
is far harder to unpick than a refusal.

### Chains and swaps

A target that already exists is only a problem when the file holding it is
**not** itself being renamed. When it is, that is a chain or a swap, and it is
worked out rather than refused:

```
A -> B, B -> C        B is renamed first, freeing the name for A
A -> B, B -> A        both are parked under temporary names, then placed
```

Files are renamed in any order whose target is free, repeatedly, so a chain
unwinds from its end. Whatever is left after that is a genuine loop, where
every target is held by another file in the batch; those are moved to
`_renaming1_<name>` first to break it, then put in place. Every step is on the
Log sheet, including the temporary names, so an interrupted run can be unpicked
by reading it.

### Why it reopens each file

Renaming does not break any links: schedules link to the MPI, never to each
other, and the MPI finds them by scanning the folder.

It is not cosmetic though. `Metadata!DocumentNumber` is derived from the file
name via `CELL("filename")`, and that flows through to `Document no` on the
Revision Page. The value stored in the file would keep showing the old document
number until somebody happened to open it. So each renamed file is reopened,
recalculated and saved, which puts the right number in the title block
immediately.

## Copying the cover and revision page

**Copy cover & revision page** replaces a schedule's `Front Cover` and
`Revision Page` with the ones from the reference schedule, sheet and all. Use
it when the pages themselves change: dropping a security classification on a
project that does not need one, moving a logo, rewording the cover. Set one
schedule up by hand, then push those two sheets out.

Each target keeps what belongs to the document rather than the template:

| Kept | Why |
|---|---|
| the schedule title (`Revision Page!A4`) | it is that document's name |
| the **entire revision history** | the table is a record of what was issued |
| Document type, Delref Classification, BSUID, Trigger Events | typed per document |

Everything else on those two sheets comes from the reference. The revision
history is written back **by column header**, so it survives the reference
reordering its columns, and the table is always cleared first, so a schedule
with no history does not inherit the reference's.

**A schedule with no Front Cover or Revision Page gets them.** Drop a bare
schedule into the folder, press this, and it is given both sheets from the
reference with an empty revision table. Its title is taken from the
`SCHEDULE OF ...` heading on its own sheet, or failing that from the file name
after the last ` - `. The schedule sheet itself is not otherwise touched.

Existing Front Cover and Revision Page sheets are **replaced**, not left alone.
The log says which it did for every file, and how many revision lines it kept,
because a silent success reads exactly like having done nothing.

Afterwards the normal repair runs on the file, which is what makes the links
local: nothing ends up pointing back at the reference workbook. That includes
formulas **you** wrote by hand on the cover. Copying a sheet turns
`='Revision Page'!B26` into `='[Golden.xlsx]Revision Page'!B26`, and the repair
now points every such reference back at this workbook's own sheet, as long as a
sheet of that name exists here. References to sheets this workbook does not
have are left alone, because those are real links somewhere else.

A reference to **somebody else's MAINPROJECTINFO** is repointed at this one
rather than reported. Schedules have no `Setup` sheet of their own, so a
formula reading `'[OLD-MPI.xlsm]Setup'!$B$3` can only have been meant for the
current MPI, and it is rewritten. That covers a renamed MPI and a schedule
copied in from another project, both of which used to survive every pass and
be flagged as "STILL LINKED" without anything being done about it.

A link the tool genuinely cannot resolve is still only reported, never broken,
but the message now names the sheet and cells holding it so there is somewhere
to look. Breaking one would hardcode whatever value was cached, which is worse
than a warning.

Both passes read references in the **path-qualified** form Excel uses whenever
the other workbook is closed:

```excel
='C:\Projects\[SRP1141-RAM-XX-XX-T-M-0023.xlsx]Revision Page'!B12
```

That is always the form here, since the workbook a sheet was copied from is
closed by the time the repair runs on the copy.

Only `Metadata!B4:B6` and the project fields still reach outside, to the MPI,
exactly as everywhere else.

It works on the ticked rows, like every other button. The reference itself is
always skipped. A file that fails is closed **without saving**, so a
half-copied schedule never gets written. Its confirmation box defaults to No
even when rows are ticked.

This is the most destructive thing the tool does. Leave `Backup before changes`
on, and the confirmation box says so in as many words when it is off.

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

## Title block

**Title block** draws the block at the top of each ticked schedule sheet so
nobody merges cells by hand:

```
| SCHEDULE OF ...  | DOCUMENTS REFERENCE: ...  |                     |
| CLIENT: ...      | PURPOSE OF DOCUMENT: ...  | REVISION: ...       |
| PROJECT: ...     | PROJECT NUMBER: ...       | REVISION DATE: ...  |
| PREPARED BY: ... | CHECKED BY: ...           | APPROVED BY: ...    |

| General Notes: ... merged across the full width                   |
```

It stays in cells rather than the page header on purpose. A native header is
capped at 255 characters (this block is already about 240), cannot hold a
formula so it would go stale the moment anyone edited a revision by hand,
shows only in print preview, and cannot have borders. In cells every value
is a live formula on the local Metadata sheet, and it shows in a screenshot.

What it does to each schedule sheet:

- **Width**: the block spans the schedule's own columns, the print area if one
  is set, otherwise the rightmost column with anything in it below the notes.
- **Splits**: where the three groups divide is **kept** from the existing block
  when it still spans exactly those columns and every group's text fits.
  Otherwise the columns are divided in proportion to each group's longest line.
  So a repair never reshapes a block that already looks right, and adding a
  column re-fits it.
- **Style**: the same fields in the same places as the project template, laid
  out for reading on paper: the schedule title larger and bold, the revision
  bold, the three groups aligned left / centre / right instead of each in its
  own box, one box round the whole block with a rule under the title row.
  Verdana throughout, and Shrink to fit on every block cell, so text that
  would overrun its group gets slightly smaller instead of being cut off. A
  block drawn in another style is brought into line.
- **Size on paper**: Excel shrinks a wide schedule to fit the page, and a title
  block in cells shrinks with it. That is how the widest schedules printed
  theirs at about 3pt. So the tool works out how far Excel will shrink each
  sheet (from the zoom, or from the paper, margins and column widths when it
  is set to fit to a page) and enlarges the block and the notes to cancel it
  out. Every schedule then prints its title at 11pt, the rest of the block at
  7.5pt and the notes at 7pt, however many columns it has. On a wide schedule
  that makes the block look large in Normal view; that is expected. The notes
  box grows by the same factor as its text, so text that fitted still fits.
- **Notes**: the General Notes box (a merged cell in column A just below the
  block, starting `General Notes` or `Notes`) is merged across the full width.
  Its text is never rewritten, so bold and line breaks in it survive.
- **Every page**: print titles become rows 1 to the end of the notes box, or
  further if they already reached further (that is where the schedule's own
  column headings are). Notes repeat on every page on purpose: whoever has page
  3 of 5 needs them too.
- **Page numbers**: `Page &P of &N` goes in the right footer when the sheet has
  no page number anywhere and that section is empty.

**A schedule with no title block gets one.** 12 rows are inserted above it:
the block, a gap, an empty `General Notes:` box six rows tall, and a gap.
Everything below moves down, formulas, print area and print titles included,
as Excel does for any inserted row. Grow the notes box by inserting rows
inside it. The Log reminds you to check that the print titles take in the
schedule's own column headings, since the tool cannot tell which rows those
are on a sheet it has never seen.

**It stops rather than guesses.** A sheet is left alone, and named as a
failure on the Log, when:

- rows 1-4 hold text that is neither a title block label nor the schedule title
- title block labels sit anywhere other than rows 1-4
- there are fewer than six labels in rows 1-4, too few to be sure it is a block
- the schedule is under three columns wide

A schedule whose block, sizes and print titles are already exactly right is
not saved, so its modified date does not move.

**Press it again after adding or removing columns, or changing the paper size,
orientation or margins.** Those change how far Excel shrinks the sheet, and so
the size the block needs to be. Hand edits to the revision table need nothing:
every value in the block is a live formula.

## Schedules that print their own title block

Some schedules carry the title block across the top of the schedule sheet
itself, in merged cells, repeating on every printed page through Print Titles.
That is handy for sending one page as a screenshot, and the repair keeps it
live.

Those cells hold the label and the value **in one cell**:

```
"CLIENT: Laing O'Rourke"
```

so the repair turns each into a concatenation:

```excel
="CLIENT: "&Metadata!B6
="REVISION DATE: "&TEXT(Metadata!B9,"dd/mm/yyyy")
```

It reads exactly as before and keeps its formatting; it just stops being typed.
Cells are found by the label text, not by address, so the merge widths can
differ from file to file, which in practice they do.

Recognised labels, and what each reads:

| Label | From |
|---|---|
| `DOCUMENTS REFERENCE` | the file name |
| `CLIENT`, `PROJECT`, `PROJECT NUMBER` | the MPI |
| `REVISION`, `REVISION DATE`, `PURPOSE OF DOCUMENT` | the revision table |
| `PREPARED BY`, `CHECKED BY`, `APPROVED BY` | the revision table |

Spelling variants are matched and **corrected to one spelling**, with every
change named on the Log sheet: `CLIENTS:` becomes `CLIENT:`,
`DOCUMENT REFERENCE:` becomes `DOCUMENTS REFERENCE:`, and a stray character
after the colon is dropped. Add a spelling to the `Select Case` in
`TitleBlockField` to have it recognised.

A cell that already holds a formula is left alone, so the repair is safe to
re-run and will not fight an edit made on purpose.

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

## Extra project fields

Anything else the whole project shares — a DfE code, a framework reference,
whatever a client wants on every title block — goes in the **PROJECT FIELDS**
block at the bottom of the Setup sheet. Name in column A, value in column B.

The repair then does two things with each one:

1. adds it to every schedule's Metadata sheet, one row after the fixed ones,
   linked to the MPI exactly like Project Name is;
2. if the Revision Page has a **label in column A with that exact name**, points
   the cell beside it at that Metadata row.

So adding a field is: type `DfE Code` and its value on the Setup sheet, add a
`DfE Code` label to the reference schedule's Revision Page, then with nothing
ticked press **Set up / repair links** and **Copy cover & revision page** to
give every other schedule the same label. After that the value comes from the MPI everywhere.

The rows are rewritten in full on every repair, so deleting a field on the MPI
deletes it from every schedule too.

## What gets written to Metadata

| Row | Header | Source |
|---|---|---|
| 2 | DocumentNumber | the file name: up to ` - ` when there is one, otherwise the whole name |
| 3 | ScheduleName | `Revision Page!A4` |
| 4 | Project Name | **MPI** |
| 5 | Project Number | **MPI** |
| 6 | Client | **MPI** |
| 7-14 | DocumentType, Revision, Date, Prepared/Checked/Approved by, Suitability Status, Suitability Description | the local Revision Page title block |

Rows 7-14 exist so the MPI can read one predictable sheet per schedule instead
of hunting for labels in a formatted page. Rows 2-6 keep the header names the
old files already used, so nothing existing breaks.

## Progress, speed and the log

The **Log** sheet keeps the last **5 runs**, newest at the top. It has to: every
button finishes by refreshing the list, so a log holding one run only ever
showed the refresh, never the thing you actually pressed. Each run is a
title row with the time, then a line per file. Change `LOG_RUNS` at the top of
`modMain.bas` to keep more or fewer.

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
re-run **Set up / repair links** (nothing ticked, Yes to all) to push the new
formulas into every file.

The repair rewrites these formulas **everywhere in the workbook**, not just the
seven cells on the Revision Page, so a copy of the same formula sitting on a
front cover or in a schedule header gets the new ranking too. It only touches
formulas that already rank `RevisionTable[Revision]` with `MAX()`, works out
what each one returns, and leaves anything it cannot identify alone.

## What the repair does and does not touch

It **does** rebuild every cell that should be derived: the Metadata sheet, the
linked cells on the Front Cover and Revision Page, the revision formulas, the
suitability dropdown, leftover links and dead defined names.

It **does not** copy layout or static text between workbooks. That is what
**Copy cover & revision page** is for, deliberately separate because it
overwrites whatever is there now.

### The one cell it writes on a schedule sheet

Only the title, the cell reading `SCHEDULE OF ...`, and only when it already
matches the Revision Page title. If they disagree it is reported and left
alone. Nothing else on a schedule sheet is ever touched: not a row, not a
column, not a format.

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

## Exporting to PDF

**Export ticked to PDF** with rows ticked asks one thing:

> Export the 6 ticked schedules to `<project folder>\_pdf`, one PDF per
> workbook, replacing any PDF of the same name?
>
> Yes - export / No - change where they go, or split them by sheet / Cancel

So re-issuing the same set is: tick, press, Yes. **No** opens the full set of
choices below for that run.

With **nothing ticked** it asks the old four questions, which also covers
schedules outside the project folder:

1. **Which schedules** — every workbook in a folder you pick, or the files you
   pick yourself (Ctrl or Shift to select several, from anywhere).
2. **Where the PDFs go** — `_pdf` inside the folder they came from, created if
   it is not there, or a folder you pick.
3. **One PDF per workbook or one per sheet** — per workbook is the usual
   answer: cover, revision page and schedule come out as a single document with
   continuous page numbers.
4. **Overwrite existing PDFs** — only asked when the destination already holds
   some.

Every sheet in the workbook is exported, in tab order, except `Metadata` and
any sheet with nothing on it to print. An empty sheet in the selection fails
the whole export, which is why they are filtered out rather than left to break
a file that would otherwise have been fine.

A PDF already in the way is deleted before the export writes, with its
read-only attribute cleared if it has one. Excel refuses to overwrite a
read-only file and reports it only as "Document not saved", so the file is
moved out of the way first and a PDF that genuinely cannot be replaced, one
open in a viewer or held by the Explorer preview pane, is named on the Log
sheet.

Two files of the same name, picked from different folders, are refused before
anything is written: they would produce one PDF, and the second would silently
replace the first.

A workbook already open in Excel is skipped rather than closed underneath you.
Nothing is ever written back to a schedule.

### Why this one opens files differently

Every other button opens a schedule with **links not updated and calculation
off**, so it reads what is actually saved in the file. This one opens them
**read only but otherwise the ordinary way**: links updated, calculation on.

That is not a preference, it is a bug that took a while to find. A schedule
opened cold with calculation manual, events off and screen updating off
exported with **a line struck through every calculated value**, while the same
file exported by hand was clean. No cell reported `Font.Strikethrough` and
there was no conditional formatting, so nothing in the file explained it.
Bisecting the settings one at a time: each is harmless on its own, the
combination is not. So the export turns off only `DisplayAlerts` and
`AskToUpdateLinks`, which cannot affect what is rendered.

`BeginQuiet` in `modMain` takes a `keepEvents` flag for exactly this, and
`OpenQuiet` takes an `updateLinks` argument that defaults to 0 for everything
else.

Exporting several sheets as one document is done the way the Publish dialog
does it: select them, then export the **active sheet**, which means "active
sheet(s)" and takes the whole selected group. `ExportAsFixedFormat` is a method
of a `Workbook`, a `Worksheet` or a `Chart`, and **not** of the `Sheets`
collection, so `ActiveWindow.SelectedSheets.ExportAsFixedFormat` does not
compile. When the wanted sheets are every visible sheet there is, which is what
a tidy schedule looks like, the workbook is exported directly and the selection
is left alone.

## Known limits

- The title block's split points are chosen from an estimate of how wide
  Verdana capitals are (`TB_CHAR_W` in `modSchedule.bas`), and its size from an
  estimate of how far Excel will shrink the sheet. Both are within a few
  percent; Shrink to fit covers the first. The on-paper sizes are the `TB_..._PT`
  constants at the top of `modSchedule.bas`.
- The ScheduleList columns changed in 1.2 (`RevDescription` and
  `Classification` were added). `InstallTool` empties the list, and the first
  refresh after it is a full one. Ticks on the old list are not carried over.
- `DocumentNumber` comes from `CELL("filename")`, so it is blank until the file
  has been saved once.
- The title cell is located within the first 60 rows / 10 columns of a sheet.
- Revision numbers above 999 would collide across families.
- A PDF path longer than 255 characters is refused rather than truncated, which
  a deep synced folder plus `_pdf` can reach.
