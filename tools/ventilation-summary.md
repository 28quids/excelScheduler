# Ventilation calc: Plant List, Summary and colour coding

`build_vent_summary.py` wires up a Stage 4 ventilation calculation workbook so
the Summary fills itself in from the Ventilation sheet, and so adding a new
unit is one line of typing rather than a formatting job.

```
python3 tools/build_vent_summary.py "Stage 4i Ventilation Calcs.xlsx" out.xlsx
```

It is a one-off conversion, not something to run every time. Once a workbook
has been through it, the wiring lives in the workbook itself.

**There are no macros.** Everything below is formulas, conditional formatting
and data validation, so the file stays a plain `.xlsx` that opens on any
machine with nothing to enable and nothing to trust.

## How it hangs together

Three sheets, in the order you use them.

**Plant List** is the one you type into. One row per unit: the name, what it
serves, a note. Everything else follows from it.

**Ventilation** is the calc, untouched apart from column T. That column is now
a dropdown of the units on the Plant List, and every row takes the colour of
the unit it is on.

**Summary** is output only. One row per Plant List row, totalling the supply
and extract of every room on that unit and applying the margins.

So adding `MVHR-003` is: type it on the next free Plant List row, then pick it
from the dropdown in column T against the rooms it serves. The colour, the
Summary row and the totals are already there.

## Adding a unit

1. Type the name on the next free row of the **Plant List**. It picks up the
   next colour in the palette straight away.
2. On the **Ventilation** sheet, pick it from the dropdown in column T for
   each room it serves.

That is the whole job. The **Summary** row for that unit was created the
moment the name went on the Plant List; the flow rates arrive as the rooms
are assigned to it.

Delete a unit by clearing its name on the Plant List. Its Summary row goes
blank and the rooms it served stop being counted, which will show up in the
checks.

## Where the numbers come from

Per unit, on the Summary:

| Column | What it is |
|---|---|
| Fan Design, Supply | `SUMIF` of Ventilation column **O** for every room on that unit |
| Fan Design, Extract | the same over column **P** |
| Fan Duty | design flow x `Brief and Assumptions!M36` (1.16, duct leakage 6% x spare capacity 10%) |
| Rooms | how many Ventilation rows are on that unit |

The formulas read rows 10 to 500 of the Ventilation sheet, so rooms added
below the current last row are picked up with no change to the Summary.

## Checks

Under the Summary table:

- **Rooms on a system that is not on the Plant List** - should be 0. This is
  the one that catches a mistyped unit name. A room typed as `MVHR-3` when the
  Plant List says `MVHR-03` is in nobody's total, and nothing else on the
  sheet would tell you.
- **Rooms whose name appears more than once** - should be 0. Duplicated room
  names are also highlighted orange in column D.
- **Rooms with no mechanical ventilation** - for information, not an error.

The dropdown in column T deliberately does not block a name that is not on the
Plant List, because `-` for a room with no mechanical ventilation is a normal
thing to write. The check above is what catches the real mistakes.

## Colour

Eighteen colours: Excel's six theme accents at lighter 80%, 60% and 40%. A
unit gets the colour of the Plant List row it sits on, shown as a swatch in the
Colour column, and that colour is applied to the room details and the system
cell of every Ventilation row on that unit, and to the bands either side of its
Summary row.

The colour is conditional formatting keyed to the Plant List, not paint. Move a
room to a different unit and it recolours itself; there is nothing to tidy up
afterwards.

A nineteenth unit still totals correctly, it just comes out uncoloured.

## What the conversion changes in the calc

The arithmetic is left alone. Three things around it are not:

- **Merged system cells are split.** `T29:T34` and `T35:T43` were single merged
  cells spanning several rooms. A unit assigned to one of those blocks would
  have counted once rather than once per room, so every row carries its own
  value now.
- **Calculated cells are wrapped in `IFERROR(..., "-")`.** Rooms with no rate
  decided yet were showing `#VALUE!` through columns N, O and P, and an error
  in one room breaks any total that includes it.
- **The old column D conditional formatting is replaced.** Those rules were
  comparing each room name against one arbitrary other cell, and one of them
  had decayed to `#REF!`. One rule using `COUNTIF` does what they were reaching
  for.

## How the file is written

openpyxl silently drops the parts of an `.xlsx` it does not model: embedded
images, printer settings, the header logo, custom XML from the document
management system. So the script edits a copy, then builds the output by
taking the original file and swapping in only the parts it actually changed -
the three sheets, the style table, and the workbook's defined names. Every
other part is copied across byte for byte.

The one part that is dropped is `calcChain.xml`, which only records the order
Excel last calculated cells in. Excel rebuilds it on open.
