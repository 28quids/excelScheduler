"""Rebuild the Summary and Plant List sheets of a Ventilation Calcs workbook.

The workbook keeps its own calculations. What this adds is the wiring:

  Plant List   one row per ventilation unit. Type a unit here and it gets a
               colour, a row on the Summary and a place in the dropdown on
               the Ventilation sheet.
  Ventilation  column T (System) becomes a dropdown fed by the Plant List,
               and every row takes the colour of the unit it is on.
  Summary      one row per Plant List entry, totalling the supply and extract
               of every room on that unit, plus the fan duty with margins.

Nothing here is a macro. It is formulas, conditional formatting and data
validation, so the workbook stays a plain .xlsx.

Run:  python3 tools/build_vent_summary.py <in.xlsx> <out.xlsx>

The write is surgical: only the Ventilation, Summary and Plant List sheets,
the style table and the string table are replaced in the .xlsx package. Every
other part (images, page setup, comments, document properties) is copied over
byte for byte, because openpyxl drops parts it does not understand.
"""

import re
import shutil
import sys
import zipfile
from copy import copy

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

# --- layout constants -------------------------------------------------------

VENT = "Ventilation"
SUMM = "Summary"
PLANT = "Plant List"
BRIEF = "Brief and Assumptions"

VENT_FIRST = 10          # first room row on the Ventilation sheet
VENT_LAST = 500          # how far the summary formulas look
VENT_SYS_COL = 20        # column T, System

PLANT_FIRST = 5          # first unit row on the Plant List
SLOTS = 18               # units the colour palette covers

SUMM_FIRST = 10          # Summary row showing Plant List row PLANT_FIRST
SUMM_LAST = SUMM_FIRST + SLOTS - 1
SUMM_TOTAL = SUMM_LAST + 2
SUMM_CHECKS = SUMM_TOTAL + 2

MARGIN = "'%s'!$M$36" % BRIEF     # duct leakage x spare capacity, 1.16

# Accent 1-6 at "lighter 80/60/40%", the tints Excel's own palette uses. The
# first three are the colours the workbook already gave AHU-01, MVHR-01 and
# MVHR-02, so those units keep the colour the drawings were marked up in.
TINTS = (0.7999816888943144, 0.5999938962981048, 0.3999755851924192)
PALETTE = [(theme, tint) for tint in TINTS for theme in range(4, 10)]

THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def band(slot):
    """The solid fill for colour slot `slot`."""
    theme, tint = PALETTE[slot]
    return PatternFill("solid", fgColor=Color(theme=theme, tint=tint))


def dxf_band(slot):
    """The same colour as a conditional-formatting differential style."""
    theme, tint = PALETTE[slot]
    return DifferentialStyle(fill=PatternFill(bgColor=Color(theme=theme, tint=tint)))


def unmerge_rows(ws, first, last):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= first and rng.max_row <= last:
            ws.unmerge_cells(str(rng))


# --- Ventilation ------------------------------------------------------------

def do_ventilation(wb):
    ws = wb[VENT]

    # 1. One system per row. T29:T34 and T35:T43 were merged blocks, so a
    #    system typed into one of them would only have counted once.
    for rng in ("T29:T34", "T35:T43", "U29:U34", "U35:U43"):
        if rng in [str(r) for r in ws.merged_cells.ranges]:
            ws.unmerge_cells(rng)
    style_src = ws["T28"]
    for r in range(29, 44):
        for col, src in ((20, ws["T28"]), (21, ws["U28"])):
            cell = ws.cell(r, col)
            cell._style = copy(src._style)
            if cell.value is None:
                cell.value = "-"

    # 2. A rate that has not been decided yet reads "-" rather than #VALUE!.
    #    Rows with no air change rate and no litres per person were erroring
    #    all the way through N, O and P, which also poisons any total that
    #    includes them.
    for r in range(VENT_FIRST, 44):
        for col in "JLNOPR":
            cell = ws["%s%d" % (col, r)]
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "IFERROR" not in v.upper():
                cell.value = '=IFERROR(%s,"-")' % v[1:]

    # 3. The system colour comes from the Plant List now, so drop the fills
    #    that were painted on by hand.
    for r in range(VENT_FIRST, VENT_LAST + 1):
        for col in (1, 2, 3, 4, VENT_SYS_COL):
            ws.cell(r, col).fill = PatternFill()

    # 4. Pick the unit from a dropdown instead of typing it.
    # A dropdown, not a rule: typing something else is allowed, because "-"
    # for a room with no mechanical ventilation is a normal thing to write.
    # The Summary carries a check that counts anything typed here that is not
    # on the Plant List, which is the safety net for a mistyped unit.
    dv = DataValidation(
        type="list",
        formula1="SystemList",
        allow_blank=True,
        showErrorMessage=False,
        showInputMessage=True,
    )
    dv.promptTitle = "System"
    dv.prompt = ("Pick a unit from the Plant List, or type - for a room with no "
                 "mechanical ventilation.")
    ws.add_data_validation(dv)
    dv.add("T%d:T%d" % (VENT_FIRST, VENT_LAST))

    # 5. Colour. One rule per Plant List row; a room takes the colour of the
    #    unit in column T. Room names that appear twice stay flagged orange,
    #    which is what the old (broken, #REF!) rules were trying to do.
    ws.conditional_formatting = type(ws.conditional_formatting)()

    def rule(formula, fill, priority, stop=False):
        r = FormulaRule(formula=[formula], fill=fill, stopIfTrue=stop or None)
        r.priority = priority
        return r

    ws.conditional_formatting.add(
        "D%d:D%d" % (VENT_FIRST, VENT_LAST),
        rule('AND($D%d<>"",COUNTIF($D$%d:$D$%d,$D%d)>1)'
             % (VENT_FIRST, VENT_FIRST, VENT_LAST, VENT_FIRST),
             PatternFill(bgColor="FFFFC000"), 1, stop=True))

    for slot in range(SLOTS):
        test = 'AND($T%d<>"",$T%d=\'%s\'!$A$%d)' % (
            VENT_FIRST, VENT_FIRST, PLANT, PLANT_FIRST + slot)
        fill = dxf_band(slot).fill
        ws.conditional_formatting.add("A%d:D%d" % (VENT_FIRST, VENT_LAST),
                                      rule(test, fill, 2 + slot))
        ws.conditional_formatting.add("T%d:T%d" % (VENT_FIRST, VENT_LAST),
                                      rule(test, fill, 2 + slot))


# --- Plant List -------------------------------------------------------------

HEAD = ["System", "Type", "Serves / location", "Colour", "Rooms", "Notes"]
SEED = [("AHU-01", ""), ("MVHR-01", ""), ("MVHR-02", "")]


def do_plant_list(wb):
    ws = wb[PLANT]
    ws.sheet_view.showGridLines = False

    title = Font(name="Arial", size=11, bold=True)
    small = Font(name="Arial", size=8)
    head = Font(name="Arial", size=8, bold=True)
    yellow = PatternFill("solid", fgColor="FFFF00")
    grey = PatternFill("solid", fgColor="F2F2F2")

    ws["A1"] = "Ventilation Plant List"
    ws["A1"].font = title
    ws["A2"] = ("Add a unit here and it gets a colour, a row on the Summary and a place in the "
                "System dropdown on the Ventilation sheet. Yellow cells are yours to fill in.")
    ws["A2"].font = small
    ws["A3"] = ("Rooms counts how many rows on the Ventilation sheet are on that unit. A new unit "
                "showing 0 means the name here and the name in column T do not match.")
    ws["A3"].font = small

    for i, text in enumerate(HEAD):
        cell = ws.cell(PLANT_FIRST - 1, 1 + i, text)
        cell.font = head
        cell.border = BOX
        cell.fill = grey
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for slot in range(SLOTS):
        r = PLANT_FIRST + slot
        for col in range(1, len(HEAD) + 1):
            cell = ws.cell(r, col)
            cell.font = small
            cell.border = BOX
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(r, 1).fill = yellow                       # System
        ws.cell(r, 3).fill = yellow                       # Serves
        ws.cell(r, 6).fill = yellow                       # Notes
        ws.cell(r, 3).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(r, 6).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(r, 4).fill = band(slot)                   # Colour swatch
        ws.cell(r, 2).value = '=IF($A{r}="","",LEFT($A{r},FIND("-",$A{r}&"-")-1))'.format(r=r)
        ws.cell(r, 5).value = (
            '=IF($A{r}="","",COUNTIF({v}!$T${f}:$T${l},$A{r}))'
            .format(r=r, v=VENT, f=VENT_FIRST, l=VENT_LAST))
        ws.cell(r, 2).fill = grey
        ws.cell(r, 5).fill = grey

    for slot, (name, serves) in enumerate(SEED):
        ws.cell(PLANT_FIRST + slot, 1).value = name
        ws.cell(PLANT_FIRST + slot, 3).value = serves

    foot = PLANT_FIRST + SLOTS + 1
    ws.cell(foot, 1, "Colours run out after %d units. A %dth unit still totals correctly on the "
                     "Summary, it just comes out uncoloured." % (SLOTS, SLOTS + 1)).font = small

    for col, width in zip("ABCDEF", (12, 10, 34, 8, 7, 46)):
        ws.column_dimensions[col].width = width


# --- Summary ----------------------------------------------------------------

def do_summary(wb):
    ws = wb[SUMM]

    unmerge_rows(ws, SUMM_FIRST, 46)
    for row in ws.iter_rows(min_row=SUMM_FIRST, max_row=46, max_col=12):
        for cell in row:
            cell.value = None

    # Header. Column G was an unused spacer; it carries the room count now,
    # and the Notes heading moves across to sit over the notes cells.
    ws["G4"] = None
    ws["H4"] = "Notes"
    ws["G5"] = "Rooms"
    ws["G5"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    base = {col: copy(ws["%s11" % col]._style) for col in "ABCDEFGHIJ"}

    for slot in range(SLOTS):
        r = SUMM_FIRST + slot
        plant_row = PLANT_FIRST + slot
        for col in "ABCDEFGHIJ":
            ws["%s%d" % (col, r)]._style = copy(base[col])
        ws.merge_cells("H%d:J%d" % (r, r))

        ws.cell(r, 1).fill = band(slot)
        ws.cell(r, 6).fill = band(slot)

        a = "$A%d" % r
        ws.cell(r, 1).value = "=IF('{p}'!$A{k}=\"\",\"\",'{p}'!$A{k})".format(p=PLANT, k=plant_row)
        ws.cell(r, 2).value = flow_formula(a, "O", r)
        ws.cell(r, 3).value = flow_formula(a, "P", r)
        ws.cell(r, 4).value = '=IF({a}="","",ROUND($B{r}*{m},0))'.format(a=a, r=r, m=MARGIN)
        ws.cell(r, 5).value = '=IF({a}="","",ROUND($C{r}*{m},0))'.format(a=a, r=r, m=MARGIN)
        ws.cell(r, 7).value = ('=IF({a}="","",COUNTIF({v}!$T${f}:$T${l},{a}))'
                               .format(a=a, v=VENT, f=VENT_FIRST, l=VENT_LAST))
        for col in (2, 3, 4, 5, 7):
            ws.cell(r, col).number_format = "0"

    # Total
    for col in "ABCDEFGHIJ":
        ws["%s%d" % (col, SUMM_TOTAL)]._style = copy(base[col])
    ws.cell(SUMM_TOTAL, 1, "TOTAL").font = Font(name="Arial", size=8, bold=True)
    ws.cell(SUMM_TOTAL, 1).fill = PatternFill()
    ws.cell(SUMM_TOTAL, 6).fill = PatternFill()
    for col in (2, 3, 4, 5, 7):
        c = ws.cell(SUMM_TOTAL, col)
        letter = c.column_letter
        c.value = "=SUM({0}{1}:{0}{2})".format(letter, SUMM_FIRST, SUMM_LAST)
        c.number_format = "0"
        c.font = Font(name="Arial", size=8, bold=True)

    # Checks. Each one should read 0; anything else is a room the Summary is
    # not seeing.
    checks = [
        ("Rooms on a system that is not on the Plant List - should be 0",
         '=SUMPRODUCT(--({v}!$T${f}:$T${l}<>""),--({v}!$T${f}:$T${l}<>"-"),'
         "--(COUNTIF('" + PLANT + "'!$A${pf}:$A${pl},{v}!$T${f}:$T${l})=0))"),
        ("Rooms whose name appears more than once - should be 0",
         '=SUMPRODUCT(--({v}!$D${f}:$D${l}<>""),'
         "--(COUNTIF({v}!$D${f}:$D${l},{v}!$D${f}:$D${l})>1))"),
        ("Rooms with no mechanical ventilation (column T blank or -) - for information",
         '=SUMPRODUCT(--({v}!$D${f}:$D${l}<>""),--(({v}!$T${f}:$T${l}="")+({v}!$T${f}:$T${l}="-")>0))'),
    ]
    head = Font(name="Arial", size=8, bold=True)
    small = Font(name="Arial", size=8)
    ws.cell(SUMM_CHECKS, 1, "Checks").font = head
    for i, (label, formula) in enumerate(checks):
        r = SUMM_CHECKS + 1 + i
        ws.merge_cells("A%d:F%d" % (r, r))
        c = ws.cell(r, 1, label)
        c.font = small
        c.alignment = Alignment(horizontal="left", vertical="center")
        v = ws.cell(r, 7)
        v.value = formula.format(v=VENT, f=VENT_FIRST, l=VENT_LAST,
                                 pf=PLANT_FIRST, pl=PLANT_FIRST + SLOTS - 1)
        v.font = small
        v.border = BOX
        v.number_format = "0"
        v.alignment = Alignment(horizontal="center", vertical="center")

    for col, width in zip("ABCDEFG", (14, 11, 11, 11, 11, 3, 8)):
        ws.column_dimensions[col].width = width
    for r in range(SUMM_FIRST, SUMM_CHECKS + len(checks) + 1):
        ws.row_dimensions[r].height = None


def flow_formula(a, vent_col, r):
    """Total one unit's design flow.

    SUMIF skips the rooms that read "-", and every calculated cell on the
    Ventilation sheet is wrapped in IFERROR, so a rate that has not been
    decided yet drops out of the total instead of breaking it.
    """
    return ('=IF({a}="","",ROUND(SUMIF({v}!$T${f}:$T${l},{a},'
            '{v}!${c}${f}:${c}${l}),0))'
            ).format(a=a, v=VENT, c=vent_col, f=VENT_FIRST, l=VENT_LAST)


# --- package surgery --------------------------------------------------------

EDITED = ("xl/worksheets/sheet3.xml", "xl/worksheets/sheet4.xml", "xl/worksheets/sheet5.xml")


def splice_tail(new_xml, old_xml):
    """Keep the original page setup, header/footer and drawing references.

    openpyxl rewrites a sheet without the r:id references to printer settings
    and the header logo. The original rels files are kept, so pasting the
    original tail back on keeps those working.
    """
    i = new_xml.find("<pageMargins")
    j = old_xml.find("<pageMargins")
    if i < 0 or j < 0:
        return new_xml
    spliced = new_xml[:i] + old_xml[j:old_xml.rfind("</worksheet>")] + "</worksheet>"
    # The tail carries r:id references; openpyxl's root does not declare r.
    if "xmlns:r=" not in spliced[:spliced.find(">")]:
        spliced = spliced.replace(
            "<worksheet ",
            '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument'
            '/2006/relationships" ', 1)
    return spliced


def repackage(original, edited, out):
    src = zipfile.ZipFile(original)
    new = zipfile.ZipFile(edited)

    # openpyxl writes its strings inline rather than into the shared string
    # table, so the original table can be left exactly as it is for the sheets
    # that are not being touched.
    sheets = {}
    for name in EDITED:
        sheets[name] = splice_tail(new.read(name).decode("utf-8"),
                                   src.read(name).decode("utf-8"))

    # The Plant List range the Ventilation dropdown reads.
    wbxml = src.read("xl/workbook.xml").decode("utf-8")
    named = "<definedName name=\"SystemList\">'%s'!$A$%d:$A$%d</definedName>" % (
        PLANT, PLANT_FIRST, PLANT_FIRST + SLOTS - 1)
    wbxml = wbxml.replace("<definedNames>", "<definedNames>" + named, 1)

    # calcChain lists which cells to recalculate in which order. It is stale
    # now, and Excel rebuilds it on open, so drop it and every reference to it.
    ct = src.read("[Content_Types].xml").decode("utf-8")
    ct = re.sub(r'<Override[^>]*calcChain[^>]*/>', "", ct)
    rels = src.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rels = re.sub(r'<Relationship[^>]*calcChain[^>]*/>', "", rels)

    replaced = {
        "[Content_Types].xml": ct,
        "xl/_rels/workbook.xml.rels": rels,
        "xl/workbook.xml": wbxml,
        "xl/styles.xml": new.read("xl/styles.xml").decode("utf-8"),
    }
    replaced.update(sheets)

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            if item.filename == "xl/calcChain.xml":
                continue
            if item.filename in replaced:
                dst.writestr(item.filename, replaced[item.filename])
            else:
                dst.writestr(item, src.read(item.filename))


def main(src_path, out_path):
    work = re.sub(r"\.xlsx$", "", out_path) + ".tmp.xlsx"
    shutil.copyfile(src_path, work)
    wb = openpyxl.load_workbook(work)
    do_ventilation(wb)
    do_plant_list(wb)
    do_summary(wb)
    wb.defined_names.add(DefinedName(
        "SystemList", attr_text="'%s'!$A$%d:$A$%d" % (PLANT, PLANT_FIRST, PLANT_FIRST + SLOTS - 1)))
    wb.save(work)
    repackage(src_path, work, out_path)
    print("wrote", out_path)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
