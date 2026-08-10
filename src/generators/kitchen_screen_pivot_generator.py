import io
from collections import defaultdict
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT


PIVOT_VIEW       = "Pivot Table"
FIELD_DISH       = "Dish Name"
FIELD_INGREDIENT = "Ingredient"
FIELD_GRAMS      = "Grams"
FIELD_COMPONENT  = "Component"

COMPONENT_ORDER  = ["meat", "veggies", "starch", "sauce", "garnish"]
GRAMS_TO_LBS     = 1 / 453.592

# PDF Template colors
NAVY       = colors.HexColor("#1A3A5C")
COMP_COLORS = {
    "meat":    colors.HexColor("#FFE4E1"),
    "veggies": colors.HexColor("#E8F5E9"),
    "starch":  colors.HexColor("#FFF9C4"),
    "sauce":   colors.HexColor("#E3F2FD"),
    "garnish": colors.HexColor("#F3E5F5"),
    "other":   colors.HexColor("#F5F5F5"),
}


def _parse_records(records):
    pivot_data        = defaultdict(dict)
    ingredient_totals = defaultdict(float)

    for record in records:
        fields     = record.get("fields", {})
        dish       = str(fields.get(FIELD_DISH, "")).strip()
        ingredient = str(fields.get(FIELD_INGREDIENT, "")).strip()
        component  = str(fields.get(FIELD_COMPONENT, "")).strip().lower()
        try:
            grams = float(fields.get(FIELD_GRAMS, 0) or 0)
        except (TypeError, ValueError):
            grams = 0.0

        if not dish or not ingredient:
            continue

        if ingredient not in pivot_data[dish]:
            pivot_data[dish][ingredient] = {"component": component, "grams": 0.0}
        pivot_data[dish][ingredient]["grams"] += grams
        ingredient_totals[ingredient] += grams

    return dict(pivot_data), dict(ingredient_totals)


def _build_pdf(pivot_data, ingredient_totals):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=0.75 * inch,
        bottomMargin=0.75 * inch,
    )

    styles   = getSampleStyleSheet()
    story    = []

    # Styling
    def make_style(name, parent="Normal", **kwargs):
        return ParagraphStyle(name, parent=styles[parent], **kwargs)

    title_s   = make_style("TitleS",  "Heading1", fontSize=13, textColor=NAVY, spaceAfter=4)
    sub_s     = make_style("SubS",    "Normal",   fontSize=8,  textColor=colors.grey, spaceAfter=10)
    dish_s    = make_style("DishS",   "Heading2", fontSize=11, textColor=NAVY, spaceAfter=4)
    hdr_s     = make_style("HdrS",    "Normal",   fontSize=9,  textColor=colors.white, alignment=TA_CENTER)
    cell_l    = make_style("CellL",   "Normal",   fontSize=9,  alignment=TA_LEFT)
    cell_c    = make_style("CellC",   "Normal",   fontSize=9,  alignment=TA_CENTER)
    cell_r    = make_style("CellR",   "Normal",   fontSize=9,  alignment=TA_RIGHT)

    # Page title
    story.append(Paragraph("Square Fare — Kitchen Screen Production Report", title_s))
    story.append(HRFlowable(width="100%", thickness=1, color=NAVY, spaceAfter=14))

    if not pivot_data:
        story.append(Paragraph(
            "⚠️ No data found in the 'Pivot Table' view. "
            "Please apply a Dish Name filter in Airtable first, then regenerate.",
            styles["Normal"]
        ))
        doc.build(story)
        buf.seek(0)
        return buf

    col_widths = [3.2 * inch, 1.2 * inch, 1.3 * inch, 1.8 * inch]
    dishes     = sorted(pivot_data.keys())

    for d_idx, dish in enumerate(dishes):
        ingredients = pivot_data[dish]

        story.append(Paragraph(f"Dish: {dish}", dish_s))

        # Group by component and sort alphabetically
        grouped = defaultdict(list)
        for ing, data in ingredients.items():
            if data["grams"] > 0:
                comp = data["component"] or "other"
                grouped[comp].append((ing, data))
        for comp in grouped:
            grouped[comp].sort(key=lambda x: x[0].lower())

        # Header row
        tbl_data = [[
            Paragraph("Ingredients (alphabetically by component)", hdr_s),
            Paragraph("Component", hdr_s),
            Paragraph("lbs (cooked)", hdr_s),
            Paragraph("% Overall Production", hdr_s),
        ]]

        per_row_bg = []
        row_idx    = 1

        for comp in COMPONENT_ORDER + [c for c in grouped if c not in COMPONENT_ORDER]:
            if comp not in grouped:
                continue
            bg = COMP_COLORS.get(comp, COMP_COLORS["other"])

            for ing, data in grouped[comp]:
                grams  = data["grams"]
                lbs    = grams * GRAMS_TO_LBS
                lbs_display = max(lbs, 0.1) if lbs > 0 else 0.0
                total  = ingredient_totals.get(ing, grams)
                pct    = (grams / total * 100) if total > 0 else 0

                tbl_data.append([
                    Paragraph(ing, cell_l),
                    Paragraph(comp.title(), cell_c),
                    Paragraph(f"{lbs_display:.1f}", cell_r),
                    Paragraph(f"{pct:.0f}%", cell_r),
                ])
                per_row_bg.append(("BACKGROUND", (0, row_idx), (-1, row_idx), bg))
                row_idx += 1

        if len(tbl_data) == 1:
            story.append(Paragraph("  No ingredients with grams > 0 for this dish.", styles["Normal"]))
        else:
            tbl = Table(tbl_data, colWidths=col_widths, repeatRows=1)
            tbl_style = TableStyle([
                # Header
                ("BACKGROUND",    (0, 0), (-1, 0), NAVY),
                ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE",      (0, 0), (-1, 0), 9),
                ("TOPPADDING",    (0, 0), (-1, 0), 7),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 7),
                ("LINEBELOW",     (0, 0), (-1, 0), 1, NAVY),
                # Data
                ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE",      (0, 1), (-1, -1), 9),
                ("TOPPADDING",    (0, 1), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
                ("LEFTPADDING",   (0, 0), (-1, -1), 8),
                ("RIGHTPADDING",  (0, 0), (-1, -1), 8),
                ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
                # Grid
                ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
                # Alignment
                ("ALIGN",         (0, 1), (0, -1), "LEFT"),
                ("ALIGN",         (1, 1), (1, -1), "CENTER"),
                ("ALIGN",         (2, 1), (3, -1), "RIGHT"),
            ])
            for cmd in per_row_bg:
                tbl_style.add(*cmd)
            tbl.setStyle(tbl_style)
            story.append(tbl)

        story.append(Spacer(1, 18))
        if d_idx < len(dishes) - 1:
            story.append(PageBreak())

    doc.build(story)
    buf.seek(0)
    return buf


def generate_kitchen_pivot(db) -> io.BytesIO:
    # Filtered records
    filtered_records = db.get_kitchen_screen_data(view=PIVOT_VIEW)

    if not filtered_records:
        raise ValueError(
            "No records found in the 'Pivot Table' view. "
            "Please apply a filter in Airtable first, then regenerate."
        )

    all_records = db.get_kitchen_screen_data()

    pivot_data, _ = _parse_records(filtered_records)

    _, ingredient_totals = _parse_records(all_records)

    return _build_pdf(pivot_data, ingredient_totals)