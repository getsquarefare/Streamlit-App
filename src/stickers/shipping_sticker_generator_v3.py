import sys
import copy
import hashlib
import logging
from io import BytesIO
from pathlib import Path
from datetime import datetime

import barcode
from barcode.writer import ImageWriter
from pptx import Presentation
from pptx.util import Pt, Inches

BASE_DIR = Path(__file__).resolve().parents[2]  # Streamlit-App
sys.path.append(str(BASE_DIR))

from src.data.exceptions import AirTableError
from src.data.store_access import new_database_access
from src.stickers.dish_barcode_ids import dish_barcode_from_open_order_fields


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

VIEW = "viwDpTtU0qaT9NcvG"


class PPTGenerationError(Exception):
    pass


def unwrap(value, default=""):
    if isinstance(value, list):
        return value[0] if value else default
    return value if value is not None else default


def format_phone(phone):
    if not phone:
        return ""

    phone_digits = "".join(filter(str.isdigit, str(phone)))
    if len(phone_digits) == 10:
        return f"{phone_digits[:3]}-{phone_digits[3:6]}-{phone_digits[6:]}"
    return str(phone)


def make_bag_group_key(record):
    return (
        str(record.get("Delivery Date", "")).strip(),
        str(record.get("Shipping Name", "")).strip().lower(),
        str(record.get("Shipping Address 1", "")).strip().lower(),
        str(record.get("Shipping Address 2", "")).strip().lower(),
        str(record.get("Shipping City", "")).strip().lower(),
        str(record.get("Shipping Province", "")).strip().lower(),
        str(record.get("Shipping Postal Code", "")).strip(),
        str(record.get("Zone Number", "")).strip(),
    )


def customization_tags_from_fields(fields):
    """
    Airtable field names vary (e.g. 'Customization Tags' vs
    'Customization Tags (from To_Match_Client_Nutrition)').
    Merge string values from every field whose name contains 'Customization Tags'.
    """
    parts = []
    for key, val in (fields or {}).items():
        if "Customization Tags" not in str(key):
            continue
        if isinstance(val, list):
            parts.extend(str(x) for x in val if str(x).strip())
        elif val is not None and str(val).strip():
            parts.append(str(val))
    return " ".join(parts)


def make_bag_barcode(shipping_info, chunk_index=0):
    delivery_date = str(shipping_info.get("Delivery Date", "")).replace("-", "")
    if not delivery_date:
        delivery_date = datetime.now().strftime("%Y%m%d")

    zone = str(shipping_info.get("Zone Number", "X")).replace("ZONE", "").replace("Zone", "").strip()

    name_addr = "|".join([
        str(shipping_info.get("Shipping Name", "")).strip().lower(),
        str(shipping_info.get("Shipping Address 1", "")).strip().lower(),
        str(shipping_info.get("Shipping Address 2", "")).strip().lower(),
        str(shipping_info.get("Shipping Postal Code", "")).strip(),
        str(chunk_index),
    ])
    short_hash = hashlib.sha1(name_addr.encode("utf-8")).hexdigest()[:6].upper()

    return f"BAG-{delivery_date}-Z{zone}-{short_hash}"


def process_order_data(db):
    """
    Read open orders and group them into shipping/bag records.
    """
    logger.info("Starting shipping sticker barcode data processing")

    orders = db.get_all_open_orders(view=VIEW)

    if not orders:
        raise AirTableError("No open orders found")

    logger.info(f"Found {len(orders)} open orders")

    shipping_data = []

    for order in orders:
        fields = order.get("fields", {})

        required_fields = ["Shipping Name", "Shipping Address 1", "Quantity"]
        if not all(fields.get(field) for field in required_fields):
            logger.warning(f"Skipping order with missing fields: {order.get('id', 'unknown')}")
            continue

        meal_portion = unwrap(fields.get("Meal Portion", ""))
        quantity = unwrap(fields.get("Quantity", 0), 0)

        try:
            quantity = float(quantity)
        except Exception:
            quantity = 0

        if meal_portion == "Breakfast":
            adjusted_quantity = quantity * 0.8
        elif meal_portion == "Snack":
            adjusted_quantity = quantity * 0.1
        else:
            adjusted_quantity = quantity

        zone = unwrap(fields.get("Zone Number (from Delivery Zone)", "N/A"))
        delivery_date = unwrap(fields.get("Delivery Date", ""))

        dish_barcode = dish_barcode_from_open_order_fields(fields)

        shipping_record = {
            "Delivery Date": delivery_date,
            "Shipping Name": unwrap(fields.get("Shipping Name", "")),
            "Shipping Address 1": unwrap(fields.get("Shipping Address 1", "")),
            "Shipping Address 2": unwrap(fields.get("Shipping Address 2", "")),
            "Shipping City": unwrap(fields.get("Shipping City", "")),
            "Shipping Province": str(unwrap(fields.get("Shipping Province", ""))).upper(),
            "Shipping Postal Code": unwrap(fields.get("Shipping Postal Code", "")),
            "Shipping Phone": format_phone(unwrap(fields.get("Shipping Phone", ""))),
            "Zone Number": str(zone),
            "Quantity": adjusted_quantity,

            # Mapping data for bag scan: same id as Open Orders "Portion Result (in ClientServings)" (Client Servings record id).
            "Dish Barcode": dish_barcode,
            "Customer Name": unwrap(fields.get("Customer Name", "")),
            "Meal Portion": meal_portion,
            "Meal Sticker": unwrap(fields.get("Meal Sticker", "")),
            "Customization Tags": customization_tags_from_fields(fields),
        }

        shipping_data.append(shipping_record)

    grouped_shipping = {}

    for record in shipping_data:
        key = make_bag_group_key(record)

        if key not in grouped_shipping:
            grouped_shipping[key] = {
                "Delivery Date": record["Delivery Date"],
                "Shipping Name": record["Shipping Name"],
                "Shipping Address 1": record["Shipping Address 1"],
                "Shipping Address 2": record["Shipping Address 2"],
                "Shipping City": record["Shipping City"],
                "Shipping Province": record["Shipping Province"],
                "Shipping Postal Code": record["Shipping Postal Code"],
                "Shipping Phone": record["Shipping Phone"],
                "Zone Number": record["Zone Number"],
                "Total Quantity": 0,
                "Household Members": set(),
                "Dishes": [],
                "Ice Pack Required": False,
            }

        grouped_shipping[key]["Total Quantity"] += record["Quantity"]

        if record["Customer Name"]:
            grouped_shipping[key]["Household Members"].add(record["Customer Name"])

        dish_line = " - ".join(
            part for part in [
                record.get("Customer Name", ""),
                record.get("Meal Portion", ""),
                record.get("Meal Sticker", ""),
            ]
            if str(part).strip()
        )

        if dish_line:
            grouped_shipping[key]["Dishes"].append(
                {
                    "dishBarcode": record.get("Dish Barcode", ""),
                    "customerName": record.get("Customer Name", ""),
                    "mealPortion": record.get("Meal Portion", ""),
                    "mealSticker": record.get("Meal Sticker", ""),
                    "displayText": dish_line,
                    "status": "unscanned",
                    "adjustedQuantity": record["Quantity"],
                }
            )

        if "Ice Pack" in str(record.get("Customization Tags", "")):
            grouped_shipping[key]["Ice Pack Required"] = True

    shipping_list = []
    portion_per_bag = 6.8

    for shipping_info in grouped_shipping.values():
        dishes = shipping_info["Dishes"]
        household_members = sorted(list(shipping_info["Household Members"]))

        chunks = []
        current_chunk = []
        current_qty = 0.0
        for dish in dishes:
            dish_qty = float(dish.get("adjustedQuantity", 0) or 0)
            if current_chunk and current_qty + dish_qty > portion_per_bag:
                chunks.append(current_chunk)
                current_chunk = []
                current_qty = 0.0
            current_chunk.append(dish)
            current_qty += dish_qty
        if current_chunk:
            chunks.append(current_chunk)
        if not chunks:
            chunks = [[]]

        for chunk_index, chunk_dishes in enumerate(chunks):
            bag = {
                **shipping_info,
                "Dishes": chunk_dishes,
                "Household Members": household_members,
                "Stickers Needed": 1,
                "Bag Barcode": make_bag_barcode(shipping_info, chunk_index),
            }
            shipping_list.append(bag)

    # Sort by zone (numeric when possible), then delivery date, then name (A-Z)
    def sort_key(info):
        zone_raw = str(info.get("Zone Number", ""))
        try:
            zone_sort = (0, int(float(zone_raw)))
        except (ValueError, TypeError):
            zone_sort = (1, zone_raw)
        return (
            zone_sort,
            str(info.get("Delivery Date", "")),
            str(info.get("Shipping Name", "")).strip().lower(),
        )

    shipping_list.sort(key=sort_key)

    logger.info(f"Processed {len(shipping_list)} bag records")

    return shipping_list


def copy_slide(template_slide, target_prs):
    new_slide = target_prs.slides.add_slide(template_slide.slide_layout)

    for shape in new_slide.shapes:
        sp = shape._element
        sp.getparent().remove(sp)

    for shape in template_slide.shapes:
        new_shape = copy.deepcopy(shape)
        new_slide.shapes._spTree.insert_element_before(new_shape._element, "p:extLst")

    return new_slide


def populate_sticker(slide, shipping_info):
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue

        text_frame = shape.text_frame

        for paragraph in text_frame.paragraphs:
            text = paragraph.text

            if "Shipping Name" in text:
                paragraph.text = shipping_info["Shipping Name"]
                paragraph.font.size = Pt(28)
                paragraph.font.name = "Lato"

            elif "Address" in text:
                if shipping_info["Shipping Address 2"]:
                    paragraph.text = f"{shipping_info['Shipping Address 1']}, {shipping_info['Shipping Address 2']}"
                else:
                    paragraph.text = shipping_info["Shipping Address 1"]
                paragraph.font.size = Pt(24)
                paragraph.font.name = "Lato"

            elif "City" in text:
                paragraph.text = (
                    f"{shipping_info['Shipping City']}, "
                    f"{shipping_info['Shipping Province']} "
                    f"{shipping_info['Shipping Postal Code']}"
                )
                paragraph.font.size = Pt(24)
                paragraph.font.name = "Lato"

            elif "Shipping Phone" in text:
                paragraph.text = str(shipping_info["Shipping Phone"])
                paragraph.font.size = Pt(24)
                paragraph.font.name = "Lato"

            elif "ZONE" in text:
                paragraph.text = f"ZONE {shipping_info['Zone Number']}"
                paragraph.font.size = Pt(28)
                paragraph.font.name = "Lato"


def add_code128_barcode(slide, prs, barcode_value):
    CODE128 = barcode.get_barcode_class("code128")

    barcode_obj = CODE128(barcode_value, writer=ImageWriter(mode="RGBA"))
    barcode_buffer = BytesIO()

    barcode_obj.write(
        barcode_buffer,
        {
            "quiet_zone": 2,
            "font_size": 9,
            "text_distance": 6,
            "module_height": 11,
            "module_width": 0.3,
            "background": "white",
            "foreground": "black",
        },
    )

    barcode_buffer.seek(0)

    # Adjust these values after checking first PPT output.
    slide.shapes.add_picture(
        barcode_buffer,
        Inches(5.55),
        Inches(3.9),
        width=Inches(3.2),
        height=Inches(0.95),
    )


def upsert_bag_to_airtable(db, shipping_info):
    dish_ids = [
        str(d["dishBarcode"])
        for d in shipping_info["Dishes"]
        if isinstance(d, dict) and d.get("dishBarcode")
    ]
    try:
        db.upsert_bag_record(
            bag_barcode=shipping_info["Bag Barcode"],
            dish_ids=dish_ids,
            ice_pack_required=shipping_info["Ice Pack Required"],
        )
    except Exception as e:
        logger.warning(f"Could not write bag {shipping_info['Bag Barcode']} to Airtable: {e}")


def create_shipping_stickers_barcode_ppt(db, shipping_list, template_path=None):
    if template_path is None:
        template_path = BASE_DIR / "template" / "Shipping_Sticker_Template.pptx"

    prs = Presentation(str(template_path))

    if len(prs.slides) == 0:
        raise PPTGenerationError("Template presentation has no slides")

    template_slide = prs.slides[0]
    total_stickers = 0

    for shipping_info in shipping_list:
        stickers_needed = shipping_info["Stickers Needed"]

        for _ in range(stickers_needed):
            slide = copy_slide(template_slide, prs)
            populate_sticker(slide, shipping_info)
            add_code128_barcode(slide, prs, shipping_info["Bag Barcode"])
            total_stickers += 1

        upsert_bag_to_airtable(db, shipping_info)

    # Remove original template slide.
    r_id = prs.slides._sldIdLst[0].rId
    prs.part.drop_rel(r_id)
    del prs.slides._sldIdLst[0]

    logger.info(f"Generated {total_stickers} shipping stickers with barcode")

    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output


def generate_shipping_stickers_barcode(db):
    template_path = BASE_DIR / "template" / "Shipping_Sticker_Template.pptx"

    shipping_list = process_order_data(db)

    if not shipping_list:
        raise AirTableError("No shipping records to process")

    ppt_file = create_shipping_stickers_barcode_ppt(
        db,
        shipping_list,
        template_path=template_path,
    )

    return ppt_file, shipping_list


if __name__ == "__main__":
    try:
        db = new_database_access()

        ppt_file, shipping_list = generate_shipping_stickers_barcode(db)

        output_path = BASE_DIR / f"shipping_stickers_barcode_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"

        with open(output_path, "wb") as f:
            f.write(ppt_file.getvalue())

        logger.info(f"Generated {len(shipping_list)} unique bag records")
        logger.info(f"Saved PPT: {output_path}")

    except AirTableError as e:
        logger.critical(f"Airtable error: {str(e)}")
    except Exception as e:
        logger.critical(f"Unexpected error: {str(e)}")