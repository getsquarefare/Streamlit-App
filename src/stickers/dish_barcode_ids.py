"""
Dish identifier for bag/shipping JSON and kitchen scans.

Open Orders field "Portion Result (in ClientServings)" links to the Client Servings
row; its value is the linked record id (e.g. rec…). That is the single dish id —
no fallbacks to order "#" or Position Id.
"""

PORTION_RESULT_IN_CLIENT_SERVINGS = "Portion Result (in ClientServings)"


def dish_barcode_from_open_order_fields(fields):
    """
    Return the Client Servings record id from Open Orders, or "" if missing.
    """
    if not fields:
        return ""
    raw = fields.get(PORTION_RESULT_IN_CLIENT_SERVINGS)
    if isinstance(raw, list):
        v = raw[0] if raw else None
    else:
        v = raw
    if v is None:
        return ""
    s = str(v).strip()
    return s if s and s.lower() != "none" else ""
