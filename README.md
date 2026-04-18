# Square Fare Toolkits

Internal Streamlit app for Square Fare meal-planning operations. Runs the weekly portioning algorithm and generates the supporting production artifacts (stickers, to-make sheets, one-pagers) from Airtable data.

## What it does

- **Portioning algorithm** — computes per-client dish portions and pushes results back to Airtable.
- **Dish stickers** — PPT labels for each dish (Airtable-driven and barcode variants).
- **Shipping stickers** — shipping-label PPTs (two generators; `v2` is the current one).
- **One-pager** — per-client PPT summary.
- **To-make sheet** — Excel production sheet for the kitchen.
- **Client servings export** — Excel report of client serving data.
- **Landing-page cache refresh** — button that pings the external `orders.getsquarefare.com` backend.

## Running locally

```bash
pip install -r requirements.txt
streamlit run streamlitController.py
```

Requires Python 3.10 (`runtime.txt`). Secrets (Airtable token, OpenAI key, etc.) go in `.streamlit/secrets.toml` — not committed.

Always launch from the repo root; generators load templates from `template/` via relative paths.

## Layout

| Path | What's in it |
|---|---|
| `streamlitController.py` | Streamlit entry point |
| `src/portioning/` | Portion algorithm + dish optimizers (rule-based and LLM) |
| `src/data/` | Airtable access layer and shared exceptions |
| `src/generators/` | Excel + PPT output generators (to-make sheet, one-pager, client servings) |
| `src/stickers/` | Dish and shipping sticker PPT generators |
| `template/` | `.pptx` and `.csv` templates used at runtime |
| `legacy/` | Code not wired into the app; kept for reference |

See `CLAUDE.md` for the Claude-specific version with more details on imports and data flow.
