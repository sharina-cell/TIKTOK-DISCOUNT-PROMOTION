# Marketplace Discount Promotion Upload File Builder

Streamlit app that turns a marketplace Seller Center report + a Masterfile
into ready-to-upload discount promo files, with any number of discount tiers
in one pass (e.g. BAU / D-day / SWFS on Shopee, BAU / 9.9 / MM & PD on TikTok).

Built generically across two axes:
- **Account** - EWG, DBC, FYW, GSK, etc. Only the Masterfile column positions
  differ per account; map once, save a preset, reuse it.
- **Marketplace** - Shopee and TikTok today, Lazada/Zalora ready to add the
  same way. Each platform has real differences (which column is the SKU used
  to match the Masterfile, what the final upload template looks like) - those
  live in one config block (`PLATFORM_SPECS` in `discount_engine.py`), not
  scattered through the code.

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```

## What it does
1. **Marketplace report** - pick Shopee or TikTok, then upload that
   marketplace's raw Seller Center export. The header row is auto-detected
   (marketplace exports have junk rows above the real header). Which column
   is used to match products against the Masterfile is platform-specific:
   - **Shopee**: SKU (col F), falling back to Parent SKU (col E) when blank.
     RRP = live "Price" (col G).
   - **TikTok**: Seller SKU (col G). Live price = "Retail Price (Local
     Currency)" (col F). (TikTok's upload file is keyed by Product ID + SKU
     ID instead, which are pulled through separately.)
2. **Masterfile** - pick which tab to use, then map:
   - SKU column
   - RRP column
   - One or more discount tiers, each with its own name + source column
     (add/remove tiers freely)
3. **Working sheet** - merges the report + Masterfile on SKU, adds:
   - `RRP CHECK` (Masterfile RRP vs live marketplace price)
   - `<tier> DISCOUNT` = 1 - SRP / live price, for every tier
   - `remarks` = "no discount remarks" when every tier's discount computes to
     0% or negative for a row
   - Products with no Masterfile match (bundles, GWP items, etc.) show
     `#N/A` in RRP, RRP CHECK, every tier price, and every tier discount -
     matching the original working-file convention.
4. **RRP MISMATCH tab** - every row where RRP CHECK is False, pulled out for review.
5. **Exclusions (optional)** - type/paste SKUs to exclude, and/or upload a
   file of them (`.xlsx` or `.csv` - if a column header contains "sku" that
   column is used, otherwise the first column). Excluded SKUs stay visible in
   the working sheet (flagged `EXCLUDED = True`, remarks = "excluded from
   promo") but are left out of every `TO UPLOAD` tab.
6. **TO UPLOAD `<tier>` tabs** - one per tier, in that marketplace's official
   upload template layout. A row is only included for a tier if it has a
   valid price for that tier, a **positive discount** for that tier (0% or
   negative discounts are left out), and isn't excluded.

All outputs land in one downloadable workbook: `Working Sheet`,
`RRP MISMATCH`, and one `TO UPLOAD <tier>` tab per configured tier.

## Multi-account use
- Step 0 has **Marketplace** and **Account name** fields - the account name
  labels the downloaded workbook and mapping preset
  (e.g. `EWG_Shopee_Discount_Promo_Upload.xlsx`).
- After mapping the Masterfile columns, click **"Save this column mapping"**
  to download a small `.json` preset for that account + platform
  (`EWG_Shopee_discount_mapping.json`, `DBC_TikTok_discount_mapping.json`,
  etc). Keep one per account/platform combo.
- Next time you run that combo, upload its `.json` in Step 0 and the
  sheet/SKU/RRP/tier columns pre-fill automatically.

## Adding a new marketplace (e.g. Lazada, Zalora)
Add a new entry to `PLATFORM_SPECS` in `discount_engine.py`:
- `header_markers` - a few header texts (lowercased) unique enough to find
  the real header row in that export.
- `columns` - logical field name -> (header text, fallback column letter).
- `match_key` / `match_key_fallback` - which field(s) match against the
  Masterfile.
- `display_order` - which fields appear (and in what order) in the working sheet.
- `upload_template` - output column header -> logical field name,
  `"__tier_price__"`, or `None` for a blank optional column.
Nothing else needs to change - the UI and merge/output logic are generic.

## Files
- `app.py` - Streamlit UI
- `discount_engine.py` - all parsing/merge/output logic, platform-agnostic,
  reusable and testable independent of Streamlit
- Handles the marketplace exports' `activePane` XML defect automatically
  (same fix used across the other GSK/EWG/FYW report tools) - no manual
  patching needed before upload.

## Notes / things to double check each run
- If the Masterfile layout changes (different tab name or column positions),
  just re-pick them in the UI - nothing is hardcoded.
- Bundle SKUs and GWP items generally won't match the Masterfile 1:1 and
  will show blank tier/discount columns - that's expected, they're excluded
  from the upload tabs automatically.
- The "no discount remarks" rule fires when *all* configured tiers compute to
  0% or negative for a row. Per-tier, any tier at 0% or negative is also
  silently dropped from that tier's own upload sheet (the product can still
  appear in a different tier's upload if that one has a real discount).
