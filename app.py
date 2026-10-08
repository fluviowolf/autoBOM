import json
from collections import defaultdict
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


CATALOG_DIR = Path(__file__).parent / "catalogs"
BRAND_FILES = {
    "ZimVie Dental": CATALOG_DIR / "ZimVie_Dental.xlsx",
    "Dentsply Sirona": CATALOG_DIR / "Dentsply_Sirona.xlsx",
}
SHEETS = [
    "Implants",
    "Drills",
    "Accessories",
    "Drivers",
    "Abutments",
    "Healing Abutments",
    "Tools",
    "Sleeves",
    "Prosthesis",
    "Pilot Only",
]


def load_catalogs():
    catalogs = {}
    for brand, path in BRAND_FILES.items():
        if not path.exists():
            raise FileNotFoundError(f"Missing catalog workbook: {path}")
        workbook = pd.ExcelFile(path)
        catalogs[brand] = {
            sheet: pd.read_excel(workbook, sheet_name=sheet).where(pd.notna, None).to_dict("records")
            for sheet in SHEETS
            if sheet in workbook.sheet_names
        }
        for sheet in SHEETS:
            catalogs[brand].setdefault(sheet, [])
    return catalogs


def unpack(value, fallback):
    if value is None or value == "":
        return fallback
    if isinstance(value, str):
        try:
            return json.loads(value)
        except json.JSONDecodeError:
            return fallback
    return value


def part_cost(catalogs, part_number):
    for brand_data in catalogs.values():
        for sheet in ("Drills", "Tools", "Drivers", "Abutments", "Healing Abutments", "Prosthesis", "Pilot Only"):
            for item in brand_data[sheet]:
                if str(item.get("partNumber", "")) == str(part_number):
                    return float(item.get("unitCost") or 0)
    for brand_data in catalogs.values():
        for item in brand_data["Implants"]:
            sku = item.get("articleNumber") or item.get("id")
            if str(sku) == str(part_number):
                return float(item.get("unitCost") or 0)
    return 0.0


def item_details(catalogs, brand, part_number):
    for sheet in ("Drills", "Tools", "Sleeves", "Accessories", "Drivers", "Abutments", "Healing Abutments", "Prosthesis", "Pilot Only"):
        for item in catalogs[brand][sheet]:
            if str(item.get("partNumber", "")) == str(part_number):
                return item.get("description") or str(part_number), bool(str(part_number).startswith("PN-PLACEHOLDER"))
    return f"{part_number} (not in catalog)", True


def implant_sku(implant):
    return implant.get("articleNumber") or implant.get("id")


def implant_label(implant):
    line = implant.get("productLine") or implant.get("model") or "Implant"
    return f"O{implant['diameter']} mm x {implant['length']} mm | {line} ({implant['id']})"


def driver_for(brand_data, implant):
    for driver in brand_data["Drivers"]:
        allowed_lines = unpack(driver.get("allowedProductLines"), [])
        if (
            float(driver.get("diameter") or 0) == float(implant.get("diameter") or 0)
            and (not allowed_lines or implant.get("productLine") in allowed_lines)
            and (not driver.get("productLine") or driver.get("productLine") == implant.get("productLine"))
        ):
            return driver
    return None


def protocol_steps(brand, brand_data, implant, density, pilot_only):
    if pilot_only:
        pilot = next(iter(brand_data["Pilot Only"]), None)
        if pilot:
            return [(pilot["partNumber"], 1, False, False)]
        return [("33107", 1, False, False)]

    line = implant.get("productLine") or ""
    standard = unpack(implant.get("standard"), [])
    dense_extra = unpack(implant.get("denseExtra"), [])
    result = []
    if line == "PrimeTaper EV":
        for key in standard:
            tool = next((item for item in brand_data["Tools"] + brand_data["Sleeves"] if item.get("key") == key), None)
            if tool:
                result.append((tool["partNumber"], 1, False, bool(tool.get("isCortical"))))
            else:
                result.append((f"PN-PLACEHOLDER-PTEV-{key}", 1, False, False))
        return result

    if brand == "Dentsply Sirona":
        replaced = unpack(implant.get("denseReplaces"), []) if density == "dense" and implant.get("denseExclusive") else []
        result.extend((str(pn), 1, False, False) for pn in standard if str(pn) not in [str(value) for value in replaced])
        if density == "dense":
            result.extend((str(pn), 1, True, False) for pn in dense_extra)
        return result

    drills = brand_data["Drills"]
    pilot_drill = next((d for d in drills if float(d.get("diaLarge") or 0) == 2.4 and float(d.get("length") or 0) == 6), None)
    if pilot_drill:
        result.append((pilot_drill["partNumber"], 1, False, False))
    for diameter, length in standard:
        drill = next((d for d in drills if float(d.get("diaLarge") or 0) == float(diameter) and float(d.get("length") or 0) == float(length)), None)
        result.append((drill["partNumber"] if drill else f"PN-PLACEHOLDER-{diameter}x{length}", 1, False, False))
    if density == "dense":
        for diameter, length in dense_extra:
            drill = next((d for d in drills if float(d.get("diaLarge") or 0) == float(diameter) and float(d.get("length") or 0) == float(length)), None)
            result.append((drill["partNumber"] if drill else f"PN-PLACEHOLDER-{diameter}x{length}", 1, True, False))
    return result


def make_bom(catalogs, rows, include_implants, include_prosthesis, include_abutments, include_healing):
    totals = {}

    def add(brand, part_number, description, quantity, dense=False, cortical=False, placeholder=False):
        key = (brand, str(part_number))
        if key not in totals:
            totals[key] = {"Brand": brand, "Part Number": part_number, "Description": description, "Quantity": 0, "Dense bone only": True, "Cortical": False, "Placeholder": placeholder}
        item = totals[key]
        item["Quantity"] += quantity
        item["Dense bone only"] = item["Dense bone only"] and dense
        item["Cortical"] = item["Cortical"] or cortical
        item["Placeholder"] = item["Placeholder"] or placeholder

    for row in rows:
        implant = row["implant"]
        brand_data = catalogs[row["brand"]]
        quantity = row["quantity"]
        if include_implants:
            sku = implant_sku(implant)
            add(row["brand"], sku, implant_label(implant), quantity, placeholder=not implant.get("articleNumber"))
        for part_number, step_quantity, dense, cortical in protocol_steps(row["brand"], brand_data, implant, row["density"], row["pilot_only"]):
            description, placeholder = item_details(catalogs, row["brand"], part_number)
            add(row["brand"], part_number, description, quantity * step_quantity, dense, cortical, placeholder)
        if row["include_driver"]:
            driver = driver_for(brand_data, implant)
            if driver:
                add(row["brand"], driver["partNumber"], driver["description"], quantity, placeholder=str(driver["partNumber"]).startswith("PN-PLACEHOLDER"))
        if include_abutments and row.get("abutment"):
            abutment = row["abutment"]
            add(row["brand"], abutment["partNumber"], abutment["description"], quantity, placeholder=str(abutment["partNumber"]).startswith("PN-PLACEHOLDER"))
        if include_healing and row.get("healing_abutment"):
            healing = row["healing_abutment"]
            add(row["brand"], healing["partNumber"], healing["description"], quantity, placeholder=str(healing["partNumber"]).startswith("PN-PLACEHOLDER"))
        if include_prosthesis:
            for item in brand_data["Prosthesis"]:
                add(row["brand"], item["partNumber"], item["description"], quantity)

    result = []
    for item in totals.values():
        item["Unit cost"] = part_cost(catalogs, item["Part Number"])
        item["Line total"] = item["Quantity"] * item["Unit cost"]
        result.append(item)
    return result


def excel_bytes(bom, case_reference, udi_values):
    output = BytesIO()
    rows = []
    for item in bom:
        rows.append({
            "Case Reference": case_reference,
            "Brand": item["Brand"],
            "QTY": item["Quantity"],
            "SKU": item["Part Number"],
            "ITEM": str(item["Description"]).upper(),
            "UDI": udi_values.get((item["Brand"], item["Part Number"]), ""),
            "UNIT COST": item["Unit cost"],
            "LINE SUBTOTAL": item["Line total"],
            "DENSE BONE ONLY": "Yes" if item["Dense bone only"] else "No",
            "CORTICAL (MANDATORY)": "Yes" if item["Cortical"] else "No",
            "PLACEHOLDER": "Yes" if item["Placeholder"] else "No",
        })
    rows.append({"Case Reference": case_reference, "Brand": bom[0]["Brand"] if bom else "", "ITEM": "BOM ADD-ON SUBTOTAL", "LINE SUBTOTAL": sum(item["Line total"] for item in bom)})
    pd.DataFrame(rows).to_excel(output, index=False, sheet_name="Generated BOM")
    output.seek(0)
    return output.getvalue()


st.set_page_config(page_title="Implant BOM Generator", page_icon="📋", layout="wide")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
:root { --ink: #17332a; --leaf: #287353; --mist: #eff4ee; --paper: #fbfcf8; --line: #d7e1d9; }
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
h1, h2, h3 { font-family: 'Manrope', sans-serif; color: var(--ink); letter-spacing: 0; }
.stApp { background: radial-gradient(ellipse at 100% 0%, #e3eee4 0, transparent 34%), var(--paper); }
[data-testid="stHeader"] { background: transparent; }
div[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 6px; border-color: var(--line); }
div.stButton > button[kind="primary"] { background: var(--leaf); border-color: var(--leaf); }
div[data-testid="stMetric"] { background: var(--mist); padding: .8rem 1rem; border-left: 3px solid var(--leaf); }
</style>
""", unsafe_allow_html=True)

st.title("Implant BOM Generator")
st.caption("Build a concise, case-specific implant and drill bill of materials.")
st.warning("Tissue punches, bone levelers, and taps are excluded. Verify each generated BOM against current controlled protocol documentation. Catalog records marked as placeholders or with zero unit cost need review.")

try:
    catalogs = load_catalogs()
except Exception as error:
    st.error(f"Unable to read the Excel catalogs: {error}")
    st.stop()

brand_names = list(catalogs)
all_implants = [(brand, item) for brand in brand_names for item in catalogs[brand]["Implants"]]
if "bom_rows" not in st.session_state:
    first_brand, first_implant = all_implants[0]
    st.session_state.bom_rows = [{"brand": first_brand, "implant_id": first_implant["id"], "quantity": 1, "density": "standard", "pilot_only": False, "include_driver": False, "abutment_id": "", "healing_head": "", "healing_id": ""}]
if "udi_values" not in st.session_state:
    st.session_state.udi_values = {}

with st.container(border=True):
    case_reference = st.text_input("Case reference", placeholder="e.g. CASE-2026-0417")
    filter_brand = st.selectbox("Brand for new implant rows", ["All brands", *brand_names])

options_by_brand = {brand: catalogs[brand]["Implants"] for brand in brand_names}
with st.container(border=True):
    title_col, add_col = st.columns([5, 1])
    title_col.subheader("Implant selections")
    if add_col.button("＋ Add implant", type="primary", use_container_width=True):
        allowed = all_implants if filter_brand == "All brands" else [(filter_brand, item) for item in options_by_brand[filter_brand]]
        brand, implant = allowed[0]
        st.session_state.bom_rows.append({"brand": brand, "implant_id": implant["id"], "quantity": 1, "density": "standard", "pilot_only": False, "include_driver": False, "abutment_id": "", "healing_head": "", "healing_id": ""})
        st.rerun()

    include_implants = st.checkbox("Include implants as BOM items")
    include_abutments = st.checkbox("Include abutments (Astra EV only)")
    include_healing = st.checkbox("Include healing abutments (Dentsply EV only)")
    include_prosthesis = st.checkbox("Include prosthesis-level components")

    for index, row in enumerate(st.session_state.bom_rows):
        with st.container(border=True):
            columns = st.columns([2, 1.1, 1, 0.9, 1, 0.5])
            row_brand = columns[0].selectbox("Brand", brand_names, index=brand_names.index(row["brand"]), key=f"brand_{index}")
            implants = options_by_brand[row_brand]
            implant_ids = [item["id"] for item in implants]
            selected_id = row["implant_id"] if row["implant_id"] in implant_ids else implant_ids[0]
            selected_index = implant_ids.index(selected_id)
            selected_id = columns[1].selectbox("Implant", implant_ids, index=selected_index, format_func=lambda value, items=implants: implant_label(next(item for item in items if item["id"] == value)), key=f"implant_{index}")
            implant = next(item for item in implants if item["id"] == selected_id)
            product_line = implant.get("productLine") or ""
            density_disabled = product_line == "PrimeTaper EV"
            density = columns[2].selectbox("Bone", ["standard", "dense"], index=0 if row["density"] == "standard" else 1, format_func=lambda value: "Soft" if value == "standard" else "Dense", disabled=density_disabled, key=f"density_{index}")
            pilot_only = columns[3].checkbox("Pilot only", value=row["pilot_only"], key=f"pilot_{index}")
            quantity = columns[4].number_input("Quantity", min_value=1, step=1, value=max(1, int(row["quantity"])), key=f"quantity_{index}")
            delete_disabled = len(st.session_state.bom_rows) == 1
            if columns[5].button("✕", key=f"remove_{index}", disabled=delete_disabled, help="Remove this implant row"):
                st.session_state.bom_rows.pop(index)
                st.rerun()

            driver = driver_for(catalogs[row_brand], implant)
            include_driver = st.checkbox("Include implant driver", value=row["include_driver"], disabled=driver is None, key=f"driver_{index}", help=driver["description"] if driver else "No matching driver in catalog")

            abutment = None
            if include_abutments:
                available_abutments = [item for item in catalogs[row_brand]["Abutments"] if float(item.get("diameter") or 0) == float(implant.get("diameter") or 0) and item.get("productLine") == product_line]
                angles = sorted({item.get("angle") for item in available_abutments if item.get("angle") is not None})
                if angles:
                    old_abutment = next((item for item in available_abutments if item.get("id") == row.get("abutment_id")), None)
                    angle = st.selectbox("Abutment angle", angles, index=angles.index(old_abutment["angle"]) if old_abutment and old_abutment["angle"] in angles else 0, key=f"abut_angle_{index}")
                    angle_options = [item for item in available_abutments if item.get("angle") == angle]
                    abutment_id = st.selectbox("Abutment", [item["id"] for item in angle_options], format_func=lambda value, choices=angle_options: f"{next(item for item in choices if item['id'] == value)['height']} mm | {next(item for item in choices if item['id'] == value)['partNumber']}", key=f"abutment_{index}")
                    abutment = next(item for item in angle_options if item["id"] == abutment_id)
                else:
                    st.caption("No abutments cataloged for this implant platform.")

            healing_abutment = None
            if include_healing and product_line in ("Astra EV", "PrimeTaper EV"):
                healing_options = [item for item in catalogs[row_brand]["Healing Abutments"] if float(item.get("connectionDiameter") or 0) == float(implant.get("diameter") or 0)]
                heads = sorted({float(item["headDiameter"]) for item in healing_options if item.get("headDiameter") is not None})
                if heads:
                    old_healing = next((item for item in healing_options if item.get("id") == row.get("healing_id")), None)
                    head = st.selectbox("Healing head diameter (mm)", heads, index=heads.index(float(old_healing["headDiameter"])) if old_healing and float(old_healing["headDiameter"]) in heads else 0, key=f"heal_head_{index}")
                    head_options = [item for item in healing_options if float(item["headDiameter"]) == float(head)]
                    healing_id = st.selectbox("Healing abutment", [item["id"] for item in head_options], format_func=lambda value, choices=head_options: f"{next(item for item in choices if item['id'] == value)['height']} mm | {next(item for item in choices if item['id'] == value)['partNumber']}", key=f"healing_{index}")
                    healing_abutment = next(item for item in head_options if item["id"] == healing_id)

            row.update({"brand": row_brand, "implant_id": selected_id, "quantity": quantity, "density": density, "pilot_only": pilot_only, "include_driver": include_driver, "abutment_id": abutment["id"] if abutment else "", "healing_head": healing_abutment.get("headDiameter") if healing_abutment else "", "healing_id": healing_abutment["id"] if healing_abutment else ""})
            row["implant"] = implant
            row["abutment"] = abutment
            row["healing_abutment"] = healing_abutment

resolved_rows = []
for row in st.session_state.bom_rows:
    brand = row["brand"]
    implant = next(item for item in options_by_brand[brand] if item["id"] == row["implant_id"])
    abutment = next((item for item in catalogs[brand]["Abutments"] if item.get("id") == row.get("abutment_id")), None) if include_abutments else None
    healing = next((item for item in catalogs[brand]["Healing Abutments"] if item.get("id") == row.get("healing_id")), None) if include_healing else None
    resolved_rows.append({**row, "implant": implant, "abutment": abutment, "healing_abutment": healing, "density": "standard" if row["pilot_only"] or implant.get("productLine") == "PrimeTaper EV" else row["density"]})

bom = make_bom(catalogs, resolved_rows, include_implants, include_prosthesis, include_abutments, include_healing)
grand_total = sum(item["Line total"] for item in bom)
st.divider()
head_col, total_col = st.columns([3, 1])
head_col.subheader(case_reference or "Generated BOM")
total_col.metric("BOM add-on subtotal", f"${grand_total:,.2f}")

if bom:
    table_rows = []
    for item in bom:
        table_rows.append({
            "BRAND": item["Brand"],
            "QTY": item["Quantity"],
            "SKU": item["Part Number"],
            "ITEM": item["Description"],
            "UDI": st.session_state.udi_values.get((item["Brand"], item["Part Number"]), ""),
            "UNIT COST": f"${item['Unit cost']:,.2f}",
            "LINE SUBTOTAL": f"${item['Line total']:,.2f}",
            "DENSE ONLY": "Yes" if item["Dense bone only"] else "",
            "CORTICAL": "Yes" if item["Cortical"] else "",
            "STATUS": "Placeholder" if item["Placeholder"] else "",
        })
    edited = st.data_editor(pd.DataFrame(table_rows), hide_index=True, width="stretch", disabled=["BRAND", "QTY", "SKU", "ITEM", "UNIT COST", "LINE SUBTOTAL", "DENSE ONLY", "CORTICAL", "STATUS"], column_config={"UDI": st.column_config.TextColumn("UDI", help="Enter or scan the item UDI")})
    st.session_state.udi_values.update({(item["Brand"], item["Part Number"]): udi for item, udi in zip(bom, edited["UDI"].fillna(""))})
    for brand in brand_names:
        brand_bom = [item for item in bom if item["Brand"] == brand]
        if brand_bom:
            safe_case = "_".join((case_reference or "case").split())
            st.download_button(
                f"Download {brand} BOM (Excel)",
                data=excel_bytes(brand_bom, case_reference, st.session_state.udi_values),
                file_name=f"BOM_{brand.replace(' ', '_')}_{safe_case}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key=f"download_{brand}",
            )
else:
    st.info("Select an implant above to generate a BOM.")

with st.expander("Catalog status"):
    st.write("Catalog data is read from the brand-specific Excel workbooks in the `catalogs` folder.")
    st.write("Product protocols, placeholder SKUs, and costs should be verified before operational use.")