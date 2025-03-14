import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import re
import requests
from PIL import Image
from copy import deepcopy

# --- Filstier (tilpas efter behov) ---
MAPPING_FILE_PATH = "mapping-file.xlsx"
STOCK_FILE_PATH = "stock.xlsx"
TEMPLATE_FILE_PATH = "template-generator.pptx"

# --- Forventede kolonner ---
# Mapping-filens originale kolonnenavne (vi normaliserer senere)
REQUIRED_MAPPING_COLS_ORIG = [
    "{{Product name}}",
    "{{Product code}}",
    "{{Product country of origin}}",
    "{{Product height}}",
    "{{Product width}}",
    "{{Product length}}",
    "{{Product depth}}",
    "{{Product seat height}}",
    "{{Product diameter}}",
    "{{CertificateName}}",
    "{{Product Consumption COM}}",
    "{{Product Fact Sheet link}}",
    "{{Product configurator link}}",
    "{{Product Packshot1}}",
    "{{Product Lifestyle1}}",
    "{{Product Lifestyle2}}",
    "{{Product Lifestyle3}}",
    "{{Product Lifestyle4}}",
    "ProductKey"  # Mapping-filens ProductKey (uden klammer)
]

# Stock-filens originale kolonnenavne
REQUIRED_STOCK_COLS_ORIG = [
    "productkey",    # kolonne B: ProductKey
    "variantname",   # kolonne D: VariantName
    "rts",           # kolonne H: RTS
    "mto"            # kolonne I: MTO
]

# --- Placeholders til erstatning i templaten ---
TEXT_PLACEHOLDERS_ORIG = {
    "{{Product name}}": "Product Name:",
    "{{Product code}}": "Product Code:",
    "{{Product country of origin}}": "Country of origin:",
    "{{Product height}}": "Height:",
    "{{Product width}}": "Width:",
    "{{Product length}}": "Length:",
    "{{Product depth}}": "Depth:",
    "{{Product seat height}}": "Seat Height:",
    "{{Product diameter}}": "Diameter:",
    "{{CertificateName}}": "Test & certificates for the product:",
    "{{Product Consumption COM}}": "Consumption information for COM:"
}

HYPERLINK_PLACEHOLDERS_ORIG = {
    "{{Product Fact Sheet link}}": "Download Product Fact Sheet",
    "{{Product configurator link}}": "Click to configure product"
}

IMAGE_PLACEHOLDERS_ORIG = [
    "{{Product Packshot1}}",
    "{{Product Lifestyle1}}",
    "{{Product Lifestyle2}}",
    "{{Product Lifestyle3}}",
    "{{Product Lifestyle4}}",
]

# --- Hjælpefunktioner ---

def normalize_text(s):
    """Fjerner alle mellemrum (inklusiv ikke-brydende) og konverterer til små bogstaver."""
    return re.sub(r"\s+", "", str(s).replace("\u00A0", " ")).lower()

def normalize_col(col):
    """Normaliserer et kolonnenavn ved at anvende normalize_text."""
    return normalize_text(col)

def find_mapping_row(item_no, mapping_df, mapping_prod_key):
    """
    Finder den række i mapping_df, hvor kolonnen for produktkode (mapping_prod_key) matcher 'Item no'.
    Sammenligningen foretages ved at normalisere begge værdier.
    """
    norm_item = normalize_text(item_no)
    for idx, row in mapping_df.iterrows():
        code = row.get(mapping_prod_key, "")
        if normalize_text(code) == norm_item:
            return row
    if "-" in str(item_no):
        partial = normalize_text(item_no.split("-")[0])
        for idx, row in mapping_df.iterrows():
            code = row.get(mapping_prod_key, "")
            if normalize_text(code).startswith(partial):
                return row
    return None

def process_stock_rts_alternative(mapping_row, stock_df):
    """
    Logik for {{Product RTS}}:
      1. Hent 'ProductKey' fra mapping_row.
      2. Filtrer stock_df, så kun rækker, hvor 'productkey' matcher mapping_row's ProductKey, er med.
      3. Filtrer herefter, så kun rækker, hvor kolonnen 'rts' IKKE er tom, medtages.
      4. Udtræk de unikke værdier fra kolonnen 'variantname' i stock, fjern dubletter,
         og sammenkæd dem med ", " (komma og mellemrum).
      5. Returnér den sammenkædte streng.
    """
    product_key = mapping_row.get("productkey", "")
    if not product_key or pd.isna(product_key):
        return ""
    norm_product_key = normalize_text(product_key)
    try:
        filtered = stock_df[stock_df["productkey"].apply(lambda x: normalize_text(x) == norm_product_key)]
    except KeyError as e:
        st.error(f"KeyError i process_stock_rts_alternative (productkey): {e}")
        return ""
    if filtered.empty:
        return ""
    filtered = filtered[filtered["rts"].notna() & (filtered["rts"] != "")]
    if filtered.empty:
        return ""
    try:
        variant_names = filtered["variantname"].dropna().astype(str).tolist()
    except KeyError as e:
        st.error(f"KeyError i process_stock_rts_alternative (variantname): {e}")
        return ""
    unique_variant_names = list(dict.fromkeys(variant_names))
    return ", ".join(unique_variant_names)

def process_stock_mto_alternative(mapping_row, stock_df):
    """
    Logik for {{Product MTO}}:
      1. Hent 'ProductKey' fra mapping_row.
      2. Filtrer stock_df, så kun rækker, hvor 'productkey' matcher mapping_row's ProductKey, er med.
      3. Filtrer herefter, så kun rækker, hvor kolonnen 'mto' IKKE er tom, medtages.
      4. Udtræk de unikke værdier fra kolonnen 'variantname' i stock, fjern dubletter,
         og sammenkæd dem med ", " (komma og mellemrum).
      5. Returnér den sammenkædte streng.
    """
    product_key = mapping_row.get("productkey", "")
    if not product_key or pd.isna(product_key):
        return ""
    norm_product_key = normalize_text(product_key)
    try:
        filtered = stock_df[stock_df["productkey"].apply(lambda x: normalize_text(x) == norm_product_key)]
    except KeyError as e:
        st.error(f"KeyError i process_stock_mto_alternative (productkey): {e}")
        return ""
    if filtered.empty:
        return ""
    filtered = filtered[filtered["mto"].notna() & (filtered["mto"] != "")]
    if filtered.empty:
        return ""
    try:
        variant_names = filtered["variantname"].dropna().astype(str).tolist()
    except KeyError as e:
        st.error(f"KeyError i process_stock_mto_alternative (variantname): {e}")
        return ""
    unique_variant_names = list(dict.fromkeys(variant_names))
    return ", ".join(unique_variant_names)

def fetch_and_process_image(url, quality=70, max_size=(1200, 1200)):
    """
    Henter billede fra en URL. Hvis billedet er i TIFF-format eller har gennemsigtighed (RGBA/LA),
    konverteres det til RGB. Billedet komprimeres ved at sætte en lavere JPEG-kvalitet og begrænse størrelsen.
    Timeout er øget til 30 sekunder.
    Returnerer et BytesIO-objekt med billedet, eller None hvis hentning fejler.
    """
    try:
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            img = Image.open(io.BytesIO(response.content))
            if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info) or (img.format and img.format.lower() == "tiff"):
                img = img.convert("RGB")
            img.thumbnail(max_size, Image.LANCZOS)
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format="JPEG", quality=quality, optimize=True)
            img_byte_arr.seek(0)
            return img_byte_arr
    except Exception as e:
        st.warning(f"Fejl ved hentning af billede fra {url}: {e}")
    return None

def duplicate_slide(prs, slide):
    """Duplicer en slide ved at kopiere dens elementer – svarende til Ctrl+D."""
    slide_layout = slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    new_slide.shapes._spTree.clear()
    for shape in slide.shapes:
        new_slide.shapes._spTree.append(deepcopy(shape._element))
    return new_slide

def replace_text_placeholders(slide, placeholder_values):
    """
    Erstatter tekstplaceholders i en slide ved at kombinere alle løb i hvert afsnit til én streng,
    udfører udskiftningen med regex (som ignorerer ekstra mellemrum) og sætter den samlede tekst tilbage
    i det første run for at bevare den overordnede formatering.
    """
    import re
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                full_text = "".join([run.text for run in paragraph.runs])
                new_text = full_text
                for placeholder, replacement in placeholder_values.items():
                    key = placeholder.strip("{}").strip()
                    pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
                    new_text = re.sub(pattern, replacement, new_text)
                if paragraph.runs:
                    first_run = paragraph.runs[0]
                    for i in range(len(paragraph.runs)-1, -1, -1):
                        paragraph.runs[i].text = ""
                    first_run.text = new_text

def replace_hyperlink_placeholders(slide, hyperlink_values):
    """Erstatter hyperlink-placeholders i en slide med display-tekst og tilhørende URL."""
    import re
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for placeholder, (display_text, url) in hyperlink_values.items():
                        key = placeholder.strip("{}").strip()
                        pattern = r"\{\{\s*" + re.escape(key) + r"\s*\}\}"
                        if re.search(pattern, run.text):
                            run.text = re.sub(pattern, display_text, run.text)
                            try:
                                run.hyperlink.address = url
                            except Exception as e:
                                st.warning(f"Hyperlink for {placeholder} kunne ikke indsættes: {e}")

def replace_image_placeholders(slide, image_values):
    """
    Erstatter billedplaceholders med billeder hentet fra URL'er (komprimeret) i en slide.
    Billedet skaleres, så det bevarer sit aspect ratio.
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            tekst = shape.text
            for ph in IMAGE_PLACEHOLDERS_ORIG:
                norm_ph = normalize_text(ph)
                if norm_ph in normalize_text(tekst):
                    url = image_values.get(ph, "")
                    if url:
                        img_stream = fetch_and_process_image(url)
                        if img_stream:
                            img = Image.open(img_stream)
                            original_width, original_height = img.size
                            target_width = shape.width
                            target_height = shape.height
                            scale = min(target_width / original_width, target_height / original_height)
                            new_width = int(original_width * scale)
                            new_height = int(original_height * scale)
                            new_img_stream = io.BytesIO()
                            img.save(new_img_stream, format="JPEG")
                            new_img_stream.seek(0)
                            slide.shapes.add_picture(new_img_stream, shape.left, shape.top, width=new_width, height=new_height)
                            shape.text = ""
                    break

# --- Main Streamlit App ---
def main():
    st.title("PowerPoint Generator App")
    st.write("Upload din brugerfil (Excel) med kolonnerne 'Item no' og 'Product name'")
    
    # Hvis vi allerede har genereret PPT, så brug session_state
    if "ppt_generated" in st.session_state and st.session_state.ppt_generated:
        st.success("PowerPoint genereret succesfuldt!")
        st.download_button("Download PowerPoint", st.session_state.generated_ppt,
                           file_name="generated_presentation.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        st.progress(100)
        return

    # Upload brugerfil med header (første række bruges som header)
    uploaded_file = st.file_uploader("Upload din bruger Excel-fil", type=["xlsx"])
    if uploaded_file is None:
        st.info("Vent venligst på at uploade brugerfilen.")
        return

    try:
        user_df = pd.read_excel(uploaded_file, header=0)
    except Exception as e:
        st.error(f"Fejl ved læsning af brugerfil: {e}")
        return

    # Valider brugerfilens kolonner
    required_user_cols = {"Item no", "Product name"}
    if not required_user_cols.issubset(set(user_df.columns)):
        st.error(f"Brugerfilen skal indeholde kolonnerne: {required_user_cols}. Fundne kolonner: {list(user_df.columns)}")
        return

    st.write("Brugerfil indlæst succesfuldt!")
    st.info("Validerer filer...")
    progress_bar = st.progress(10)

    # Indlæs og normalisér mapping-filen
    try:
        mapping_df = pd.read_excel(MAPPING_FILE_PATH)
        mapping_df.columns = [normalize_col(col) for col in mapping_df.columns]
    except Exception as e:
        st.error(f"Fejl ved læsning af mapping-fil: {e}")
        return

    normalized_required_mapping_cols = [normalize_col(col) for col in REQUIRED_MAPPING_COLS_ORIG]
    missing_mapping_cols = [req for req in normalized_required_mapping_cols if req not in mapping_df.columns]
    if missing_mapping_cols:
        st.error(f"Mapping-filen mangler følgende kolonner (efter normalisering): {missing_mapping_cols}. Fundne kolonner: {mapping_df.columns.tolist()}")
        return

    st.write("Mapping-fil indlæst succesfuldt!")
    progress_bar.progress(30)
    MAPPING_PRODUCT_CODE_KEY = normalize_col("{{Product code}}")

    # Indlæs og normalisér stock-filen
    try:
        stock_df = pd.read_excel(STOCK_FILE_PATH)
        stock_df.columns = [normalize_col(col) for col in stock_df.columns]
    except Exception as e:
        st.error(f"Fejl ved læsning af stock-fil: {e}")
        return

    normalized_required_stock_cols = [normalize_col(col) for col in REQUIRED_STOCK_COLS_ORIG]
    missing_stock_cols = [req for req in normalized_required_stock_cols if req not in stock_df.columns]
    if missing_stock_cols:
        st.error(f"Stock-filen mangler følgende kolonner (efter normalisering): {missing_stock_cols}. Fundne kolonner: {stock_df.columns.tolist()}")
        return

    st.write("Stock-fil indlæst succesfuldt!")
    progress_bar.progress(50)

    # Indlæs PowerPoint templaten
    try:
        prs = Presentation(TEMPLATE_FILE_PATH)
    except Exception as e:
        st.error(f"Fejl ved læsning af template-fil: {e}")
        return

    if len(prs.slides) < 1:
        st.error("Template-filen skal indeholde mindst én slide.")
        return

    st.write("Template-fil indlæst succesfuldt!")
    progress_bar.progress(70)

    # Brug første slide som template og fjern den fra præsentationen
    template_slide = prs.slides[0]
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    total_products = len(user_df)
    for index, product in user_df.iterrows():
        item_no = product["Item no"]
        slide = duplicate_slide(prs, template_slide)

        mapping_row = find_mapping_row(item_no, mapping_df, MAPPING_PRODUCT_CODE_KEY)
        if mapping_row is None:
            st.warning(f"Ingen match fundet i mapping-fil for Item no: {item_no}")
            continue

        placeholder_texts = {}
        for ph, label in TEXT_PLACEHOLDERS_ORIG.items():
            norm_ph = normalize_col(ph)
            value = mapping_row.get(norm_ph, "")
            if pd.isna(value):
                value = ""
            if ph in ("{{Product code}}", "{{Product name}}", "{{Product country of origin}}"):
                placeholder_texts[ph] = f"{label} {value}"
            else:
                placeholder_texts[ph] = f"{label}\n{value}"

        product_code = mapping_row.get(MAPPING_PRODUCT_CODE_KEY, "")
        rts_text = process_stock_rts_alternative(mapping_row, stock_df)
        mto_text = process_stock_mto_alternative(mapping_row, stock_df)
        placeholder_texts["{{Product RTS}}"] = f"Product in stock versions:\n{rts_text}"
        placeholder_texts["{{Product MTO}}"] = f"Avilable for made to order:\n{mto_text}"

        replace_text_placeholders(slide, placeholder_texts)

        hyperlink_vals = {}
        for ph, display_text in HYPERLINK_PLACEHOLDERS_ORIG.items():
            norm_ph = normalize_col(ph)
            url = mapping_row.get(norm_ph, "")
            if pd.isna(url):
                url = ""
            hyperlink_vals[ph] = (display_text, url)
        replace_hyperlink_placeholders(slide, hyperlink_vals)

        image_vals = {}
        for ph in IMAGE_PLACEHOLDERS_ORIG:
            norm_ph = normalize_col(ph)
            url = mapping_row.get(norm_ph, "")
            if pd.isna(url):
                url = ""
            image_vals[ph] = url
        replace_image_placeholders(slide, image_vals)

        progress_value = 70 + min(int((index + 1) / total_products * 30), 30)
        progress_bar.progress(progress_value)

    ppt_io = io.BytesIO()
    try:
        prs.save(ppt_io)
        ppt_io.seek(0)
    except Exception as e:
        st.error(f"Fejl ved gemning af PowerPoint: {e}")
        return

    st.success("PowerPoint genereret succesfuldt!")
    st.download_button("Download PowerPoint", ppt_io,
                       file_name="generated_presentation.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    
    # Gem den genererede PPT i session_state, så progress baren forbliver komplet
    st.session_state.ppt_generated = True
    st.session_state.generated_ppt = ppt_io

if __name__ == '__main__':
    if 'ppt_generated' not in st.session_state:
        st.session_state.ppt_generated = False
    main()
