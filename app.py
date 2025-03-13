import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import re
import requests
from PIL import Image
from copy import deepcopy

# Filstier – juster efter behov (fx "data/mapping-file.xlsx" osv.)
MAPPING_FILE_PATH = "mapping-file.xlsx"
STOCK_FILE_PATH = "stock.xlsx"
TEMPLATE_FILE_PATH = "template-generator.pptx"

# --- Forventede kolonner ---
# Vi definerer de forventede kolonnenavne, som vi efter normalisering skal have i mapping- og stock-filerne.

# Mapping-filens krævede kolonner (brug originalt format – vi normaliserer efterfølgende)
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
    "{{Product Lifestyle4}}"
]

# Stock-filens krævede kolonner (brug originalt format – vi normaliserer efterfølgende)
REQUIRED_STOCK_COLS_ORIG = [
    "{{productcode}}",  # vi antager, at stock-filen har dette i små bogstaver
    "variantfamily",
    "variantcommercialname",
    "rts",
    "mto"
]

# Vi definerer konstanter til at tilgå stock-data
STOCK_CODE_COL = "{{productcode}}"   # Dette skal normaliseres til det samme format som i stock_df
STOCK_GROUP_COL = "variantfamily"
STOCK_VALUE_COL = "variantcommercialname"
STOCK_RTS_FILTER_COL = "rts"
STOCK_MTO_FILTER_COL = "mto"

# --- Placeholders til erstatning i templaten ---
# Tekstfelter: mapping fra placeholder til foruddefineret label
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

# Hyperlink felter: nøglen er placeholder, og værdien er display-tekst
HYPERLINK_PLACEHOLDERS_ORIG = {
    "{{Product Fact Sheet link}}": "Download Product Fact Sheet",
    "{{Product configurator link}}": "Click to configure product"
}

# Billed placeholders: der indsættes billeder fra URL
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
    """Normaliserer et kolonnenavn: fjerner mellemrum (inklusiv ikke-brydende) og konverterer til små bogstaver."""
    return normalize_text(col)

# Efter indlæsning vil vi erstatte kolonnenavnene med deres normaliserede versioner.

def find_mapping_row(item_no, mapping_df, mapping_prod_key):
    """
    Finder den række i mapping_df, hvor kolonnen for produktkode (mapping_prod_key) matcher item_no.
    Prøver først et eksakt match; hvis ikke og item_no indeholder '-', matches delstrengen før '-'.
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

def process_stock_rts(stock_df, product_code):
    """Behandler RTS-data: Filtrér stock_df for matchende produktkode og ikke-tomme RTS-celler,
    gruppering på 'variantfamily' og samler værdier fra 'variantcommercialname' med linjeskift."""
    norm_code = normalize_text(product_code)
    try:
        filtered = stock_df[stock_df[STOCK_CODE_COL].apply(lambda x: normalize_text(x) == norm_code)]
    except KeyError as e:
        st.error(f"KeyError i process_stock_rts: {e}")
        return ""
    if filtered.empty:
        return ""
    filtered = filtered[filtered[STOCK_RTS_FILTER_COL].notna() & (filtered[STOCK_RTS_FILTER_COL] != "")]
    if filtered.empty:
        return ""
    result_lines = []
    grouped = filtered.groupby(STOCK_GROUP_COL)
    for group_name, group in grouped:
        values = group[STOCK_VALUE_COL].dropna().astype(str).tolist()
        if values:
            group_text = f"{group_name}:"
            values_text = "\n".join(values)
            result_lines.append(f"{group_text}\n{values_text}")
    return "\n".join(result_lines)

def process_stock_mto(stock_df, product_code):
    """Behandler MTO-data: Filtrér stock_df for matchende produktkode og ikke-tomme MTO-celler,
    gruppering på 'variantfamily' og samler værdier fra 'variantcommercialname' med komma og mellemrum."""
    norm_code = normalize_text(product_code)
    try:
        filtered = stock_df[stock_df[STOCK_CODE_COL].apply(lambda x: normalize_text(x) == norm_code)]
    except KeyError as e:
        st.error(f"KeyError i process_stock_mto: {e}")
        return ""
    if filtered.empty:
        return ""
    filtered = filtered[filtered[STOCK_MTO_FILTER_COL].notna() & (filtered[STOCK_MTO_FILTER_COL] != "")]
    if filtered.empty:
        return ""
    result_lines = []
    grouped = filtered.groupby(STOCK_GROUP_COL)
    for group_name, group in grouped:
        values = group[STOCK_VALUE_COL].dropna().astype(str).tolist()
        if values:
            group_text = f"{group_name}:"
            values_text = ", ".join(values)
            result_lines.append(f"{group_text}\n{values_text}")
    return "\n".join(result_lines)

def fetch_and_process_image(url, quality=70, max_size=(1200, 1200)):
    """
    Henter billede fra en URL. Hvis billedet er i TIFF-format, konverteres det til JPEG.
    Billedet komprimeres ved at sætte en lavere JPEG-kvalitet og begrænse størrelsen.
    Returnerer et BytesIO-objekt med billedet.
    """
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            img = Image.open(io.BytesIO(response.content))
            if img.format.lower() == "tiff":
                img = img.convert("RGB")
            img.thumbnail(max_size, Image.LANCZOS)  # Ændret fra ANTIALIAS til LANCZOS
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format="JPEG", quality=quality, optimize=True)
            img_byte_arr.seek(0)
            return img_byte_arr
    except Exception as e:
        st.error(f"Fejl ved hentning af billede fra {url}: {e}")
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
    """Erstatter tekstplaceholders i en slide med de leverede værdier."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            tekst = shape.text
            for placeholder, ny_tekst in placeholder_values.items():
                if placeholder in tekst:
                    shape.text = ny_tekst

def replace_hyperlink_placeholders(slide, hyperlink_values):
    """Erstatter hyperlink-placeholders i en slide med display-tekst og tilhørende URL."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    for placeholder, (display_text, url) in hyperlink_values.items():
                        if placeholder in run.text:
                            run.text = display_text
                            try:
                                run.hyperlink.address = url
                            except Exception as e:
                                st.warning(f"Hyperlink for {placeholder} kunne ikke indsættes: {e}")

def replace_image_placeholders(slide, image_values):
    """Erstatter billedplaceholders med billeder hentet fra URL'er (komprimeret) i en slide."""
    for shape in slide.shapes:
        if shape.has_text_frame:
            tekst = shape.text
            for ph in IMAGE_PLACEHOLDERS_ORIG:
                # Vi normaliserer ph for at sikre, at match sker uafhængigt af mellemrum
                norm_ph = normalize_text(ph)
                if norm_ph in normalize_text(tekst):
                    url = image_values.get(ph, "")
                    if url:
                        img_stream = fetch_and_process_image(url)
                        if img_stream:
                            left = shape.left
                            top = shape.top
                            width = shape.width
                            height = shape.height
                            slide.shapes.add_picture(img_stream, left, top, width=width, height=height)
                            shape.text = ""  # Fjern placeholder-teksten
                    break

# --- Main Streamlit App ---

def main():
    st.title("PowerPoint Generator App")
    st.write("Upload din brugerfil (Excel) med kolonnerne 'Item no' og 'Product name'")
    
    # Upload brugerfil
    uploaded_file = st.file_uploader("Upload din bruger Excel-fil", type=["xlsx"])
    if uploaded_file is None:
        st.info("Vent venligst på at uploade brugerfilen.")
        return

    try:
        user_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Fejl ved læsning af brugerfil: {e}")
        return

    # Valider brugerfilens kolonner
    required_user_cols = {"Item no", "Product name"}
    if not required_user_cols.issubset(set(user_df.columns)):
        st.error(f"Brugerfilen skal indeholde kolonnerne: {required_user_cols}. Fundne kolonner: {list(user_df.columns)}")
        return

    st.write("Brugerfil indlæst succesfuldt!")

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

    # Definér den normaliserede nøgle for mapping-filens produktkode
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

    # Indlæs PowerPoint templaten
    try:
        prs = Presentation(TEMPLATE_FILE_PATH)
    except Exception as e:
        st.error(f"Fejl ved læsning af template-fil: {e}")
        return

    if len(prs.slides) < 1:
        st.error("Template-filen skal indeholde mindst én slide.")
        return

    # Brug første slide som template og fjern den fra præsentationen
    template_slide = prs.slides[0]
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])

    # For hvert produkt i brugerfilen
    for index, product in user_df.iterrows():
        item_no = product["Item no"]
        slide = duplicate_slide(prs, template_slide)

        # Find match i mapping-filen baseret på den normaliserede produktkode
        mapping_row = find_mapping_row(item_no, mapping_df, MAPPING_PRODUCT_CODE_KEY)
        if mapping_row is None:
            st.warning(f"Ingen match fundet i mapping-fil for Item no: {item_no}")
            continue

        # Opret dictionary for tekst placeholders – brug de originale nøgler, da de forventes i templaten
        placeholder_texts = {}
        for ph, label in TEXT_PLACEHOLDERS_ORIG.items():
            norm_ph = normalize_col(ph)
            value = mapping_row.get(norm_ph, "")
            if pd.isna(value):
                value = ""
            placeholder_texts[ph] = f"{label}\n{value}"

        # Hent produktkode fra mapping_row
        product_code = mapping_row.get(MAPPING_PRODUCT_CODE_KEY, "")
        # Behandl stock-data for RTS og MTO
        rts_text = process_stock_rts(stock_df, product_code)
        mto_text = process_stock_mto(stock_df, product_code)
        placeholder_texts["{{Product RTS}}"] = f"Product in stock versions:\n{rts_text}"
        placeholder_texts["{{Product MTO}}"] = f"Avilable for made to order:\n{mto_text}"

        # Erstat tekst, hyperlinks og billeder i sliden
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

    # Gem den genererede præsentation i en BytesIO-strøm og gør den klar til download
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

if __name__ == '__main__':
    main()
