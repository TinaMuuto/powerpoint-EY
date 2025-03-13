import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import re
import requests
from PIL import Image
from copy import deepcopy

# Filstier til de kendte filer – juster eventuelt stierne hvis nødvendigt
MAPPING_FILE_PATH = "mapping-file.xlsx"
STOCK_FILE_PATH = "stock.xlsx"
TEMPLATE_FILE_PATH = "template-generator.pptx"

# Konstanter for stock-filens kolonner – tilpas disse hvis dine headers er anderledes
STOCK_CODE_COL = "Product code"     # Skal matche stock-filen
STOCK_GROUP_COL = "Group"           
STOCK_VALUE_COL = "Value"           
STOCK_RTS_FILTER_COL = "RTS"        
STOCK_MTO_FILTER_COL = "MTO"        

# Definer placeholder mappings for tekstfelter – nøglerne skal matche de kolonneoverskrifter, der findes i mapping-filen og template
TEXT_PLACEHOLDERS = {
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
    "{{Product Consumption COM}}": "Consumption information for COM:",
}

# Hyperlink placeholders med foruddefineret linktekst
HYPERLINK_PLACEHOLDERS = {
    "{{Product Fact Sheet link}}": "Download Product Fact Sheet",
    "{{Product configurator link}}": "Click to configure product"
}

# Billed placeholders – de her skal erstattes med et billede fra URL
IMAGE_PLACEHOLDERS = [
    "{{Product Packshot1}}",
    "{{Product Lifestyle1}}",
    "{{Product Lifestyle2}}",
    "{{Product Lifestyle3}}",
    "{{Product Lifestyle4}}",
]

def duplicate_slide(prs, slide):
    """
    Dupliker en slide i et Presentation-objekt.
    Bemærk: python-pptx understøtter ikke native slide-duplication, så vi kopierer slide-indholdet.
    """
    slide_layout = slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)
    # Ryd de eksisterende shapes (en ny slide kommer normalt med standardobjekter)
    new_slide.shapes._spTree.clear()
    # Kopier alle shapes fra template-sliden til den nye slide
    for shape in slide.shapes:
        new_slide.shapes._spTree.append(deepcopy(shape._element))
    return new_slide

def normalize_text(s):
    """Fjerner mellemrum og konverterer til små bogstaver for at lette sammenligninger."""
    return re.sub(r"\s+", "", str(s)).lower()

def find_mapping_row(item_no, mapping_df):
    """
    Find den række i mapping-filen, hvor kolonnen '{{Product code}}' matcher item_no.
    Der tages højde for case, mellemrum og evt. delvist match, hvis der findes en '-' i item_no.
    """
    norm_item = normalize_text(item_no)
    # Forsøg først et eksakt match (brug hele nøgle-strengen)
    for idx, row in mapping_df.iterrows():
        code = row.get("{{Product code}}", "")
        if normalize_text(code) == norm_item:
            return row
    # Hvis ikke eksakt match, og hvis item_no indeholder '-', så match på delstreng før '-'
    if "-" in str(item_no):
        partial = normalize_text(item_no.split("-")[0])
        for idx, row in mapping_df.iterrows():
            code = row.get("{{Product code}}", "")
            if normalize_text(code).startswith(partial):
                return row
    return None

def process_stock_rts(stock_df, product_code):
    """
    Henter og grupperer data til feltet {{Product RTS}} fra stock-filen.
    Filtrerer stock_df for matchende produktkode og ikke-tomme celler i RTS-kolonnen.
    Gruppér derefter på 'Group' og returnér en streng med gruppenavn og tilhørende værdier (fra 'Value').
    """
    norm_code = normalize_text(product_code)
    filtered = stock_df[stock_df[STOCK_CODE_COL].apply(lambda x: normalize_text(x) == norm_code)]
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
    """
    Henter og grupperer data til feltet {{Product MTO}} fra stock-filen.
    Filtrerer stock_df for matchende produktkode og ikke-tomme celler i MTO-kolonnen.
    Gruppér på 'Group' og returnér en streng med gruppenavn og tilhørende værdier (fra 'Value') sammenkædet med komma og mellemrum.
    """
    norm_code = normalize_text(product_code)
    filtered = stock_df[stock_df[STOCK_CODE_COL].apply(lambda x: normalize_text(x) == norm_code)]
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
    Henter et billede fra en URL. Hvis billedet er TIFF, konverteres det til JPEG.
    Billedet komprimeres ved at sætte en lavere JPEG-kvalitet og begrænse størrelsen.
    Returnerer et BytesIO-objekt med billedet.
    """
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            img = Image.open(io.BytesIO(response.content))
            if img.format.lower() == "tiff":
                img = img.convert("RGB")
            img.thumbnail(max_size, Image.ANTIALIAS)
            img_byte_arr = io.BytesIO()
            img.save(img_byte_arr, format="JPEG", quality=quality, optimize=True)
            img_byte_arr.seek(0)
            return img_byte_arr
    except Exception as e:
        st.error(f"Fejl ved hentning af billede fra {url}: {e}")
    return None

def replace_text_placeholders(slide, placeholder_values):
    """
    Gennemløber alle shapes i en slide og erstatter tekst, hvis placeholderen findes.
    placeholder_values: dictionary med nøglen (f.eks. "{{Product name}}") og værdi, der skal indsættes (inkl. label og ny linje).
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            tekst = shape.text
            for placeholder, ny_tekst in placeholder_values.items():
                if placeholder in tekst:
                    shape.text = ny_tekst

def replace_hyperlink_placeholders(slide, hyperlink_values):
    """
    Erstat hyperlink placeholders.
    hyperlink_values: dictionary med nøglen (f.eks. "{{Product Fact Sheet link}}") og værdi som et tuple (display_text, url).
    """
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
                                st.warning(f"Hyperlink kunne ikke indsættes for {placeholder}: {e}")

def replace_image_placeholders(slide, image_values):
    """
    Erstat billedplaceholders med billeder fra URL.
    image_values: dictionary med nøgle (f.eks. "{{Product Packshot1}}") og værdi som billede-URL.
    For hvert match hentes billedet, komprimeres og indsættes i den angivne shape.
    """
    for shape in slide.shapes:
        if shape.has_text_frame:
            tekst = shape.text
            for ph in IMAGE_PLACEHOLDERS:
                if ph in tekst:
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

def main():
    st.title("PowerPoint Generator App")
    st.write("Upload din brugerfil (Excel) med kolonnerne 'Item no' og 'Product name'")
    
    uploaded_file = st.file_uploader("Upload din bruger Excel-fil", type=["xlsx"])
    
    if uploaded_file is not None:
        try:
            user_df = pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Fejl ved læsning af brugerfil: {e}")
            return
        
        # Tjek om de nødvendige kolonner findes
        if not {"Item no", "Product name"}.issubset(user_df.columns):
            st.error("Filen skal indeholde kolonnerne 'Item no' og 'Product name'")
            return
        
        st.write("Brugerfil indlæst succesfuldt!")
        
        # Indlæs mapping-fil
        try:
            mapping_df = pd.read_excel(MAPPING_FILE_PATH)
        except Exception as e:
            st.error(f"Fejl ved læsning af mapping-fil: {e}")
            return
        
        # Indlæs stock-fil
        try:
            stock_df = pd.read_excel(STOCK_FILE_PATH)
        except Exception as e:
            st.error(f"Fejl ved læsning af stock-fil: {e}")
            return
        
        # Indlæs template PowerPoint
        try:
            prs = Presentation(TEMPLATE_FILE_PATH)
        except Exception as e:
            st.error(f"Fejl ved læsning af template-fil: {e}")
            return
        
        if len(prs.slides) < 1:
            st.error("Template-filen skal indeholde mindst én slide")
            return
        
        # Brug den første slide som template
        template_slide = prs.slides[0]
        # Fjern den oprindelige slide, så den ikke medtages i den endelige præsentation
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])
        
        # For hvert produkt (hver række i brugerfilen undtagen header) laves der én slide
        for index, product in user_df.iterrows():
            item_no = product["Item no"]
            slide = duplicate_slide(prs, template_slide)
            
            # Find match i mapping-filen baseret på Item no (bruger mapping-kolonnen med krølleparenteser)
            mapping_row = find_mapping_row(item_no, mapping_df)
            if mapping_row is None:
                st.warning(f"Ingen match fundet i mapping-fil for Item no: {item_no}")
                continue
            
            # Hent tekstværdier direkte med placeholder-nøglerne
            placeholder_texts = {}
            for ph, label in TEXT_PLACEHOLDERS.items():
                value = mapping_row.get(ph, "")
                if pd.isna(value):
                    value = ""
                placeholder_texts[ph] = f"{label}\n{value}"
            
            # Hent produktkoden fra mapping-filen (brug den nøjagtige nøgle)
            product_code = mapping_row.get("{{Product code}}", "")
            rts_text = process_stock_rts(stock_df, product_code)
            mto_text = process_stock_mto(stock_df, product_code)
            placeholder_texts["{{Product RTS}}"] = f"Product in stock versions:\n{rts_text}"
            placeholder_texts["{{Product MTO}}"] = f"Avilable for made to order:\n{mto_text}"
            
            # Erstat tekst placeholders i sliden
            replace_text_placeholders(slide, placeholder_texts)
            
            # Hyperlink felter – hent URL fra mapping-filen med de fulde nøgler
            hyperlink_vals = {}
            for ph, display_text in HYPERLINK_PLACEHOLDERS.items():
                url = mapping_row.get(ph, "")
                if pd.isna(url):
                    url = ""
                hyperlink_vals[ph] = (display_text, url)
            replace_hyperlink_placeholders(slide, hyperlink_vals)
            
            # Billedfelter – hent URL og indsæt billedet i den angivne shape
            image_vals = {}
            for ph in IMAGE_PLACEHOLDERS:
                url = mapping_row.get(ph, "")
                if pd.isna(url):
                    url = ""
                image_vals[ph] = url
            replace_image_placeholders(slide, image_vals)
        
        # Gem den genererede præsentation til en BytesIO-strøm
        ppt_io = io.BytesIO()
        try:
            prs.save(ppt_io)
            ppt_io.seek(0)
        except Exception as e:
            st.error(f"Fejl ved gemning af PowerPoint: {e}")
            return
        
        st.success("PowerPoint genereret succesfuldt!")
        st.download_button("Download PowerPoint", ppt_io, file_name="generated_presentation.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == '__main__':
    main()
