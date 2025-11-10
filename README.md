# 🖼️ PowerPoint Generator til præsentationer til EY

Dette projekt er en **Streamlit-applikation** designet til automatisk at generere en PowerPoint-præsentation baseret på en liste af varenumre, som beriges med data fra et **Mapping File** og et **Stock File**. Appen udfylder pladsholdere (tekst, links og billeder) i en skabelon-slide, og duplikerer sliden for hvert unikt varenummer.

Det er en app der er lavet til sales operations
---

## 🔗 Adgang til Appen

Applikationen er hostet via Streamlit Sharing og kan tilgås direkte her:

* **App Link:** https://power-point.streamlit.app/

---

## 🚀 Funktioner og Formål

Hovedformålet er at automatisere oprettelsen af datablad-lignende slides til et sæt produkter:

1.  **Input:** Brugeren indsætter en liste af varenumre via et tekstfelt i Streamlit-grænsefladen.
2.  **Dataopslag:** For hvert varenummer slås relevante produktinformationer op i det lokale **Mapping File**.
3.  **Lagerdata:** Appen henter og grupperer lagerinformation (**RTS** / **MTO**) fra det lokale **Stock File** baseret på `ProductKey`.
4.  **Slide Generation:** Den første slide i skabelonen bruges til det første produkt, og derefter **duplikeres** skabelon-sliden for hvert efterfølgende produkt.
5.  **Indholdserstatning:** Tekstpladsholdere, hyperlinks og billedpladsholdere erstattes med data. Billeder hentes fra URL'er, komprimeres og skaleres, før de indsættes.

---

## 🛠️ Opsætning og Nødvendige Filer

For at køre eller vedligeholde appen lokalt, skal følgende filer være til stede i rodmappen:

### Statiske Filer

| Filnavn | Formål | Vigtige Kolonner/Krav |
| :--- | :--- | :--- |
| `**mapping-file.xlsx**` | Primær datakilde for produktinfo, links og billed-URL'er. | Skal indeholde: `{{Product code}}` og `ProductKey`. |
| `**stock.xlsx**` | Data for lagerstatus. | Skal indeholde: `productkey`, `variantname`, `rts`, `mto`. |
| `**template-generator.pptx**` | PowerPoint-skabelon. | Skal indeholde mindst **én slide** med de relevante pladsholdere. |

### Python-Biblioteker

Installer de nødvendige afhængigheder for lokal kørsel:

```bash
pip install streamlit pandas openpyxl python-pptx requests Pillow
````

### Kørsel (Lokal)

Start appen ved hjælp af Streamlit:

```bash
streamlit run <dit_scriptnavn>.py
```

-----

## ⚙️ Logik for Datamatching og Erstatning

### A. Opslagslogik (`find_mapping_row`)

Brugerens varenummer (`Item no`) matches mod `mapping-file.xlsx` efter denne prioritet:

1.  **Eksakt match** mod den normaliserede værdi i kolonnen `{{Product code}}`.
2.  **Partielt match** (kun basis-SKU, f.eks. før bindestreg) hvis eksakt match fejler.

### B. Lagerdata Gruppering (RTS/MTO)

Lagerinformationen grupperes baseret på `ProductKey` og Variantnavn:

  * Variantnavne filtreres på ikke-tomme **`rts`** eller **`mto`** værdier.
  * Varianterne grupperes efter **præfiks** (delen før `" - "`), hvor suffikserne samles på samme linje.

### C. Pladsholdere

Skabelonen skal indeholde pladsholdere i følgende formater:

| Type | Pladsholder Eksempel | Data Kilde |
| :--- | :--- | :--- |
| **Tekst** | `{{Product name}}`, `{{Product height}}` | `mapping-file.xlsx` |
| **Lagerstatus** | `{{Product RTS}}`, `{{Product MTO}}` | `stock.xlsx` (Grupperet data) |
| **Hyperlink** | `{{Product Fact Sheet link}}` | `mapping-file.xlsx` (URL) |
| **Billede** | `{{Product Packshot1}}`, `{{Product Lifestyle1}}` | `mapping-file.xlsx` (URL) |

### D. Billedbehandling

Billeder hentes via URL'er. Funktionen `fetch_and_process_image` optimerer billederne for PowerPoint-præsentationens filstørrelse og ydeevne:

  * Billedet hentes.
  * Konverteres til **RGB** (hvis nødvendigt).
  * **Resizes** (`thumbnail`) til maksimalt **1200x1200** pixels.
  * Gemmes som **JPEG** med `quality=70` for komprimering.

<!-- end list -->

```
```
