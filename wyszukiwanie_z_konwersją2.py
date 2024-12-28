import fitz  # PyMuPDF
import os
import re
import pandas as pd

#zmiany
#kolejne zmiany w gałęzi poprawki

def pdf_to_images_and_html_with_responsive_anchors(pdf_path, pattern, output_html_dir):
    pdf_document = fitz.open(pdf_path)
    output_html = os.path.join(output_html_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}.html")
    image_dir = os.path.join(output_html_dir, "images")
    os.makedirs(image_dir, exist_ok=True)  # Tworzymy folder na obrazy
    
    # CSS i JavaScript dla przewijania
    html_content = """<html><head>
    <title>PDF to HTML</title>
    <style>
        .page-container {
            position: relative;
            width: fit-content;
            margin-bottom: 20px;
        }
        .page-image {
            display: block;
            width: 100%;
        }
        .anchor {
            position: absolute;
            background-color: rgba(255, 255, 0, 0.5);
            border: 1px solid red;
            font-size: 12px;
            pointer-events: none;
        }
    </style>
    <script>

    function applyZoomAndScroll() {
        const hash = window.location.hash.substring(1); // Pobierz fragment po #
        const urlParams = new URLSearchParams(hash);
        const anchor = urlParams.get('anchor'); // Pobierz wartość "anchor"
        const zoom = parseFloat(urlParams.get('zoom')) || 1; // Pobierz wartość "zoom" (domyślnie 1x)

        // Ustaw skalowanie całej strony
        document.body.style.transform = `scale(${zoom})`;
        document.body.style.transformOrigin = "top left";

        // Przelicz pozycje kotwic
        const containers = document.querySelectorAll('.page-container');
        containers.forEach(container => {
            const anchors = container.querySelectorAll('.anchor');

            anchors.forEach(anchorElem => {
                const left = parseFloat(anchorElem.getAttribute('data-original-left'));
                const top = parseFloat(anchorElem.getAttribute('data-original-top'));
                const width = parseFloat(anchorElem.getAttribute('data-original-width'));
                const height = parseFloat(anchorElem.getAttribute('data-original-height'));

                anchorElem.style.left = `${left}px`;
                anchorElem.style.top = `${top}px`;
                anchorElem.style.width = `${width}px`;
                anchorElem.style.height = `${height}px`;
            });
        });

        // Przewiń do kotwicy i ustaw ją w centrum
        setTimeout(() => {
            if (anchor) {
                const targetAnchor = document.getElementById(anchor);
                if (targetAnchor) {
                    const rect = targetAnchor.getBoundingClientRect();
                    const offsetX = window.pageXOffset + rect.left - (window.innerWidth / 2) + (rect.width / 2);
                    const offsetY = window.pageYOffset + rect.top - (window.innerHeight / 2) + (rect.height / 2);
                    window.scrollTo({ left: offsetX, top: offsetY, behavior: 'smooth' });
                }
            }
        }, 100); // Opóźnienie 100 ms na przeliczenie
    }

    // Wykonaj funkcję po załadowaniu całej strony
    window.onload = applyZoomAndScroll;

    </script>
    </head><body>\n"""
    
    anchor_counter = 1  # Licznik kotwic
    matches = []  # Lista znalezionych tekstów
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        
        # Generowanie obrazu strony PDF
        pix = page.get_pixmap(dpi=150)  # Lepsza jakość (150 DPI)
        image_width, image_height = pix.width, pix.height
        image_path = os.path.join(image_dir, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page{page_num + 1}.png")
        pix.save(image_path)
        
        # Dodanie kontenera strony do HTML
        html_content += f'<div class="page-container" style="width: {image_width}px; height: {image_height}px;">\n'
        html_content += f'<img src="{os.path.relpath(image_path, output_html_dir)}" alt="Strona {page_num + 1}" class="page-image">\n'
        
        # Przetwarzanie tekstu i dodanie kotwic z pozycjonowaniem proporcjonalnym
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        text = span["text"]
                        if re.search(pattern, text):
                            # Współrzędne tekstu w jednostkach obrazu
                            bbox = span["bbox"]  # (x0, y0, x1, y1) w jednostkach PDF
                            left = (bbox[0] / page.rect.width) * 100
                            top = (bbox[1] / page.rect.height) * 100
                            width = ((bbox[2] - bbox[0]) / page.rect.width) * 100
                            height = ((bbox[3] - bbox[1]) / page.rect.height) * 100
                            
                            # Dodaj kotwicę z ID
                            anchor_name = f"tekst{anchor_counter}"
                            html_content += f'<a id="{anchor_name}" class="anchor" style="left: {left}%; top: {top}%; width: {width}%; height: {height}%;"></a>\n'
                            matches.append({
                                "Tekst": text,
                                "Strona": page_num + 1,
                                "Hiperłącze": f"http://localhost:8000/{os.path.basename(output_html)}#anchor=tekst{anchor_counter}"
                            })
                            anchor_counter += 1
        
        html_content += "</div>\n"
    
    html_content += "</body></html>"
    pdf_document.close()
    
    # Zapisz plik HTML
    with open(output_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"HTML zapisany jako: {output_html}")
    return matches

def save_to_excel(data, output_excel):
    df = pd.DataFrame(data)
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Znalezione Teksty", index=False)
        workbook = writer.book
        sheet = writer.sheets["Znalezione Teksty"]
        
        # Dodanie hiperłączy do kolumny "Hiperłącze"
        for row_num, link in enumerate(df["Hiperłącze"], start=2):
            sheet.cell(row=row_num, column=3).hyperlink = link  # Kolumna Hiperłącze
    
    print(f"Dane zapisano do pliku Excel: {output_excel}")

def process_pdfs_and_generate_html_and_excel(pdf_dir, pattern, output_dir, excel_output_path):
    all_matches = []
    for file_name in os.listdir(pdf_dir):
        if file_name.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_dir, file_name)
            matches = pdf_to_images_and_html_with_responsive_anchors(pdf_path, pattern, output_dir)
            all_matches.extend(matches)
    
    # Zapisz wszystkie wyniki do pliku Excel
    save_to_excel(all_matches, excel_output_path)

# Ścieżki
pdf_dir = r"C:\lech_dane\python\wszystkie"  # Katalog z PDF
output_dir = r"C:\lech_dane\python\wszystkie\output"  # Katalog wyjściowy na HTML i obrazy
excel_output_path = os.path.join(pdf_dir, "znalezione_teksty.xlsx")
os.makedirs(output_dir, exist_ok=True)

# Wzorzec do wyszukiwania
pattern = r'^(([A-Z]|PU)\d{5}(?!VM|SB|ST|BV|RO|VOD)[A-Z]{2,3}\d{3})|^(\d{3}(?!VM|SB|ST|BV|RO)[A-Z]{2}\d{3})'

# Przetwarzanie
process_pdfs_and_generate_html_and_excel(pdf_dir, pattern, output_dir, excel_output_path)
