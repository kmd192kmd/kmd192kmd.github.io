import pypdf

pdf_path = r"C:\Users\a\Desktop\쇼핑몰 프로젝트 자료모음\6_자소서, 포폴 1차 완성본_260525\포트폴리오(박지명).pdf"
reader = pypdf.PdfReader(pdf_path)

output_file = r"C:\Users\a\Desktop\coding\forGithub\kmd192kmd.github.io\.antigravitycli\pdf_park_text.txt"

with open(output_file, 'w', encoding='utf-8') as f:
    f.write(f"Number of pages: {len(reader.pages)}\n\n")
    for idx, page in enumerate(reader.pages):
        f.write(f"=== Page {idx + 1} ===\n")
        f.write(page.extract_text() or "")
        f.write("\n\n")

print("Inspection output written to pdf_park_text.txt successfully!")
