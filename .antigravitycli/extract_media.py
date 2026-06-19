import os
import zipfile

pptx_path = r"C:\Users\a\Desktop\쇼핑몰 프로젝트 자료모음\6_자소서, 포폴 1차 완성본_260525\포트폴리오(김민도).pptx"
extract_dir = r"C:\Users\a\Desktop\coding\forGithub\kmd192kmd.github.io\images\extracted"

os.makedirs(extract_dir, exist_ok=True)

with zipfile.ZipFile(pptx_path, 'r') as z:
    for file in z.namelist():
        if file.startswith("ppt/media/"):
            basename = os.path.basename(file)
            if basename:
                out_path = os.path.join(extract_dir, basename)
                with open(out_path, 'wb') as out_f:
                    out_f.write(z.read(file))
                print(f"Extracted: {basename} to {out_path}")

print("Media extraction complete!")
