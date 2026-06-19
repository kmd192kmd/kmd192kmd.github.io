libs = ['pypdf', 'PyPDF2', 'pdfplumber', 'fitz', 'reportlab', 'pptx']
for lib in libs:
    try:
        __import__(lib)
        print(f"Library {lib} is INSTALLED")
    except ImportError:
        print(f"Library {lib} is NOT installed")
