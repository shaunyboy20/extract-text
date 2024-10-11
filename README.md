# extract-text

This project provides resources for extracting text from files. Supported file types are XLS, XLSX, CSV, DOCX, HTM, HTML, TIF, JPG, PNG, JPEG, TIFF, TXT, and PDF.

OCR is used to extract text from PDF files with no text data and image files. pytesseract must be installed for OCR operations, and the location of the executable must be set as pytesseract.pytesseract.tesseract_cmd


## Example use

Get text from "test.pdf"
```python
extract_text("test.pdf")

# OR

stream = get_stream("test.pdf")
extract_text(stream)
```

Force OCR on "test.pdf"
```python
# default dpi is 300
extract_text("test.pdf", dpi=300, force_ocr=True)
```

Get text from each file in "test.zip"
```python
streams = get_stream("test.zip")
for s in streams:
  extract_text(s)
```

Get text from "test.png" (uses OCR)
```python
extract_text("test.png")
```

Get text from "test" which is an extensionless PDF file
```python
extract_text("test", ext="PDF")
```
