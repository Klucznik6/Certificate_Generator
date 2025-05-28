# Certificate Generator

Certificate Generator is a Python application created for the League of Extraordinary Minds during a programming internship. It automates the process of generating personalized certificates from a template and participant data in an Excel file.

## Features

- Import a PDF certificate template and convert it for editing
- Load participant data from an XLSX file
- Automatically generate certificates with personalized data
- Export certificates as PDF files
- Send certificates via email
- User-friendly GUI with CustomTkinter

## Requirements

- Python 3.x
- opencv-python
- Pillow
- img2pdf
- openpyxl
- pdf2image
- customtkinter
- Poppler for Windows (included in `poppler-23.07.0`)

## Folder Structure

- `main.py` – Main application logic and GUI
- `pathsAndMail.py` – Paths and email configuration
- `fontStyle/BAHNSCHRIFT.TTF` – Font used for certificates
- `poppler-23.07.0/` – Poppler binaries for PDF processing
- `templates/` – Certificate templates (created at runtime)
- `Wygenerowane zaswiadczenia/` – Generated certificates (created at runtime)

## Usage

1. **Install dependencies:**
   ```
   pip install opencv-python Pillow img2pdf openpyxl pdf2image customtkinter
   ```
2. **Run the application:**
   ```
   python main.py
   ```
3. **Load your template and XLSX file using the GUI.**
4. **Generate and (optionally) send certificates.**

## License

This project is licensed under the GNU General Public License v3.0. See [LICENSE](LICENSE) for details.

---

Created for the League of Extraordinary Minds on a programming internship.