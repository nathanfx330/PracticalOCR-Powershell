# Practical OCR - PowerShell Version

This PowerShell script provides a workflow for converting standard PDF files into searchable PDFs using external command-line tools for image conversion and Optical Character Recognition (OCR).

## Features

*   Converts selected PDFs from a `./pdfs` input directory.
*   Extracts pages as JPG images using ImageMagick.
*   Optionally adjusts image levels (contrast) before OCR.
*   Performs OCR on images using Tesseract OCR to create searchable PDF layers.
*   Merges the processed pages into a final searchable PDF using PDFtk.
*   Checks for required tools based on paths configured *within the script*.
*   Supports resuming partially completed jobs.
*   Creates `./output`, `./ocr_output`, and `./final_merge` directories for intermediate and final files.

## Requirements

This script **requires** the following external command-line tools to be installed separately:

1.  **ImageMagick (Version 7+ Recommended)**
    *   Provides `magick.exe` for image conversion.
    *   Download: [https://imagemagick.org/script/download.php](https://imagemagick.org/script/download.php)
2.  **Poppler (Windows Binaries)**
    *   Provides `pdfinfo.exe` to get PDF page counts.
    *   Download: Search for "poppler windows binaries". A common source is [https://github.com/oschwartz10612/poppler-windows/releases](https://github.com/oschwartz10612/poppler-windows/releases). You will need to extract the archive.
3.  **Tesseract OCR**
    *   Provides `tesseract.exe` for OCR.
    *   **Requires Language Data:** You MUST install the language data files (e.g., `eng.traineddata` for English) for the languages you intend to OCR.
    *   Download: The installer from [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki) is recommended as it helps manage language data installation.
    *   **Important:** Tesseract needs to find its language data. Ensure the installed language files (e.g., `eng.traineddata`) are located within a subfolder named `tessdata` inside the main Tesseract installation directory (e.g., `C:\Program Files\Tesseract-OCR\tessdata\`).
4.  **PDFtk Server (Free command-line version)**
    *   Provides `pdftk.exe` for merging PDFs.
    *   Download: [https://www.pdflabs.com/tools/pdftk-server/](https://www.pdflabs.com/tools/pdftk-server/) (Note: Check license terms). Alternatively, Java-based ports exist.

## Setup - IMPORTANT!

1.  **Install Requirements:** Install all four tools listed above.
2.  **Locate Executables:** Find the exact installation location (the full path including the `.exe` filename) for:
    *   `magick.exe` (from ImageMagick 7+)
    *   `pdfinfo.exe` (usually in the `bin` folder of your extracted Poppler files)
    *   `tesseract.exe`
    *   `pdftk.exe` (often in a `bin` folder)
3.  **Configure Script Paths:**
    *   Open the `PracticalOCR.ps1` script file in a text editor (like VS Code, Notepad++, or even Notepad).
    *   Locate the section marked `--- !! IMPORTANT: Tool Path Configuration !! ---`.
    *   **Replace the placeholder paths** for `$ConvertPath`, `$PdfInfoPath`, `$TesseractPath`, and `$PdftkPath` with the **full paths** you found in step 2. Ensure the paths are enclosed in double quotes.
    *   **Example:**
        ```powershell
        $ConvertPath   = "C:\Program Files\ImageMagick-7.1.1-Q16\magick.exe"
        $PdfInfoPath   = "C:\tools\poppler-23.11.0\Library\bin\pdfinfo.exe"
        $TesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
        $PdftkPath     = "C:\Program Files (x86)\PDFtk\bin\pdftk.exe"
        ```
    *   Save the script file.

## Usage

1.  **Create Input Folder:** In the same directory where you saved `PracticalOCR.ps1`, create a subfolder named `pdfs`.
2.  **Add PDFs:** Place the PDF files you want to process into the `pdfs` folder.
3.  **Run from PowerShell:**
    *   Open PowerShell.
    *   Navigate (`cd`) to the directory containing the script.
    *   You may need to adjust your PowerShell Execution Policy to allow local scripts to run. For the current session only, you can often use:
        ```powershell
        Set-ExecutionPolicy Bypass -Scope Process -Force
        ```
        *(Use `Bypass` with caution, or try `RemoteSigned`)*.
    *   Execute the script:
        ```powershell
        .\PracticalOCR.ps1
        ```
4.  **Follow Prompts:**
    *   The script will first verify the tool paths you configured.
    *   It will list the PDFs found in the `./pdfs` directory and ask you to select one by number.
    *   It will ask whether you want to apply image levels adjustment (useful for scanned documents, slower).
5.  **Processing:** The script will proceed through image extraction, OCR, and merging stages, showing progress.
6.  **Output:**
    *   Intermediate JPG images are stored in `./output`.
    *   Intermediate single-page OCR PDFs are stored in `./ocr_output`.
    *   The final merged, searchable PDF is saved in `./final_merge` with `_final.pdf` appended to the original base name.

## Troubleshooting

*   **"File not found at specified path" / "term ... is not recognized":** Double-check the paths configured in the script match the actual installation locations *exactly*. Test paths manually in PowerShell using `& "C:\Full\Path\To\Tool.exe" --version` (or similar version flag).
*   **Tesseract Errors about Language Data / `TESSDATA_PREFIX`:** Ensure the required `.traineddata` files are in the `tessdata` subfolder of your main Tesseract installation directory.
*   **PDFtk Merge Errors:** Ensure `pdftk.exe` is working correctly. Try merging a few small PDFs manually using the `pdftk` command line to isolate issues.
*   **PowerShell Execution Policy Errors:** See the `Set-ExecutionPolicy` command under "Usage".

## License

MIT License

Copyright (c) 2025 nathanfx330

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
