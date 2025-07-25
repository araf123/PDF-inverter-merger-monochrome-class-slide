# PDF Processor Suite

A Python desktop application built with Tkinter to perform batch operations on PDF files. It allows for merging, splitting into chunks, inverting colors, and applying a smart monochrome filter.

![Screenshot of App](https://i.imgur.com/YOUR_SCREENSHOT_URL.png)

## Features

- **Merge Multiple PDFs:** Combine several PDF files into one, in any order.
- **Page Editor:** Specify pages or page ranges to exclude from any file.
- **Batch Processing:** Processes hundreds of pages without running out of memory by using a "divide and conquer" chunking strategy.
- **Invert Colors:** Inverts the colors of PDFs, ideal for documents with a dark background.
- **Smart Monochrome:** Converts pages to pure black-and-white for crisp, clean text and smaller file sizes.
- **N-Up Layout:** Save the final output as 1, 2, 3, or 4 pages per sheet.

## Requirements

*   Python 3.x
*   Ghostscript (must be installed and accessible in the system's PATH)
*   The Python libraries listed in `requirements.txt`. Install them using:
    ```
    pip install -r requirements.txt
    ```

## How to Run

1.  Ensure you have [Ghostscript](https://www.ghostscript.com/releases/gsdnld.html) installed.
2.  Clone the repository: `git clone https://github.com/YourUsername/pdf-processor-suite.git`
3.  Navigate to the project directory: `cd pdf-processor-suite`
4.  Install dependencies: `pip install -r requirements.txt`
5.  Run the application: `python pdf_app.py`