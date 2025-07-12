# OCR and Document Processing Tool

This is a versatile desktop application built with Python and Tkinter, designed to streamline document processing workflows. It combines Optical Character Recognition (OCR), advanced text editing, language translation, and text-to-speech functionalities into a single, intuitive interface.

## Application Variants

This repository contains three main Python scripts, each offering a slightly different set of functionalities:

*   **`new.py` (Recommended)**: This is the most comprehensive version of the application. It provides a full-featured document editor with integrated OCR, translation, and text-to-speech capabilities. It's designed for users who need a single tool for all their document processing needs.

*   **`editor.py`**: This script focuses primarily on document editing, offering a rich text editor with basic formatting, file management, and some voice typing and text-to-speech features. It also includes a basic OCR function.

*   **`gui.py`**: This script is a dedicated OCR and conversion tool. Its main purpose is to extract text from PDF files (with page range selection) and convert them into Word documents (`.docx`) or spoken audio (`.mp3`). It also includes translation capabilities for the extracted text.

## Features

### Document Editing (primarily in `new.py` and `editor.py`)
*   **Rich Text Editor**: A modern, Word-like interface for creating and editing documents.
*   **File Management**: Open, save, and create new documents in `.docx`, `.txt`, and `.pdf` formats.
*   **Basic Formatting**: Apply bold, italic, underline, change font family and size.
*   **Advanced Formatting**: Text alignment (left, center, right), text color, highlighting, bullet and numbered lists, indentation, and line spacing.
*   **Find and Replace**: Easily search for and replace text within your document.
*   **Autosave**: Automatically saves your work at regular intervals.
*   **Word Count**: Real-time display of word count.

### Optical Character Recognition (OCR) (available in all variants, most advanced in `new.py` and `gui.py`)
*   **Image OCR**: Extract text from various image formats (PNG, JPG, BMP, TIFF).
*   **PDF OCR**: Extract text from PDF documents, with options to specify page ranges.
*   **Multi-language Support**: Supports OCR for English, Tamil, Hindi, Telugu, Malayalam, Kannada, French, Spanish, German, Chinese (Simplified), and Japanese.
*   **Advanced Preprocessing**: Utilizes image processing techniques (grayscale, adaptive thresholding, Otsu's thresholding) for improved OCR accuracy.
*   **Post-processing**: Cleans and refines extracted text to fix common OCR errors.
*   **Progress Tracking**: Monitors and displays the progress of OCR operations.

### Language Translation (primarily in `new.py` and `gui.py`)
*   **Translate Text**: Translate selected text or entire documents between multiple languages using Google Translate.
*   **Supported Languages**: English, Tamil, Hindi, Telugu, Malayalam, Kannada, French, Spanish, German, Chinese, and Japanese.
*   **Progress Tracking**: Shows translation progress for large texts.

### Text-to-Speech (TTS) (primarily in `new.py` and `gui.py`)
*   **Read Aloud**: Convert selected text or entire documents into spoken audio.
*   **Save Audio**: Save the generated speech as an MP3 file.
*   **Multi-language Support**: Supports text-to-speech for English, Tamil, Hindi, Telugu, Malayalam, Kannada, French, Spanish, German, Chinese, and Japanese.
*   **Progress Tracking**: Displays progress during audio generation.

### User Interface (UI/UX)
*   **Modern Design**: Utilizes `ttkthemes` for a contemporary look and feel.
*   **Ribbon Interface**: Organized features in a familiar tabbed ribbon layout (Home, View, Tools) in `new.py`.
*   **Light/Dark Theme**: Toggle between light and dark modes for comfortable viewing.
*   **Tooltips**: Provides helpful tooltips for buttons and features.

## Installation

### Prerequisites
*   **Python 3.x**: Ensure you have Python installed.
*   **Tesseract OCR**: This application relies on Tesseract OCR. Download and install it from the [official Tesseract GitHub page](https://tesseract-ocr.github.io/tessdoc/Installation.html).
    *   **Windows**: The installer `tesseract-ocr-w64-setup-5.5.0.20241111.exe` is included in the repository for convenience. Make sure to add Tesseract to your system's PATH during installation, or update the `pytesseract.pytesseract.tesseract_cmd` variable in `editor.py`, `gui.py`, and `new.py` to point to your Tesseract executable.

### Setup
1.  **Clone the repository**:
    ```bash
    git clone https://github.com/your-username/ocr-tool.git
    cd ocr-tool
    ```
2.  **Create a virtual environment** (recommended):
    ```bash
    python -m venv myenv
    ```
3.  **Activate the virtual environment**:
    *   **Windows**:
        ```bash
        .\myenv\Scripts\activate
        ```
    *   **macOS/Linux**:
        ```bash
        source myenv/bin/activate
        ```
4.  **Install dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

To run the main application (recommended), execute `new.py`:

```bash
python new.py
```

Alternatively, you can run the other variants:

*   **Document Editor**: `python editor.py`
*   **OCR Conversion Tool**: `python gui.py`

## Technologies Used

*   **Python 3**
*   **Tkinter**: For the graphical user interface.
*   **Pillow (PIL)**: For image processing.
*   **PyMuPDF (fitz)**: For PDF handling.
*   **pytesseract**: Python wrapper for Tesseract OCR.
*   **python-docx**: For Word document (`.docx`) creation and manipulation.
*   **gTTS**: Google Text-to-Speech for audio generation.
*   **langdetect**: For language detection.
*   **deep-translator**: For text translation.
*   **speechRecognition**: For voice typing (speech-to-text).
*   **opencv-python (cv2)**: For advanced image preprocessing in OCR.
*   **numpy**: For numerical operations, especially with OpenCV.
*   **ttkthemes**: For modern Tkinter themes.
*   **pygame**: For audio playback (in `editor.py`).
*   **tkinterdnd2**: For drag-and-drop functionality (in `gui.py`).

## Contributing

Contributions are welcome! Please feel free to fork the repository, make your changes, and submit a pull request.

## License

[Specify your license here, e.g., MIT License]