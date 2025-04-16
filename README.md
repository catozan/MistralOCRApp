# 📄 MistralOCRApp

A powerful **VB.NET**-based OCR utility for extracting structured text and images from PDFs and scanned documents using the **Mistral AI API**.

---

## 🚀 Features

- 🔍 **OCR from Local Files** – Upload and process PDF or image documents.
- 🌐 **OCR from URLs** – Extract text from documents hosted online.
- 🆔 **OCR from File ID** – Resume or reprocess previously uploaded files.
- 📥 **File Upload Support** – Automatically uploads and registers documents via Mistral API.
- 🖼️ **Image Extraction** – Decodes and saves base64 images from OCR output.
- 🧾 **Multi-format Output** – Generates `.txt`, `.md`, and `.json` files.
- 📁 **Organized Output** – Automatically saves outputs to `~/Desktop/MistralOCR_Output/`.
- 🖱️ **Interactive CLI** – Minimal setup, easy to run from terminal.

---

## 🛠 Requirements

- Windows system with .NET Framework 4.8+
- [Newtonsoft.Json](https://www.nuget.org/packages/Newtonsoft.Json) (for JSON handling)
- Internet connection

---

## 📦 Project Structure

```
MistralOCRApp.vb
└── (Main entry point)

📁 MistralOCR_Output/
├── output.txt       => Full extracted text
├── output.md        => Combined markdown content
├── response.json    => Raw OCR JSON response
├── page_X_images/   => Base64-decoded images from each page
```

---

## 🔧 How to Use

### 🏁 Run the App

1. Build the app in Visual Studio or run via terminal.
2. Follow CLI prompts:
   - Enter your **Mistral API Key**
   - Choose document input method:
     - `1` Local File (path)
     - `2` URL to a document
     - `3` Existing Mistral File ID

```bash
> Enter your Mistral API key: sk-xxxxxxx
> Select document source (1 for local file, 2 for URL, 3 for file ID): 1
> Enter the path to your document file: C:\Documents\Report.pdf
```

### 📂 Output Directory

After processing completes, all results will be saved under:
```
C:\Users\<YourName>\Desktop\MistralOCR_Output\
```

You will be prompted to open the folder:
```
Do you want to open the output folder? (y/n): y
```

---

## 📜 Output Samples

### ✅ Markdown Output (`output.md`)
```md
# Introduction

## Section Title

Paragraph content from page.

---
```

### 🖼️ Image Output
Images are saved as `page_X_images/img_*.png`.

---

## 🧠 API Documentation Reference

- [Mistral AI OCR API](https://docs.mistral.ai/capabilities/document/)

---

## 💬 Feedback & Issues

If you encounter bugs or want to suggest features, open an [issue on GitHub](https://github.com/catozan/MistralOCRApp).

---

## 📌 Future Improvements
- Add GUI (WinForms/WPF)
- Add language/model selection
- Export to Word or Excel
- Add PDF merge tool from images

---

## 🔐 Disclaimer
Make sure not to expose your API keys. Always store them securely (e.g., using secrets manager or config file).

