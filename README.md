# 🔒 DOCX Anonymizer

A tool for anonymizing sensitive information in Word documents. Built with Python, Streamlit, and python-docx.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.41+-red.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

## ✨ Features

- **🔑 Keyword Replacement** - Replace specific names, companies, or phrases with `[REDACTED_X]` placeholders
- **📚 Persistent Dictionary** - Keywords are saved and reused across sessions
- **💰 Financial Data Detection** - Auto-detect and anonymize currency amounts ($, €, £, ¥, etc.)
- **🔐 PII Auto-Detection** - Automatically find and redact:
  - Email addresses
  - Phone numbers
  - Social Security Numbers
  - Credit card numbers
  - IP addresses
  - Dates
- **📁 Multi-File Processing** - Upload and process up to 20 files at once
- **📦 Batch Download** - Download all processed files as a ZIP
- **🎨 Modern UI** - Dark theme with glassmorphism design

## 🚀 Quick Start

### Prerequisites
- Python 3.8 or higher

### Installation

```bash
# Clone the repository
git clone https://github.com/albertcande/docx-anonymizer.git
cd docx-anonymizer

# Install dependencies
pip install -r requirements.txt

# Run the application
streamlit run app.py
```

Open your browser to `http://localhost:8501`

## 📖 Usage

1. **Upload** - Drag and drop one or more `.docx` files
2. **Configure** - Enable options in the sidebar:
   - ☑️ Include dictionary keywords
   - ☑️ Anonymize financial data
   - ☑️ Anonymize PII data
3. **Add Keywords** - Enter additional keywords (comma-separated)
4. **Process** - Click "Process Document(s)"
5. **Download** - Get your anonymized files

## 🔧 Configuration

### File Limits
| Setting | Value |
|---------|-------|
| Max file size | 50 MB |
| Max files per batch | 20 |
| Max keywords | 100 |
| Max keyword length | 200 chars |

Configure in `.streamlit/config.toml`:
```toml
[server]
maxUploadSize = 50
```

## � Project Structure

```
docx_processor/
├── app.py                 # Streamlit UI
├── processor.py           # Core anonymization logic
├── requirements.txt       # Python dependencies
├── README.md              # This file
├── keyword_dictionary.json # Persistent keyword storage (auto-generated)
├── .streamlit/
│   └── config.toml        # Streamlit configuration
└── test_docs/             # Sample test documents
```

## �️ Security Features

- **Thread-safe dictionary** - File locking prevents race conditions
- **Input validation** - Size limits and keyword sanitization
- **No data retention** - Processed files exist only in memory

## 🧪 Testing

Generate sample test documents:
```bash
python generate_test_docs.py
```

This creates 5 test files in `test_docs/`:
- `01_simple_letter.docx` - Names, emails, phone numbers
- `02_financial_report.docx` - Financial data with tables
- `03_employee_directory.docx` - Employee info table
- `04_service_contract.docx` - Contract with multiple parties
- `05_nested_tables.docx` - Nested table structures

## 🔌 API Reference

### `anonymize_docx()`

```python
from processor import anonymize_docx

result, stats = anonymize_docx(
    file_stream,                    # BinaryIO - DOCX file
    keywords=["John Doe"],          # List or Dict of keywords
    include_dictionary=True,        # Use saved keywords
    anonymize_financial=True,       # Detect currency amounts
    anonymize_pii=True,             # Detect PII patterns
    placeholder_template="[REDACTED_{n}]"
)
```

### `ProcessingStats`

```python
stats.keywords_replaced   # int - Number of keyword replacements
stats.financial_replaced  # int - Number of financial amounts replaced
stats.pii_replaced        # Dict[str, int] - PII counts by type
stats.total_replacements() # int - Total replacements made
```

## 🛠️ Technologies

- **[Streamlit](https://streamlit.io/)** - Web UI framework
- **[python-docx](https://python-docx.readthedocs.io/)** - Word document processing
- **[filelock](https://py-filelock.readthedocs.io/)** - Thread-safe file operations

## 📄 License

MIT License - see [LICENSE](LICENSE) for details.

## 🤝 Contributing

Contributions welcome! Please open an issue or submit a pull request.

---

Built with ❤️ using Python


