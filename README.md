# Universal File Analyzer and Converter

A multi-purpose tool to read, analyze, visualize, and convert files between many formats.

## Supported formats

**Data:** CSV, Excel (XLSX, XLS), JSON

**Documents:** TXT, PDF, DOCX, DOC

**Images:** JPG, JPEG, PNG, GIF, BMP, WEBP

## Conversions

**Data files**
- Excel ↔ CSV
- JSON ↔ CSV ↔ Excel
- Any format → PDF

**Documents**
- PDF → TXT
- DOCX → TXT, PDF
- TXT → PDF

**Images**
- Any format → PDF

## Analysis and visualization

- **Interactive charts**: line, bar, histogram
- **Statistics**: mean, median, standard deviation
- **Correlation matrix**: see how columns relate
- **AI assistant**: chat about your file using the Groq API

## Tech stack

- **Streamlit** — web interface
- **pandas** — data processing
- **Plotly** — interactive charts
- **FPDF** — PDF creation
- **Pillow** — image handling
- **PyPDF2** — read PDFs
- **python-docx** — read Word files
- **Groq** — AI chat

## Setup

Requirements: Python 3.8+

```bash
git clone https://github.com/dastonbek8747/ai-file-analyzer.git
cd ai-file-analyzer

python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate

pip install -r requirements.txt

cp .env.example .env            # then set GROQ_API_KEY
```

Get a Groq API key at [groq.com](https://groq.com).

## Run

```bash
streamlit run main.py
```

The app opens at http://localhost:8501

## Usage

1. **Upload a file** in the left sidebar. It is analyzed automatically.
2. **View the data**:
   - Table tab — rows and columns
   - Visualization tab — charts
   - Info tab — summary statistics
3. **Convert**: pick a target format in the sidebar and download the result.
4. **Ask the AI**: type a question in the chat box at the bottom. The AI uses the file as context.

Example questions:
- "Which value is the highest in this data?"
- "What is the main idea of this text?"
- "What conclusions can we draw from this data?"

## Project structure

```
main.py           # Streamlit app
services.py       # conversion helpers
.env              # environment variables (not committed)
requirements.txt  # dependencies
Files/            # uploaded files
Output/           # converted files
```

## Configuration

- **Retries**: file operations retry up to 3 times on error.
- **PDF limits**: data is shown up to 100 rows; more than 8 columns switch to vertical layout; images fit an A4 page.

## Troubleshooting

**Permission error** — the file is open in another program. Close it and try again.

**Groq API error** — check `GROQ_API_KEY` in `.env` and your internet connection.

**File will not upload** — check the file size and format, then clear the browser cache.

## License

MIT
