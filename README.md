# Company Document Generator

A Streamlit-based AI document generation platform that creates **Business Intelligence reports**, **Executive Storylines**, **ADM Financial Summaries**, and **Application Discovery & Modernization (ADM) proposals** using Gemini, LangGraph, Pydantic validation, and deterministic financial logic.

The application is designed for consulting-style enterprise deliverables where structure, financial consistency, and repeatability matter as much as the generated content.

---

## 1. Project Overview

This project generates structured company analysis and ADM proposal documents through a controlled multi-step workflow.

The platform currently supports:

- Business Intelligence report generation
- Leadership-based executive storyline generation
- ADM financial summary generation
- ROI-controlled ADM financial modeling
- Six-batch ADM proposal generation
- Structure and number validation
- TXT and DOCX downloads

The application is built as a single Streamlit app, currently contained in `doc.py` or `app.py`.

---

## 2. Core Workflow

```text
User Inputs
↓
Company Name + Gemini API Key + Optional Leadership Mapping
↓
Generate Business Intelligence
↓
Generate Financial Summary
↓
Generate ADM Batch 1
↓
Continue ADM Batches 2 to 6
↓
Validate Output
↓
Download TXT / DOCX Files
```

---

## 3. Main Features

### 3.1 Business Intelligence Generation

The BI module generates a structured company-level report with:

- Functional business lines
- Market leaders
- Current-state analysis
- Competitive benchmarks
- Business challenges
- AI reinvention opportunities
- Quantified annual impact table

Relevant code sections:

```python
BI_PROMPT
BI_STRUCTURE_FIX_PROMPT
generate_bi()
validate_bi_structure()
```

### 3.2 Executive Storylines

The Executive Storyline module creates meeting narratives for executives based on leadership mapping input.

It supports profiles such as:

- CEO
- CFO
- CMO
- CIO
- Business Line Head
- Business Line Technology Head
- Board Member

Relevant code sections:

```python
LEADERSHIP_EXTRACTION_PROMPT
STORYLINE_PROMPT
parse_exec_profiles_from_json()
generate_storylines()
```

### 3.3 ADM Financial Summary

The ADM Financial Summary is generated using a combination of:

- LLM-based extraction
- Pydantic schema validation
- Deterministic financial computation
- LangGraph orchestration

The financial engine calculates:

- Employee count
- Revenue in millions
- Sector classification
- Estimated application count
- Annual maintenance cost
- Modernization backlog
- Annual value drivers
- Five-year value
- Investment requirement
- ROI
- Payback period
- Partner and client investment split

Relevant code sections:

```python
FINANCIAL_EXTRACTION_PROMPT
CompanyFactsSchema
BusinessUnitSchema
ValueDriverSchema
FinancialExtractionSchema
financial_extract_node()
financial_compute_node()
run_financial_graph()
```

### 3.4 ROI Control

The application includes a deterministic ROI solver to prevent negative or unrealistic ROI values.

The ROI target range is:

```text
Minimum ROI: 150%
Maximum ROI: 300%
Target ROI: 180%
```

Relevant code section:

```python
solve_investment_for_roi()
```

The ROI formula used is:

```text
ROI = ((Five-Year Value - Investment) / Investment) × 100
```

This ensures the generated ADM financial summary stays commercially usable and does not produce negative ROI.

### 3.5 ADM Proposal Generation

The ADM proposal is generated in batches to reduce structure drift and improve reliability.

| Batch | Section Generated |
|---|---|
| Batch 1 | Executive Summary and Part 1: Detailed Application Portfolio Analysis |
| Batch 2 | Part 2: Competitive Benchmarking Against Market Leaders |
| Batch 3 | Part 3.1 and 3.2: Partnership Overview and Roadmap |
| Batch 4 | Part 3.3 and 3.4: Financial Model and Offshore Delivery |
| Batch 5 | Part 3.5 to 3.8: Governance, Risk, Transition, and Success Metrics |
| Batch 6 | Part 4, Appendices, and Footer |

Relevant code sections:

```python
ADM_BATCH1_PROMPT
ADM_CONTINUE_PROMPT
generate_adm_batch1()
generate_adm_next_batch()
```

### 3.6 Validation Engine

The validation layer checks:

- BI structure
- Financial math
- ADM section order
- ADM batch completion
- Required locked numbers
- Missing footer after Batch 6
- Unsupported zero values

Relevant code sections:

```python
validate_financial_math()
validate_bi_structure()
validate_adm_structure_and_numbers()
build_validation_report()
render_validation_report_text()
```

The ADM validator is batch-aware, meaning it only checks for sections that should exist at the current generation stage.

---

## 4. Project Structure

The current implementation is contained in one main Python file.

Recommended repository structure:

```text
project-root/
│
├── app.py
│   Main Streamlit application. This may also be named doc.py.
│
├── requirements.txt
│   Python dependencies required to run the app.
│
├── README.md
│   Project documentation.
│
└── .streamlit/
    └── secrets.toml
    Optional location for storing Gemini API key in Streamlit Cloud.
```

If the project is later refactored, a cleaner structure would be:

```text
project-root/
│
├── app.py
│
├── prompts/
│   ├── bi_prompts.py
│   ├── adm_prompts.py
│   └── storyline_prompts.py
│
├── financial/
│   ├── extraction.py
│   ├── compute.py
│   └── tables.py
│
├── validation/
│   ├── financial_validation.py
│   ├── adm_validation.py
│   └── report.py
│
├── utils/
│   ├── formatting.py
│   ├── parsing.py
│   └── docx_export.py
│
├── requirements.txt
└── README.md
```

---

## 5. Key Code Sections

### 5.1 Configuration

```python
HARDCODED_GEMINI_API_KEY = ""
DEFAULT_MODEL_NAME = "gemini-3.1-pro-preview"
```

The API key can be supplied through:

1. Hardcoded variable
2. Streamlit secrets
3. Sidebar input

Recommended production approach:

```toml
# .streamlit/secrets.toml
GEMINI_API_KEY = "your-api-key-here"
```

### 5.2 Data Models

The app uses dataclasses, TypedDicts, and Pydantic models to structure data.

Important models:

```python
ExecProfile
AdmFinancialState
CompanyFactsSchema
BusinessUnitSchema
ValueDriverSchema
FinancialExtractionSchema
ValidationReportSchema
```

### 5.3 Helper Functions

Helper functions handle:

- Filename sanitization
- DOCX export
- JSON cleaning
- Number formatting
- Safe numeric parsing
- Zero-value detection
- Pydantic v1 and v2 compatibility

Examples:

```python
sanitize_filename()
save_docx_bytes()
clean_json_response()
safe_float()
safe_int()
mfmt()
pfmt()
pydantic_to_dict()
```

### 5.4 Gemini Client

The Gemini client wrapper handles model calls and response extraction.

```python
class GeminiClient:
    def __init__(self, api_key: str, model: str)
    def generate(self, prompt: str) -> str
```

### 5.5 LangGraph Financial Agent

LangGraph is used to orchestrate the ADM financial flow.

```text
extract_inputs
↓
compute_financials
↓
validate_financials
↓
END
```

Relevant function:

```python
run_financial_graph()
```

---

## 6. Installation

Install dependencies:

```bash
pip install streamlit google-genai langgraph langchain-core pydantic python-docx
```

Or create a `requirements.txt` file:

```text
streamlit
google-genai
langgraph
langchain-core
pydantic
python-docx
```

Then install:

```bash
pip install -r requirements.txt
```

---

## 7. Running the App

Run locally:

```bash
streamlit run app.py
```

If your file is named `doc.py`, run:

```bash
streamlit run doc.py
```

---

## 8. Streamlit Cloud Deployment

For Streamlit Cloud:

1. Push the repo to GitHub.
2. Add `requirements.txt`.
3. Add the Gemini API key in Streamlit secrets:

```toml
GEMINI_API_KEY = "your-api-key"
```

4. Set the main file as:

```text
app.py
```

If the file is named differently, update Streamlit Cloud settings accordingly.

---

## 9. User Interface

The app provides six main actions:

| Button | Purpose |
|---|---|
| Generate BI | Creates the Business Intelligence report |
| Generate Storylines | Creates executive storylines from leadership mapping |
| Generate Financial Summary | Runs financial extraction, computation, and validation |
| Generate ADM Batch 1 | Creates the first ADM section |
| Continue ADM | Generates the next ADM batch |
| Validate | Runs BI, financial, and ADM validation |

Output tabs include:

- Business Intelligence
- Executive Storylines
- ADM Financial Summary
- ADM
- Validation Report

---

## 10. ADM Document Structure

The final ADM document follows this structure:

```text
COMPREHENSIVE APPLICATION PORTFOLIO ANALYSIS & 5-YEAR TRANSFORMATION PARTNERSHIP

EXECUTIVE SUMMARY: THE STRATEGIC IMPERATIVE

PART 1: DETAILED APPLICATION PORTFOLIO ANALYSIS
1.1 Application Portfolio Composition & Characteristics
1.2 Business Unit Deep Dive
1.3 Business Unit Deep Dive
1.4 Business Unit Deep Dive
1.5 Business Unit Deep Dive
1.6 Business Unit Deep Dive, if applicable

PART 2: COMPETITIVE BENCHMARKING AGAINST MARKET LEADERS

PART 3: 5-YEAR TRANSFORMATION PARTNERSHIP DEAL STRUCTURE
3.1 Partnership Overview & Commercial Terms
3.2 Year-by-Year Investment & Delivery Roadmap
3.3 Detailed Financial Model
3.4 Offshore Delivery Model & Cost Advantage
3.5 Governance & Operating Model
3.6 Risk Mitigation Framework
3.7 Transition Approach
3.8 Success Metrics & Performance Dashboard

PART 4: CONCLUSION & STRATEGIC IMPERATIVES
4.1 The Competitive Imperative
4.2 The Partnership Advantage
4.3 Critical Success Factors
4.4 Recommended Next Steps
4.5 Final Investment Summary

APPENDICES

Prepared for:
Prepared by:
Date:
Tholons Contacts:
```

---

## 11. Common Issues and Fixes

### 11.1 `ADM_BATCH1_PROMPT is not defined`

Cause:

The Batch 1 prompt was deleted or placed incorrectly.

Fix:

Ensure this appears before `ADM_CONTINUE_PROMPT`:

```python
ADM_BATCH1_PROMPT = """
...
"""
```

### 11.2 SyntaxError near `PART 2`

Cause:

Prompt text was placed outside triple quotes.

Fix:

Make sure all prompt text is inside:

```python
ADM_CONTINUE_PROMPT = """
...
"""
```

### 11.3 Negative ROI

Cause:

Investment was previously calculated using an old multiplier-only method.

Fix:

Use:

```python
solve_investment_for_roi()
```

and ensure validation matches the ROI-based investment logic.

### 11.4 Validation Fails After Batch 1

Cause:

Validator expects all ADM sections before all batches are complete.

Fix:

Use batch-aware validation:

```python
validate_adm_structure_and_numbers(adm_text, fs, adm_batch=adm_batch)
```

### 11.5 Missing Footer

Cause:

Batch 6 did not generate the footer.

Fix:

Ensure `ADM_CONTINUE_PROMPT` includes:

```text
Prepared for:
{company_name} Executive Leadership Team

Prepared by:
Deloitte Consulting LLP & Tholons Inc.

Date: March 2026

Tholons Contacts:
Abhay Anant Vashistha; abhay@tholons.com
Frank Pendle; frank@tholons.com
Avinash Vashistha; avi@tholons.com
```

---

## 12. Development Guidelines

When contributing to this project:

### Do

- Keep financial logic deterministic.
- Keep prompts strict and structured.
- Update validation logic when changing output structure.
- Test each batch before changing the next.
- Use clear function names.
- Keep generated financial numbers tied to `financial_summary`.

### Do Not

- Let the LLM freely calculate ROI.
- Put prompt text outside triple quotes.
- Remove batch-aware validation.
- Use zero-value placeholders.
- Mix Batch 1 logic into `ADM_CONTINUE_PROMPT`.
- Mix Batch 2 to 6 logic into `ADM_BATCH1_PROMPT`.

---

## 13. Recommended Improvements

Future contributors can improve the project by:

- Splitting the single file into modules
- Adding unit tests for financial math
- Adding retry logic for failed LLM generations
- Adding PDF export
- Adding HTML dashboard export
- Adding structured JSON output for ADM
- Improving UI state management
- Adding progress indicators per batch
- Making the footer dynamic by month and year

---

## 14. Testing Checklist

Before committing changes:

- [ ] App starts without syntax errors
- [ ] BI generation works
- [ ] Financial Summary generation works
- [ ] ROI remains between 150% and 300%
- [ ] ADM Batch 1 generates successfully
- [ ] Continue ADM generates Batches 2 to 6
- [ ] Batch 6 includes footer
- [ ] Validation passes after full ADM generation
- [ ] TXT downloads work
- [ ] DOCX downloads work

---

## 15. Contribution Workflow

Recommended Git workflow:

```bash
git checkout -b feature/your-feature-name
git add .
git commit -m "Describe your change clearly"
git push origin feature/your-feature-name
```

Then open a pull request with:

- What changed
- Why it changed
- How it was tested
- Screenshots if UI was changed

---

## 16. Notes for Contributors

This project depends heavily on prompt structure. A small formatting change can break downstream validation.

Before editing prompts, check:

1. Where the prompt is defined
2. Which function uses it
3. Which validator expects its output
4. Whether generated sections are batch-specific

The most important principle:

```text
The LLM writes the document.
The code owns the numbers.
The validator protects the structure.
```

---

## 17. License

Add your preferred license here.

Common options:

- MIT License
- Apache 2.0
- Proprietary internal use only

---

## 18. Maintainers

Add maintainer details here.

```text
Maintainer: [Your Name / Team]
Organization: Tholons
Contact: [Email]
```
