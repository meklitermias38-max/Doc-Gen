from __future__ import annotations

import io
import json
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, TypedDict

import streamlit as st
from docx import Document
from docx.shared import Pt

try:
    from google import genai
except ImportError:
    genai = None

try:
    from langgraph.graph import StateGraph, END
except ImportError:
    StateGraph = None
    END = None


# ============================================================
# CONFIG
# ============================================================

HARDCODED_GEMINI_API_KEY = ""
DEFAULT_MODEL_NAME = "gemini-3.1-pro-preview"


st.set_page_config(page_title="Company Document Generator", layout="wide")
st.title("Company Document Generator")
st.caption("Generate Business Intelligence, Executive Storylines, and ADM using Gemini")


# ============================================================
# DATA MODELS
# ============================================================

@dataclass
class ExecProfile:
    name: str
    title: str
    linkedin: str
    type: str
    business_line: Optional[str] = None


class AdmFinancialState(TypedDict, total=False):
    company_name: str
    bi_text: str
    extracted_inputs: Dict[str, Any]
    financial_summary: Dict[str, Any]
    error: str


# ============================================================
# HELPERS
# ============================================================

def sanitize_filename(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._ -]+", "", name).strip().replace(" ", "_")


def save_docx_bytes(title: str, body: str) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    doc.add_heading(title, level=0)
    for line in body.split("\n"):
        doc.add_paragraph(line)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


def clean_json_response(text: str) -> str:
    text = text.strip()

    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:].strip()

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    return text.strip()


def parse_exec_profiles_from_json(raw_json: str) -> List[ExecProfile]:
    cleaned = clean_json_response(raw_json)
    data = json.loads(cleaned)

    items = data.get("executives", [])
    profiles: List[ExecProfile] = []

    for item in items:
        profiles.append(
            ExecProfile(
                name=item.get("name", ""),
                title=item.get("title", ""),
                linkedin=item.get("linkedin", ""),
                type=item.get("type", ""),
                business_line=item.get("business_line"),
            )
        )

    return profiles


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if value is None:
            return default
        if isinstance(value, (int, float)):
            return float(value)
        cleaned = str(value).replace(",", "").replace("$", "").replace("%", "").strip()
        return float(cleaned)
    except Exception:
        return default


def safe_int(value: Any, default: int = 0) -> int:
    try:
        if value is None:
            return default
        if isinstance(value, int):
            return value
        return int(round(safe_float(value, default)))
    except Exception:
        return default


def round1(x: float) -> float:
    return round(float(x), 1)


def mfmt(x: float) -> str:
    return f"${round1(x):,.1f}M"


def pfmt(x: float) -> str:
    return f"{round1(x)}%"


def mfmt_or_na(x: Any) -> str:
    try:
        if x is None:
            return "N/A"
        x = float(x)
        if abs(x) < 0.0001:
            return "N/A"
        return f"${round(x, 1):,.1f}M"
    except Exception:
        return "N/A"


def pfmt_or_na(x: Any) -> str:
    try:
        if x is None:
            return "N/A"
        x = float(x)
        if abs(x) < 0.0001:
            return "N/A"
        return f"{round(x, 1)}%"
    except Exception:
        return "N/A"


def contains_bad_zero_values(text: str) -> bool:
    bad_patterns = [
        r"\$0(?:\.0+)?M",
        r"\$0(?:\.0+)?K",
        r"\b0(?:\.0+)?%\b",
        r"\|\s*0(?:\.0+)?\s*\|",
    ]
    return any(re.search(p, text) for p in bad_patterns)


def assert_no_bad_zero_values(text: str, label: str) -> None:
    if contains_bad_zero_values(text):
        raise ValueError(f"{label} contains unsupported zero values. Recalculate and regenerate.")


def allocate_component_total(total: float, weights: List[float]) -> List[float]:
    total = round1(total)
    raw = [total * w for w in weights]
    rounded = [round1(x) for x in raw]
    diff = round1(total - sum(rounded))
    rounded[-1] = round1(rounded[-1] + diff)
    return rounded


def find_payback_years(investment: float, annual_value: float) -> float:
    if investment <= 0 or annual_value <= 0:
        return 0.0

    yearly_values = [
        annual_value * 0.15,
        annual_value * 0.40,
        annual_value * 0.70,
        annual_value * 0.90,
        annual_value * 1.00,
    ]
    cumulative = 0.0
    for idx, yv in enumerate(yearly_values, start=1):
        prev = cumulative
        cumulative += yv
        if cumulative >= investment:
            remaining = investment - prev
            fraction = remaining / yv if yv > 0 else 0
            return round(idx - 1 + fraction, 2)
    return 5.0


# ============================================================
# GEMINI CLIENT
# ============================================================

class GeminiClient:
    def __init__(self, api_key: str, model: str) -> None:
        if genai is None:
            raise ImportError("google-genai is not installed. Run: python -m pip install -U google-genai")
        self.client = genai.Client(api_key=api_key)
        self.model = model

    def generate(self, prompt: str) -> str:
        response = self.client.models.generate_content(
            model=self.model,
            contents=prompt,
        )

        text = getattr(response, "text", None)
        if text and text.strip():
            return text.strip()

        try:
            candidates = getattr(response, "candidates", []) or []
            parts: List[str] = []
            for cand in candidates:
                content = getattr(cand, "content", None)
                if not content:
                    continue
                for part in getattr(content, "parts", []) or []:
                    maybe_text = getattr(part, "text", None)
                    if maybe_text:
                        parts.append(maybe_text)
            return "\n".join(parts).strip()
        except Exception:
            return ""


# ============================================================
# PROMPTS
# ============================================================

BI_PROMPT = """
You are a senior enterprise strategy consultant.

Your task is to generate a complete Business Intelligence document for the company below.

Company Name: {company_name}

STRICT OUTPUT RULES
- Identify exactly 3 to 5 top business lines only
- Business lines must be FUNCTIONAL, not geographic
- If business lines are based on geography, regenerate internally before answering
- Do not include any introduction, disclaimer, methodology note, or conclusion before the business lines
- Start directly with the business lines
- Keep the structure exactly as specified below
- Use the exact section labels below
- Do not skip any section
- Do not merge sections
- Do not change the order
- Do not use placeholder text
- Do not use generic filler
- Quantify impact wherever possible
- Be specific to the company

REQUIRED OUTPUT STRUCTURE

First, output this heading exactly:
TOP BUSINESS LINES AND REVENUE

Then provide a table in this exact format:
| Business Line | Estimated Revenue | Notes |
|---|---:|---|

After the table, create a separate full section for EACH business line using this exact structure:

## [Business Line Name]

**Market Leaders:** [List named competitors]

**What "Good" Looks Like Today in {company_name}:**
- [Bullet 1]
- [Bullet 2]
- [Bullet 3]

**What "Good" Looks Like Today Across Market Leaders:**
I. [Competitor 1]
[Detailed explanation]

II. [Competitor 2]
[Detailed explanation]

III. [Competitor 3]
[Detailed explanation]

IV. [Competitor 4 if relevant]
[Detailed explanation]

**Challenges faced by {company_name} in [Business Line Name]:**
- [Challenge 1]
- [Challenge 2]
- [Challenge 3]

**Strategic AI Reinvention and ROI:**
- **Strategic Theme:** [One line]
- **Tangible Value / ROI:** [Quantified business impact]
- **5 Daily AI-Driven Nudges:**
  1. [Nudge 1]
  2. [Nudge 2]
  3. [Nudge 3]
  4. [Nudge 4]
  5. [Nudge 5]
- **What to do to deliver:** [Execution recommendation]

After ALL business lines, output this heading exactly:
SUMMARY OF QUANTIFIED IMPACT (ANNUAL)

Then provide this exact table:
| Business Unit | Primary Hard ROI Metric | Estimated Annual Dollar Impact |
|---|---|---:|

FINAL RULES
- The answer must contain 3 to 5 business lines only
- Every business line must follow the exact same structure
- Do not output any text after the final summary table
"""

BI_STRUCTURE_FIX_PROMPT = """
You are a strict formatting and structure editor.

Your task is to rewrite the Business Intelligence content below so that it follows the REQUIRED STRUCTURE exactly.

Do not change the company-specific substance unless necessary.
Do not add new fluff.
Do not shorten depth.
Do not remove quantified insights.
Only fix structure, labels, order, and consistency.

REQUIRED STRUCTURE:
1. Heading: TOP BUSINESS LINES AND REVENUE
2. Revenue table:
| Business Line | Estimated Revenue | Notes |
|---|---:|---|
3. For each business line:
## [Business Line Name]

**Market Leaders:**
**What "Good" Looks Like Today in [Company]:**
**What "Good" Looks Like Today Across Market Leaders:**
**Challenges faced by [Company] in [Business Line Name]:**
**Strategic AI Reinvention and ROI:**
4. Final heading: SUMMARY OF QUANTIFIED IMPACT (ANNUAL)
5. Final summary table:
| Business Unit | Primary Hard ROI Metric | Estimated Annual Dollar Impact |
|---|---|---:|

COMPANY NAME:
{company_name}

BUSINESS INTELLIGENCE DRAFT:
{bi_text}
"""

LEADERSHIP_EXTRACTION_PROMPT = """
You are a strict JSON generator.

Return ONLY valid JSON.
NO explanations. NO markdown. NO text outside JSON.

Schema:
{{
  "executives": [
    {{
      "name": "Full Name",
      "title": "Role Title",
      "linkedin": "",
      "type": "CEO | CFO | CMO | CIO | BUSINESS_LINE_HEAD | BUSINESS_LINE_TECH_HEAD | BOD",
      "business_line": "Optional"
    }}
  ]
}}

Rules:
- Always include "executives"
- If no executives found, return: {{"executives": []}}
- Do not invent names
- Use empty string if linkedin is not provided
- Use BUSINESS_LINE_HEAD for business line leaders
- Use BUSINESS_LINE_TECH_HEAD for technology heads of business lines
- Use BOD for board members

LEADERSHIP MAPPING TEXT:
{leadership_text}
"""

STORYLINE_PROMPT = """
You are a senior strategy advisor creating a high-level executive meeting storyline.

TARGET COMPANY:
{company_name}

BUSINESS INTELLIGENCE:
{business_intelligence}

EXECUTIVE PROFILE:
Name: {name}
Title: {title}
LinkedIn: {linkedin}
Type: {exec_type}
Business Line: {business_line}

TASK:
Create a detailed executive storyline customized for this person.

OUTPUT REQUIREMENTS:
- Tailored intro paragraph
- Named meeting storyline theme
- The Hook (Minutes 1-3)
- Proof of Knowledge
- The Pivot (Minutes 4-5)
- Competitive contrast with named peers
- The Close for Action (Minutes 6-7)
- Value proposition with quantified impact
- Detailed meeting structure
- ROI table
"""

FINANCIAL_EXTRACTION_PROMPT = """
You are a strict extraction engine for ADM financial inputs.

The Business Intelligence document below has already been generated.
Do NOT rewrite it. Do NOT summarize it.

Return ONLY valid JSON.
NO markdown.
NO explanations.

Schema:
{{
  "company_facts": {{
    "employee_count": 0,
    "annual_revenue_m": 0,
    "sector": "",
    "legacy_level": "high | moderate | low",
    "scope_preference": "light | medium | heavy"
  }},
  "business_units": [
    {{
      "name": "",
      "estimated_weight_pct": 0
    }}
  ],
  "value_drivers": [
    {{
      "business_unit": "",
      "driver_name": "",
      "revenue_or_cost_base_m": 0,
      "improvement_pct": 0,
      "annual_impact_m": 0,
      "source_logic": ""
    }}
  ],
  "error": ""
}}

Rules:
- annual_revenue_m must be in millions
- sector must be one of: Financial Services, Semiconductor, Media, Telecom, Manufacturing, Healthcare, Retail
- Extract ALL meaningful value drivers from the BI
- estimated_weight_pct across business_units should sum close to 100
- Do not invent fake value drivers just to reach a target count
- Output JSON only
- NEVER output 0 for employee_count, annual_revenue_m, revenue_or_cost_base_m, annual_impact_m, or improvement_pct unless explicitly stated in the BI
- If a company fact is missing, estimate it from the BI and public scale cues
- If a value driver is weak or incomplete, derive the annual impact using the best available base and percentage
- If you cannot support at least 4 value drivers with non-zero annual impact, return an error object instead of zero-filled rows

COMPANY NAME:
{company_name}

BUSINESS INTELLIGENCE:
{business_intelligence}
"""

NUMERIC_CORRECTION_PROMPT = """
You are a numerical consistency editor.

TASK:
Correct only number discrepancies in the ADM text below using the Financial Summary JSON and the exact tables.
Do NOT rewrite style.
Do NOT change structure.
Do NOT shorten text.
Only replace incorrect numbers, percentages, totals, and repeated values.
Do not add zeros.
Do not add $0.0M or 0.0% placeholders.

BUSINESS INTELLIGENCE CONTEXT:
{business_intelligence}

FINANCIAL SUMMARY JSON:
{financial_summary_json}

VERBATIM TABLES:
{financial_tables_text}

ADM TEXT TO CORRECT:
{adm_text}

Return only the corrected ADM text.
"""

ADM_BATCH1_PROMPT = """
You are writing an ADM proposal.

IMPORTANT SOURCE RULES
- The Business Intelligence document below has ALREADY been generated.
- Do NOT regenerate the Business Intelligence.
- Do NOT rewrite or replace the Business Intelligence.
- Use it only as source context.

- The Financial Summary JSON below is the ONLY source of truth for all numbers.
- If a number is needed, copy it exactly.
- Do NOT recalculate.
- Do NOT invent.
- Do NOT change any number.

- The financial tables below are prebuilt.
- Insert them exactly where relevant.
- Do not modify table values.

- If a section does not have sufficient support from the Business Intelligence or Financial Summary, do not invent content and do not insert zero-value placeholders.
- Do not include any table row with 0, 0.0, or $0.0M unless that value is explicitly supported by the Financial Summary JSON.

CLIENT NAME:
{company_name}

BUSINESS INTELLIGENCE DOCUMENT (ALREADY GENERATED - USE AS CONTEXT ONLY):
{business_intelligence}

FINANCIAL SUMMARY JSON:
{financial_summary_json}

VERBATIM FINANCIAL TABLES:
{financial_tables_text}

Write EXACTLY this structure and do not skip headings:

TITLE PAGE

EXECUTIVE SUMMARY
- Strategic Imperative
- Portfolio Facts
- Investment Overview
- Value Proposition

PART 1: APPLICATION PORTFOLIO ANALYSIS
1.1 Application Portfolio Composition & Characteristics
1.2 Technology Stack Distribution
1.3 First Business Unit Deep Dive
1.4 Second Business Unit Deep Dive
1.5 Third Business Unit Deep Dive if applicable
1.6 Fourth Business Unit Deep Dive if applicable

Rules for Batch 1:
- Use the BI for narrative only
- Use the financial summary JSON for every number
- Insert Table 1 in 1.1
- Insert business unit allocation table in 1.1
- Insert technology stack distribution table in 1.2
- For each business unit deep dive, write 3 to 6 systems
- End each business unit with a Quantifiable Impact table using value drivers relevant to that unit
- Do not add extra headings
- Do not add methodology
- Do not add explanatory notes before or after the requested sections

START with:
BATCH 1: Writing Executive Summary and Part 1. All numbers from Financial Summary.

END with:
BATCH 1 complete. Say 'continue' for the next batch.
"""

ADM_CONTINUE_PROMPT = """
You are continuing an ADM proposal.

IMPORTANT SOURCE RULES
- The Business Intelligence document below has ALREADY been generated.
- Do NOT regenerate the Business Intelligence.
- Do NOT rewrite or replace the Business Intelligence.
- Use it only as source context.

- The Financial Summary JSON below is the ONLY source of truth for all numbers.
- If a number is needed, copy it exactly.
- Do NOT recalculate.
- Do NOT invent.
- Do NOT change any number.

- The financial tables below are prebuilt.
- Insert them exactly where relevant.
- Do not modify table values.

- If a section does not have sufficient support from the Business Intelligence or Financial Summary, do not invent content and do not insert zero-value placeholders.
- Do not include any table row with 0, 0.0, or $0.0M unless that value is explicitly supported by the Financial Summary JSON.

CLIENT NAME:
{company_name}

BUSINESS INTELLIGENCE DOCUMENT:
{business_intelligence}

FINANCIAL SUMMARY JSON:
{financial_summary_json}

VERBATIM FINANCIAL TABLES:
{financial_tables_text}

ALREADY GENERATED ADM CONTENT:
{current_adm_text}

TASK
Continue with BATCH {next_batch_number} only.

Required hardcoded structure by batch:

BATCH 2
PART 2: COMPETITIVE BENCHMARKING
Create a separate benchmarking section and comparison table for EACH business unit.

BATCH 3
PART 3: 5-YEAR PARTNERSHIP DEAL STRUCTURE
3.1 Partnership Overview & Commercial Terms
3.2 Year-by-Year Investment & Delivery Roadmap
Write Year 1, Year 2, Year 3, Year 4, Year 5 separately.

BATCH 4
3.3 Detailed Financial Model
- Insert Table 5-Year Investment Profile exactly
- Insert Business Value Creation table
- Insert Return on Investment Analysis table
3.4 Offshore Delivery Model
- Insert blended rate table exactly
- Write 3 delivery centers

BATCH 5
3.5 Governance & Operating Model
3.6 Risk Mitigation
3.7 Transition Approach
3.8 Success Metrics

BATCH 6
PART 4: CONCLUSION
4.1 Competitive Imperative
4.2 Partnership Advantage
4.3 Critical Success Factors
4.4 Next Steps
4.5 Final Investment Summary
APPENDICES
FOOTER

START with:
BATCH {next_batch_number}: Writing requested sections. All numbers from Financial Summary.

END with:
BATCH {next_batch_number} complete. Say 'continue' for the next batch.
"""


# ============================================================
# NUMERICAL AGENT LOGIC
# ============================================================

SECTOR_APP_RATIOS = {
    "Financial Services": 15,
    "Semiconductor": 22,
    "Media": 20,
    "Telecom": 18,
    "Manufacturing": 25,
    "Healthcare": 18,
    "Retail": 22,
}

SECTOR_MAINT_RATIOS = {
    "Financial Services": 0.025,
    "Semiconductor": 0.015,
    "Media": 0.020,
    "Telecom": 0.020,
    "Manufacturing": 0.015,
    "Healthcare": 0.022,
    "Retail": 0.018,
}

LEGACY_MULTIPLIERS = {
    "high": 2.0,
    "moderate": 1.5,
    "low": 1.2,
}


def financial_extract_node(state: AdmFinancialState) -> AdmFinancialState:
    client = st.session_state._gemini_client

    base_prompt = FINANCIAL_EXTRACTION_PROMPT.format(
        company_name=state["company_name"],
        business_intelligence=state["bi_text"],
    )

    raw = client.generate(base_prompt)
    cleaned = clean_json_response(raw)

    try:
        extracted = json.loads(cleaned)
    except Exception as e:
        return {
            "error": f"JSON parsing failed: {str(e)}\nRaw output:\n{cleaned}"
        }

    def extracted_has_bad_zeros(data: Dict[str, Any]) -> bool:
        facts = data.get("company_facts", {})
        if safe_float(facts.get("employee_count")) <= 0:
            return True
        if safe_float(facts.get("annual_revenue_m")) <= 0:
            return True

        drivers = data.get("value_drivers", [])
        non_zero_drivers = 0
        for d in drivers:
            if safe_float(d.get("annual_impact_m")) > 0:
                non_zero_drivers += 1
        return non_zero_drivers < 4

    if extracted.get("error") or extracted_has_bad_zeros(extracted):
        retry_prompt = base_prompt + """

STRICT RETRY INSTRUCTION:
- Your previous output contained zero, missing, or unusable financial values.
- Recalculate and re-estimate all missing company facts.
- Return at least 4 non-zero value drivers.
- Do NOT output any driver with annual_impact_m = 0.
- Do NOT output employee_count = 0 or annual_revenue_m = 0.
"""
        raw = client.generate(retry_prompt)
        cleaned = clean_json_response(raw)

        try:
            extracted = json.loads(cleaned)
        except Exception as e:
            return {
                "error": f"Retry JSON parsing failed: {str(e)}\nRaw output:\n{cleaned}"
            }

    if extracted.get("error"):
        return {
            "error": f"Extraction failed: {extracted['error']}"
        }

    return {"extracted_inputs": extracted}


def build_business_unit_allocations(
    business_units: List[Dict[str, Any]],
    value_drivers: List[Dict[str, Any]],
    app_count: int,
    maintenance_m: float,
    tech_debt_m: float,
) -> List[Dict[str, Any]]:
    impacts_by_unit: Dict[str, float] = {}
    for vd in value_drivers:
        unit = vd.get("business_unit", "Other")
        impacts_by_unit[unit] = impacts_by_unit.get(unit, 0.0) + safe_float(vd.get("annual_impact_m"))

    allocations: List[Dict[str, Any]] = []
    total_weight = 0.0

    for bu in business_units:
        name = bu.get("name", "Business Unit")
        weight = safe_float(bu.get("estimated_weight_pct"))
        if weight <= 0 and name in impacts_by_unit:
            weight = impacts_by_unit[name]
        if weight <= 0:
            weight = 1.0
        allocations.append({"name": name, "weight": weight})
        total_weight += weight

    if total_weight <= 0:
        total_weight = float(len(allocations)) if allocations else 1.0

    for item in allocations:
        pct = item["weight"] / total_weight
        item["portfolio_pct"] = round1(pct * 100)
        item["app_count"] = max(1, int(round(app_count * pct)))
        item["annual_maintenance_m"] = round1(maintenance_m * pct)
        item["modernization_backlog_m"] = round1(tech_debt_m * pct)

    if allocations:
        app_diff = app_count - sum(x["app_count"] for x in allocations)
        allocations[-1]["app_count"] += app_diff

        maint_diff = round1(maintenance_m - sum(x["annual_maintenance_m"] for x in allocations))
        allocations[-1]["annual_maintenance_m"] = round1(allocations[-1]["annual_maintenance_m"] + maint_diff)

        debt_diff = round1(tech_debt_m - sum(x["modernization_backlog_m"] for x in allocations))
        allocations[-1]["modernization_backlog_m"] = round1(allocations[-1]["modernization_backlog_m"] + debt_diff)

        pct_diff = round1(100.0 - sum(x["portfolio_pct"] for x in allocations))
        allocations[-1]["portfolio_pct"] = round1(allocations[-1]["portfolio_pct"] + pct_diff)

    return allocations


def build_blended_rates(sector: str) -> List[Dict[str, Any]]:
    roles = [
        ("Enterprise Architect", 210, 82, "40/60"),
        ("Business Analyst", 155, 62, "35/65"),
        ("Senior Engineer", 175, 68, "25/75"),
        ("Cloud Engineer", 180, 72, "25/75"),
        ("Data Engineer", 160, 60, "20/80"),
        ("Full Stack Developer", 145, 55, "20/80"),
        ("QA Automation Engineer", 120, 45, "10/90"),
        ("Legacy Support Specialist", 110, 40, "10/90"),
    ]

    if sector == "Healthcare":
        roles[2] = ("Senior Platform Engineer", 180, 70, "25/75")
    elif sector == "Financial Services":
        roles[1] = ("Business Analyst", 165, 65, "35/65")
        roles[2] = ("Senior Engineering Lead", 185, 72, "25/75")
    elif sector == "Manufacturing":
        roles[3] = ("Cloud / IoT Engineer", 175, 68, "25/75")

    out: List[Dict[str, Any]] = []
    for role, us, india, mix in roles:
        us_pct, india_pct = mix.split("/")
        us_share = safe_float(us_pct) / 100.0
        india_share = safe_float(india_pct) / 100.0
        blended = round1((us * us_share) + (india * india_share))
        savings_pct = round1(((us - blended) / us) * 100)
        out.append(
            {
                "role": role,
                "us_k": us,
                "india_k": india,
                "mix": mix,
                "blended_k": blended,
                "savings_pct": savings_pct,
                "formula": f"({us} x {us_share:.2f}) + ({india} x {india_share:.2f}) = {blended}",
            }
        )
    return out


def financial_compute_node(state: AdmFinancialState) -> AdmFinancialState:
    extracted = state["extracted_inputs"]

    facts = extracted.get("company_facts", {})
    business_units = extracted.get("business_units", [])
    value_drivers_raw = extracted.get("value_drivers", [])

    employee_count = safe_int(facts.get("employee_count"), 0)
    annual_revenue_m = safe_float(facts.get("annual_revenue_m"), 0.0)
    sector = facts.get("sector", "Manufacturing")

    if employee_count <= 0:
        raise ValueError("Employee count was not extracted with a valid non-zero value.")
    if annual_revenue_m <= 0:
        raise ValueError("Annual revenue was not extracted with a valid non-zero value.")
    if sector not in SECTOR_APP_RATIOS:
        sector = "Manufacturing"

    legacy_level = str(facts.get("legacy_level", "moderate")).lower().strip()
    if legacy_level not in LEGACY_MULTIPLIERS:
        legacy_level = "moderate"

    scope_preference = str(facts.get("scope_preference", "medium")).lower().strip()
    if scope_preference not in {"light", "medium", "heavy"}:
        scope_preference = "medium"

    sector_ratio = SECTOR_APP_RATIOS[sector]
    maintenance_ratio = SECTOR_MAINT_RATIOS[sector]
    legacy_multiplier = LEGACY_MULTIPLIERS[legacy_level]

    lower_bound = employee_count / 30
    upper_bound = employee_count / 12
    raw_app_count = employee_count / sector_ratio
    used_ratio = sector_ratio
    if raw_app_count < lower_bound or raw_app_count > upper_bound:
        used_ratio = 20
        raw_app_count = employee_count / used_ratio

    app_count = int(round(raw_app_count))
    annual_maintenance_m = round1(annual_revenue_m * maintenance_ratio)
    tech_debt_m = round1(annual_maintenance_m * legacy_multiplier)

    value_drivers: List[Dict[str, Any]] = []
    for idx, vd in enumerate(value_drivers_raw, start=1):
        bu = vd.get("business_unit", "Business Unit")
        name = vd.get("driver_name", f"Driver {idx}")
        base_m = safe_float(vd.get("revenue_or_cost_base_m"))
        improvement_pct = safe_float(vd.get("improvement_pct"))
        if improvement_pct > 1:
            improvement_decimal = improvement_pct / 100.0
            improvement_pct_display = improvement_pct
        else:
            improvement_decimal = improvement_pct
            improvement_pct_display = improvement_pct * 100.0

        extracted_annual_impact = safe_float(vd.get("annual_impact_m"))
        computed_annual_impact = round1(base_m * improvement_decimal) if base_m > 0 and improvement_decimal > 0 else 0.0
        annual_impact_m = computed_annual_impact if computed_annual_impact > 0 else extracted_annual_impact
        annual_impact_m = round1(annual_impact_m)

        if annual_impact_m <= 0:
            continue

        if base_m > 0 and improvement_decimal > 0:
            formula = f"{round1(base_m)} x {round1(improvement_pct_display)}% = {annual_impact_m}"
        else:
            formula = vd.get("source_logic", "") or f"Annual impact estimated at {annual_impact_m}"

        value_drivers.append(
            {
                "business_unit": bu,
                "driver_name": name,
                "revenue_or_cost_base_m": round1(base_m),
                "improvement_pct": round1(improvement_pct_display),
                "annual_impact_m": annual_impact_m,
                "formula": formula,
            }
        )

    if len(value_drivers) < 4:
        raise ValueError(
            "Not enough valid value drivers were extracted from the Business Intelligence. "
            "Please improve the BI output before generating the ADM."
        )

    total_annual_value_m = round1(sum(v["annual_impact_m"] for v in value_drivers))
    five_year_value_m = round1(total_annual_value_m * 3.15)

    base_multiplier_map = {"light": 3.5, "medium": 4.0, "heavy": 4.5}
    multiplier = base_multiplier_map[scope_preference]

    def roi_for(mult: float) -> float:
        inv = annual_maintenance_m * mult
        if inv <= 0:
            return 0.0
        return round1(((five_year_value_m - inv) / inv) * 100)

    roi_pct = roi_for(multiplier)
    while roi_pct < 150.0 and multiplier > 2.0:
        multiplier -= 0.5
        roi_pct = roi_for(multiplier)

    while roi_pct > 300.0 and multiplier < 6.0:
        multiplier += 0.5
        roi_pct = roi_for(multiplier)

    investment_m = round1(annual_maintenance_m * multiplier)

    cost_savings = {
        "y1_m": round1(annual_maintenance_m * 0.12),
        "y2_m": round1(annual_maintenance_m * 0.22),
        "y3_m": round1(annual_maintenance_m * 0.30),
        "y4_m": round1(annual_maintenance_m * 0.35),
        "y5_m": round1(annual_maintenance_m * 0.38),
    }
    cost_savings["five_year_total_m"] = round1(
        cost_savings["y1_m"] + cost_savings["y2_m"] + cost_savings["y3_m"] + cost_savings["y4_m"] + cost_savings["y5_m"]
    )

    blended_rates = build_blended_rates(sector)

    legacy_total = round1(investment_m * 0.595)
    modernization_total = round1(investment_m * 0.155)
    digital_total = round1(investment_m * 0.19)
    innovation_total = round1(investment_m * 0.06)

    legacy_y = allocate_component_total(legacy_total, [0.24, 0.22, 0.20, 0.18, 0.16])
    modernization_y = allocate_component_total(modernization_total, [0.14, 0.22, 0.30, 0.20, 0.14])
    digital_y = allocate_component_total(digital_total, [0.14, 0.22, 0.28, 0.22, 0.14])
    innovation_y = allocate_component_total(innovation_total, [0.10, 0.15, 0.20, 0.25, 0.30])

    total_y = [
        round1(legacy_y[i] + modernization_y[i] + digital_y[i] + innovation_y[i])
        for i in range(5)
    ]
    total_y[-1] = round1(investment_m - sum(total_y[:-1]))

    partner_y = [round1(y * 0.42) for y in total_y]
    client_y = [round1(y * 0.58) for y in total_y]
    partner_total = round1(sum(partner_y))
    client_total = round1(sum(client_y))

    partner_margin_low_m = round1(partner_total * 0.18)
    partner_margin_high_m = round1(partner_total * 0.22)

    business_unit_allocations = build_business_unit_allocations(
        business_units=business_units if business_units else [{"name": "Business Unit 1", "estimated_weight_pct": 100}],
        value_drivers=value_drivers,
        app_count=app_count,
        maintenance_m=annual_maintenance_m,
        tech_debt_m=tech_debt_m,
    )

    payback_years = find_payback_years(investment_m, total_annual_value_m)
    annualized_return_pct = round1((roi_pct / 5.0) if roi_pct else 0.0)

    financial_summary = {
        "base_data": {
            "employee_count": employee_count,
            "annual_revenue_m": round1(annual_revenue_m),
            "sector": sector,
            "sector_app_ratio_used": used_ratio,
            "sector_maintenance_ratio_pct": round1(maintenance_ratio * 100),
            "legacy_level": legacy_level,
            "legacy_multiplier": legacy_multiplier,
            "app_count": app_count,
            "annual_maintenance_m": annual_maintenance_m,
            "tech_debt_m": tech_debt_m,
        },
        "business_unit_allocations": business_unit_allocations,
        "value_drivers": value_drivers,
        "total_annual_value_m": total_annual_value_m,
        "five_year_value_m": five_year_value_m,
        "investment_m": investment_m,
        "investment_multiplier_used": multiplier,
        "roi_pct": roi_pct,
        "payback_years": payback_years,
        "annualized_return_pct": annualized_return_pct,
        "cost_savings": cost_savings,
        "blended_rates": blended_rates,
        "investment_schedule": {
            "legacy_total_m": legacy_total,
            "modernization_total_m": modernization_total,
            "digital_total_m": digital_total,
            "innovation_total_m": innovation_total,
            "legacy_y_m": legacy_y,
            "modernization_y_m": modernization_y,
            "digital_y_m": digital_y,
            "innovation_y_m": innovation_y,
            "total_y_m": total_y,
        },
        "partner_split": {
            "partner_y_m": partner_y,
            "client_y_m": client_y,
            "partner_total_m": partner_total,
            "client_total_m": client_total,
            "partner_margin_low_m": partner_margin_low_m,
            "partner_margin_high_m": partner_margin_high_m,
        },
    }

    return {"financial_summary": financial_summary}


def financial_validate_node(state: AdmFinancialState) -> AdmFinancialState:
    fs = state["financial_summary"]
    schedule = fs["investment_schedule"]
    partner = fs["partner_split"]

    errors: List[str] = []

    if len(fs["value_drivers"]) < 4:
        errors.append("Less than 4 value drivers found.")

    if round1(sum(v["annual_impact_m"] for v in fs["value_drivers"])) != round1(fs["total_annual_value_m"]):
        errors.append("Total annual value does not match sum of value drivers.")

    if round1(sum(schedule["legacy_y_m"])) != round1(schedule["legacy_total_m"]):
        errors.append("Legacy schedule does not sum correctly.")
    if round1(sum(schedule["modernization_y_m"])) != round1(schedule["modernization_total_m"]):
        errors.append("Modernization schedule does not sum correctly.")
    if round1(sum(schedule["digital_y_m"])) != round1(schedule["digital_total_m"]):
        errors.append("Digital schedule does not sum correctly.")
    if round1(sum(schedule["innovation_y_m"])) != round1(schedule["innovation_total_m"]):
        errors.append("Innovation schedule does not sum correctly.")
    if round1(sum(schedule["total_y_m"])) != round1(fs["investment_m"]):
        errors.append("Total yearly schedule does not equal total investment.")

    if round1(sum(partner["partner_y_m"])) != round1(partner["partner_total_m"]):
        errors.append("Partner yearly split does not sum correctly.")
    if round1(sum(partner["client_y_m"])) != round1(partner["client_total_m"]):
        errors.append("Client yearly split does not sum correctly.")

    return {"error": " | ".join(errors)} if errors else {"error": ""}


def run_financial_graph(client: GeminiClient, company_name: str, bi_text: str) -> Dict[str, Any]:
    if StateGraph is None:
        raise ImportError("langgraph is not installed. Run: python -m pip install -U langgraph langchain-core")

    st.session_state._gemini_client = client

    workflow = StateGraph(AdmFinancialState)
    workflow.add_node("extract_inputs", financial_extract_node)
    workflow.add_node("compute_financials", financial_compute_node)
    workflow.add_node("validate_financials", financial_validate_node)

    workflow.set_entry_point("extract_inputs")
    workflow.add_edge("extract_inputs", "compute_financials")
    workflow.add_edge("compute_financials", "validate_financials")
    workflow.add_edge("validate_financials", END)

    graph = workflow.compile()
    result = graph.invoke({"company_name": company_name, "bi_text": bi_text})

    if result.get("error"):
        raise ValueError(result["error"])

    return result["financial_summary"]


# ============================================================
# TABLE BUILDERS
# ============================================================

def build_table_1_text(fs: Dict[str, Any]) -> str:
    b = fs["base_data"]
    return "\n".join([
        "| Item | Value | Source |",
        "|---|---:|---|",
        f"| Employee Count | {b['employee_count']:,} | Extracted from BI/public basis |",
        f"| Annual Revenue | {mfmt(b['annual_revenue_m'])} | Extracted from BI/public basis |",
        f"| Sector | {b['sector']} | Classification |",
        "",
        "| Metric | Formula | Result |",
        "|---|---|---:|",
        f"| App Count | {b['employee_count']:,} / {b['sector_app_ratio_used']} | {b['app_count']:,} apps |",
        f"| Annual Maintenance | {mfmt(b['annual_revenue_m'])} x {b['sector_maintenance_ratio_pct']}% | {mfmt(b['annual_maintenance_m'])} |",
        f"| Tech Debt | {mfmt(b['annual_maintenance_m'])} x {b['legacy_multiplier']} | {mfmt(b['tech_debt_m'])} |",
    ])


def build_business_unit_allocation_table(fs: Dict[str, Any]) -> str:
    rows = [
        "| Business Unit | App Count | % of Portfolio | Annual Maintenance | Modernization Backlog |",
        "|---|---:|---:|---:|---:|",
    ]
    for bu in fs["business_unit_allocations"]:
        rows.append(
            f"| {bu['name']} | {bu['app_count']} | {bu['portfolio_pct']}% | "
            f"{mfmt(bu['annual_maintenance_m'])} | {mfmt(bu['modernization_backlog_m'])} |"
        )
    rows.append(
        f"| TOTAL | {fs['base_data']['app_count']} | 100% | "
        f"{mfmt(fs['base_data']['annual_maintenance_m'])} | {mfmt(fs['base_data']['tech_debt_m'])} |"
    )
    return "\n".join(rows)


def build_technology_stack_distribution_table(fs: Dict[str, Any]) -> str:
    total_apps = fs["base_data"]["app_count"]
    sector = fs["base_data"]["sector"]

    if sector == "Healthcare":
        cats = [
            ("Legacy Clinical / Commercial Platforms", 0.28, "15-22", "Critical Shortage", "High"),
            ("Mid-Life ERP / Supply / Trial Platforms", 0.26, "10-15", "Declining", "Medium-High"),
            ("Modern Data / AI / Patient Platforms", 0.18, "3-8", "High Demand", "Low"),
            ("Regulatory / Quality / Manufacturing Systems", 0.18, "8-12", "Available", "Medium"),
            ("SaaS / Enterprise Support Systems", 0.10, "3-5", "Vendor Managed", "Low-Medium"),
        ]
    elif sector == "Financial Services":
        cats = [
            ("Legacy Core Banking / Ledger Platforms", 0.30, "15-25", "Critical Shortage", "High"),
            ("Mid-Life CRM / Ops / Reporting Systems", 0.24, "10-15", "Declining", "Medium-High"),
            ("Modern Cloud / API / Digital Platforms", 0.18, "3-8", "High Demand", "Low"),
            ("Risk / Compliance / Treasury Systems", 0.18, "8-12", "Available", "Medium"),
            ("SaaS / Enterprise Support Systems", 0.10, "3-5", "Vendor Managed", "Low-Medium"),
        ]
    elif sector == "Manufacturing":
        cats = [
            ("Legacy Manufacturing / Dealer Platforms", 0.28, "15-25", "Critical Shortage", "High"),
            ("Mid-Life ERP / MES / Supply Systems", 0.30, "10-15", "Declining", "Medium-High"),
            ("Modern Cloud / IoT / Connected Platforms", 0.14, "3-8", "High Demand", "Low"),
            ("Engineering / Product / Quality Systems", 0.18, "8-12", "Available", "Medium"),
            ("SaaS / Enterprise Support Systems", 0.10, "3-5", "Vendor Managed", "Low-Medium"),
        ]
    else:
        cats = [
            ("Legacy Core Platforms", 0.28, "15-25", "Critical Shortage", "High"),
            ("Mid-Life Operational Platforms", 0.28, "10-15", "Declining", "Medium-High"),
            ("Modern Cloud / Digital Platforms", 0.18, "3-8", "High Demand", "Low"),
            ("Industry-Specific Systems", 0.16, "8-12", "Available", "Medium"),
            ("SaaS / Enterprise Support Systems", 0.10, "3-5", "Vendor Managed", "Low-Medium"),
        ]

    counts = [int(round(total_apps * c[1])) for c in cats]
    counts[-1] += total_apps - sum(counts)

    rows = [
        "| Technology Category | App Count | Average Age (Years) | Skills Availability | Risk Level |",
        "|---|---:|---|---|---|",
    ]
    for idx, c in enumerate(cats):
        rows.append(f"| {c[0]} | {counts[idx]} | {c[2]} | {c[3]} | {c[4]} |")
    return "\n".join(rows)


def build_table_2_text(fs: Dict[str, Any]) -> str:
    rows = [
        "| # | Business Unit | Driver Name | Revenue/Cost Base | Improvement % | Annual Impact | Full Formula |",
        "|---|---|---|---:|---:|---:|---|",
    ]
    for i, vd in enumerate(fs["value_drivers"], start=1):
        rows.append(
            f"| {i} | {vd['business_unit']} | {vd['driver_name']} | {mfmt(vd['revenue_or_cost_base_m'])} | "
            f"{vd['improvement_pct']}% | {mfmt(vd['annual_impact_m'])} | {vd['formula']} |"
        )
    rows.append(
        f"| TOTAL ANNUAL VALUE |  |  |  |  | {mfmt(fs['total_annual_value_m'])} | Sum of all drivers |"
    )
    rows.append("")
    rows.append(f"5-Year Value = {mfmt(fs['total_annual_value_m'])} x 3.15 = {mfmt(fs['five_year_value_m'])}")
    rows.append(f"Final Investment = {mfmt(fs['investment_m'])} | ROI = {pfmt(fs['roi_pct'])}")
    return "\n".join(rows)


def build_table_3_text(fs: Dict[str, Any]) -> str:
    c = fs["cost_savings"]
    b = fs["base_data"]["annual_maintenance_m"]
    return "\n".join([
        "| Year | % | Savings | Formula |",
        "|---|---:|---:|---|",
        f"| 1 | 12% | {mfmt(c['y1_m'])} | {mfmt(b)} x 0.12 |",
        f"| 2 | 22% | {mfmt(c['y2_m'])} | {mfmt(b)} x 0.22 |",
        f"| 3 | 30% | {mfmt(c['y3_m'])} | {mfmt(b)} x 0.30 |",
        f"| 4 | 35% | {mfmt(c['y4_m'])} | {mfmt(b)} x 0.35 |",
        f"| 5 | 38% | {mfmt(c['y5_m'])} | {mfmt(b)} x 0.38 |",
        f"| 5-Year Total |  | {mfmt(c['five_year_total_m'])} | Sum |",
    ])


def build_table_4_text(fs: Dict[str, Any]) -> str:
    rows = [
        "| Role | US ($K) | India ($K) | Mix (US/India) | Blended ($K) | Savings% | Formula |",
        "|---|---:|---:|---|---:|---:|---|",
    ]
    for r in fs["blended_rates"]:
        rows.append(
            f"| {r['role']} | {r['us_k']} | {r['india_k']} | {r['mix']} | {r['blended_k']} | {r['savings_pct']}% | {r['formula']} |"
        )
    return "\n".join(rows)


def build_table_5_text(fs: Dict[str, Any]) -> str:
    s = fs["investment_schedule"]
    return "\n".join([
        "| Component | Y1 | Y2 | Y3 | Y4 | Y5 | Total |",
        "|---|---:|---:|---:|---:|---:|---:|",
        f"| Legacy (59.5%) | {mfmt(s['legacy_y_m'][0])} | {mfmt(s['legacy_y_m'][1])} | {mfmt(s['legacy_y_m'][2])} | {mfmt(s['legacy_y_m'][3])} | {mfmt(s['legacy_y_m'][4])} | {mfmt(s['legacy_total_m'])} |",
        f"| Modernization (15.5%) | {mfmt(s['modernization_y_m'][0])} | {mfmt(s['modernization_y_m'][1])} | {mfmt(s['modernization_y_m'][2])} | {mfmt(s['modernization_y_m'][3])} | {mfmt(s['modernization_y_m'][4])} | {mfmt(s['modernization_total_m'])} |",
        f"| Digital Pods (19%) | {mfmt(s['digital_y_m'][0])} | {mfmt(s['digital_y_m'][1])} | {mfmt(s['digital_y_m'][2])} | {mfmt(s['digital_y_m'][3])} | {mfmt(s['digital_y_m'][4])} | {mfmt(s['digital_total_m'])} |",
        f"| Innovation (6%) | {mfmt(s['innovation_y_m'][0])} | {mfmt(s['innovation_y_m'][1])} | {mfmt(s['innovation_y_m'][2])} | {mfmt(s['innovation_y_m'][3])} | {mfmt(s['innovation_y_m'][4])} | {mfmt(s['innovation_total_m'])} |",
        f"| TOTAL | {mfmt(s['total_y_m'][0])} | {mfmt(s['total_y_m'][1])} | {mfmt(s['total_y_m'][2])} | {mfmt(s['total_y_m'][3])} | {mfmt(s['total_y_m'][4])} | {mfmt(fs['investment_m'])} |",
    ])


def build_roi_table_text(fs: Dict[str, Any]) -> str:
    p = fs["partner_split"]
    return "\n".join([
        "| Metric | Value |",
        "|---|---:|",
        f"| Total 5-Year Investment | {mfmt(fs['investment_m'])} |",
        f"| Total 5-Year Value | {mfmt(fs['five_year_value_m'])} |",
        f"| Net Value Created | {mfmt(fs['five_year_value_m'] - fs['investment_m'])} |",
        f"| ROI | {pfmt(fs['roi_pct'])} |",
        f"| Payback Period | {fs['payback_years']} years |",
        f"| Annualized Return | {pfmt(fs['annualized_return_pct'])} |",
        f"| Partner Revenue | {mfmt(p['partner_total_m'])} |",
        f"| Partner Margin Range | {mfmt(p['partner_margin_low_m'])} to {mfmt(p['partner_margin_high_m'])} |",
    ])


def build_business_value_creation_table(fs: Dict[str, Any]) -> str:
    unit_names = [u["name"] for u in fs["business_unit_allocations"]]
    if len(unit_names) < 2:
        unit_names = unit_names + ["Business Unit 2"]

    rev_map: Dict[str, float] = {u: 0.0 for u in unit_names}
    cost_map: Dict[str, float] = {u: 0.0 for u in unit_names}
    risk_map: Dict[str, float] = {u: 0.0 for u in unit_names}
    retain_map: Dict[str, float] = {u: 0.0 for u in unit_names}

    for vd in fs["value_drivers"]:
        unit = vd["business_unit"] if vd["business_unit"] in rev_map else unit_names[0]
        name = vd["driver_name"].lower()
        annual = vd["annual_impact_m"] * 3.15
        if any(k in name for k in ["retention", "churn", "renewal", "retained"]):
            retain_map[unit] += annual
        elif any(k in name for k in ["risk", "delinquency", "compliance", "fraud"]):
            risk_map[unit] += annual
        elif any(k in name for k in ["cost", "efficiency", "savings", "productivity", "cycle"]):
            cost_map[unit] += annual
        else:
            rev_map[unit] += annual

    header = f"| Value Driver | {unit_names[0]} | {unit_names[1]} | Total |"
    sep = "|---|---:|---:|---:|"

    def total_of(mp: Dict[str, float]) -> float:
        return sum(mp.values())

    def row(label: str, mp: Dict[str, float]) -> str:
        return (
            f"| {label} | {mfmt_or_na(mp.get(unit_names[0], 0.0))} | "
            f"{mfmt_or_na(mp.get(unit_names[1], 0.0))} | "
            f"{mfmt_or_na(total_of(mp))} |"
        )

    rows = []
    if total_of(rev_map) > 0:
        rows.append(row("Revenue Growth", rev_map))
    if total_of(cost_map) > 0:
        rows.append(row("Cost Reduction", cost_map))
    if total_of(risk_map) > 0:
        rows.append(row("Risk Mitigation", risk_map))
    if total_of(retain_map) > 0:
        rows.append(row("Asset Retention", retain_map))

    total_map = {
        unit_names[0]: rev_map.get(unit_names[0], 0.0) + cost_map.get(unit_names[0], 0.0) + risk_map.get(unit_names[0], 0.0) + retain_map.get(unit_names[0], 0.0),
        unit_names[1]: rev_map.get(unit_names[1], 0.0) + cost_map.get(unit_names[1], 0.0) + risk_map.get(unit_names[1], 0.0) + retain_map.get(unit_names[1], 0.0),
    }

    rows.append(row("TOTAL VALUE", total_map))

    return "\n".join([header, sep] + rows)


def build_all_financial_tables_text(fs: Dict[str, Any]) -> str:
    chunks = [
        "TABLE 1: BASE DATA\n" + build_table_1_text(fs),
        "BUSINESS UNIT ALLOCATION TABLE\n" + build_business_unit_allocation_table(fs),
        "TECHNOLOGY STACK DISTRIBUTION TABLE\n" + build_technology_stack_distribution_table(fs),
        "TABLE 2: VALUE AND INVESTMENT\n" + build_table_2_text(fs),
        "TABLE 3: COST SAVINGS\n" + build_table_3_text(fs),
        "TABLE 4: BLENDED RATES\n" + build_table_4_text(fs),
        "TABLE 5: INVESTMENT SCHEDULE\n" + build_table_5_text(fs),
        "BUSINESS VALUE CREATION TABLE\n" + build_business_value_creation_table(fs),
        "ROI TABLE\n" + build_roi_table_text(fs),
    ]
    return "\n\n".join(chunks)


def render_financial_summary_text(company_name: str, fs: Dict[str, Any]) -> str:
    text = [
        f"ADM Financial Summary for {company_name}",
        "",
        build_all_financial_tables_text(fs),
        "",
        "Financial summary complete. All checks passed.",
    ]
    return "\n".join(text)


# ============================================================
# GENERATION FUNCTIONS
# ============================================================

def extract_leadership_json(client: GeminiClient, leadership_text: str) -> str:
    prompt = LEADERSHIP_EXTRACTION_PROMPT.format(leadership_text=leadership_text)
    return client.generate(prompt)


def generate_bi(client: GeminiClient, company_name: str) -> str:
    raw_bi = client.generate(BI_PROMPT.format(company_name=company_name))
    fixed_bi = client.generate(
        BI_STRUCTURE_FIX_PROMPT.format(
            company_name=company_name,
            bi_text=raw_bi
        )
    )
    return fixed_bi


def generate_storylines(
    client: GeminiClient,
    profiles: List[ExecProfile],
    company_name: str,
    business_intelligence: str,
) -> Dict[str, str]:
    results: Dict[str, str] = {}
    progress = st.progress(0)
    total = max(len(profiles), 1)

    for idx, profile in enumerate(profiles, start=1):
        prompt = STORYLINE_PROMPT.format(
            company_name=company_name,
            business_intelligence=business_intelligence,
            name=profile.name,
            title=profile.title,
            linkedin=profile.linkedin,
            exec_type=profile.type,
            business_line=profile.business_line or "N/A",
        )
        results[f"{profile.type}__{profile.name}"] = client.generate(prompt)
        progress.progress(idx / total)

    return results


def run_numeric_correction(
    client: GeminiClient,
    business_intelligence: str,
    financial_summary: Dict[str, Any],
    financial_tables_text: str,
    adm_text: str,
) -> str:
    prompt = NUMERIC_CORRECTION_PROMPT.format(
        business_intelligence=business_intelligence,
        financial_summary_json=json.dumps(financial_summary, indent=2),
        financial_tables_text=financial_tables_text,
        adm_text=adm_text,
    )
    corrected = client.generate(prompt)
    final_text = corrected if corrected.strip() else adm_text

    if contains_bad_zero_values(final_text):
        retry_prompt = prompt + """

STRICT RETRY:
- The corrected ADM still contains invalid zero values.
- Replace every unsupported zero with the correct number from the Financial Summary JSON.
- If a section is unsupported, remove the zero-filled row rather than leaving $0.0M or 0.0%.
"""
        final_text = client.generate(retry_prompt).strip()

    assert_no_bad_zero_values(final_text, "ADM output")
    return final_text


def generate_adm_batch1(
    client: GeminiClient,
    company_name: str,
    business_intelligence: str,
    financial_summary: Dict[str, Any],
    financial_tables_text: str,
) -> str:
    prompt = ADM_BATCH1_PROMPT.format(
        company_name=company_name,
        business_intelligence=business_intelligence,
        financial_summary_json=json.dumps(financial_summary, indent=2),
        financial_tables_text=financial_tables_text,
    )
    raw = client.generate(prompt)
    corrected = run_numeric_correction(
        client=client,
        business_intelligence=business_intelligence,
        financial_summary=financial_summary,
        financial_tables_text=financial_tables_text,
        adm_text=raw,
    )
    assert_no_bad_zero_values(corrected, "ADM Batch 1")
    return corrected


def generate_adm_next_batch(
    client: GeminiClient,
    company_name: str,
    business_intelligence: str,
    financial_summary: Dict[str, Any],
    financial_tables_text: str,
    current_adm_text: str,
    next_batch_number: int,
) -> str:
    prompt = ADM_CONTINUE_PROMPT.format(
        company_name=company_name,
        business_intelligence=business_intelligence,
        financial_summary_json=json.dumps(financial_summary, indent=2),
        financial_tables_text=financial_tables_text,
        current_adm_text=current_adm_text,
        next_batch_number=next_batch_number,
    )
    raw = client.generate(prompt)
    corrected = run_numeric_correction(
        client=client,
        business_intelligence=business_intelligence,
        financial_summary=financial_summary,
        financial_tables_text=financial_tables_text,
        adm_text=raw,
    )
    assert_no_bad_zero_values(corrected, f"ADM Batch {next_batch_number}")
    return corrected


# ============================================================
# SESSION STATE
# ============================================================

if "leadership_json" not in st.session_state:
    st.session_state.leadership_json = ""

if "bi_text" not in st.session_state:
    st.session_state.bi_text = ""

if "storylines" not in st.session_state:
    st.session_state.storylines = {}

if "financial_summary" not in st.session_state:
    st.session_state.financial_summary = None

if "financial_summary_text" not in st.session_state:
    st.session_state.financial_summary_text = ""

if "financial_tables_text" not in st.session_state:
    st.session_state.financial_tables_text = ""

if "adm_text" not in st.session_state:
    st.session_state.adm_text = ""

if "adm_batch" not in st.session_state:
    st.session_state.adm_batch = 0


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("Configuration")

api_key = HARDCODED_GEMINI_API_KEY.strip()
if not api_key:
    api_key = st.secrets.get("GEMINI_API_KEY", "")
if not api_key:
    api_key = st.sidebar.text_input("Gemini API Key", type="password")

model_name = DEFAULT_MODEL_NAME
st.sidebar.text_input("Model", value=model_name, disabled=True)

company_name = st.sidebar.text_input("Target Company Name")

st.sidebar.markdown("### Leadership Mapping Text")
leadership_text = st.sidebar.text_area(
    "Paste Leadership Mapping Text",
    height=350,
    placeholder="Paste CEO, CFO, CMO, CIO, business line heads, business line tech heads, and board members here..."
)


# ============================================================
# VALIDATION
# ============================================================

def validate_base() -> bool:
    if not api_key:
        st.error("Enter your Gemini API key.")
        return False
    if not company_name:
        st.error("Enter a target company name.")
        return False
    return True


def validate_leadership() -> bool:
    if not leadership_text.strip():
        st.error("Paste the leadership mapping text first.")
        return False
    return True


def validate_bi() -> bool:
    if not st.session_state.bi_text:
        st.error("Generate Business Intelligence first.")
        return False
    return True


def validate_financial_summary() -> bool:
    if not st.session_state.financial_summary:
        st.error("Generate the ADM Financial Summary first.")
        return False
    return True


# ============================================================
# UI
# ============================================================

st.subheader("Inputs")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**Company**")
    st.write(company_name or "Not set")

with col2:
    st.markdown("**Model**")
    st.write(model_name)

st.markdown("**Leadership Mapping Status**")
if leadership_text.strip():
    st.success("Leadership mapping text loaded.")
else:
    st.info("No leadership mapping text pasted yet.")

st.markdown("**ADM Numerical Agent Status**")
if StateGraph is None:
    st.warning("LangGraph not installed yet. Install with: python -m pip install -U langgraph langchain-core")
else:
    st.success("LangGraph available. ADM financial agent ready.")

col_a, col_b, col_c, col_d, col_e = st.columns(5)

with col_a:
    bi_btn = st.button("Generate Business Intelligence", use_container_width=True)

with col_b:
    storyline_btn = st.button("Generate Executive Storylines", use_container_width=True)

with col_c:
    financial_btn = st.button("Generate ADM Financial Summary", use_container_width=True)

with col_d:
    adm_batch1_btn = st.button("Generate ADM Batch 1", use_container_width=True)

with col_e:
    adm_continue_btn = st.button("Continue ADM", use_container_width=True)


# ============================================================
# ACTIONS
# ============================================================

if bi_btn:
    if validate_base():
        try:
            client = GeminiClient(api_key=api_key, model=model_name)
            with st.spinner("Generating Business Intelligence..."):
                st.session_state.bi_text = generate_bi(client, company_name)
            st.success("Business Intelligence generated.")
        except Exception as e:
            st.error(f"BI generation failed: {e}")

if storyline_btn:
    if validate_base() and validate_leadership() and validate_bi():
        try:
            client = GeminiClient(api_key=api_key, model=model_name)

            with st.spinner("Extracting leadership structure from text..."):
                leadership_json = extract_leadership_json(client, leadership_text)
                st.session_state.leadership_json = leadership_json

            st.subheader("Leadership Extraction Debug")
            st.code(leadership_json)

            profiles = parse_exec_profiles_from_json(leadership_json)

            with st.spinner("Generating executive storylines..."):
                st.session_state.storylines = generate_storylines(
                    client=client,
                    profiles=profiles,
                    company_name=company_name,
                    business_intelligence=st.session_state.bi_text,
                )

            st.success(f"Executive storylines generated for {len(profiles)} profiles.")
        except Exception as e:
            st.error(f"Storyline generation failed: {e}")
            if st.session_state.leadership_json:
                st.subheader("Leadership Extraction Debug")
                st.code(st.session_state.leadership_json)

if financial_btn:
    if validate_base() and validate_bi():
        try:
            client = GeminiClient(api_key=api_key, model=model_name)
            with st.spinner("Running LangGraph ADM financial agent..."):
                financial_summary = run_financial_graph(
                    client=client,
                    company_name=company_name,
                    bi_text=st.session_state.bi_text,
                )
                st.session_state.financial_summary = financial_summary
                st.session_state.financial_tables_text = build_all_financial_tables_text(financial_summary)
                st.session_state.financial_summary_text = render_financial_summary_text(company_name, financial_summary)
            st.success("ADM Financial Summary generated and locked as source of truth.")
        except Exception as e:
            st.error(f"Financial summary generation failed: {e}")

if adm_batch1_btn:
    if validate_base() and validate_bi():
        try:
            if not st.session_state.financial_summary:
                client = GeminiClient(api_key=api_key, model=model_name)
                with st.spinner("Financial summary not found. Running LangGraph ADM financial agent first..."):
                    financial_summary = run_financial_graph(
                        client=client,
                        company_name=company_name,
                        bi_text=st.session_state.bi_text,
                    )
                    st.session_state.financial_summary = financial_summary
                    st.session_state.financial_tables_text = build_all_financial_tables_text(financial_summary)
                    st.session_state.financial_summary_text = render_financial_summary_text(company_name, financial_summary)

            client = GeminiClient(api_key=api_key, model=model_name)
            with st.spinner("Generating ADM Batch 1 using existing BI and locked financial summary..."):
                batch1_text = generate_adm_batch1(
                    client=client,
                    company_name=company_name,
                    business_intelligence=st.session_state.bi_text,
                    financial_summary=st.session_state.financial_summary,
                    financial_tables_text=st.session_state.financial_tables_text,
                )
                st.session_state.adm_text = st.session_state.financial_summary_text + "\n\n" + batch1_text
                st.session_state.adm_batch = 1
            st.success("ADM Financial Summary + Batch 1 generated.")
        except Exception as e:
            st.error(f"ADM batch 1 generation failed: {e}")

if adm_continue_btn:
    if validate_base() and validate_bi() and validate_financial_summary():
        if st.session_state.adm_batch <= 0:
            st.error("Generate ADM Batch 1 first.")
        elif st.session_state.adm_batch >= 6:
            st.info("ADM is already complete.")
        else:
            try:
                client = GeminiClient(api_key=api_key, model=model_name)
                next_batch = st.session_state.adm_batch + 1
                with st.spinner(f"Generating ADM Batch {next_batch} using existing BI and locked numbers..."):
                    new_batch_text = generate_adm_next_batch(
                        client=client,
                        company_name=company_name,
                        business_intelligence=st.session_state.bi_text,
                        financial_summary=st.session_state.financial_summary,
                        financial_tables_text=st.session_state.financial_tables_text,
                        current_adm_text=st.session_state.adm_text,
                        next_batch_number=next_batch,
                    )
                    st.session_state.adm_text += "\n\n" + new_batch_text
                    st.session_state.adm_batch = next_batch
                st.success(f"ADM Batch {next_batch} generated.")
            except Exception as e:
                st.error(f"ADM continuation failed: {e}")


# ============================================================
# OUTPUTS
# ============================================================

st.divider()
st.header("Outputs")

tab1, tab2, tab3, tab4 = st.tabs(
    ["Business Intelligence", "Executive Storylines", "ADM Financial Summary", "ADM"]
)

with tab1:
    st.text_area("Business Intelligence Output", value=st.session_state.bi_text, height=500)

    if st.session_state.bi_text:
        st.download_button(
            "Download BI TXT",
            data=st.session_state.bi_text,
            file_name=f"{sanitize_filename(company_name)}_Business_Intelligence.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.download_button(
            "Download BI DOCX",
            data=save_docx_bytes(f"{company_name} Business Intelligence", st.session_state.bi_text),
            file_name=f"{sanitize_filename(company_name)}_Business_Intelligence.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

with tab2:
    if st.session_state.storylines:
        selected = st.selectbox("Select Storyline", list(st.session_state.storylines.keys()))
        selected_text = st.session_state.storylines[selected]

        st.text_area("Storyline Output", value=selected_text, height=500)

        st.download_button(
            "Download Selected Storyline TXT",
            data=selected_text,
            file_name=f"{sanitize_filename(selected)}.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.download_button(
            "Download Selected Storyline DOCX",
            data=save_docx_bytes(selected, selected_text),
            file_name=f"{sanitize_filename(selected)}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

        combined = []
        for k, v in st.session_state.storylines.items():
            combined.append(f"\n{'=' * 80}\n{k}\n{'=' * 80}\n{v}\n")
        combined_text = "\n".join(combined)

        st.download_button(
            "Download All Storylines TXT",
            data=combined_text,
            file_name=f"{sanitize_filename(company_name)}_Executive_Storylines.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.download_button(
            "Download All Storylines DOCX",
            data=save_docx_bytes(f"{company_name} Executive Storylines", combined_text),
            file_name=f"{sanitize_filename(company_name)}_Executive_Storylines.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    else:
        st.info("No storylines generated yet.")

with tab3:
    st.text_area("ADM Financial Summary Output", value=st.session_state.financial_summary_text, height=650)

    if st.session_state.financial_summary_text:
        st.download_button(
            "Download Financial Summary TXT",
            data=st.session_state.financial_summary_text,
            file_name=f"{sanitize_filename(company_name)}_ADM_Financial_Summary.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.download_button(
            "Download Financial Summary DOCX",
            data=save_docx_bytes(f"{company_name} ADM Financial Summary", st.session_state.financial_summary_text),
            file_name=f"{sanitize_filename(company_name)}_ADM_Financial_Summary.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

with tab4:
    st.markdown(f"**ADM Progress:** Batch {st.session_state.adm_batch} / 6")
    st.text_area("ADM Output", value=st.session_state.adm_text, height=700)

    if st.session_state.adm_text:
        st.download_button(
            "Download ADM TXT",
            data=st.session_state.adm_text,
            file_name=f"{sanitize_filename(company_name)}_ADM.txt",
            mime="text/plain",
            use_container_width=True,
        )
        st.download_button(
            "Download ADM DOCX",
            data=save_docx_bytes(f"{company_name} ADM", st.session_state.adm_text),
            file_name=f"{sanitize_filename(company_name)}_ADM.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
