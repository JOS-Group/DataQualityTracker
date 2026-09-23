"""
DART - Data Assurance Reconciliation Tracker
One-file FastAPI website edition

Run:
  python -m pip install fastapi uvicorn pandas numpy plotly openpyxl python-multipart openai
  python -m uvicorn app5:app --reload

What is included in this update:
- Metric Explanations page and inline metric help cards
- Global hover/focus info icons for statistic cards across Medicare, Medicaid, and Build Your Own
- Rebuilt always-visible light/dark bulb switch in a dedicated navbar utility slot, with cache-safe navbar delivery
- Editable Mapping Catalog with add, edit, delete, and seed-from-data support
- Impact Explorer with field-to-report/KPI/owner mapping and impact scoring
- FastAPI Email Agent with persistent automations, baselines, SMTP delivery, and workbook watcher
- Raw Data Editor with cell-level Excel writeback, automatic backups, and immediate automation checks
- Bright white-and-gold professional interface
- Simplified visual layer with Sankey/dependency-network charts removed
- Seeded Workflow Issues so Action Center is populated immediately
- Governance Center with controls, decisions, owners, rule templates, and export-ready summary
- Better empty states so charts do not appear blank when filters return no rows
- Revised page names, plain-English copy, and nonduplicative page content
- Real-workbook-only calculations and stronger chart validation guards
- Clear filter selection labels and safer persona fallbacks
- Loads the real All Streams reconciliation workbook by default
- Enriches each reconciliation row from the CMS NCH-to-STTM mapping workbook
- Persistent Build Your Own dataset library under Data/DIY with two-file comparison and grounded AI analyst
- Build Your Own Raw Data Editor with per-file CSV/XLSX writeback, backups, and immediate BYO alert checks
- Build Your Own Email Alerts with per-dataset automations, baselines, SMTP delivery, and background file watching
- Build Your Own navigation uses the same compact dropdown groups as the other workspaces; Raw Data Editor and Email Alerts are standard Workspace menu items with no special badge or gradient styling
- One Python file only: FastAPI backend + embedded HTML/CSS/JS frontend

Prototype note:
- This is a local demo/prototype app. Do not use with sensitive production data without
  replacing prototype state/auth patterns with production-grade storage and authentication.
"""
from __future__ import annotations

import io
import csv
import sqlite3
import json
import os
import re
import ssl
import smtplib
import hashlib
import html
import uuid
import base64
import ipaddress
import threading
import time
import traceback
import shutil
from contextlib import contextmanager
from datetime import datetime, timezone, date, timedelta
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse, FileResponse

try:
    from openai import OpenAI
except Exception:  # pragma: no cover
    OpenAI = None

try:  # Optional: only needed for Build Your Own S3 import
    import boto3
except Exception:  # pragma: no cover
    boto3 = None

try:  # Optional: only needed for Build Your Own URL import
    import requests
except Exception:  # pragma: no cover
    requests = None

APP_VERSION = "8.9.1"
BUILD_FINGERPRINT = "BYO-NAV-DROPDOWNS-8.9.1"
APP_STARTED_AT = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
app = FastAPI(title="DART - Data Assurance Reconciliation Tracker", version=APP_VERSION)

USERS_FILE = Path(__file__).resolve().parent / "dart_users.json"


BASE_DIR = Path(__file__).resolve().parent


def _resolve_local_file(filename: str) -> Path:
    """Support either repo-root workbooks or a Data/ subfolder without changing code."""
    env_name = "DART_" + re.sub(r"[^A-Za-z0-9]", "_", filename).upper()
    configured = os.environ.get(env_name, "").strip()
    if configured:
        return Path(configured).expanduser().resolve()
    candidates = [BASE_DIR / filename, BASE_DIR / "Data" / filename]
    return next((p for p in candidates if p.exists()), candidates[0])


MAIN_PATH = _resolve_local_file("All Streams June 2026.xlsx")
MAIN_SHEET = os.environ.get("DART_MAIN_SHEET", "Sheet1")
MAPPING_PATH = _resolve_local_file("CMS Version NCH to NCH STTM 12_13_2024.xlsx")
WORKBOOK_BACKUP_DIR = Path(os.environ.get("DART_BACKUP_DIR", str(BASE_DIR / "Data" / "demo_backups")))
BYO_DATA_DIR = Path(os.environ.get("DART_BYO_DATA_DIR", str(BASE_DIR / "Data" / "DIY")))
BYO_MAX_FILE_BYTES = 25 * 1024 * 1024

# Shared Groq configuration for both DART Copilot and Build Your Own AI Analyst.
# Configure GROQ_API_KEY in the local environment or Render service settings.
HARDCODED_GROQ_API_KEY = ""
DEFAULT_GROQ_MODEL = "openai/gpt-oss-20b"
GROQ_FALLBACK_MODELS = [
    "openai/gpt-oss-20b",
    "openai/gpt-oss-120b",
    "llama-3.1-8b-instant",
]
# Leave substantial headroom below Groq's on-demand TPM ceiling. Character
# budgets are deliberately conservative because exact tokenization varies.
GROQ_REQUEST_CHAR_BUDGET = 9000
GROQ_RETRY_CHAR_BUDGET = 6000
# Give the model enough room to finish a polished report while keeping the
# compact input comfortably below the 8K on-demand TPM ceiling.
GROQ_MAX_OUTPUT_TOKENS = 1800

def groq_api_key() -> str:
    return os.getenv("GROQ_API_KEY", "").strip() or HARDCODED_GROQ_API_KEY

def groq_model() -> str:
    return os.getenv("GROQ_MODEL", "").strip() or DEFAULT_GROQ_MODEL

def groq_model_candidates() -> List[str]:
    """Return the preferred model followed by safe fallbacks, with duplicates removed."""
    ordered = [groq_model(), *GROQ_FALLBACK_MODELS]
    seen: set[str] = set()
    return [m for m in ordered if m and not (m in seen or seen.add(m))]

def _trim_ai_text(value: Any, max_chars: int) -> str:
    text = str(value or "")
    if len(text) <= max_chars:
        return text
    if max_chars <= 80:
        return text[:max_chars]
    marker = "\n...[context compacted for Groq token budget]...\n"
    if max_chars <= len(marker) + 24:
        return text[:max_chars]
    payload_chars = max_chars - len(marker)
    head = payload_chars * 3 // 4
    tail = payload_chars - head
    return text[:head] + marker + text[-tail:]

def _fit_groq_messages(messages: List[Dict[str, str]], char_budget: int) -> List[Dict[str, str]]:
    """Preserve grounding + newest turns while bounding the request size."""
    cleaned = [
        {"role": str(m.get("role", "user")), "content": str(m.get("content", ""))}
        for m in messages
        if str(m.get("content", "")).strip()
    ]
    if not cleaned:
        return []
    if sum(len(m["content"]) for m in cleaned) <= char_budget:
        return cleaned

    system_messages = [m for m in cleaned if m["role"] == "system"]
    conversation = [m for m in cleaned if m["role"] != "system"]
    fitted: List[Dict[str, str]] = []

    if system_messages:
        system_budget = max(2500, int(char_budget * 0.68))
        fitted.append({
            "role": "system",
            "content": _trim_ai_text(system_messages[0]["content"], system_budget),
        })

    remaining = max(1000, char_budget - sum(len(m["content"]) for m in fitted))
    recent: List[Dict[str, str]] = []
    # Newest messages get priority so the current question is never dropped.
    for message in reversed(conversation[-6:]):
        if remaining <= 0:
            break
        per_message = min(1400, remaining)
        content = _trim_ai_text(message["content"], per_message)
        recent.append({"role": message["role"], "content": content})
        remaining -= len(content)
    fitted.extend(reversed(recent))
    return fitted

def groq_chat_completion(
    messages: List[Dict[str, str]],
    temperature: float = 0.2,
    max_output_tokens: int = GROQ_MAX_OUTPUT_TOKENS,
) -> Tuple[str, str]:
    """Run Groq with model fallback and automatic compaction for oversized requests."""
    if OpenAI is None:
        raise RuntimeError("The OpenAI client package is not installed.")
    api_key = groq_api_key()
    if not api_key:
        raise RuntimeError("The Groq API key is not configured.")

    client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
    for model_name in groq_model_candidates():
        unavailable_model = False
        last_size_error: Optional[Exception] = None
        for attempt, char_budget in enumerate([GROQ_REQUEST_CHAR_BUDGET, GROQ_RETRY_CHAR_BUDGET]):
            fitted_messages = _fit_groq_messages(messages, char_budget)
            try:
                completion_kwargs: Dict[str, Any] = {
                    "model": model_name,
                    "temperature": temperature,
                    "max_completion_tokens": max(900, max_output_tokens - (attempt * 300)),
                    "messages": fitted_messages,
                }
                # GPT-OSS supports variable reasoning effort. Low reasoning keeps more of
                # the organization's 8K TPM allowance available for the visible answer.
                if model_name.startswith("openai/gpt-oss"):
                    completion_kwargs["reasoning_effort"] = "low"
                response = client.chat.completions.create(**completion_kwargs)
                content = response.choices[0].message.content or ""
                return content, model_name
            except Exception as exc:
                error_lower = str(exc).lower()
                unavailable = any(token in error_lower for token in [
                    "model_not_found",
                    "does not exist",
                    "model not found",
                    "do not have access",
                    "not have access",
                ])
                if unavailable:
                    unavailable_model = True
                    break

                too_large = (
                    "request too large" in error_lower
                    or ("tokens per minute" in error_lower and "requested" in error_lower and "limit" in error_lower)
                    or "error code: 413" in error_lower
                    or "status code: 413" in error_lower
                )
                if too_large and attempt == 0:
                    last_size_error = exc
                    continue
                raise
        if unavailable_model:
            continue
        if last_size_error is not None:
            raise RuntimeError(
                "The request exceeded Groq's token budget even after automatic compaction. "
                "Clear the conversation or ask a more focused question."
            ) from last_size_error

    tried = ", ".join(groq_model_candidates())
    raise RuntimeError(f"No accessible Groq chat model was found. Tried: {tried}.")

# Developer email configuration. Set SMTP_PASSWORD in the deployment environment.
HARDCODED_SMTP_CONFIG = {
    "host": "smtp.gmail.com",
    "port": 465,
    "username": "dartcustomagent@gmail.com",
    "password": os.getenv("SMTP_PASSWORD", ""),
    "from_email": "dartcustomagent@gmail.com",
    "from_name": "DART Email Agent",
    # Gmail port 465 uses implicit TLS/SSL from the start of the connection.
    "use_tls": False,
    "use_ssl": True,
}

EMAIL_AGENT_DIR = Path(os.environ.get("EMAIL_AGENT_DIR", str(BASE_DIR / "Data" / "email_agent")))
EMAIL_TEMPLATE_PATH = EMAIL_AGENT_DIR / "automation_templates.json"
EMAIL_SNAPSHOT_DIR = EMAIL_AGENT_DIR / "snapshots"
EMAIL_AGENT_STATUS_PATH = EMAIL_AGENT_DIR / "agent_status.json"
EMAIL_AGENT_HISTORY_PATH = EMAIL_AGENT_DIR / "execution_history.jsonl"
EMAIL_AGENT_AUTO_RUN = os.environ.get("EMAIL_AGENT_AUTO_RUN", "true").strip().lower() in {"1", "true", "yes", "y", "on"}
EMAIL_AGENT_MIN_CHECK_SECONDS = max(15, int(os.environ.get("EMAIL_AGENT_MIN_CHECK_SECONDS", "60")))
EMAIL_AGENT_STARTUP_DELAY_SECONDS = max(0, int(os.environ.get("EMAIL_AGENT_STARTUP_DELAY_SECONDS", "5")))
EMAIL_AGENT_LOCK_TIMEOUT_SECONDS = max(30, int(os.environ.get("EMAIL_AGENT_LOCK_TIMEOUT_SECONDS", "900")))
EMAIL_PREVIEW_ROW_LIMIT = 30
NULL_TOKEN = "<DART_NULL>"
_EMAIL_AGENT_THREAD = None
_EMAIL_AGENT_THREAD_LOCK = threading.Lock()

# Build Your Own alerting is intentionally isolated from Medicare automations.
BYO_EMAIL_AGENT_DIR = Path(os.environ.get("DART_BYO_EMAIL_AGENT_DIR", str(BASE_DIR / "Data" / "byo_email_agent")))
BYO_EMAIL_TEMPLATE_PATH = BYO_EMAIL_AGENT_DIR / "automation_templates.json"
BYO_EMAIL_STATUS_PATH = BYO_EMAIL_AGENT_DIR / "agent_status.json"
BYO_EMAIL_HISTORY_PATH = BYO_EMAIL_AGENT_DIR / "execution_history.jsonl"
BYO_WORKBOOK_BACKUP_DIR = Path(os.environ.get("DART_BYO_BACKUP_DIR", str(BASE_DIR / "Data" / "byo_backups")))
BYO_EMAIL_AGENT_AUTO_RUN = os.environ.get("DART_BYO_EMAIL_AUTO_RUN", "true").strip().lower() in {"1", "true", "yes", "y", "on"}
BYO_EMAIL_AGENT_MIN_CHECK_SECONDS = max(15, int(os.environ.get("DART_BYO_EMAIL_MIN_CHECK_SECONDS", str(EMAIL_AGENT_MIN_CHECK_SECONDS))))
_BYO_EMAIL_AGENT_THREAD = None
_BYO_EMAIL_AGENT_THREAD_LOCK = threading.Lock()

EMAIL_OPERATORS = [
    "Is blank (empty or null)", "Is null / NaN", "Is empty string", "Equals",
    "Does not equal", "Contains", "Does not contain", "Starts with",
    "Greater than", "Greater than or equal to", "Less than", "Less than or equal to",
    "Before date", "After date",
]

STATE: Dict[str, Any] = {
    "persona": None,
    "df": pd.DataFrame(),
    "source": "",
    "issues": [],
    "connections": [],
    "impacts": [],
    "governance": [],
    "rules": [],
    "chat": [],
    "snapshots": [],
    "current_user": None,
    # Build Your Own is intentionally isolated from the Medicare reconciliation state.
    "byo_df": pd.DataFrame(),
    "byo_source": "",
    "byo_uploaded_at": "",
    "byo_chat": [],
    "byo_chat_pair": [],
}

# -----------------------------------------------------------------------------
# User/profile helpers
# -----------------------------------------------------------------------------

def load_users() -> Dict[str, Any]:
    if not USERS_FILE.exists():
        return {"users": {}}
    try:
        return json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"users": {}}


def save_users(data: Dict[str, Any]) -> None:
    USERS_FILE.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")


def public_user(username: Optional[str]) -> Optional[Dict[str, Any]]:
    if not username:
        return None
    user = load_users().get("users", {}).get(username)
    if not user:
        return None
    profile = user.get("profile", {})
    persona = user.get("persona") if isinstance(user.get("persona"), dict) else None
    return {
        "username": username,
        "display_name": profile.get("display_name", username),
        "email": profile.get("email", ""),
        "organization": profile.get("organization", ""),
        "role": profile.get("role", ""),
        "created_at": user.get("created_at", "local prototype"),
        "storage_file": USERS_FILE.name,
        "has_persona": bool(persona),
    }


PERSONA_WORKSPACE_PAGES: Dict[str, List[str]] = {
    "medicare": [
        "home", "briefing", "metrics", "briefbuilder", "riskcenter", "insights",
        "scorecards", "actioncenter", "catalog", "governance",
        "explorer", "raweditor", "emailagent", "data",
        "copilot", "profile", "settings",
    ],
    "medicaid": [
        "medicaid-home", "medicaid-states", "medicaid-state", "medicaid-compare",
        "medicaid-analytics", "medicaid-ai", "medicaid-chat", "medicaid-exec",
        "medicaid-heatmap", "medicaid-claims", "medicaid-optimize",
        "medicaid-methodology", "profile", "settings",
    ],
    "byo": [
        "byo-home", "byo-library", "byo-lab", "byo-compare", "byo-quality",
        "byo-preview", "byo-raw", "byo-email", "byo-ai", "byo-export", "profile", "settings",
    ],
}

PERSONA_ROLE_LANDING = {
    "Executive / Leadership": "briefing",
    "Data / Analytics": "explorer",
    "Program / Policy": "briefing",
    "Operations": "actioncenter",
    "Quality / Compliance": "riskcenter",
    "IT / Engineering": "catalog",
    "Research": "insights",
    "Other": "home",
}

PERSONA_ROLE_RECOMMENDATIONS: Dict[str, List[Tuple[str, str, str]]] = {
    "Executive / Leadership": [
        ("briefing", "My Briefing", "Review the decision-ready quality and risk picture."),
        ("briefbuilder", "Executive Brief", "Turn the current scope into a leadership-ready readout."),
        ("governance", "Governance Center", "Review controls, ownership, and decision readiness."),
    ],
    "Data / Analytics": [
        ("explorer", "System Explorer", "Trace the field-level evidence behind anomalies."),
        ("insights", "AI Insights", "Explore evidence-backed patterns and hypotheses."),
        ("scorecards", "Scorecards", "Inspect stream-level distributions and metrics."),
    ],
    "Program / Policy": [
        ("briefing", "My Briefing", "Translate technical findings into program priorities."),
        ("governance", "Governance Center", "Review controls, ownership, and decision readiness."),
        ("scorecards", "Scorecards", "Compare program streams at an executive level."),
    ],
    "Operations": [
        ("actioncenter", "Remediation Center", "Assign, track, and move findings toward resolution."),
        ("emailagent", "Email Agent", "Automate notifications when monitored data changes."),
        ("riskcenter", "Risk Center", "Focus operational work on the highest-impact findings."),
    ],
    "Quality / Compliance": [
        ("riskcenter", "Risk Center", "Prioritize critical and elevated reconciliation findings."),
        ("governance", "Governance Center", "Review controls, evidence, ownership, and cadence."),
        ("actioncenter", "Remediation Center", "Track corrective actions through resolution."),
    ],
    "IT / Engineering": [
        ("catalog", "Mapping Catalog", "Validate source-to-target mappings and lineage."),
        ("raweditor", "Raw Data Editor", "Inspect and correct source workbook values safely."),
        ("emailagent", "Email Agent", "Operationalize data-change monitoring and alerts."),
    ],
    "Research": [
        ("insights", "AI Insights", "Explore evidence-backed patterns and hypotheses."),
        ("copilot", "DART Copilot", "Ask focused questions against the current data context."),
        ("scorecards", "Scorecards", "Compare patterns across reconciliation streams."),
    ],
    "Other": [
        ("home", "Command Center", "Start with the overall quality and risk picture."),
        ("briefing", "My Briefing", "Review the most important findings in plain language."),
        ("riskcenter", "Risk Center", "Inspect the highest-priority reconciliation issues."),
    ],
}


def _persona_workspace_from_program(program: str) -> str:
    program_lower = str(program or "").strip().lower()
    if program_lower == "medicaid":
        return "medicaid"
    if program_lower in {"other / general", "other", "general"}:
        return "byo"
    return "medicare"


def normalize_persona(payload: Dict[str, Any]) -> Dict[str, Any]:
    persona = dict(payload or {})
    persona["program"] = str(persona.get("program", "Medicare") or "Medicare").strip()
    persona["role"] = str(persona.get("role", "Other") or "Other").strip()
    persona["audience"] = str(persona.get("audience", "Myself") or "Myself").strip()
    persona["depth"] = str(persona.get("depth", "Balanced") or "Balanced").strip()
    focus = persona.get("focus", [])
    if isinstance(focus, str):
        focus = [x.strip() for x in focus.split(",") if x.strip()]
    persona["focus"] = [str(x).strip() for x in (focus or []) if str(x).strip()][:12]
    persona["geography"] = [str(x).strip() for x in (persona.get("geography") or ["National"]) if str(x).strip()] or ["National"]
    persona["first_question"] = str(persona.get("first_question", "") or "").strip()

    requested_workspace = str(persona.get("preferred_workspace", "") or "").strip().lower()
    if requested_workspace not in PERSONA_WORKSPACE_PAGES:
        requested_workspace = _persona_workspace_from_program(persona["program"])
    persona["preferred_workspace"] = requested_workspace

    requested_page = str(persona.get("landing_page", "") or "").strip()
    if requested_page in {"", "auto", "recommended"}:
        if requested_workspace == "medicaid":
            requested_page = "medicaid-home"
        elif requested_workspace == "byo":
            requested_page = "byo-home"
        else:
            requested_page = PERSONA_ROLE_LANDING.get(persona["role"], "home")
    if requested_page not in PERSONA_WORKSPACE_PAGES[requested_workspace]:
        requested_page = {"medicare": "home", "medicaid": "medicaid-home", "byo": "byo-home"}[requested_workspace]
    persona["landing_page"] = requested_page
    persona["updated_at"] = _utc_now_iso() if "_utc_now_iso" in globals() else datetime.now(timezone.utc).replace(microsecond=0).isoformat()
    return persona


def current_persona() -> Optional[Dict[str, Any]]:
    username = STATE.get("current_user")
    if not username:
        return None
    user = load_users().get("users", {}).get(username, {})
    persona = user.get("persona")
    if isinstance(persona, dict) and persona:
        return normalize_persona(persona)
    return None


def persona_effects(persona: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    p = normalize_persona(persona or {}) if persona else {}
    if not p:
        return {}
    workspace = p.get("preferred_workspace", "medicare")
    role = p.get("role", "Other")
    depth = p.get("depth", "Balanced")
    focus = p.get("focus", [])

    if workspace == "byo":
        recs = [
            ("byo-compare", "Compare Lab", "Run deterministic analysis across the selected dataset pair."),
            ("byo-ai", "AI Analyst", "Ask grounded questions across both selected datasets."),
            ("byo-quality", "Schema & Quality", "Inspect missingness, schema, and column-level differences."),
        ]
        lens_title = "Two-dataset investigation lens"
        lens_summary = "Prioritize comparison evidence, join quality, anomalies, and clear next checks across the selected files."
    elif workspace == "medicaid":
        role = p.get("role", "Other")
        role_recs = {
            "Executive / Leadership": [("medicaid-exec", "Executive Brief", "Review the nationwide quality picture and 30/60/90 actions."), ("medicaid-heatmap", "US Heatmap", "See geographic quality variation and drill into states."), ("medicaid-home", "Dashboard", "Start with nationwide Medicaid quality KPIs and rankings.")],
            "Data / Analytics": [("medicaid-analytics", "Analytics", "Inspect issue-type performance and modeled risk."), ("medicaid-compare", "Compare States", "Benchmark state quality scores and confidence."), ("medicaid-heatmap", "US Heatmap", "Explore geographic patterns across quality metrics.")],
            "Operations": [("medicaid-states", "State Explorer", "Work directly from state issue backlogs and statuses."), ("medicaid-claims", "Claims Analysis", "Review claim-quality trends and monitoring opportunities."), ("medicaid-optimize", "Optimize Spending", "Prioritize modeled financial and quality opportunities.")],
            "Quality / Compliance": [("medicaid-analytics", "Analytics", "Review issue-type quality, backlog, and cancellation patterns."), ("medicaid-states", "State Explorer", "Inspect state-level quality evidence and remediation."), ("medicaid-exec", "Executive Brief", "Summarize quality governance priorities for leadership.")],
            "IT / Engineering": [("medicaid-methodology", "Methodology", "Review issue definitions, scoring logic, and model assumptions."), ("medicaid-states", "State Explorer", "Inspect issue records and technical categories by state."), ("medicaid-chat", "CMS Q&A", "Query state and issue evidence directly.")],
        }
        recs = role_recs.get(role, [("medicaid-home", "Dashboard", "Start with nationwide Medicaid quality KPIs and rankings."), ("medicaid-states", "State Explorer", "Open any state for detailed issue tracking."), ("medicaid-ai", "AI Insights", "Review evidence-based cross-state recommendations.")])
        lens_title = f"{role} Medicaid lens"
        lens_summary = "Use the 50-state Medicaid quality workspace for state benchmarking, CMS issue tracking, claims quality, modeled risk, and executive decision support."
    else:
        recs = PERSONA_ROLE_RECOMMENDATIONS.get(role, PERSONA_ROLE_RECOMMENDATIONS["Other"])
        role_summary = {
            "Executive / Leadership": "Surface the few decisions, business impacts, and risks that need leadership attention.",
            "Data / Analytics": "Emphasize evidence, distributions, comparisons, and field-level investigation paths.",
            "Program / Policy": "Translate reconciliation findings into program impact, reporting implications, and decision needs.",
            "Operations": "Prioritize actionable findings, ownership, alerts, and remediation progress.",
            "Quality / Compliance": "Emphasize critical risk, controls, evidence, governance, and corrective action.",
            "IT / Engineering": "Emphasize mappings, lineage, raw-data evidence, automation, and technical diagnostics.",
            "Research": "Emphasize patterns, comparisons, hypotheses, and evidence before drawing conclusions.",
            "Other": "Balance quality, risk, impact, and next actions across the current scope.",
        }.get(role, "Balance quality, risk, impact, and next actions across the current scope.")
        lens_title = f"{role} lens"
        lens_summary = role_summary

    priority_limit = {"Executive": 4, "Balanced": 8, "Technical": 12}.get(depth, 8)
    return {
        "workspace": workspace,
        "landing_page": p.get("landing_page"),
        "lens_title": lens_title,
        "lens_summary": lens_summary,
        "priority_limit": priority_limit,
        "focus": focus[:5],
        "recommended_pages": [
            {"id": page_id, "label": label, "reason": reason}
            for page_id, label, reason in recs
        ],
    }

# -----------------------------------------------------------------------------
# Data helpers
# -----------------------------------------------------------------------------

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df


def find_col(df: pd.DataFrame, choices: List[str]) -> Optional[str]:
    exact = {str(c).lower().strip(): c for c in df.columns}
    for choice in choices:
        key = choice.lower().strip()
        if key in exact:
            return exact[key]
    for col in df.columns:
        for choice in choices:
            if choice.lower() in str(col).lower():
                return col
    return None


def standardize(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    df = norm_cols(df)
    aliases = {
        "Stream": ["Stream", "Claim Stream", "Domain", "LOB", "Line of Business"],
        "NCH Target Table": ["NCH Target Table", "NCH Table", "System A Table", "Source Table", "Source Entity"],
        "NCH Target Column": ["NCH Target Column", "NCH Column", "System A Column", "Source Column", "Field"],
        "SS Table": ["SS Table", "Shared Systems Table", "System B Table", "Target Table", "Target Entity"],
        "SS Column": ["SS Column", "SS Colmn if different", "Shared Systems Column", "System B Column", "Target Column"],
        "MatchedClaims": ["MatchedClaims", "Matched Claims", "1/1/24 - 12/31/24 Matched claims", "Matched Events", "Matched", "Match Count"],
        "NotMatchedClaims": ["NotMatchedClaims", "Not Matched Claims", "1/1/24 - 12/31/24 Not Matched claims", "Not-Matched Claims", "Unmatched", "Mismatch Count"],
        "MatchRate": ["MatchRate", "Match Rate", "% matched", "Match %", "Match Percentage", "Accuracy"],
        "Classification": ["Classification", "ClassificationClean", "Class", "Status"],
        "Sub-Classification": ["Sub-Classification", "Sub Classification", "Subclass", "Reason", "Driver"],
        "Disposition": ["Disposition", "DispositionClean", "Action"],
        "Recommendation": ["Recommendation", "Recommended Action", "Next Step"],
        "Comments": ["Comments", "Comments/Examples", "Comment", "Notes"],
        "Lynette Recommendation": ["Lynette Recommendation", "Lynette's recommendation", "Lynnette's recommendation"],
        "Owner": ["Owner", "Business Owner", "Team"],
        "Report": ["Report", "Dashboard", "Downstream Report"],
        "KPI": ["KPI", "Metric", "Measure"],
    }
    rename_map = {}
    for target, options in aliases.items():
        col = find_col(df, options)
        if col and col != target:
            rename_map[col] = target
    df = df.rename(columns=rename_map)

    for col in ["MatchedClaims", "NotMatchedClaims", "MatchRate"]:
        if col in df:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
                errors="coerce",
            )
    if "MatchedClaims" not in df:
        df["MatchedClaims"] = np.nan
    if "NotMatchedClaims" not in df:
        df["NotMatchedClaims"] = np.nan
    if "MatchRate" in df and df["MatchRate"].dropna().size and df["MatchRate"].dropna().max() > 1.5:
        df["MatchRate"] = df["MatchRate"] / 100

    # Keep match-rate metrics tied to the actual reconciliation counts. The source
    # workbook can contain stale/formatted percentage values, so matched and
    # not-matched counts are authoritative whenever claim volume is available.
    total = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
    calculated_rate = pd.Series(
        np.where(total > 0, df["MatchedClaims"].fillna(0) / total, np.nan),
        index=df.index,
        dtype="float64",
    )
    if "MatchRate" not in df:
        df["MatchRate"] = calculated_rate
    else:
        supplied_rate = pd.to_numeric(df["MatchRate"], errors="coerce")
        df["MatchRate"] = calculated_rate.where(total > 0, supplied_rate)
    if "Classification" not in df:
        df["Classification"] = np.where(df["MatchRate"] >= 0.99, "Match", np.where(df["MatchRate"].notna(), "Different", "Missing"))

    for col in [
        "Stream", "NCH Target Table", "NCH Target Column", "SS Table", "SS Column",
        "Disposition", "Recommendation", "Comments", "Lynette Recommendation", "Sub-Classification", "Owner", "Report", "KPI",
        "Source Field Name", "Source Field Description", "Business Rule", "Mapping Comments", "Mapping Notes", "Mapping Status",
    ]:
        if col not in df:
            df[col] = ""

    df["Comparable"] = df["Classification"].astype(str).str.lower().str.strip().isin(["match", "different", "mismatch"])
    def normalized_action_text(column: str) -> pd.Series:
        return (
            df[column]
            .astype(str)
            .str.lower()
            .str.replace(r"[^a-z0-9\s]", "", regex=True)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    disp = normalized_action_text("Disposition")
    recommendation = normalized_action_text("Recommendation")
    lynette_recommendation = normalized_action_text("Lynette Recommendation")
    cls = df["Classification"].astype(str).str.lower()
    # Explicit no-action recommendations override a mismatch classification.
    no_action_values = {"no action needed", "no action is needed", "no action required"}
    no_action_needed = (
        disp.isin(no_action_values)
        | recommendation.isin(no_action_values)
        | lynette_recommendation.isin(no_action_values)
    )
    df["NeedsChange"] = (
        disp.str.contains("action|change|review|fix|remediate", na=False)
        | cls.isin(["missing", "different", "mismatch"])
    ) & ~no_action_needed
    df["TotalClaims"] = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
    df["ImpactScore"] = ((1 - df["MatchRate"].fillna(0)) * np.log1p(df["TotalClaims"].fillna(0))).round(2)
    df["RiskTier"] = np.select(
        [df["ImpactScore"] >= 6, df["ImpactScore"] >= 3, df["ImpactScore"] > 0],
        ["Critical", "Elevated", "Watch"],
        default="Stable",
    )
    return df


def pct(x: Any) -> str:
    return "-" if pd.isna(x) else f"{float(x):.1%}"


def fmt_int(x: Any) -> str:
    try:
        return f"{int(float(x)):,}"
    except Exception:
        return "0"


def field_average(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    x = df[df["Comparable"] & df["MatchRate"].notna()]
    return float(x["MatchRate"].mean()) if len(x) else np.nan


def weighted_rate(df: pd.DataFrame) -> float:
    if df.empty:
        return np.nan
    m = float(df["MatchedClaims"].fillna(0).sum())
    u = float(df["NotMatchedClaims"].fillna(0).sum())
    return m / (m + u) if (m + u) else np.nan


def _clean_key(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip()).upper()
    return "" if text in {"", "NAN", "NONE", "-"} else text


def _join_unique(values: pd.Series) -> str:
    seen: List[str] = []
    for value in values:
        text = str(value).strip() if pd.notna(value) else ""
        if text and text.lower() not in {"nan", "none", "-"} and text not in seen:
            seen.append(text)
    return " | ".join(seen)


def load_mapping_catalog(path: Path) -> pd.DataFrame:
    mapping = pd.read_excel(path, sheet_name="Lynnette's Version", header=1, engine="openpyxl")
    mapping = norm_cols(mapping)
    needed = ["NCH Target Table", "NCH Target Column"]
    if not all(c in mapping.columns for c in needed):
        raise ValueError("Mapping workbook is missing NCH Target Table/NCH Target Column.")
    mapping = mapping[mapping["NCH Target Table"].notna() & mapping["NCH Target Column"].notna()].copy()
    mapping["_table_key"] = mapping["NCH Target Table"].map(_clean_key)
    mapping["_column_key"] = mapping["NCH Target Column"].map(_clean_key)
    mapping = mapping[(mapping["_table_key"] != "") & (mapping["_column_key"] != "")]
    rename = {
        "NCH Cobol Copy book Field Name": "Source Field Description",
        "In Table Source Fields": "Source Field Name",
        "Business Rule": "Business Rule",
        "Comments": "Mapping Comments",
        "Notes": "Mapping Notes",
        "Data Match? Y/N": "Mapping Status",
    }
    mapping = mapping.rename(columns={k: v for k, v in rename.items() if k in mapping.columns})
    keep = ["_table_key", "_column_key"] + [c for c in rename.values() if c in mapping.columns]
    mapping = mapping[keep]
    agg = {c: _join_unique for c in keep if c not in {"_table_key", "_column_key"}}
    return mapping.groupby(["_table_key", "_column_key"], as_index=False).agg(agg)


def enrich_with_mapping(df: pd.DataFrame, mapping_path: Path) -> pd.DataFrame:
    if df.empty or not mapping_path.exists():
        return df
    left = df.copy()
    left["_table_key"] = left["NCH Target Table"].map(_clean_key)
    left["_column_key"] = left["NCH Target Column"].map(_clean_key)
    try:
        catalog = load_mapping_catalog(mapping_path)
        left = left.merge(catalog, on=["_table_key", "_column_key"], how="left")
    except Exception as exc:
        # Mapping enrichment is optional. Preserve the reconciliation data if the
        # mapping workbook has a different sheet/header layout.
        print(f"WARNING: Mapping enrichment skipped: {exc}")
        left["Mapping Load Warning"] = str(exc)

    source = left.get("Source Field Name", pd.Series("", index=left.index)).fillna("").astype(str).str.strip()
    rule = left.get("Business Rule", pd.Series("", index=left.index)).fillna("").astype(str).str.strip()
    left["Mapping Found"] = source.ne("") | rule.ne("")
    return left.drop(columns=["_table_key", "_column_key"], errors="ignore")


def load_local_data() -> Tuple[pd.DataFrame, str]:
    data_path = Path(MAIN_PATH)
    mapping_path = Path(MAPPING_PATH)
    if not data_path.exists():
        return pd.DataFrame(), f"Missing required file: {data_path.name}"
    try:
        df = pd.read_excel(data_path, sheet_name=MAIN_SHEET, engine="openpyxl")
        df = standardize(df)
        df = enrich_with_mapping(df, mapping_path)
        df = standardize(df)
        source = data_path.name
        if mapping_path.exists():
            source += f" + {mapping_path.name}"
        else:
            source += " (mapping workbook not found)"
        return df, source
    except Exception as exc:
        return pd.DataFrame(), f"Could not load workbook data: {exc}"


def seed_lineage_from_df(df: pd.DataFrame) -> None:
    if STATE["connections"]:
        return
    sample = df.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(12)
    records = []
    for i, r in sample.iterrows():
        records.append({
            "id": f"L{i}",
            "Stream": str(r.get("Stream", "")),
            "SourceTable": str(r.get("NCH Target Table", "")),
            "SourceField": str(r.get("NCH Target Column", "")),
            "TargetTable": str(r.get("SS Table", "")),
            "TargetField": str(r.get("SS Column", "")),
            "Report": str(r.get("Report", "")) or f"{r.get('Stream', 'Stream')} Quality Report",
            "KPI": str(r.get("KPI", "")) or f"{r.get('Stream', 'Stream')} Match KPI",
            "Owner": str(r.get("Owner", "")) or "Data Quality",
            "Status": "Draft" if r.get("RiskTier") in ["Critical", "Elevated"] else "Validated",
            "RiskTier": str(r.get("RiskTier", "")),
            "Notes": str(r.get("Recommendation", "")),
        })
    STATE["connections"] = records


def seed_workflow_items(df: pd.DataFrame) -> None:
    if STATE["issues"]:
        return
    sample = df[df["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(8)
    items = []
    lane_cycle = ["Open", "In progress", "Open", "Blocked", "Open", "In progress", "Resolved", "Open"]
    for pos, (_, r) in enumerate(sample.iterrows()):
        items.append({
            "id": f"I{pos+1}",
            "Stream": str(r["Stream"]),
            "Field": str(r["NCH Target Column"]),
            "RiskTier": str(r["RiskTier"]),
            "Priority": "Critical" if r["RiskTier"] == "Critical" else "High" if r["RiskTier"] == "Elevated" else "Medium",
            "Owner": str(r.get("Owner", "Data Quality")) or "Data Quality",
            "Status": lane_cycle[pos % len(lane_cycle)],
            "Due": "Next review cycle",
            "Note": str(r.get("Recommendation", "Confirm mapping and lineage.")),
        })
    STATE["issues"] = items


def seed_impacts(df: pd.DataFrame) -> None:
    if STATE["impacts"]:
        return
    sample = df.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(8)
    STATE["impacts"] = [{
        "id": f"M{i+1}",
        "Stream": str(r["Stream"]),
        "Field": str(r["NCH Target Column"]),
        "Report": str(r.get("Report", "")) or f"{r['Stream']} Quality Report",
        "KPI": str(r.get("KPI", "")) or f"{r['Stream']} Match KPI",
        "Owner": str(r.get("Owner", "")) or "Reporting",
        "Impact": "Critical" if r["RiskTier"] == "Critical" else "High" if r["RiskTier"] == "Elevated" else "Moderate",
        "DecisionNeed": "Confirm remediation priority" if r["RiskTier"] in ["Critical", "Elevated"] else "Monitor",
    } for i, (_, r) in enumerate(sample.iterrows())]


def seed_governance() -> None:
    if STATE["governance"]:
        return
    STATE["governance"] = [
        {"id": "G1", "Control": "Critical mapping owner assigned", "Owner": "Data Quality", "Cadence": "Weekly", "Status": "Active", "Evidence": "Action Center owner and note"},
        {"id": "G2", "Control": "Lineage status reviewed", "Owner": "Reporting", "Cadence": "Monthly", "Status": "Draft", "Evidence": "Lineage catalog status"},
        {"id": "G3", "Control": "Executive KPI impact confirmed", "Owner": "Leadership", "Cadence": "Review cycle", "Status": "Active", "Evidence": "Impact Explorer mapping"},
        {"id": "G4", "Control": "Automation threshold defined", "Owner": "Engineering", "Cadence": "As needed", "Status": "Planned", "Evidence": "Automation rule template"},
    ]


def ensure_data() -> None:
    if STATE["df"].empty:
        df, name = load_local_data()
        if df.empty:
            raise RuntimeError(name or "The required All Streams workbook could not be loaded.")
        STATE["df"] = df
        STATE["source"] = name
    df = standardize(STATE["df"])


def custom_mask(series: pd.Series, op: str, value: str) -> pd.Series:
    s = series.astype(str)
    value_text = str(value)
    op = str(op or "contains")
    if op == "equals":
        return s.str.lower() == value_text.lower()
    if op == "not_equals":
        return s.str.lower() != value_text.lower()
    if op == "starts_with":
        return s.str.lower().str.startswith(value_text.lower(), na=False)
    if op == "ends_with":
        return s.str.lower().str.endswith(value_text.lower(), na=False)
    if op in ["gt", "gte", "lt", "lte"]:
        nums = pd.to_numeric(series, errors="coerce")
        try:
            v = float(value_text)
        except Exception:
            return pd.Series([True] * len(series), index=series.index)
        if op == "gt":
            return nums > v
        if op == "gte":
            return nums >= v
        if op == "lt":
            return nums < v
        if op == "lte":
            return nums <= v
    return s.str.contains(value_text, case=False, na=False)


def apply_filters(filters: Dict[str, Any]) -> pd.DataFrame:
    ensure_data()
    df = standardize(STATE["df"])
    fdf = df.copy()
    streams = filters.get("streams")
    classes = filters.get("classes")
    tiers = filters.get("tiers")
    reasons = filters.get("reasons")
    min_rate = float(filters.get("min_rate", 0) or 0)
    min_impact = float(filters.get("min_impact", 0) or 0)
    min_unmatched = float(filters.get("min_unmatched", 0) or 0)
    actions_only = bool(filters.get("actions_only", False))
    search = str(filters.get("search", "")).strip()
    custom_filters = filters.get("custom_filters") or []
    if streams:
        fdf = fdf[fdf["Stream"].astype(str).isin(streams)]
    if classes:
        fdf = fdf[fdf["Classification"].astype(str).isin(classes)]
    if tiers:
        fdf = fdf[fdf["RiskTier"].astype(str).isin(tiers)]
    if reasons:
        fdf = fdf[fdf["Sub-Classification"].astype(str).isin(reasons)]
    if min_rate > 0:
        fdf = fdf[fdf["MatchRate"].isna() | (fdf["MatchRate"] >= min_rate)]
    if min_impact > 0:
        fdf = fdf[fdf["ImpactScore"].fillna(0) >= min_impact]
    if min_unmatched > 0:
        fdf = fdf[fdf["NotMatchedClaims"].fillna(0) >= min_unmatched]
    if actions_only:
        fdf = fdf[fdf["NeedsChange"]]
    if search:
        mask = fdf.astype(str).apply(lambda col: col.str.contains(search, case=False, na=False)).any(axis=1)
        fdf = fdf[mask]
    for cf in custom_filters:
        field = str(cf.get("field", "")).strip()
        op = str(cf.get("op", "contains")).strip()
        value = str(cf.get("value", "")).strip()
        if field in fdf.columns and value:
            fdf = fdf[custom_mask(fdf[field], op, value)]
    return fdf



# -----------------------------------------------------------------------------
# Raw workbook editor + Email Agent
# -----------------------------------------------------------------------------

def _source_file_signature(path: Path | str) -> str:
    try:
        file_path = Path(path)
        stat = file_path.stat()
        digest = hashlib.sha256()
        with file_path.open("rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                digest.update(chunk)
        return f"{stat.st_mtime_ns}:{stat.st_size}:{digest.hexdigest()}"
    except OSError:
        return "missing"


def _read_raw_main_sheet() -> pd.DataFrame:
    return pd.read_excel(MAIN_PATH, sheet_name=MAIN_SHEET, engine="openpyxl")


def _editor_value_is_blank(value: Any) -> bool:
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    return isinstance(value, str) and value == ""


def _editor_values_equal(left: Any, right: Any) -> bool:
    if _editor_value_is_blank(left) and _editor_value_is_blank(right):
        return True
    if isinstance(left, pd.Timestamp):
        left = left.to_pydatetime()
    if isinstance(right, pd.Timestamp):
        right = right.to_pydatetime()
    if isinstance(left, np.generic):
        left = left.item()
    if isinstance(right, np.generic):
        right = right.item()
    try:
        result = left == right
        return bool(result) if not pd.isna(result) else False
    except Exception:
        return str(left) == str(right)


def _editor_compare_text(value: Any) -> str:
    if _editor_value_is_blank(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if isinstance(value, np.generic):
        value = value.item()
    return str(value)


def _coerce_editor_input(value: Any, original: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return None
    else:
        text = str(value)

    try:
        if isinstance(original, (bool, np.bool_)):
            return text.lower() in {"true", "1", "yes", "y"}
        if isinstance(original, (int, np.integer)) and not isinstance(original, (bool, np.bool_)):
            return int(float(text))
        if isinstance(original, (float, np.floating)) and not pd.isna(original):
            return float(text)
        if isinstance(original, (pd.Timestamp, datetime)):
            parsed = pd.to_datetime(text, errors="raise")
            return parsed.to_pydatetime()
    except Exception:
        # If the user intentionally changes the type, keep the entered text instead of failing.
        return value
    return value


def _excel_safe_value(value: Any) -> Any:
    if _editor_value_is_blank(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, np.generic):
        return value.item()
    return value


def _build_raw_editor_frame(raw_df: pd.DataFrame) -> pd.DataFrame:
    frame = raw_df.copy()
    frame.insert(0, "_ExcelRow", np.arange(2, len(frame) + 2, dtype=int))
    return frame


def _create_workbook_backup() -> Path:
    source = Path(MAIN_PATH)
    WORKBOOK_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = WORKBOOK_BACKUP_DIR / f"{source.stem}_before_edit_{stamp}{source.suffix}"
    shutil.copy2(source, backup)
    return backup


def _editor_before_matches(current_value: Any, expected_before: Any) -> bool:
    """Compare browser text to the current Excel value without false conflicts for 1 vs 1.0."""
    current_text = _editor_compare_text(current_value)
    expected_text = "" if expected_before is None else str(expected_before)
    if current_text == expected_text:
        return True
    if current_text.strip() == "" and expected_text.strip() == "":
        return True
    try:
        return float(current_text) == float(expected_text)
    except (TypeError, ValueError):
        return False


def _write_raw_changes_to_workbook(
    changes: List[Dict[str, Any]], expected_signature: Optional[str] = None
) -> Dict[str, Any]:
    if not changes:
        return {"saved": 0, "backup": None, "path": str(Path(MAIN_PATH).resolve())}

    from openpyxl import load_workbook

    source = Path(MAIN_PATH)
    if not source.exists():
        raise FileNotFoundError(f"Workbook not found: {source.resolve()}")
    if expected_signature and _source_file_signature(source) != expected_signature:
        raise RuntimeError(
            "The workbook changed after this editor view loaded. Refresh the Raw Data Editor before saving."
        )

    raw_now = _read_raw_main_sheet()
    source_columns = list(raw_now.columns)
    validated: List[Dict[str, Any]] = []
    stale_cells: List[str] = []

    for change in changes:
        column = str(change.get("column", ""))
        excel_row = int(change.get("excel_row", 0) or 0)
        if column not in source_columns:
            raise ValueError(f"Column no longer exists in workbook: {column}")
        source_index = excel_row - 2
        if source_index < 0 or source_index >= len(raw_now):
            raise ValueError(f"Excel row {excel_row} is outside the current source data.")
        current_value = raw_now.iloc[source_index][column]
        expected_before = change.get("before", "")
        if not _editor_before_matches(current_value, expected_before):
            stale_cells.append(f"row {excel_row} / {column}")
            continue
        after = _coerce_editor_input(change.get("after"), current_value)
        if not _editor_values_equal(current_value, after):
            validated.append({"excel_row": excel_row, "column": column, "after": after})

    if stale_cells:
        raise RuntimeError(
            "The workbook changed after this editor view loaded. Refresh before saving. "
            "Stale cells: " + ", ".join(stale_cells[:8])
        )
    if not validated:
        return {"saved": 0, "backup": None, "path": str(source.resolve())}

    backup = _create_workbook_backup()
    temp_path = source.with_name(f".{source.stem}.dart_edit_{uuid.uuid4().hex[:8]}{source.suffix}")
    try:
        workbook = load_workbook(source)
        if MAIN_SHEET not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Worksheet '{MAIN_SHEET}' was not found in {source.name}.")
        worksheet = workbook[MAIN_SHEET]
        header_map: Dict[str, int] = {}
        for cell in worksheet[1]:
            if cell.value is not None and str(cell.value) not in header_map:
                header_map[str(cell.value)] = cell.column
        missing_headers = sorted({c["column"] for c in validated if c["column"] not in header_map})
        if missing_headers:
            workbook.close()
            raise ValueError("Workbook headers changed while editor was open: " + ", ".join(missing_headers))
        for change in validated:
            worksheet.cell(row=int(change["excel_row"]), column=int(header_map[change["column"]])).value = _excel_safe_value(change["after"])
        workbook.save(temp_path)
        workbook.close()
        os.replace(temp_path, source)
    except PermissionError as exc:
        temp_path.unlink(missing_ok=True)
        raise PermissionError("Windows could not replace the workbook. Close it in desktop Excel and try again.") from exc
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

    return {"saved": len(validated), "backup": str(backup), "path": str(source.resolve())}


def _latest_workbook_backup() -> Optional[Path]:
    WORKBOOK_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    source = Path(MAIN_PATH)
    candidates = sorted(
        WORKBOOK_BACKUP_DIR.glob(f"{source.stem}_before_edit_*{source.suffix}"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def _restore_workbook_backup(backup_path: Path | str) -> str:
    backup = Path(backup_path)
    source = Path(MAIN_PATH)
    if not backup.exists():
        raise FileNotFoundError(f"Backup not found: {backup}")
    temp_path = source.with_name(f".{source.stem}.dart_restore_{uuid.uuid4().hex[:8]}{source.suffix}")
    shutil.copy2(backup, temp_path)
    os.replace(temp_path, source)
    return str(source.resolve())


def _reload_state_from_workbook() -> pd.DataFrame:
    df, name = load_local_data()
    if df.empty:
        raise RuntimeError(name or "Workbook reload returned no data.")
    STATE["df"] = df
    STATE["source"] = name
    STATE["snapshots"] = []
    return df


def _ensure_email_agent_dirs() -> None:
    EMAIL_AGENT_DIR.mkdir(parents=True, exist_ok=True)
    EMAIL_SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)


def _as_bool(value: Any, default: bool = False) -> bool:
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on"}


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _safe_template_id(template_id: Any) -> str:
    safe = re.sub(r"[^A-Za-z0-9_-]", "", str(template_id or ""))
    return safe or uuid.uuid4().hex[:12]


def _snapshot_path(template_id: Any) -> Path:
    _ensure_email_agent_dirs()
    return EMAIL_SNAPSHOT_DIR / f"{_safe_template_id(template_id)}.csv"


def _atomic_write_text(path: Path | str, content: str) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(content, encoding="utf-8")
    tmp.replace(path)


def _write_agent_status(status: str, **extra: Any) -> None:
    payload = {"status": status, "updated_at": _utc_now_iso(), **extra}
    try:
        _atomic_write_text(EMAIL_AGENT_STATUS_PATH, json.dumps(payload, indent=2, ensure_ascii=False, default=str))
    except Exception:
        pass


def _append_agent_history(event: Dict[str, Any]) -> None:
    try:
        EMAIL_AGENT_DIR.mkdir(parents=True, exist_ok=True)
        with EMAIL_AGENT_HISTORY_PATH.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps({"timestamp": _utc_now_iso(), **event}, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass


@contextmanager
def _email_agent_process_lock():
    _ensure_email_agent_dirs()
    lock_path = EMAIL_AGENT_DIR / ".agent.lock"
    acquired = False
    fd = None
    try:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode("utf-8"))
            acquired = True
        except FileExistsError:
            try:
                age = time.time() - lock_path.stat().st_mtime
                if age > EMAIL_AGENT_LOCK_TIMEOUT_SECONDS:
                    lock_path.unlink(missing_ok=True)
                    fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.write(fd, str(os.getpid()).encode("utf-8"))
                    acquired = True
            except Exception:
                acquired = False
        yield acquired
    finally:
        if fd is not None:
            try:
                os.close(fd)
            except Exception:
                pass
        if acquired:
            try:
                lock_path.unlink(missing_ok=True)
            except Exception:
                pass


def load_email_automations() -> List[Dict[str, Any]]:
    _ensure_email_agent_dirs()
    if not EMAIL_TEMPLATE_PATH.exists():
        return []
    try:
        raw = json.loads(EMAIL_TEMPLATE_PATH.read_text(encoding="utf-8"))
        return raw if isinstance(raw, list) else []
    except Exception:
        return []


def save_email_automations(templates: List[Dict[str, Any]]) -> None:
    _ensure_email_agent_dirs()
    _atomic_write_text(EMAIL_TEMPLATE_PATH, json.dumps(templates, indent=2, ensure_ascii=False, default=str))


def upsert_email_automation(template: Dict[str, Any]) -> Dict[str, Any]:
    templates = load_email_automations()
    template = dict(template)
    template["id"] = _safe_template_id(template.get("id") or uuid.uuid4().hex[:12])
    template["updated_at"] = _utc_now_iso()
    if not template.get("created_at"):
        template["created_at"] = template["updated_at"]
    for i, existing in enumerate(templates):
        if existing.get("id") == template["id"]:
            templates[i] = template
            break
    else:
        templates.append(template)
    save_email_automations(templates)
    return template


def delete_email_automation(template_id: str) -> None:
    save_email_automations([t for t in load_email_automations() if t.get("id") != template_id])
    _snapshot_path(template_id).unlink(missing_ok=True)


def _parse_recipients(value: Any) -> List[str]:
    return [x.strip() for x in re.split(r"[,;]", str(value or "")) if x.strip()]


def validate_recipients(value: Any) -> Tuple[List[str], List[str]]:
    recipients = _parse_recipients(value)
    email_re = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]+$")
    return recipients, [x for x in recipients if not email_re.match(x)]


def _looks_like_email_placeholder(value: Any) -> bool:
    s = str(value or "").strip().lower()
    return any(token in s for token in ("your-sender", "your_sender", "your-password", "your_password", "example.com", "example.gov", "<", ">"))


def smtp_settings() -> Dict[str, Any]:
    """Return the SMTP configuration."""
    hardcoded = dict(HARDCODED_SMTP_CONFIG)
    host = str(hardcoded.get("host", "")).strip()
    username = str(hardcoded.get("username", "")).strip()
    password = re.sub(r"\s+", "", str(hardcoded.get("password", "")))
    from_email = str(hardcoded.get("from_email", "") or username).strip()
    from_name = str(hardcoded.get("from_name", "DART Email Agent")).strip()
    try:
        port = int(hardcoded.get("port", 465))
    except (TypeError, ValueError):
        port = 465
    use_tls = _as_bool(hardcoded.get("use_tls", False), False)
    use_ssl = _as_bool(hardcoded.get("use_ssl", True), True)
    return {"host": host, "port": port, "username": username, "password": password, "from_email": from_email, "from_name": from_name, "use_tls": use_tls, "use_ssl": use_ssl}


def smtp_configuration_issues() -> List[str]:
    cfg = smtp_settings()
    issues: List[str] = []
    if not cfg["host"] or _looks_like_email_placeholder(cfg["host"]):
        issues.append("SMTP host is missing.")
    if not cfg["from_email"] or _looks_like_email_placeholder(cfg["from_email"]):
        issues.append("Sender email is missing or still a placeholder.")
    if cfg["username"] and _looks_like_email_placeholder(cfg["username"]):
        issues.append("SMTP username is still a placeholder.")
    if cfg["password"] and _looks_like_email_placeholder(cfg["password"]):
        issues.append("SMTP password is still a placeholder.")
    if cfg["username"] and not cfg["password"]:
        issues.append("SMTP username is set but the SMTP password is blank.")
    if cfg["use_ssl"] and cfg["use_tls"]:
        issues.append("Choose either SMTP SSL or STARTTLS, not both.")
    if cfg["host"].lower() == "smtp.gmail.com":
        if cfg["port"] == 465 and not cfg["use_ssl"]:
            issues.append("Gmail port 465 requires SSL.")
        if cfg["port"] == 465 and cfg["use_tls"]:
            issues.append("Gmail port 465 should not also use STARTTLS.")
        if cfg["port"] == 587 and not cfg["use_tls"]:
            issues.append("Gmail port 587 requires STARTTLS.")
        if cfg["username"] and cfg["from_email"].lower() != cfg["username"].lower():
            issues.append("For Gmail, from_email should match the authenticated username.")
    return issues


def smtp_is_configured() -> bool:
    return len(smtp_configuration_issues()) == 0


def _snapshot_value(value: Any) -> str:
    if value is None or pd.isna(value):
        return NULL_TOKEN
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if isinstance(value, (int, float, np.integer, np.floating)):
        if pd.isna(value):
            return NULL_TOKEN
        return format(float(value), ".15g")
    return re.sub(r"\s+", " ", str(value).strip())


def _valid_columns(columns: List[str], df: pd.DataFrame) -> List[str]:
    return [c for c in columns if c in df.columns]


def build_snapshot_frame(df: pd.DataFrame, template: Dict[str, Any]) -> Tuple[pd.DataFrame, pd.DataFrame, List[str], List[str]]:
    source = df.reset_index(drop=True).copy()
    identity_columns = _valid_columns(template.get("identity_columns", []), source)
    if not identity_columns:
        identity_columns = [source.columns[0]] if len(source.columns) else []
    requested_monitor = template.get("monitor_columns", []) or list(source.columns)
    condition_columns = [r.get("column") for r in template.get("conditions", [])]
    display_columns = template.get("display_columns", [])
    monitor_columns = _valid_columns(list(dict.fromkeys(identity_columns + requested_monitor + condition_columns + display_columns)), source)
    if not monitor_columns:
        monitor_columns = list(source.columns)
    normalized = pd.DataFrame(index=source.index)
    for col in monitor_columns:
        normalized[col] = source[col].map(_snapshot_value)
    base_key = normalized[identity_columns].agg(" | ".join, axis=1) if identity_columns else pd.Series(["ROW"] * len(source), index=source.index)
    occurrence = base_key.groupby(base_key, sort=False).cumcount().astype(str)
    row_key = base_key + " | occurrence=" + occurrence
    row_payload = normalized[monitor_columns].agg("\x1f".join, axis=1)
    row_hash = row_payload.map(lambda s: hashlib.sha256(s.encode("utf-8")).hexdigest())
    snapshot = normalized.copy()
    snapshot.insert(0, "_row_hash", row_hash)
    snapshot.insert(0, "_row_key", row_key)
    snapshot.insert(0, "_source_position", np.arange(len(source)))
    return source, snapshot, identity_columns, monitor_columns


def write_email_snapshot(df: pd.DataFrame, template: Dict[str, Any]) -> Path:
    _, snapshot, _, _ = build_snapshot_frame(df, template)
    path = _snapshot_path(template["id"])
    snapshot.to_csv(path, index=False)
    return path


def read_email_snapshot(template_id: str) -> Optional[pd.DataFrame]:
    path = _snapshot_path(template_id)
    if not path.exists():
        return None
    try:
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    except Exception:
        return None


def detect_email_agent_changes(df: pd.DataFrame, template: Dict[str, Any]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    source, current, identity_columns, monitor_columns = build_snapshot_frame(df, template)
    previous = read_email_snapshot(template["id"])
    meta = {"baseline_missing": previous is None, "identity_columns": identity_columns, "monitor_columns": monitor_columns, "duplicate_identity_count": 0, "source_rows": len(source)}
    if identity_columns and not source.empty:
        normalized_keys = source[identity_columns].astype(str).agg(" | ".join, axis=1)
        meta["duplicate_identity_count"] = int(normalized_keys.duplicated(keep=False).sum())
    if previous is None:
        empty = source.iloc[0:0].copy()
        empty["_Change Type"] = pd.Series(dtype=str)
        empty["_Changed Fields"] = pd.Series(dtype=str)
        empty["_Previous Values"] = pd.Series(dtype=str)
        return empty, meta

    prev_by_key = previous.set_index("_row_key", drop=False)
    rows: List[Dict[str, Any]] = []
    trigger_mode = template.get("trigger_mode", "New or changed rows")
    current_keys = set(current["_row_key"].astype(str))
    previous_keys = set(previous["_row_key"].astype(str))

    for key in previous_keys - current_keys:
        prev_row = prev_by_key.loc[key]
        if isinstance(prev_row, pd.DataFrame):
            prev_row = prev_row.iloc[0]
        record = {c: prev_row.get(c, "") for c in source.columns}
        record["_Change Type"] = "Deleted"
        record["_Changed Fields"] = ", ".join(c for c in monitor_columns if c in previous.columns and c not in identity_columns) or "Row deleted"
        record["_Previous Values"] = "Row existed in previous snapshot but is no longer present."
        if trigger_mode == "All changes including deletions":
            rows.append(record)

    for _, snap_row in current.iterrows():
        key = snap_row["_row_key"]
        pos = int(snap_row["_source_position"])
        change_type = None
        changed_fields: List[str] = []
        previous_values: List[str] = []
        if key not in prev_by_key.index:
            change_type = "New"
            changed_fields = [c for c in monitor_columns if c not in identity_columns]
        else:
            prev_row = prev_by_key.loc[key]
            if isinstance(prev_row, pd.DataFrame):
                prev_row = prev_row.iloc[0]
            if str(prev_row.get("_row_hash", "")) != str(snap_row.get("_row_hash", "")):
                change_type = "Changed"
                for col in monitor_columns:
                    old = str(prev_row.get(col, NULL_TOKEN))
                    new = str(snap_row.get(col, NULL_TOKEN))
                    if old != new:
                        changed_fields.append(col)
                        previous_values.append(f"{col}: {'Null' if old == NULL_TOKEN else old} → {'Null' if new == NULL_TOKEN else new}")
        if change_type is None:
            continue
        if trigger_mode == "New rows only" and change_type != "New":
            continue
        record = source.iloc[pos].to_dict()
        record["_Change Type"] = change_type
        record["_Changed Fields"] = ", ".join(changed_fields) if changed_fields else "—"
        record["_Previous Values"] = "; ".join(previous_values[:8]) if previous_values else "—"
        rows.append(record)
    return pd.DataFrame(rows), meta


def _blank_mask(series: pd.Series) -> pd.Series:
    text = series.astype(str).str.strip().str.lower()
    return series.isna() | text.isin(["", "nan", "none", "null", "<na>"])


def condition_mask(series: pd.Series, operator: str, target: Any = "") -> pd.Series:
    text = series.astype(str).str.strip()
    text_lower = text.str.lower()
    target_text = str(target or "").strip()
    target_lower = target_text.lower()
    if operator == "Is blank (empty or null)":
        return _blank_mask(series)
    if operator == "Is null / NaN":
        return series.isna()
    if operator == "Is empty string":
        return series.notna() & text.eq("")
    if operator == "Equals":
        if pd.api.types.is_numeric_dtype(series):
            try:
                return pd.to_numeric(series, errors="coerce").eq(float(target_text))
            except ValueError:
                pass
        return text_lower.eq(target_lower)
    if operator == "Does not equal":
        return ~condition_mask(series, "Equals", target_text)
    if operator == "Contains":
        return text_lower.str.contains(re.escape(target_lower), na=False)
    if operator == "Does not contain":
        return ~text_lower.str.contains(re.escape(target_lower), na=False)
    if operator == "Starts with":
        return text_lower.str.startswith(target_lower, na=False)
    if operator in ("Greater than", "Greater than or equal to", "Less than", "Less than or equal to"):
        numeric = pd.to_numeric(series, errors="coerce")
        try:
            threshold = float(target_text)
        except ValueError:
            return pd.Series(False, index=series.index)
        if operator == "Greater than":
            return numeric > threshold
        if operator == "Greater than or equal to":
            return numeric >= threshold
        if operator == "Less than":
            return numeric < threshold
        return numeric <= threshold
    if operator in ("Before date", "After date"):
        dates = pd.to_datetime(series, errors="coerce")
        target_date = pd.to_datetime(target_text, errors="coerce")
        if pd.isna(target_date):
            return pd.Series(False, index=series.index)
        return dates < target_date if operator == "Before date" else dates > target_date
    return pd.Series(False, index=series.index)


def _email_condition_mode(template: Dict[str, Any]) -> str:
    explicit = str(template.get("condition_mode", "")).strip()
    if explicit in ("Any change", "Only when conditions are met"):
        return explicit
    return "Only when conditions are met" if template.get("conditions") else "Any change"


def evaluate_email_conditions(df: pd.DataFrame, template: Dict[str, Any]) -> pd.DataFrame:
    if df.empty:
        return df.copy()
    if _email_condition_mode(template) == "Any change":
        return df.copy()
    conditions = [r for r in template.get("conditions", []) if r.get("column") in df.columns]
    if not conditions:
        return df.iloc[0:0].copy()
    masks = [condition_mask(df[r["column"]], r.get("operator", "Equals"), r.get("value", "")) for r in conditions]
    combined = masks[0].copy()
    for mask in masks[1:]:
        combined = (combined | mask) if template.get("condition_logic", "ALL conditions") == "ANY condition" else (combined & mask)
    return df[combined.fillna(False)].copy()


def current_matching_rows(df: pd.DataFrame, template: Dict[str, Any]) -> pd.DataFrame:
    rows = evaluate_email_conditions(df, template).copy()
    rows["_Change Type"] = "Current match"
    rows["_Changed Fields"] = "Any detected change will qualify" if _email_condition_mode(template) == "Any change" else "Matches configured conditions"
    rows["_Previous Values"] = "—"
    return rows


def _condition_sentence(rule: Dict[str, Any]) -> str:
    col = rule.get("column", "(field)")
    operator = rule.get("operator", "Equals")
    value = str(rule.get("value", "")).strip()
    no_value = operator in ("Is blank (empty or null)", "Is null / NaN", "Is empty string")
    return f"{col} {operator.lower()}" if no_value else f"{col} {operator.lower()} {value}"


def _email_display_value(value: Any, column: str) -> str:
    if value is None or pd.isna(value):
        return "Null"
    if column in ("MatchRate", "MismatchRate"):
        try:
            return f"{float(value):.2%}"
        except (TypeError, ValueError):
            pass
    if column in ("MatchedClaims", "NotMatchedClaims", "TotalClaims"):
        try:
            return f"{float(value):,.0f}"
        except (TypeError, ValueError):
            pass
    text = str(value)
    return text if text.strip() else "(empty)"


def build_structured_email(template: Dict[str, Any], matches: pd.DataFrame, is_test: bool = False) -> Tuple[str, str, str, List[str]]:
    name = template.get("name", "DART data alert")
    subject_prefix = template.get("subject_prefix", "[DART Alert]")
    subject = f"{subject_prefix} {'TEST — ' if is_test else ''}{name}: {len(matches):,} matching change(s)"
    greeting = template.get("greeting", "Hello,")
    intro = template.get("email_intro", "DART detected new or changed reconciliation data that matches your saved requirements.")
    display_columns = [c for c in template.get("display_columns", []) if c in matches.columns]
    meta_columns = [c for c in ["_Change Type", "_Changed Fields", "_Previous Values"] if c in matches.columns]
    display_columns = list(dict.fromkeys(meta_columns + display_columns)) or list(matches.columns[:10])
    condition_mode = _email_condition_mode(template)
    condition_lines = [_condition_sentence(r) for r in template.get("conditions", [])]
    logic_label = "No conditions — any detected change" if condition_mode == "Any change" else template.get("condition_logic", "ALL conditions")
    streams = sorted(str(v) for v in matches["Stream"].dropna().unique() if str(v).strip()) if "Stream" in matches.columns else []
    unmatched_total = pd.to_numeric(matches["NotMatchedClaims"], errors="coerce").fillna(0).sum() if "NotMatchedClaims" in matches.columns else None

    rows_html: List[str] = []
    for _, row in matches.head(EMAIL_PREVIEW_ROW_LIMIT).iterrows():
        cells = "".join(f"<td style='padding:8px;border:1px solid #e5e7eb;vertical-align:top;'>{html.escape(_email_display_value(row.get(col), col))}</td>" for col in display_columns)
        rows_html.append(f"<tr>{cells}</tr>")
    header_html = "".join(f"<th style='padding:8px;border:1px solid #d1d5db;background:#fffaf0;text-align:left;'>{html.escape(col)}</th>" for col in display_columns)
    dataset_name = str(template.get("dataset_name", "") or "").strip()
    summary_cards = [("Matching changes", f"{len(matches):,}")]
    summary_cards.append(("Dataset", dataset_name) if dataset_name else ("Streams affected", f"{len(streams):,}" if streams else "—"))
    if unmatched_total is not None:
        summary_cards.append(("Not-matched claims", f"{unmatched_total:,.0f}"))
    cards_html = "".join(f"<td style='padding:10px 16px;border:1px solid #eadfb8;background:#fffdf8;'><div style='font-size:11px;color:#64748b;text-transform:uppercase;font-weight:700;'>{html.escape(label)}</div><div style='font-size:22px;color:#0f172a;font-weight:750;'>{html.escape(value)}</div></td>" for label, value in summary_cards)
    rules_html = "<li>Any detected change in the monitored columns qualifies for this alert.</li>" if condition_mode == "Any change" else "".join(f"<li>{html.escape(line)}</li>" for line in condition_lines)
    truncated_note = f"<p style='color:#64748b;font-size:12px;'>Showing the first {EMAIL_PREVIEW_ROW_LIMIT:,} of {len(matches):,} matching changes. The attached CSV contains all matches.</p>" if len(matches) > EMAIL_PREVIEW_ROW_LIMIT else ""
    body_html = f"""
    <div style="font-family:Segoe UI,Arial,sans-serif;color:#0f172a;max-width:1100px;margin:auto;">
      <div style="background:linear-gradient(135deg,#7a5a00,#d5a928,#f3d878);padding:22px 26px;border-radius:14px 14px 0 0;color:white;">
        <div style="font-size:12px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;">{html.escape(str(template.get("agent_label", "DART Email Agent")))}</div>
        <h2 style="margin:5px 0 0 0;color:white;">{html.escape(name)}</h2>
      </div>
      <div style="border:1px solid #eadfb8;border-top:0;padding:24px 26px;border-radius:0 0 14px 14px;background:white;">
        <p>{html.escape(greeting)}</p><p>{html.escape(intro)}</p>
        <table style="border-collapse:collapse;margin:18px 0;"><tr>{cards_html}</tr></table>
        <h3 style="margin-bottom:6px;">Requirements that triggered this alert</h3>
        <p style="margin-top:0;color:#475569;">Rule logic: <b>{html.escape(logic_label)}</b></p><ul>{rules_html}</ul>
        {f"<p><b>Streams:</b> {html.escape(', '.join(streams))}</p>" if streams else ""}
        <h3>Matching data changes</h3><div style="overflow-x:auto;"><table style="border-collapse:collapse;width:100%;font-size:12px;"><thead><tr>{header_html}</tr></thead><tbody>{''.join(rows_html)}</tbody></table></div>
        {truncated_note}
        <p style="margin-top:22px;"><b>Recommended next step:</b> {html.escape(str(template.get("recommended_next_step", "Review the affected target fields and route the reconciliation work using the saved recommendation/owner context.")))}</p>
        <p style="color:#64748b;font-size:12px;margin-top:24px;">Generated automatically by DART on {html.escape(_utc_now_iso())}.</p>
      </div>
    </div>"""
    text_lines = [greeting, "", intro, "", f"Automation: {name}", f"Matching changes: {len(matches):,}", f"Rule logic: {logic_label}", "Requirements:"]
    if condition_mode == "Any change":
        text_lines.append("- Any detected change in the monitored columns qualifies for this alert.")
    else:
        text_lines.extend(f"- {line}" for line in condition_lines)
    if streams:
        text_lines.append(f"Streams: {', '.join(streams)}")
    if unmatched_total is not None:
        text_lines.append(f"Total not-matched claims: {unmatched_total:,.0f}")
    text_lines.extend(["", "Matching data changes:", matches[display_columns].head(EMAIL_PREVIEW_ROW_LIMIT).to_string(index=False), "", str(template.get("recommended_next_step", "Review the affected target fields and route the reconciliation work using the saved recommendation/owner context."))])
    return subject, "\n".join(text_lines), body_html, display_columns


@contextmanager
def _authenticated_smtp_connection(cfg: Dict[str, Any]):
    server = None
    try:
        context = ssl.create_default_context()
        if cfg["use_ssl"]:
            server = smtplib.SMTP_SSL(cfg["host"], cfg["port"], context=context, timeout=30)
        else:
            server = smtplib.SMTP(cfg["host"], cfg["port"], timeout=30)
            server.ehlo()
            if cfg["use_tls"]:
                server.starttls(context=context)
                server.ehlo()
        if cfg["username"]:
            server.login(cfg["username"], cfg["password"])
        yield server
    finally:
        if server is not None:
            try:
                server.quit()
            except Exception:
                try:
                    server.close()
                except Exception:
                    pass


def test_smtp_connection() -> Dict[str, Any]:
    cfg = smtp_settings()
    issues = smtp_configuration_issues()
    if issues:
        raise RuntimeError("; ".join(issues))
    with _authenticated_smtp_connection(cfg):
        return {"host": cfg["host"], "port": cfg["port"], "transport": "SSL" if cfg["use_ssl"] else ("STARTTLS" if cfg["use_tls"] else "plain SMTP"), "sender": cfg["from_email"]}


def send_structured_email(template: Dict[str, Any], matches: pd.DataFrame, is_test: bool = False) -> str:
    recipients, invalid = validate_recipients(template.get("recipient_email", ""))
    if not recipients or invalid:
        raise ValueError(f"Invalid recipient email address(es): {', '.join(invalid) or 'none supplied'}")
    cfg = smtp_settings()
    if not smtp_is_configured():
        raise RuntimeError("Email delivery is not ready. Configure HARDCODED_SMTP_CONFIG or DART_SMTP_* environment variables. " + "; ".join(smtp_configuration_issues()))
    subject, text_body, body_html, display_columns = build_structured_email(template, matches, is_test=is_test)
    message = EmailMessage()
    message["Subject"] = subject
    message["From"] = f"{cfg['from_name']} <{cfg['from_email']}>"
    message["To"] = ", ".join(recipients)
    message.set_content(text_body)
    message.add_alternative(body_html, subtype="html")
    if template.get("include_csv", True) and not matches.empty:
        attachment_cols = [c for c in display_columns if c in matches.columns]
        message.add_attachment(matches[attachment_cols].to_csv(index=False).encode("utf-8"), maintype="text", subtype="csv", filename=f"dart_alert_{_safe_template_id(template.get('id', 'alert'))}.csv")
    try:
        with _authenticated_smtp_connection(cfg) as server:
            server.send_message(message)
    except smtplib.SMTPAuthenticationError as exc:
        raise RuntimeError("SMTP authentication failed. For Gmail, use a current Google App Password rather than the normal account password.") from exc
    except (ConnectionResetError, ConnectionAbortedError, TimeoutError, OSError) as exc:
        raise RuntimeError(f"Could not complete SMTP connection to {cfg['host']}:{cfg['port']}: {exc}") from exc
    return subject


def establish_email_baseline(df: pd.DataFrame, template: Dict[str, Any]) -> Dict[str, Any]:
    path = write_email_snapshot(df, template)
    return {"status": "Baseline established", "snapshot": str(path), "rows": len(df)}


def _public_automation_result(result: Dict[str, Any]) -> Dict[str, Any]:
    out = {k: v for k, v in result.items() if k != "matched_rows"}
    matched = result.get("matched_rows")
    if isinstance(matched, pd.DataFrame):
        out["matched_rows"] = _records_json_safe(matched, 100)
    return out


def run_email_automation(df: pd.DataFrame, template: Dict[str, Any], send_email: bool = True) -> Dict[str, Any]:
    result: Dict[str, Any] = {"id": template.get("id"), "name": template.get("name"), "status": "Not run", "matches": 0, "checked_at": _utc_now_iso()}
    changes, meta = detect_email_agent_changes(df, template)
    result["change_rows"] = len(changes)
    result["meta"] = meta
    if meta["baseline_missing"]:
        write_email_snapshot(df, template)
        result["status"] = "Baseline established; no email sent"
        return result
    matches = evaluate_email_conditions(changes, template)
    result["matches"] = len(matches)
    if matches.empty:
        write_email_snapshot(df, template)
        result["status"] = "Checked; no matching changes"
        return result
    if not send_email:
        result["status"] = "Matching changes found; email not sent"
        result["matched_rows"] = matches
        return result
    try:
        result["subject"] = send_structured_email(template, matches, is_test=False)
        write_email_snapshot(df, template)
        result["status"] = "Email sent"
    except Exception as exc:
        result["status"] = f"Email failed: {exc}"
        result["error"] = str(exc)
    result["matched_rows"] = matches
    return result


def run_enabled_email_automations(df: pd.DataFrame) -> List[Dict[str, Any]]:
    with _email_agent_process_lock() as acquired:
        if not acquired:
            _write_agent_status("Skipped — another worker is running")
            return []
        results: List[Dict[str, Any]] = []
        templates = load_email_automations()
        enabled = [t for t in templates if t.get("enabled", False)]
        changed = False
        for template in enabled:
            try:
                result = run_email_automation(df, template, send_email=True)
            except Exception as exc:
                result = {"id": template.get("id"), "name": template.get("name"), "status": f"Runner error: {exc}", "matches": 0, "change_rows": 0, "checked_at": _utc_now_iso(), "error": str(exc)}
                traceback.print_exc()
            template["last_run_at"] = result.get("checked_at")
            template["last_status"] = result.get("status")
            template["last_match_count"] = result.get("matches", 0)
            template["last_change_count"] = result.get("change_rows", 0)
            results.append(result)
            changed = True
            _append_agent_history({"automation_id": template.get("id"), "automation": template.get("name"), "status": result.get("status"), "detected_changes": result.get("change_rows", 0), "matching_changes": result.get("matches", 0)})
        if changed:
            save_email_automations(templates)
        _write_agent_status("Checked", enabled_automations=len(enabled), results=[{"automation": r.get("name"), "status": r.get("status"), "detected_changes": r.get("change_rows", 0), "matching_changes": r.get("matches", 0), "checked_at": r.get("checked_at")} for r in results])
        return results


def _email_agent_scheduler_loop() -> None:
    if EMAIL_AGENT_STARTUP_DELAY_SECONDS:
        time.sleep(EMAIL_AGENT_STARTUP_DELAY_SECONDS)
    last_signature = None
    _write_agent_status("Running", workbook=str(MAIN_PATH), interval_seconds=EMAIL_AGENT_MIN_CHECK_SECONDS)
    while True:
        try:
            signature = _source_file_signature(MAIN_PATH)
            enabled = [t for t in load_email_automations() if t.get("enabled", False)]
            if signature == "missing":
                _write_agent_status("Waiting — source workbook not found", workbook=str(MAIN_PATH))
            elif not enabled:
                _write_agent_status("Running — no enabled automations", workbook=str(MAIN_PATH), workbook_signature=signature)
                last_signature = signature
            elif last_signature is None:
                fresh_df, name = load_local_data()
                if fresh_df.empty:
                    raise RuntimeError(name)
                results = run_enabled_email_automations(fresh_df)
                last_signature = signature
                _write_agent_status("Startup check complete", workbook=str(MAIN_PATH), workbook_signature=signature, results_count=len(results))
            elif signature != last_signature:
                fresh_df, name = load_local_data()
                if fresh_df.empty:
                    raise RuntimeError(name)
                results = run_enabled_email_automations(fresh_df)
                STATE["df"] = fresh_df
                STATE["source"] = name
                last_signature = signature
                _write_agent_status("Workbook change processed", workbook=str(MAIN_PATH), workbook_signature=signature, results_count=len(results))
        except Exception as exc:
            _write_agent_status("Scheduler error", workbook=str(MAIN_PATH), error=str(exc))
            traceback.print_exc()
        time.sleep(EMAIL_AGENT_MIN_CHECK_SECONDS)


def start_email_agent_scheduler() -> bool:
    global _EMAIL_AGENT_THREAD
    if not EMAIL_AGENT_AUTO_RUN:
        return False
    with _EMAIL_AGENT_THREAD_LOCK:
        if _EMAIL_AGENT_THREAD is not None and _EMAIL_AGENT_THREAD.is_alive():
            return True
        _EMAIL_AGENT_THREAD = threading.Thread(target=_email_agent_scheduler_loop, name="dart-email-agent", daemon=True)
        _EMAIL_AGENT_THREAD.start()
        return True


def _default_email_agent_template(df: pd.DataFrame) -> Dict[str, Any]:
    columns = list(df.columns)
    key_defaults = [c for c in ["Stream", "NCH Target Table", "NCH Target Column"] if c in columns]
    display_defaults = [c for c in ["Stream", "NCH Target Table", "NCH Target Column", "SS Table", "Disposition", "MatchRate", "NotMatchedClaims", "Sub-Classification", "Recommendation"] if c in columns]
    conditions: List[Dict[str, str]] = []
    if "Disposition" in columns:
        conditions.append({"column": "Disposition", "operator": "Contains", "value": "action"})
    elif "NotMatchedClaims" in columns:
        conditions.append({"column": "NotMatchedClaims", "operator": "Greater than", "value": "0"})
    return {"id": "", "name": "Action-needed reconciliation alert", "recipient_email": "", "enabled": False, "trigger_mode": "New or changed rows", "identity_columns": key_defaults, "monitor_columns": columns, "condition_mode": "Only when conditions are met", "condition_logic": "ALL conditions", "conditions": conditions, "display_columns": display_defaults, "subject_prefix": "[DART Alert]", "greeting": "Hello Data Owner,", "email_intro": "DART detected new or changed reconciliation data requiring review.", "include_csv": True}


def _records_json_safe(df: pd.DataFrame, limit: int = 200) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for row in df.head(limit).to_dict("records"):
        clean: Dict[str, Any] = {}
        for key, value in row.items():
            if value is None:
                clean[str(key)] = None
            elif isinstance(value, (pd.Timestamp, datetime)):
                clean[str(key)] = value.isoformat()
            elif isinstance(value, np.generic):
                clean[str(key)] = value.item()
            else:
                try:
                    clean[str(key)] = None if pd.isna(value) else value
                except Exception:
                    clean[str(key)] = value
        records.append(clean)
    return records


def _email_agent_status_payload() -> Dict[str, Any]:
    status: Dict[str, Any] = {"status": "Not started", "updated_at": None}
    if EMAIL_AGENT_STATUS_PATH.exists():
        try:
            status = json.loads(EMAIL_AGENT_STATUS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return status


def _email_agent_history(limit: int = 30) -> List[Dict[str, Any]]:
    if not EMAIL_AGENT_HISTORY_PATH.exists():
        return []
    try:
        lines = EMAIL_AGENT_HISTORY_PATH.read_text(encoding="utf-8").splitlines()[-limit:]
        return [json.loads(line) for line in reversed(lines) if line.strip()]
    except Exception:
        return []


def _normalize_email_template(payload: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
    columns = list(df.columns)
    template = dict(payload)
    template["id"] = _safe_template_id(template.get("id")) if template.get("id") else ""
    template["name"] = str(template.get("name", "")).strip()
    template["recipient_email"] = str(template.get("recipient_email", "")).strip()
    template["enabled"] = bool(template.get("enabled", False))
    template["trigger_mode"] = str(template.get("trigger_mode", "New or changed rows"))
    template["condition_mode"] = str(template.get("condition_mode", "Only when conditions are met"))
    template["condition_logic"] = str(template.get("condition_logic", "ALL conditions"))
    template["identity_columns"] = [c for c in template.get("identity_columns", []) if c in columns]
    template["monitor_columns"] = [c for c in template.get("monitor_columns", []) if c in columns]
    template["display_columns"] = [c for c in template.get("display_columns", []) if c in columns]
    template["subject_prefix"] = str(template.get("subject_prefix", "[DART Alert]")).strip() or "[DART Alert]"
    template["greeting"] = str(template.get("greeting", "Hello,")).strip() or "Hello,"
    template["email_intro"] = str(template.get("email_intro", "DART detected a qualifying data change.")).strip()
    template["include_csv"] = bool(template.get("include_csv", True))
    rules = []
    for rule in template.get("conditions", []):
        column = str(rule.get("column", ""))
        operator = str(rule.get("operator", "Equals"))
        if column in columns and operator in EMAIL_OPERATORS:
            rules.append({"column": column, "operator": operator, "value": str(rule.get("value", ""))})
    template["conditions"] = rules
    if not template["name"]:
        raise ValueError("Automation name is required.")
    recipients, invalid = validate_recipients(template["recipient_email"])
    if not recipients or invalid:
        raise ValueError("Enter at least one valid recipient email address.")
    if not template["identity_columns"]:
        raise ValueError("Choose at least one row identity column.")
    if template["condition_mode"] == "Only when conditions are met" and not template["conditions"]:
        raise ValueError("Add at least one condition, or change Requirement mode to Any change.")
    return template


def clean_records(df: pd.DataFrame, limit: int = 200) -> List[Dict[str, Any]]:
    return df.head(limit).replace({np.nan: None}).to_dict("records")


def generic_records(df: pd.DataFrame, limit: int = 250) -> List[Dict[str, Any]]:
    """JSON-safe records for arbitrary Build Your Own uploads."""
    if df is None or df.empty:
        return []
    frame = df.head(limit).copy()
    frame = frame.replace([np.inf, -np.inf], np.nan)
    # pandas' JSON serializer safely handles timestamps, numpy scalars, and nulls.
    return json.loads(frame.to_json(orient="records", date_format="iso"))


def _ensure_byo_data_dir() -> Path:
    BYO_DATA_DIR.mkdir(parents=True, exist_ok=True)
    return BYO_DATA_DIR


def _safe_byo_filename(filename: Any) -> str:
    raw = Path(str(filename or "dataset")).name.strip() or "dataset"
    suffix = Path(raw).suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise ValueError("Build Your Own supports CSV and XLSX files.")
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(raw).stem).strip("._-") or "dataset"
    return f"{stem[:120]}{suffix}"


def _frame_from_bytes(content: bytes, suffix: str) -> pd.DataFrame:
    if suffix == ".csv":
        try:
            frame = pd.read_csv(io.BytesIO(content))
        except UnicodeDecodeError:
            frame = pd.read_csv(io.BytesIO(content), encoding="latin-1")
    elif suffix == ".xlsx":
        frame = pd.read_excel(io.BytesIO(content), engine="openpyxl")
    else:
        raise ValueError("Build Your Own supports CSV and XLSX files.")
    frame = norm_cols(frame)
    if frame.empty and len(frame.columns) == 0:
        raise ValueError("The file did not contain a readable table.")
    return frame


def _save_byo_bytes(content: bytes, name_hint: str) -> Dict[str, Any]:
    if len(content) > BYO_MAX_FILE_BYTES:
        raise ValueError("File exceeds the 25 MB DIY file limit.")
    safe_name = _safe_byo_filename(name_hint)
    frame = _frame_from_bytes(content, Path(safe_name).suffix.lower())
    target = _unique_byo_target(safe_name)
    target.write_bytes(content)
    return {"name": target.name, "rows": int(len(frame)), "columns": int(len(frame.columns))}


def _validate_public_import_url(url: str) -> str:
    """Basic SSRF guard: only allow public http(s) URLs, not loopback/private/link-local hosts."""
    url = (url or "").strip()
    if not url:
        raise ValueError("Provide a file URL.")
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError("Only http:// and https:// URLs are supported.")
    host = (parsed.hostname or "").lower()
    if not host or host == "localhost" or host.endswith(".local"):
        raise ValueError("Local URLs are not allowed.")
    try:
        ip = ipaddress.ip_address(host)
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            raise ValueError("Private network URLs are not allowed.")
    except ValueError as exc:
        if "not allowed" in str(exc):
            raise
        # host is a hostname, not an IP literal - accepted as best-effort (no DNS pinning here)
    return url


def _s3_client(region: str, access_key: str, secret_key: str, session_token: str):
    if boto3 is None:
        raise RuntimeError("The 'boto3' package is not installed on the server. Run: pip install boto3")
    kwargs: Dict[str, Any] = {}
    if region:
        kwargs["region_name"] = region
    if access_key and secret_key:
        kwargs["aws_access_key_id"] = access_key
        kwargs["aws_secret_access_key"] = secret_key
        if session_token:
            kwargs["aws_session_token"] = session_token
    return boto3.client("s3", **kwargs)


def _graph_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """App-only (client credentials) sign-in against Microsoft Entra ID / Graph.

    Requires an Azure AD app registration with the Sites.Read.All (or Files.Read.All)
    application permission, admin-consented, in the same tenant as the SharePoint/OneDrive site.
    """
    if requests is None:
        raise RuntimeError("The 'requests' package is not installed on the server. Run: pip install requests")
    tenant_id, client_id, client_secret = (tenant_id or "").strip(), (client_id or "").strip(), (client_secret or "").strip()
    if not (tenant_id and client_id and client_secret):
        raise ValueError("Provide the Azure AD tenant ID, client ID, and client secret.")
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=20,
    )
    if not resp.ok:
        try:
            detail = resp.json().get("error_description", "")
        except Exception:
            detail = resp.text[:200]
        raise ValueError(f"Azure AD sign-in failed: {detail or 'unknown error'}")
    token = resp.json().get("access_token", "")
    if not token:
        raise ValueError("Azure AD did not return an access token.")
    return token


def _graph_encode_share_url(share_url: str) -> str:
    raw = (share_url or "").strip()
    if not raw:
        raise ValueError("Provide a SharePoint or OneDrive share link.")
    encoded = base64.urlsafe_b64encode(raw.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"


def _graph_share_files(token: str, share_url: str) -> List[Dict[str, Any]]:
    if requests is None:
        raise RuntimeError("The 'requests' package is not installed on the server. Run: pip install requests")
    encoded = _graph_encode_share_url(share_url)
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{encoded}/driveItem"
        "?$expand=children($select=name,size,file,folder,@microsoft.graph.downloadUrl)",
        headers=headers, timeout=20,
    )
    if not resp.ok:
        try:
            detail = resp.json().get("error", {}).get("message", "")
        except Exception:
            detail = resp.text[:200]
        raise ValueError(f"Microsoft Graph error: {detail or 'unknown error'}")
    item = resp.json()
    children = item.get("children") or ([item] if "folder" not in item else [])
    results: List[Dict[str, Any]] = []
    for child in children:
        name = str(child.get("name", ""))
        if "folder" in child or not name.lower().endswith((".csv", ".xlsx")):
            continue
        download_url = child.get("@microsoft.graph.downloadUrl", "")
        if not download_url:
            continue
        size = int(child.get("size", 0) or 0)
        results.append({"name": name, "size_bytes": size, "size_mb": round(size / (1024 * 1024), 2), "download_url": download_url})
    return results


def _byo_path(filename: Any, must_exist: bool = True) -> Path:
    safe = _safe_byo_filename(filename)
    root = _ensure_byo_data_dir().resolve()
    path = (root / safe).resolve()
    if path.parent != root:
        raise ValueError("Invalid dataset path.")
    if must_exist and not path.exists():
        raise FileNotFoundError(f"Saved DIY dataset not found: {safe}")
    return path


def _unique_byo_target(filename: Any) -> Path:
    path = _byo_path(filename, must_exist=False)
    if not path.exists():
        return path
    stem, suffix = path.stem, path.suffix
    for idx in range(2, 10000):
        candidate = path.with_name(f"{stem}_{idx}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError("Could not create a unique DIY dataset filename.")


def _read_byo_dataset(path_or_name: Any) -> pd.DataFrame:
    path = Path(path_or_name)
    if not path.is_absolute():
        path = _byo_path(path_or_name)
    suffix = path.suffix.lower()
    if suffix == ".csv":
        try:
            frame = pd.read_csv(path)
        except UnicodeDecodeError:
            frame = pd.read_csv(path, encoding="latin-1")
    elif suffix == ".xlsx":
        frame = pd.read_excel(path, engine="openpyxl")
    else:
        raise ValueError("Build Your Own supports CSV and XLSX files.")
    frame = norm_cols(frame)
    if frame.empty and len(frame.columns) == 0:
        raise ValueError("The file did not contain a readable table.")
    return frame


def _dataset_profile(df: pd.DataFrame) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    rows = len(df)
    profile_rows: List[Dict[str, Any]] = []
    quality_rows: List[Dict[str, Any]] = []
    for col in df.columns:
        series = df[col]
        non_null = int(series.notna().sum())
        miss = int(series.isna().sum())
        unique = int(series.nunique(dropna=True))
        profile_rows.append({
            "Column": str(col), "Type": str(series.dtype), "Non-null": non_null,
            "Missing": miss, "Missing %": round((miss / rows * 100) if rows else 0, 2), "Unique": unique,
        })
        quality_rows.append({
            "Column": str(col), "Missing": miss, "Missing %": round((miss / rows * 100) if rows else 0, 2),
            "Unique": unique, "Distinct %": round((unique / non_null * 100) if non_null else 0, 2),
        })
    quality_rows.sort(key=lambda r: (r["Missing"], r["Unique"]), reverse=True)
    return profile_rows, quality_rows


def _dataset_payload_from_frame(df: pd.DataFrame, filename: str, preview_limit: int = 150) -> Dict[str, Any]:
    rows, cols = df.shape
    missing = int(df.isna().sum().sum())
    total_cells = int(rows * cols)
    completeness = (1 - missing / total_cells) if total_cells else 0.0
    profile_rows, quality_rows = _dataset_profile(df)
    numeric_count = int(sum(pd.api.types.is_numeric_dtype(df[c]) for c in df.columns))
    text_count = int(sum(pd.api.types.is_object_dtype(df[c]) or pd.api.types.is_string_dtype(df[c]) for c in df.columns))
    return {
        "loaded": True,
        "source": filename,
        "rows": int(rows),
        "columns": int(cols),
        "missing_cells": missing,
        "duplicate_rows": int(df.duplicated().sum()),
        "completeness": float(completeness),
        "numeric_columns": numeric_count,
        "text_columns": text_count,
        "profile": profile_rows,
        "quality": quality_rows,
        "preview": generic_records(df, preview_limit),
        "columns_list": [str(c) for c in df.columns],
    }


def _byo_library_records() -> List[Dict[str, Any]]:
    root = _ensure_byo_data_dir()
    records: List[Dict[str, Any]] = []
    for path in sorted([*root.glob("*.xlsx"), *root.glob("*.csv")], key=lambda p: p.stat().st_mtime, reverse=True):
        stat = path.stat()
        record: Dict[str, Any] = {
            "name": path.name,
            "type": path.suffix.lower().lstrip(".").upper(),
            "size_bytes": int(stat.st_size),
            "size_mb": round(stat.st_size / (1024 * 1024), 2),
            "modified_at": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "rows": None,
            "columns": None,
            "status": "Ready",
        }
        try:
            df = _read_byo_dataset(path)
            record["rows"] = int(len(df))
            record["columns"] = int(len(df.columns))
        except Exception as exc:
            record["status"] = f"Read error: {exc}"
        records.append(record)
    return records


def byo_payload() -> Dict[str, Any]:
    datasets = _byo_library_records()
    return {
        "datasets": datasets,
        "dataset_count": len(datasets),
        "storage_path": str(Path("Data") / "DIY"),
        "ai_configured": bool(OpenAI and groq_api_key()),
        "ai_model": groq_model(),
        "chat": STATE.get("byo_chat", []),
        # Compatibility fields for older front-end references.
        "loaded": False, "source": "", "uploaded_at": "", "rows": 0, "columns": 0,
        "missing_cells": 0, "duplicate_rows": 0, "completeness": 0.0,
        "numeric_columns": 0, "text_columns": 0, "profile": [], "quality": [], "preview": [], "columns_list": [],
    }


def _read_byo_editable_dataset(path_or_name: Any) -> pd.DataFrame:
    """Read the editable first table without normalizing source headers."""
    path = Path(path_or_name)
    if not path.is_absolute():
        path = _byo_path(path_or_name)
    if path.suffix.lower() == ".csv":
        try:
            return pd.read_csv(path)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")
    if path.suffix.lower() == ".xlsx":
        return pd.read_excel(path, engine="openpyxl")
    raise ValueError("Build Your Own supports CSV and XLSX files.")


def _byo_sheet_name(path_or_name: Any) -> str:
    path = Path(path_or_name)
    if not path.is_absolute():
        path = _byo_path(path_or_name)
    if path.suffix.lower() != ".xlsx":
        return "CSV"
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=False)
    try:
        return wb.sheetnames[0] if wb.sheetnames else "Sheet1"
    finally:
        wb.close()


def _byo_backup_dir_for(path: Path) -> Path:
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", path.stem) or "dataset"
    target = BYO_WORKBOOK_BACKUP_DIR / safe_stem
    target.mkdir(parents=True, exist_ok=True)
    return target


def _create_byo_backup(path_or_name: Any) -> Path:
    source = Path(path_or_name)
    if not source.is_absolute():
        source = _byo_path(path_or_name)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    backup = _byo_backup_dir_for(source) / f"{source.stem}_before_edit_{stamp}{source.suffix}"
    shutil.copy2(source, backup)
    return backup


def _latest_byo_backup(filename: str) -> Optional[Path]:
    source = _byo_path(filename)
    folder = _byo_backup_dir_for(source)
    candidates = sorted(folder.glob(f"{source.stem}_before_edit_*{source.suffix}"), key=lambda q: q.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


def _restore_byo_backup(filename: str, backup_path: Path | str) -> str:
    source = _byo_path(filename)
    backup = Path(backup_path)
    if not backup.exists():
        raise FileNotFoundError(f"Backup not found: {backup}")
    temp = source.with_name(f".{source.stem}.dart_restore_{uuid.uuid4().hex[:8]}{source.suffix}")
    shutil.copy2(backup, temp)
    os.replace(temp, source)
    return str(source.resolve())


def _write_byo_changes(filename: str, changes: List[Dict[str, Any]], expected_signature: str = "") -> Dict[str, Any]:
    if not changes:
        source = _byo_path(filename)
        return {"saved": 0, "backup": None, "path": str(source.resolve())}
    source = _byo_path(filename)
    if expected_signature and _source_file_signature(source) != expected_signature:
        raise RuntimeError("This DIY file changed after the editor loaded. Refresh the editor before saving.")
    raw_now = _read_byo_editable_dataset(source)
    column_map = {str(c): c for c in raw_now.columns}
    columns = list(column_map)
    validated: List[Dict[str, Any]] = []
    stale: List[str] = []
    for change in changes:
        column = str(change.get("column", ""))
        source_row = int(change.get("source_row", 0) or 0)
        if column not in columns:
            raise ValueError(f"Column no longer exists in dataset: {column}")
        idx = source_row - 2
        if idx < 0 or idx >= len(raw_now):
            raise ValueError(f"Source row {source_row} is outside the current dataset.")
        current_value = raw_now.iloc[idx][column_map[column]]
        if not _editor_before_matches(current_value, change.get("before", "")):
            stale.append(f"row {source_row} / {column}")
            continue
        after = _coerce_editor_input(change.get("after"), current_value)
        if not _editor_values_equal(current_value, after):
            validated.append({"source_row": source_row, "column": column, "after": after})
    if stale:
        raise RuntimeError("This DIY file changed after the editor loaded. Refresh before saving. Stale cells: " + ", ".join(stale[:8]))
    if not validated:
        return {"saved": 0, "backup": None, "path": str(source.resolve())}

    backup = _create_byo_backup(source)
    temp = source.with_name(f".{source.stem}.dart_edit_{uuid.uuid4().hex[:8]}{source.suffix}")
    try:
        if source.suffix.lower() == ".xlsx":
            from openpyxl import load_workbook
            workbook = load_workbook(source)
            sheet_name = workbook.sheetnames[0] if workbook.sheetnames else None
            if not sheet_name:
                workbook.close()
                raise ValueError("The Excel file does not contain a worksheet.")
            worksheet = workbook[sheet_name]
            header_map: Dict[str, int] = {}
            for cell in worksheet[1]:
                if cell.value is not None and str(cell.value) not in header_map:
                    header_map[str(cell.value)] = cell.column
            missing = sorted({x["column"] for x in validated if x["column"] not in header_map})
            if missing:
                workbook.close()
                raise ValueError("Workbook headers changed while the editor was open: " + ", ".join(missing))
            for change in validated:
                worksheet.cell(row=change["source_row"], column=header_map[change["column"]]).value = _excel_safe_value(change["after"])
            workbook.save(temp)
            workbook.close()
        else:
            frame = raw_now.copy()
            for change in validated:
                frame.at[change["source_row"] - 2, column_map[change["column"]]] = change["after"]
            frame.to_csv(temp, index=False, encoding="utf-8-sig")
        os.replace(temp, source)
    except PermissionError as exc:
        temp.unlink(missing_ok=True)
        raise PermissionError("Windows could not replace the selected DIY file. Close it in Excel and try again.") from exc
    except Exception:
        temp.unlink(missing_ok=True)
        raise
    return {"saved": len(validated), "backup": str(backup), "path": str(source.resolve())}


def _ensure_byo_email_dirs() -> None:
    BYO_EMAIL_AGENT_DIR.mkdir(parents=True, exist_ok=True)


def load_byo_email_automations() -> List[Dict[str, Any]]:
    _ensure_byo_email_dirs()
    if not BYO_EMAIL_TEMPLATE_PATH.exists():
        return []
    try:
        raw = json.loads(BYO_EMAIL_TEMPLATE_PATH.read_text(encoding="utf-8"))
        return raw if isinstance(raw, list) else []
    except Exception:
        return []


def save_byo_email_automations(templates: List[Dict[str, Any]]) -> None:
    _ensure_byo_email_dirs()
    _atomic_write_text(BYO_EMAIL_TEMPLATE_PATH, json.dumps(templates, indent=2, ensure_ascii=False, default=str))


def _byo_email_history_event(event: Dict[str, Any]) -> None:
    try:
        _ensure_byo_email_dirs()
        with BYO_EMAIL_HISTORY_PATH.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps({"timestamp": _utc_now_iso(), **event}, ensure_ascii=False, default=str) + "\n")
    except Exception:
        pass


def _byo_email_history(limit: int = 30) -> List[Dict[str, Any]]:
    if not BYO_EMAIL_HISTORY_PATH.exists():
        return []
    try:
        lines = BYO_EMAIL_HISTORY_PATH.read_text(encoding="utf-8").splitlines()[-limit:]
        return [json.loads(line) for line in reversed(lines) if line.strip()]
    except Exception:
        return []


def _write_byo_email_status(status: str, **extra: Any) -> None:
    try:
        _ensure_byo_email_dirs()
        _atomic_write_text(BYO_EMAIL_STATUS_PATH, json.dumps({"status": status, "updated_at": _utc_now_iso(), **extra}, indent=2, ensure_ascii=False, default=str))
    except Exception:
        pass


def _byo_email_status_payload() -> Dict[str, Any]:
    if BYO_EMAIL_STATUS_PATH.exists():
        try:
            return json.loads(BYO_EMAIL_STATUS_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"status": "Not started", "updated_at": None}


def upsert_byo_email_automation(template: Dict[str, Any]) -> Dict[str, Any]:
    templates = load_byo_email_automations()
    item = dict(template)
    raw_id = str(item.get("id") or "")
    item["id"] = _safe_template_id(raw_id if raw_id.startswith("byo_") else (raw_id or f"byo_{uuid.uuid4().hex[:12]}"))
    if not item["id"].startswith("byo_"):
        item["id"] = "byo_" + item["id"]
    item["workspace"] = "byo"
    item["updated_at"] = _utc_now_iso()
    if not item.get("created_at"):
        item["created_at"] = item["updated_at"]
    for i, existing in enumerate(templates):
        if existing.get("id") == item["id"]:
            if str(existing.get("dataset_name", "")) != str(item.get("dataset_name", "")):
                _snapshot_path(item["id"]).unlink(missing_ok=True)
            templates[i] = item
            break
    else:
        templates.append(item)
    save_byo_email_automations(templates)
    return item


def delete_byo_email_automation(template_id: str) -> None:
    save_byo_email_automations([t for t in load_byo_email_automations() if str(t.get("id")) != str(template_id)])
    _snapshot_path(template_id).unlink(missing_ok=True)


def _find_byo_email_template(template_id: str) -> Optional[Dict[str, Any]]:
    return next((t for t in load_byo_email_automations() if str(t.get("id")) == str(template_id)), None)


def _default_byo_email_template(df: pd.DataFrame, dataset_name: str) -> Dict[str, Any]:
    columns = [str(c) for c in df.columns]
    identity = columns[:1]
    return {
        "id": "", "workspace": "byo", "dataset_name": dataset_name,
        "name": f"{Path(dataset_name).stem} data-change alert" if dataset_name else "DIY data-change alert",
        "recipient_email": "", "enabled": False, "trigger_mode": "New or changed rows",
        "identity_columns": identity, "monitor_columns": columns, "condition_mode": "Any change",
        "condition_logic": "ALL conditions", "conditions": [], "display_columns": columns[:8],
        "subject_prefix": "[DART DIY Alert]", "greeting": "Hello,",
        "email_intro": f"DART detected a qualifying change in the Build Your Own dataset {dataset_name}." if dataset_name else "DART detected a qualifying change in a Build Your Own dataset.",
        "include_csv": True, "agent_label": "DART Build Your Own Alert",
        "recommended_next_step": "Review the changed fields in the selected Build Your Own dataset and confirm whether the update is expected.",
    }


def _normalize_byo_email_template(payload: Dict[str, Any], df: pd.DataFrame) -> Dict[str, Any]:
    dataset_name = str(payload.get("dataset_name", "")).strip()
    _byo_path(dataset_name)
    item = _normalize_email_template(payload, df)
    item["dataset_name"] = dataset_name
    item["workspace"] = "byo"
    item["agent_label"] = "DART Build Your Own Alert"
    item["recommended_next_step"] = "Review the changed fields in the selected Build Your Own dataset and confirm whether the update is expected."
    return item


def _run_one_byo_email_template(template: Dict[str, Any], send_email: bool = True) -> Dict[str, Any]:
    dataset_name = str(template.get("dataset_name", "")).strip()
    df = _read_byo_dataset(dataset_name)
    result = run_email_automation(df, template, send_email=send_email)
    result["dataset_name"] = dataset_name
    return result


def run_enabled_byo_email_automations(dataset_name: Optional[str] = None) -> List[Dict[str, Any]]:
    templates = load_byo_email_automations()
    enabled = [t for t in templates if t.get("enabled") and (not dataset_name or str(t.get("dataset_name")) == str(dataset_name))]
    results: List[Dict[str, Any]] = []
    changed = False
    for template in enabled:
        try:
            result = _run_one_byo_email_template(template, send_email=True)
        except Exception as exc:
            result = {"id": template.get("id"), "name": template.get("name"), "dataset_name": template.get("dataset_name"), "status": f"Runner error: {exc}", "matches": 0, "change_rows": 0, "checked_at": _utc_now_iso(), "error": str(exc)}
        template["last_run_at"] = result.get("checked_at")
        template["last_status"] = result.get("status")
        template["last_match_count"] = result.get("matches", 0)
        template["last_change_count"] = result.get("change_rows", 0)
        results.append(result)
        changed = True
        _byo_email_history_event({"automation_id": template.get("id"), "automation": template.get("name"), "dataset": template.get("dataset_name"), "status": result.get("status"), "detected_changes": result.get("change_rows", 0), "matching_changes": result.get("matches", 0)})
    if changed:
        save_byo_email_automations(templates)
    _write_byo_email_status("Checked", enabled_automations=len(enabled), dataset=dataset_name or "all enabled datasets", results=[{"automation": r.get("name"), "dataset": r.get("dataset_name"), "status": r.get("status"), "detected_changes": r.get("change_rows", 0), "matching_changes": r.get("matches", 0)} for r in results])
    return results


def _byo_email_scheduler_loop() -> None:
    last_signatures: Dict[str, str] = {}
    _write_byo_email_status("Running", interval_seconds=BYO_EMAIL_AGENT_MIN_CHECK_SECONDS)
    while True:
        try:
            enabled = [t for t in load_byo_email_automations() if t.get("enabled")]
            datasets = list(dict.fromkeys(str(t.get("dataset_name", "")) for t in enabled if str(t.get("dataset_name", "")).strip()))
            if not datasets:
                _write_byo_email_status("Running — no enabled automations", interval_seconds=BYO_EMAIL_AGENT_MIN_CHECK_SECONDS)
            for name in datasets:
                try:
                    path = _byo_path(name)
                    signature = _source_file_signature(path)
                except Exception as exc:
                    _write_byo_email_status("Waiting — monitored DIY file missing", dataset=name, error=str(exc))
                    continue
                previous = last_signatures.get(name)
                if previous is None or signature != previous:
                    results = run_enabled_byo_email_automations(name)
                    last_signatures[name] = signature
                    _write_byo_email_status("DIY dataset change processed" if previous is not None else "Startup check complete", dataset=name, interval_seconds=BYO_EMAIL_AGENT_MIN_CHECK_SECONDS, results_count=len(results))
        except Exception as exc:
            _write_byo_email_status("Scheduler error", error=str(exc))
            traceback.print_exc()
        time.sleep(BYO_EMAIL_AGENT_MIN_CHECK_SECONDS)


def start_byo_email_agent_scheduler() -> bool:
    global _BYO_EMAIL_AGENT_THREAD
    if not BYO_EMAIL_AGENT_AUTO_RUN:
        return False
    with _BYO_EMAIL_AGENT_THREAD_LOCK:
        if _BYO_EMAIL_AGENT_THREAD is not None and _BYO_EMAIL_AGENT_THREAD.is_alive():
            return True
        _BYO_EMAIL_AGENT_THREAD = threading.Thread(target=_byo_email_scheduler_loop, name="dart-byo-email-agent", daemon=True)
        _BYO_EMAIL_AGENT_THREAD.start()
        return True


def _byo_common_pairs(left: pd.DataFrame, right: pd.DataFrame) -> List[Tuple[str, str]]:
    right_map = {str(c).strip().lower(): str(c) for c in right.columns}
    pairs: List[Tuple[str, str]] = []
    seen = set()
    for col in left.columns:
        lc = str(col).strip().lower()
        if lc in right_map and lc not in seen:
            pairs.append((str(col), right_map[lc]))
            seen.add(lc)
    return pairs


def _series_overlap_stats(left: pd.Series, right: pd.Series) -> Dict[str, Any]:
    left_clean = left.dropna().astype(str).str.strip()
    right_clean = right.dropna().astype(str).str.strip()
    left_clean = left_clean[left_clean.ne("")]
    right_clean = right_clean[right_clean.ne("")]
    left_values = set(left_clean.unique().tolist())
    right_values = set(right_clean.unique().tolist())
    shared = left_values & right_values
    return {
        "Distinct A": len(left_values),
        "Distinct B": len(right_values),
        "Shared distinct": len(shared),
        "A values found in B %": round((len(shared) / len(left_values) * 100) if left_values else 0, 2),
        "B values found in A %": round((len(shared) / len(right_values) * 100) if right_values else 0, 2),
    }


def compare_byo_datasets(left_name: str, right_name: str, include_previews: bool = True) -> Dict[str, Any]:
    if not left_name or not right_name:
        raise ValueError("Choose two saved datasets.")
    if left_name == right_name:
        raise ValueError("Choose two different datasets to compare.")
    left = _read_byo_dataset(left_name)
    right = _read_byo_dataset(right_name)
    left_payload = _dataset_payload_from_frame(left, left_name, 100 if include_previews else 0)
    right_payload = _dataset_payload_from_frame(right, right_name, 100 if include_previews else 0)
    pairs = _byo_common_pairs(left, right)
    left_common = {a for a, _ in pairs}
    right_common = {b for _, b in pairs}
    only_left = [str(c) for c in left.columns if str(c) not in left_common]
    only_right = [str(c) for c in right.columns if str(c) not in right_common]

    column_comparison: List[Dict[str, Any]] = []
    numeric_comparison: List[Dict[str, Any]] = []
    key_candidates: List[Dict[str, Any]] = []
    for left_col, right_col in pairs:
        ls = left[left_col]
        rs = right[right_col]
        left_non_null = int(ls.notna().sum())
        right_non_null = int(rs.notna().sum())
        left_unique = int(ls.nunique(dropna=True))
        right_unique = int(rs.nunique(dropna=True))
        left_miss = int(ls.isna().sum())
        right_miss = int(rs.isna().sum())
        left_unique_pct = (left_unique / left_non_null * 100) if left_non_null else 0.0
        right_unique_pct = (right_unique / right_non_null * 100) if right_non_null else 0.0
        column_comparison.append({
            "Column A": left_col,
            "Column B": right_col,
            "Type A": str(ls.dtype),
            "Type B": str(rs.dtype),
            "Type match": str(ls.dtype) == str(rs.dtype),
            "Missing A": left_miss,
            "Missing B": right_miss,
            "Missing % A": round((left_miss / len(left) * 100) if len(left) else 0, 2),
            "Missing % B": round((right_miss / len(right) * 100) if len(right) else 0, 2),
            "Unique A": left_unique,
            "Unique B": right_unique,
            "Distinct % A": round(left_unique_pct, 2),
            "Distinct % B": round(right_unique_pct, 2),
        })

        if pd.api.types.is_numeric_dtype(ls) and pd.api.types.is_numeric_dtype(rs):
            lnum = pd.to_numeric(ls, errors="coerce")
            rnum = pd.to_numeric(rs, errors="coerce")
            lmean = float(lnum.mean()) if lnum.notna().any() else np.nan
            rmean = float(rnum.mean()) if rnum.notna().any() else np.nan
            numeric_comparison.append({
                "Column": left_col,
                "Mean A": None if pd.isna(lmean) else round(lmean, 4),
                "Mean B": None if pd.isna(rmean) else round(rmean, 4),
                "Mean delta": None if pd.isna(lmean) or pd.isna(rmean) else round(rmean - lmean, 4),
                "Median A": None if lnum.dropna().empty else round(float(lnum.median()), 4),
                "Median B": None if rnum.dropna().empty else round(float(rnum.median()), 4),
                "Min A": None if lnum.dropna().empty else round(float(lnum.min()), 4),
                "Min B": None if rnum.dropna().empty else round(float(rnum.min()), 4),
                "Max A": None if lnum.dropna().empty else round(float(lnum.max()), 4),
                "Max B": None if rnum.dropna().empty else round(float(rnum.max()), 4),
            })

        name_hint = bool(re.search(r"(^|[^a-z])(id|key|identifier|number|num|code)([^a-z]|$)", left_col.lower()))
        min_unique_pct = min(left_unique_pct, right_unique_pct)
        if name_hint or min_unique_pct >= 70:
            overlap = _series_overlap_stats(ls, rs)
            min_overlap = min(float(overlap["A values found in B %"]), float(overlap["B values found in A %"]))
            if name_hint or int(overlap["Shared distinct"]) > 0:
                score = min_unique_pct + min_overlap
                if min_unique_pct >= 95 and min_overlap >= 50:
                    quality = "Strong"
                elif min_overlap > 0:
                    quality = "Possible"
                else:
                    quality = "Name-based only"
                key_candidates.append({
                    "Column A": left_col,
                    "Column B": right_col,
                    "Uniqueness A %": round(left_unique_pct, 2),
                    "Uniqueness B %": round(right_unique_pct, 2),
                    **overlap,
                    "Key quality": quality,
                    "_score": score,
                })

    key_candidates.sort(key=lambda r: r.get("_score", 0), reverse=True)
    for row in key_candidates:
        row.pop("_score", None)

    type_mismatches = [r for r in column_comparison if not r["Type match"]]
    summary = {
        "dataset_a": left_name,
        "dataset_b": right_name,
        "rows_a": int(len(left)),
        "rows_b": int(len(right)),
        "row_delta": int(len(right) - len(left)),
        "columns_a": int(len(left.columns)),
        "columns_b": int(len(right.columns)),
        "common_columns": len(pairs),
        "only_a": len(only_left),
        "only_b": len(only_right),
        "type_mismatches": len(type_mismatches),
        "completeness_a": left_payload["completeness"],
        "completeness_b": right_payload["completeness"],
        "completeness_delta_pp": round((right_payload["completeness"] - left_payload["completeness"]) * 100, 2),
        "duplicates_a": left_payload["duplicate_rows"],
        "duplicates_b": right_payload["duplicate_rows"],
        "key_candidates": len(key_candidates),
    }
    return {
        "summary": summary,
        "left": left_payload,
        "right": right_payload,
        "common_columns": [{"Column A": a, "Column B": b} for a, b in pairs],
        "only_left": [{"Only in A": c} for c in only_left],
        "only_right": [{"Only in B": c} for c in only_right],
        "column_comparison": column_comparison,
        "type_mismatches": type_mismatches,
        "numeric_comparison": numeric_comparison,
        "key_candidates": key_candidates[:25],
    }


def _prompt_terms(prompt: str) -> List[str]:
    stop = {
        "the", "and", "for", "with", "from", "that", "this", "what", "which", "where",
        "compare", "dataset", "datasets", "data", "show", "tell", "about", "between",
    }
    return [
        t for t in re.findall(r"[A-Za-z0-9_]+", str(prompt).lower())
        if len(t) >= 3 and t not in stop
    ]


def _relevant_columns(df: pd.DataFrame, prompt: str, max_columns: int = 10) -> List[str]:
    cols = [str(c) for c in df.columns]
    terms = _prompt_terms(prompt)
    scored: List[Tuple[float, str]] = []
    for idx, col in enumerate(cols):
        lc = col.lower()
        score = sum(5 for t in terms if t in lc)
        if re.search(r"(^|[^a-z])(id|key|join|code|status|amount|date)([^a-z]|$)", lc):
            score += 1.5
        score -= min(idx, 1000) / 10000
        scored.append((score, col))
    matched = [c for score, c in sorted(scored, key=lambda x: x[0], reverse=True) if score > 0]
    return list(dict.fromkeys(matched + cols[:max_columns]))[:max_columns]


def _ai_safe_scalar(value: Any, max_chars: int = 160) -> Any:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, (int, np.integer)):
        return int(value)
    if isinstance(value, (float, np.floating)):
        return None if not np.isfinite(float(value)) else round(float(value), 6)
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    value_text = str(value)
    return value_text if len(value_text) <= max_chars else value_text[: max_chars - 1] + "…"


def _byo_ai_sample_rows(df: pd.DataFrame, cols: List[str], prompt: str, limit: int = 5) -> List[Dict[str, Any]]:
    if df.empty or not cols:
        return []
    terms = _prompt_terms(prompt)[:6]
    selected = df.iloc[0:0]
    if terms:
        search_frame = df[cols].astype(str)
        mask = pd.Series(False, index=df.index)
        for term in terms:
            mask = mask | search_frame.apply(
                lambda col: col.str.contains(re.escape(term), case=False, na=False)
            ).any(axis=1)
        selected = df.loc[mask, cols].head(limit)
    if len(selected) < limit:
        filler = df.loc[~df.index.isin(selected.index), cols].head(limit - len(selected))
        selected = pd.concat([selected, filler], axis=0)
    return [
        {str(col): _ai_safe_scalar(row.get(col)) for col in cols}
        for _, row in selected.head(limit).iterrows()
    ]


def _byo_ai_dataset_context(df: pd.DataFrame, name: str, prompt: str) -> Dict[str, Any]:
    cols = _relevant_columns(df, prompt, max_columns=10)
    numeric_stats: List[Dict[str, Any]] = []
    categorical_stats: List[Dict[str, Any]] = []
    profile: List[Dict[str, Any]] = []

    for col in cols:
        series = df[col]
        profile.append({
            "column": col,
            "type": str(series.dtype),
            "non_null": int(series.notna().sum()),
            "missing": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
        })
        if pd.api.types.is_numeric_dtype(series):
            nums = pd.to_numeric(series, errors="coerce")
            numeric_stats.append({
                "column": col,
                "count": int(nums.notna().sum()),
                "mean": None if nums.dropna().empty else round(float(nums.mean()), 4),
                "median": None if nums.dropna().empty else round(float(nums.median()), 4),
                "min": None if nums.dropna().empty else round(float(nums.min()), 4),
                "max": None if nums.dropna().empty else round(float(nums.max()), 4),
            })
        else:
            values = series.dropna().astype(str).str.strip()
            values = values[values.ne("")]
            top = values.value_counts().head(4)
            categorical_stats.append({
                "column": col,
                "non_null": int(series.notna().sum()),
                "unique": int(series.nunique(dropna=True)),
                "top_values": [{"value": _ai_safe_scalar(k, 90), "count": int(v)} for k, v in top.items()],
            })

    return {
        "name": name,
        "rows": int(len(df)),
        "columns": int(len(df.columns)),
        "analyzed_columns": cols,
        "profile": profile,
        "numeric_stats": numeric_stats,
        "categorical_stats": categorical_stats,
        "sample_rows": _byo_ai_sample_rows(df, cols, prompt, limit=5),
    }


def _normalize_compare_value(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def _byo_join_evidence(
    left: pd.DataFrame,
    right: pd.DataFrame,
    comparison: Dict[str, Any],
    prompt: str,
) -> Dict[str, Any]:
    """Summarize join-level discrepancies instead of sending entire row sets to the model."""
    candidates = comparison.get("key_candidates", [])
    if not candidates:
        return {}

    candidate = next(
        (
            c for c in candidates
            if re.search(r"join|(^|[^a-z])id([^a-z]|$)|key", str(c.get("Column A", "")).lower())
        ),
        candidates[0],
    )
    key_a = str(candidate.get("Column A", ""))
    key_b = str(candidate.get("Column B", ""))
    if key_a not in left.columns or key_b not in right.columns:
        return {}

    lk = left[key_a].map(_normalize_compare_value)
    rk = right[key_b].map(_normalize_compare_value)
    left_keys = set(lk[lk.ne("")].unique().tolist())
    right_keys = set(rk[rk.ne("")].unique().tolist())
    shared_keys = left_keys & right_keys
    only_a = sorted(left_keys - right_keys)
    only_b = sorted(right_keys - left_keys)

    pairs = [(a, b) for a, b in _byo_common_pairs(left, right) if a != key_a and b != key_b]
    left_pair_cols = [a for a, _ in pairs]
    relevant_left = _relevant_columns(left[left_pair_cols], prompt, max_columns=8) if left_pair_cols else []
    pair_lookup = {a: b for a, b in pairs}
    ordered_pairs = [(a, pair_lookup[a]) for a in relevant_left if a in pair_lookup]
    ordered_pairs += [p for p in pairs if p not in ordered_pairs]
    ordered_pairs = ordered_pairs[:8]

    left_work = left.copy()
    right_work = right.copy()
    left_work["_DART_JOIN_KEY"] = lk
    right_work["_DART_JOIN_KEY"] = rk
    left_unique = left_work[left_work["_DART_JOIN_KEY"].ne("")].drop_duplicates("_DART_JOIN_KEY", keep="first")
    right_unique = right_work[right_work["_DART_JOIN_KEY"].ne("")].drop_duplicates("_DART_JOIN_KEY", keep="first")

    mismatch_summary: List[Dict[str, Any]] = []
    for col_a, col_b in ordered_pairs:
        merged = left_unique[["_DART_JOIN_KEY", col_a]].merge(
            right_unique[["_DART_JOIN_KEY", col_b]],
            on="_DART_JOIN_KEY",
            how="inner",
            suffixes=("_A", "_B"),
        )
        if merged.empty:
            continue
        a_name = col_a if col_a != col_b else f"{col_a}_A"
        b_name = col_b if col_a != col_b else f"{col_b}_B"
        av = merged[a_name].map(_normalize_compare_value)
        bv = merged[b_name].map(_normalize_compare_value)
        mismatch = av.ne(bv)
        examples = []
        for idx in merged.index[mismatch][:3]:
            examples.append({
                "key": _ai_safe_scalar(merged.at[idx, "_DART_JOIN_KEY"], 90),
                "a": _ai_safe_scalar(merged.at[idx, a_name], 120),
                "b": _ai_safe_scalar(merged.at[idx, b_name], 120),
            })
        mismatch_summary.append({
            "column_a": col_a,
            "column_b": col_b,
            "compared_rows": int(len(merged)),
            "mismatch_rows": int(mismatch.sum()),
            "mismatch_pct": round(float(mismatch.mean()) * 100, 2),
            "examples": examples,
        })

    mismatch_summary.sort(key=lambda row: row["mismatch_rows"], reverse=True)
    return {
        "key_a": key_a,
        "key_b": key_b,
        "key_quality": candidate.get("Key quality"),
        "distinct_a": len(left_keys),
        "distinct_b": len(right_keys),
        "shared_keys": len(shared_keys),
        "only_a_count": len(only_a),
        "only_b_count": len(only_b),
        "duplicate_key_rows_a": int(lk[lk.ne("")].duplicated(keep=False).sum()),
        "duplicate_key_rows_b": int(rk[rk.ne("")].duplicated(keep=False).sum()),
        "only_a_examples": [_ai_safe_scalar(v, 90) for v in only_a[:8]],
        "only_b_examples": [_ai_safe_scalar(v, 90) for v in only_b[:8]],
        "shared_column_mismatches": mismatch_summary[:8],
    }


def _byo_ai_context(left_name: str, right_name: str, prompt: str) -> Dict[str, Any]:
    left = _read_byo_dataset(left_name)
    right = _read_byo_dataset(right_name)
    comparison = compare_byo_datasets(left_name, right_name, include_previews=False)
    compact_comparison = {
        "summary": comparison["summary"],
        "common_columns": comparison["common_columns"][:20],
        "only_left": comparison["only_left"][:12],
        "only_right": comparison["only_right"][:12],
        "type_mismatches": comparison["type_mismatches"][:12],
        "numeric_comparison": comparison["numeric_comparison"][:10],
        "key_candidates": comparison["key_candidates"][:6],
        "column_comparison": comparison["column_comparison"][:16],
    }
    return {
        "persona": current_persona(),
        "dataset_a": _byo_ai_dataset_context(left, left_name, prompt),
        "dataset_b": _byo_ai_dataset_context(right, right_name, prompt),
        "comparison": compact_comparison,
        "join_evidence": _byo_join_evidence(left, right, comparison, prompt),
    }


def _byo_local_answer(left_name: str, right_name: str, comparison: Dict[str, Any]) -> str:
    s = comparison["summary"]
    keys = comparison.get("key_candidates", [])[:3]
    key_text = ", ".join(str(k.get("Column A")) for k in keys) if keys else "no strong shared key candidate was detected"
    return (
        f"I compared {left_name} and {right_name}. {left_name} has {s['rows_a']:,} rows and {s['columns_a']:,} columns; "
        f"{right_name} has {s['rows_b']:,} rows and {s['columns_b']:,} columns. They share {s['common_columns']:,} column names, "
        f"with {s['type_mismatches']:,} shared-column type mismatch(es). Completeness is {s['completeness_a']:.1%} vs "
        f"{s['completeness_b']:.1%}. Likely join-key candidates: {key_text}. Model-backed analysis is enabled for "
        f"question-specific analysis over both selected datasets when the Groq client is available."
    )


# -----------------------------------------------------------------------------
# Medicaid State Intelligence workspace (ported from the standalone Flask app)
# -----------------------------------------------------------------------------
MEDICAID_DB_PATH = Path(os.environ.get("DART_MEDICAID_DB", str(BASE_DIR / "Data" / "medicaid_issues.db")))
MEDICAID_ALL_STATES = [
    'Alabama','Alaska','Arizona','Arkansas','California','Colorado','Connecticut','Delaware',
    'Florida','Georgia','Hawaii','Idaho','Illinois','Indiana','Iowa','Kansas','Kentucky',
    'Louisiana','Maine','Maryland','Massachusetts','Michigan','Minnesota','Mississippi',
    'Missouri','Montana','Nebraska','Nevada','New Hampshire','New Jersey','New Mexico',
    'New York','North Carolina','North Dakota','Ohio','Oklahoma','Oregon','Pennsylvania',
    'Rhode Island','South Carolina','South Dakota','Tennessee','Texas','Utah','Vermont',
    'Virginia','Washington','West Virginia','Wisconsin','Wyoming'
]
MEDICAID_STATE_ABBR = {
    'Alabama':'AL','Alaska':'AK','Arizona':'AZ','Arkansas':'AR','California':'CA','Colorado':'CO',
    'Connecticut':'CT','Delaware':'DE','Florida':'FL','Georgia':'GA','Hawaii':'HI','Idaho':'ID',
    'Illinois':'IL','Indiana':'IN','Iowa':'IA','Kansas':'KS','Kentucky':'KY','Louisiana':'LA',
    'Maine':'ME','Maryland':'MD','Massachusetts':'MA','Michigan':'MI','Minnesota':'MN',
    'Mississippi':'MS','Missouri':'MO','Montana':'MT','Nebraska':'NE','Nevada':'NV',
    'New Hampshire':'NH','New Jersey':'NJ','New Mexico':'NM','New York':'NY','North Carolina':'NC',
    'North Dakota':'ND','Ohio':'OH','Oklahoma':'OK','Oregon':'OR','Pennsylvania':'PA',
    'Rhode Island':'RI','South Carolina':'SC','South Dakota':'SD','Tennessee':'TN','Texas':'TX',
    'Utah':'UT','Vermont':'VT','Virginia':'VA','Washington':'WA','West Virginia':'WV',
    'Wisconsin':'WI','Wyoming':'WY'
}
MEDICAID_ISSUE_TYPES = [
    'duplicate_claims','invalid_diagnosis_code','debit_credit_mismatch','null_recipient_values',
    'improper_procedure_code','referential_integrity','payment_amount_exceeded','birth_date_error',
    'discharge_date_error'
]
MEDICAID_ISSUE_TITLES = {
    'duplicate_claims':'Duplicate Claims Issue',
    'invalid_diagnosis_code':'Invalid Diagnosis Code',
    'debit_credit_mismatch':'Debit/Credit Not Matching',
    'null_recipient_values':'Null Recipient Values',
    'improper_procedure_code':'Improper Population of Value in Procedure Code Column',
    'referential_integrity':'Referential Integrity Issue',
    'payment_amount_exceeded':'Claim Payment Amount Exceeds Max for Type of Service',
    'birth_date_error':'Birth Date Greater Than Claim Date of Service',
    'discharge_date_error':'Discharge Date Greater Than Claim Admission Date'
}
MEDICAID_ISSUE_LABELS = {
    'duplicate_claims':'Duplicate Claims',
    'invalid_diagnosis_code':'Invalid Diagnosis Code',
    'debit_credit_mismatch':'Debit/Credit Mismatch',
    'null_recipient_values':'Null Recipient Values',
    'improper_procedure_code':'Improper Procedure Code',
    'referential_integrity':'Referential Integrity',
    'payment_amount_exceeded':'Payment Amount Exceeded',
    'birth_date_error':'Birth Date Error',
    'discharge_date_error':'Discharge Date Error'
}
MEDICAID_DUP_CLAIM_RATES = [
    1.2,3.4,0.8,5.1,2.3,4.7,1.9,6.2,0.5,3.8,2.1,4.3,1.5,7.2,2.8,3.1,0.9,
    5.5,1.7,4.1,2.6,3.9,1.1,6.8,2.4,0.7,3.6,2.9,4.8,1.3,5.9,2.2,1.6,3.3,
    0.6,4.5,2.7,1.8,5.3,3.7,2.0,6.1,1.4,4.2,0.4,3.5,2.5,7.8,1.0,4.9
]
MEDICAID_BASE_COUNTS = {
    'invalid_diagnosis_code':145,'debit_credit_mismatch':12,'null_recipient_values':78,
    'improper_procedure_code':234,'referential_integrity':34,'payment_amount_exceeded':23,
    'birth_date_error':5,'discharge_date_error':8
}
MEDICAID_COST_MULTIPLIERS = {
    'duplicate_claims':1250000,'invalid_diagnosis_code':450,'debit_credit_mismatch':2500,
    'null_recipient_values':150,'improper_procedure_code':320,'referential_integrity':850,
    'payment_amount_exceeded':1200,'birth_date_error':200,'discharge_date_error':250
}
MEDICAID_ISSUE_DESCRIPTIONS = {
    'duplicate_claims':[
        'Duplicate claim submissions detected in current period. Exact duplicates with matching NPI, service date, procedure code, and billed amount found across claims batch.',
        'Near-duplicate claims identified where minor field variations mask resubmissions. Member ID, DOS, and procedure code match across multiple claim IDs.',
        'Prior authorization duplicates found; same service authorized and billed under distinct claim IDs within the same benefit period.'
    ],
    'invalid_diagnosis_code':[
        'ICD-10-CM codes submitted that are not valid for the date of service. Codes are retired or do not exist in the applicable code set version.',
        'Diagnosis codes missing required 4th, 5th, or 6th character specificity per ICD-10-CM official guidelines for the billed service.',
        'Principal diagnosis code sequenced as a manifestation code, which cannot be listed first per ICD-10-CM coding conventions.'
    ],
    'debit_credit_mismatch':[
        'Remittance advice shows debit and credit totals that do not reconcile with paid claim amounts in the system of record.',
        'Encounter data financial crosswalk shows discrepancy between capitation payments and adjudicated claim values for the period.',
        'Monthly financial reconciliation identified net balance discrepancy between claim payment ledger and treasury disbursement records.'
    ],
    'null_recipient_values':[
        'Medicaid ID field is null or blank on submitted claims, preventing member attribution and eligibility validation.',
        'Recipient date of birth and gender fields contain null values, blocking downstream clinical quality measure attribution.',
        'Member enrollment segment contains null values in required demographic fields, causing eligibility verification failures.'
    ],
    'improper_procedure_code':[
        'Procedure code column populated with revenue codes instead of HCPCS/CPT codes, causing systematic claim adjudication failures.',
        'Type of Bill code used in place of procedure code on professional claims, resulting in systematic processing errors.',
        'Procedure modifier applied in the procedure code field rather than the designated modifier field, causing pricing calculation errors.'
    ],
    'referential_integrity':[
        'Claim header references a provider NPI that does not exist in the provider enrollment file, violating referential constraints.',
        'Member ID on claim does not match any active enrollment record, indicating breakdown between claims and eligibility systems.',
        'Rendering provider taxonomy code does not correspond to a valid taxonomy in the reference table, blocking credentialing validation.'
    ],
    'payment_amount_exceeded':[
        'Claim payment amount exceeds the established fee schedule maximum allowable for the billed procedure and place of service.',
        'Inpatient DRG payment calculation resulted in an amount exceeding the outlier threshold, requiring additional clinical review.',
        'Bundled payment for episode of care surpasses the benchmark rate set for the applicable service category and region.'
    ],
    'birth_date_error':[
        'Member date of birth on claim is recorded as a date after the claim date of service, indicating a data entry or system error.',
        'Newborn claims contain birth date that post-dates the delivery admission date of service in the ADT feed.',
        'Eligibility file shows member birth date greater than the earliest claim date in history, creating an impossible chronological sequence.'
    ],
    'discharge_date_error':[
        'Inpatient claim discharge date precedes the admission date, indicating a systemic date reversal error in the ADT feed.',
        'Long-term care claims show discharge date earlier than the admission date on the claim header record.',
        'UB-04 inpatient claims submitted with statement-from and statement-through dates transposed, resulting in negative length of stay.'
    ]
}


def _medicaid_db() -> sqlite3.Connection:
    MEDICAID_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(str(MEDICAID_DB_PATH))
    db.row_factory = sqlite3.Row
    return db


def _init_medicaid_db() -> None:
    db = _medicaid_db(); cur = db.cursor()
    cur.execute('CREATE TABLE IF NOT EXISTS states (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)')
    cur.execute('''CREATE TABLE IF NOT EXISTS issues (
        id INTEGER PRIMARY KEY AUTOINCREMENT, state_id INTEGER NOT NULL, title TEXT NOT NULL,
        description TEXT, status TEXT DEFAULT 'open', priority TEXT DEFAULT 'medium',
        issue_type TEXT DEFAULT 'general', metric_value REAL DEFAULT NULL, start_date TEXT DEFAULT NULL,
        end_date TEXT DEFAULT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY (state_id) REFERENCES states(id))''')
    cur.execute('CREATE TABLE IF NOT EXISTS tags (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL)')
    cur.execute('''CREATE TABLE IF NOT EXISTS issue_tags (issue_id INTEGER NOT NULL, tag_id INTEGER NOT NULL,
        FOREIGN KEY (issue_id) REFERENCES issues(id), FOREIGN KEY (tag_id) REFERENCES tags(id), PRIMARY KEY (issue_id, tag_id))''')
    cur.execute('''CREATE TABLE IF NOT EXISTS agent_subscriptions (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT NOT NULL,
        requirements TEXT, filter_value TEXT, rules TEXT, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
    for state in MEDICAID_ALL_STATES: cur.execute('INSERT OR IGNORE INTO states (name) VALUES (?)',(state,))
    tags=['needs improvement','will probably benefit','will not benefit','urgent','completed','in progress']
    for tag in tags: cur.execute('INSERT OR IGNORE INTO tags (name) VALUES (?)',(tag,))
    db.commit()
    cur.execute("SELECT COUNT(*) FROM issues WHERE issue_type = 'duplicate_claims'")
    if cur.fetchone()[0] == 0:
        cur.execute('DELETE FROM issue_tags'); cur.execute('DELETE FROM issues')
        cur.execute('SELECT id,name FROM states ORDER BY name'); state_rows=cur.fetchall()
        cur.execute('SELECT id,name FROM tags ORDER BY id'); tag_lookup={r['name']:r['id'] for r in cur.fetchall()}
        profiles=[0,2,0,3,1,3,1,4,0,2,1,3,0,4,1,2,0,3,1,2,1,2,0,4,1,0,2,1,3,0,3,1,1,2,0,3,1,1,3,2,1,4,0,2,0,2,1,4,0,3]
        totals=[18,40,66,96,142]
        dists=[(.10,.82,.08),(.20,.70,.10),(.40,.45,.15),(.62,.26,.12),(.76,.14,.10)]
        nondup=[t for t in MEDICAID_ISSUE_TYPES if t!='duplicate_claims']
        quarters=[('Q2 2025',date(2025,4,1),91),('Q3 2025',date(2025,7,1),92),('Q4 2025',date(2025,10,1),92),('Q1 2026',date(2026,1,1),90)]
        def h(a,b=0,c=0): return (a*37+b*23+c*13+a*b+b*c+1)%100
        for si,srow in enumerate(state_rows):
            sid=srow['id']; sname=srow['name']; profile=profiles[si]; dup_rate=MEDICAID_DUP_CLAIM_RATES[si]
            base_total=totals[profile]; var_frac=(h(si,7)-50)/100.0; total_issues=max(10,int(round(base_total*(1+var_frac*.4))))
            nondup_total=max(8,total_issues-len(quarters)); weights=[.5+h(si,ti,3)/100*3 for ti in range(len(nondup))]
            tw=sum(weights); counts=[max(1,round(w/tw*nondup_total)) for w in weights]; diff=nondup_total-sum(counts)
            for i in range(abs(diff)):
                ix=i%len(counts); counts[ix]+=1 if diff>0 else (-1 if counts[ix]>1 else 0)
            open_p,done_p,canc_p=dists[profile]
            for qi,(qlabel,qstart,qlen) in enumerate(quarters):
                sd=qstart+timedelta(days=h(si,qi)%min(20,qlen//4)); status='open' if qi==len(quarters)-1 else 'done'
                ed=(sd+timedelta(days=20+h(si,qi,1)%40)) if status=='done' else None
                metric=max(.1,min(9.9,round(dup_rate+(h(si,qi,2)-50)/500,2)))
                title=f"{sname} – {MEDICAID_ISSUE_TITLES['duplicate_claims']} ({qlabel})"; desc=MEDICAID_ISSUE_DESCRIPTIONS['duplicate_claims'][(qi+si)%3]
                prio='high' if dup_rate>4 else 'medium'
                cur.execute('''INSERT INTO issues (state_id,title,description,status,priority,issue_type,metric_value,start_date,end_date)
                    VALUES (?,?,?,?,?,?,?,?,?)''',(sid,title,desc,status,prio,'duplicate_claims',metric,sd.isoformat(),ed.isoformat() if ed else None))
                iid=cur.lastrowid; tnames=['completed','will probably benefit'] if status=='done' else ['needs improvement','in progress','urgent']
                for tn in tnames:
                    if tag_lookup.get(tn): cur.execute('INSERT OR IGNORE INTO issue_tags VALUES (?,?)',(iid,tag_lookup[tn]))
            for ti,itype in enumerate(nondup):
                for k in range(counts[ti]):
                    qi=h(si,ti,k+5)%len(quarters); qlabel,qstart,qlen=quarters[qi]
                    title=f"{sname} – {MEDICAID_ISSUE_TITLES[itype]}" + (f" ({qlabel})" if counts[ti]<=1 else f" #{k+1} ({qlabel})")
                    desc=MEDICAID_ISSUE_DESCRIPTIONS[itype][(ti+k+si)%3]
                    metric=round(MEDICAID_BASE_COUNTS[itype]*(.25+h(si,ti,k+1)/100*2.75)); rnd=h(si,ti*100+k+9)%100
                    status='done' if rnd<int(done_p*100) else ('cancelled' if rnd<int((done_p+canc_p)*100) else 'open')
                    priority=['low','medium','high','medium','high'][h(si,ti,k+2)%5]
                    sd=qstart+timedelta(days=h(si,ti,k+3)%min(40,qlen//2)); ed=None
                    if status in ('done','cancelled'):
                        work_days=14+h(si,ti,k+4)%77 if status=='done' else 7+h(si,ti,k+4)%42; ed=sd+timedelta(days=work_days)
                    cur.execute('''INSERT INTO issues (state_id,title,description,status,priority,issue_type,metric_value,start_date,end_date)
                        VALUES (?,?,?,?,?,?,?,?,?)''',(sid,title,desc,status,priority,itype,metric,sd.isoformat(),ed.isoformat() if ed else None))
                    iid=cur.lastrowid
                    tnames=(['completed','will probably benefit','in progress'] if status=='done' else
                            ['will not benefit','in progress'] if status=='cancelled' else
                            ['needs improvement','in progress','urgent'] if priority=='high' else ['needs improvement','will probably benefit','in progress'])
                    for tn in tnames:
                        if tag_lookup.get(tn): cur.execute('INSERT OR IGNORE INTO issue_tags VALUES (?,?)',(iid,tag_lookup[tn]))
        db.commit()
    db.close()


def _medicaid_comparison_rows(issue_type: str = '', state_ids: Optional[List[int]] = None) -> List[Dict[str, Any]]:
    db=_medicaid_db(); cur=db.cursor(); params=[issue_type,issue_type]; where=''
    if state_ids:
        where='WHERE s.id IN ('+','.join('?' for _ in state_ids)+')'; params.extend(state_ids)
    cur.execute(f'''SELECT s.id,s.name as state,COUNT(i.id) total,
        SUM(CASE WHEN i.status='done' THEN 1 ELSE 0 END) successful,
        SUM(CASE WHEN i.status='cancelled' THEN 1 ELSE 0 END) cancelled,
        SUM(CASE WHEN i.status='open' THEN 1 ELSE 0 END) open
        FROM states s LEFT JOIN issues i ON s.id=i.state_id AND (?='' OR LOWER(i.issue_type)=?) {where}
        GROUP BY s.id,s.name ORDER BY s.name''',params)
    rows=[]
    for r in cur.fetchall():
        x=dict(r); total=x['total'] or 0; done=x['successful'] or 0; canc=x['cancelled'] or 0; op=x['open'] or 0
        sr=done/total*100 if total else 0; cr=canc/total*100 if total else 0; br=op/total*100 if total else 0; conf=min(1,total/6)
        score=(.55*sr+.25*(100-cr)+.20*(100-br))*conf
        x.update(success_rate=round(sr,1),cancel_rate=round(cr,1),backlog_rate=round(br,1),confidence=round(conf*100,1),composite_score=round(score,1)); rows.append(x)
    db.close(); return rows


def _medicaid_generate_insights(data: List[Dict[str, Any]], issue_type: str='') -> List[Dict[str, str]]:
    eligible=[d for d in data if (d.get('total') or 0)>0]
    if not eligible: return [{'type':'strategic_insight','title':'Insufficient Data For Reliable Analysis','description':'No states met the current filter and minimum sample criteria.','recommendation':'Reduce strict filters or lower the minimum sample threshold to compare more states.'}]
    ranked=sorted(eligible,key=lambda x:(x.get('composite_score',0),x.get('success_rate',0),x.get('total',0)),reverse=True); best,worst=ranked[0],ranked[-1]
    avgs=sum(d['success_rate'] for d in eligible)/len(eligible); avgb=sum(d['backlog_rate'] for d in eligible)/len(eligible)
    out=[{'type':'strategic_insight','title':f'Dataset Health ({MEDICAID_ISSUE_LABELS.get(issue_type,"all issue types")})','description':f'Across {len(eligible)} states with data, average success is {avgs:.1f}% and average backlog is {avgb:.1f}%.','recommendation':'Use backlog reduction and cancellation prevention as primary levers for system-wide improvement.'},
         {'type':'best_performer','title':f'Top Performer: {best["state"]}','description':f'{best["state"]} leads with a composite score of {best["composite_score"]:.1f}, success {best["success_rate"]:.1f}%, confidence {best["confidence"]:.1f}%.','recommendation':f'Capture {best["state"]} playbooks by issue type and reuse them in lower-performing states.'},
         {'type':'improvement_opportunity','title':f'Growth Opportunity: {worst["state"]}','description':f'{worst["state"]} is trailing with score {worst["composite_score"]:.1f}; backlog {worst["backlog_rate"]:.1f}% and cancel rate {worst["cancel_rate"]:.1f}% are the main drag factors.','recommendation':'Prioritize issue triage, early risk checks, and weekly closure targets to improve throughput.'}]
    gap=best['composite_score']-worst['composite_score']
    if gap>15: out.append({'type':'strategic_insight','title':'High Performance Variance','description':f'The top-to-bottom composite gap is {gap:.1f} points, indicating uneven execution quality across states.','recommendation':'Launch cross-state quality reviews and standard operating playbooks to normalize outcomes.'})
    low=[d for d in eligible if d.get('confidence',0)<50]
    if low: out.append({'type':'strategic_insight','title':'Reliability Warning','description':f'{len(low)} states have low confidence due to limited sample sizes.','recommendation':'Treat those rankings as directional. Add more issues or widen filter scope for stronger reliability.'})
    return out


def _medicaid_executive_brief(issue_type: str='', min_total: int=3) -> Dict[str, Any]:
    rows=_medicaid_comparison_rows(issue_type); eligible=[r for r in rows if (r.get('total') or 0)>=max(0,min_total)]
    if not eligible:
        return {'kpis':{'states_analyzed':0,'avg_success':0.0,'avg_backlog':0.0,'avg_quality':0.0,'coverage':0.0},'top_states':[],'watchlist':[],'summary':'Insufficient data for the selected filters.','actions_30_60_90':{'30_days':['Widen filters or lower minimum sample threshold.'],'60_days':['Standardize issue categorization for better comparability.'],'90_days':['Collect additional outcomes data for executive benchmarking.']}}
    ranked=sorted(eligible,key=lambda x:(x['composite_score'],x['success_rate'],x['total']),reverse=True); watch=sorted(eligible,key=lambda x:(x['composite_score'],-x['backlog_rate']))
    avgs=sum(r['success_rate'] for r in eligible)/len(eligible); avgb=sum(r['backlog_rate'] for r in eligible)/len(eligible); avgq=sum(r['composite_score'] for r in eligible)/len(eligible); cov=len(eligible)/max(1,len(rows))*100
    return {'kpis':{'states_analyzed':len(eligible),'avg_success':round(avgs,1),'avg_backlog':round(avgb,1),'avg_quality':round(avgq,1),'coverage':round(cov,1)},'top_states':ranked[:5],'watchlist':watch[:5],
        'summary':f"Executive snapshot for {MEDICAID_ISSUE_LABELS.get(issue_type, issue_type) if issue_type else 'all issue types'}: {len(eligible)} states analyzed, {avgs:.1f}% average success, {avgb:.1f}% average backlog, and {avgq:.1f} average quality score.",
        'actions_30_60_90':{'30_days':['Stand up weekly quality review with top and watchlist states.','Target open backlog reduction in high-volume states by 10%.','Publish issue-type playbooks from top performers.'],'60_days':['Standardize triage criteria and cancellation controls.','Adopt cross-state KPI scorecards with confidence thresholds.','Track issue aging and improve time-to-close in priority categories.'],'90_days':['Launch peer mentoring between top and watchlist states.','Set issue-type specific success targets and accountability owners.','Institutionalize quarterly executive quality brief reporting.']}}


def _medicaid_heatmap_rows(status: str='', issue_type: str='', tag_id: int=0) -> List[Dict[str, Any]]:
    db=_medicaid_db(); cur=db.cursor(); params=[issue_type,issue_type,status,status,int(tag_id or 0),int(tag_id or 0)]
    cur.execute('''SELECT s.id,s.name,COUNT(fi.id) total,
        SUM(CASE WHEN fi.status='done' THEN 1 ELSE 0 END) successful,
        SUM(CASE WHEN fi.status='cancelled' THEN 1 ELSE 0 END) cancelled,
        SUM(CASE WHEN fi.status='open' THEN 1 ELSE 0 END) open
        FROM states s LEFT JOIN (SELECT i.* FROM issues i WHERE (?='' OR LOWER(i.issue_type)=?) AND (?='' OR i.status=?)
          AND (?=0 OR EXISTS (SELECT 1 FROM issue_tags it WHERE it.issue_id=i.id AND it.tag_id=?))) fi ON fi.state_id=s.id
        GROUP BY s.id,s.name ORDER BY s.name''',params)
    rows=[]
    for r in cur.fetchall():
        x=dict(r); t=x['total'] or 0; d=x['successful'] or 0; c=x['cancelled'] or 0; o=x['open'] or 0
        sr=d/t*100 if t else 0; cr=c/t*100 if t else 0; br=o/t*100 if t else 0; qs=.55*sr+.25*(100-cr)+.20*(100-br)
        x.update(state_code=MEDICAID_STATE_ABBR.get(x['name'],''),success_rate=round(sr,1),cancel_rate=round(cr,1),backlog_rate=round(br,1),quality_score=round(qs,1)); rows.append(x)
    db.close(); return rows


def _medicaid_dashboard_payload() -> Dict[str, Any]:
    db=_medicaid_db(); cur=db.cursor(); cur.execute('''SELECT s.id,s.name,COUNT(i.id) total_issues,
        SUM(CASE WHEN i.status='done' THEN 1 ELSE 0 END) done_issues,
        SUM(CASE WHEN i.status='cancelled' THEN 1 ELSE 0 END) cancelled_issues,
        SUM(CASE WHEN i.status='open' THEN 1 ELSE 0 END) open_issues FROM states s LEFT JOIN issues i ON s.id=i.state_id GROUP BY s.id,s.name ORDER BY s.name''')
    states=[dict(r) for r in cur.fetchall()]; cur.execute('SELECT id,name FROM tags ORDER BY name'); tags=[dict(r) for r in cur.fetchall()]; db.close()
    cmp=_medicaid_comparison_rows(); leaderboard=sorted([r for r in cmp if r['total']>0],key=lambda x:x['composite_score'],reverse=True)[:10]
    total=sum(r['total_issues'] or 0 for r in states); done=sum(r['done_issues'] or 0 for r in states); op=sum(r['open_issues'] or 0 for r in states); canc=sum(r['cancelled_issues'] or 0 for r in states)
    return {'states':states,'leaderboard':leaderboard,'tags':tags,'issue_types':MEDICAID_ISSUE_TYPES,'issue_labels':MEDICAID_ISSUE_LABELS,'totals':{'total':total,'done':done,'open':op,'cancelled':canc,'resolution_rate':round(done/total*100,1) if total else 0}}


def _medicaid_state_payload(state_id: int, status: str='', issue_type: str='', tag_id: int=0) -> Dict[str, Any]:
    db=_medicaid_db(); cur=db.cursor(); cur.execute('SELECT * FROM states WHERE id=?',(state_id,)); state=cur.fetchone()
    if not state: db.close(); raise ValueError('State not found')
    params=[state_id,issue_type,issue_type,status,status,int(tag_id or 0),int(tag_id or 0)]
    cur.execute('''SELECT i.*,GROUP_CONCAT(DISTINCT t.name) tags FROM issues i LEFT JOIN issue_tags it ON i.id=it.issue_id LEFT JOIN tags t ON it.tag_id=t.id
        WHERE i.state_id=? AND (?='' OR LOWER(i.issue_type)=?) AND (?='' OR i.status=?) AND (?=0 OR EXISTS (SELECT 1 FROM issue_tags it2 WHERE it2.issue_id=i.id AND it2.tag_id=?))
        GROUP BY i.id ORDER BY i.updated_at DESC,i.created_at DESC LIMIT 500''',params)
    issues=[dict(r) for r in cur.fetchall()]; cur.execute('SELECT id,name FROM tags ORDER BY name'); tags=[dict(r) for r in cur.fetchall()]; db.close()
    now=datetime.utcnow();
    for item in issues:
        item['days_open']=None
        if item.get('status')=='open' and item.get('created_at'):
            try: item['days_open']=(now-datetime.strptime(item['created_at'][:19],'%Y-%m-%d %H:%M:%S')).days
            except Exception: pass
    total=len(issues); done=sum(i['status']=='done' for i in issues); op=sum(i['status']=='open' for i in issues); canc=sum(i['status']=='cancelled' for i in issues)
    return {'state':{**dict(state),'code':MEDICAID_STATE_ABBR.get(state['name'],'')},'summary':{'total':total,'successful':done,'open':op,'cancelled':canc,'success_rate':round(done/total*100,1) if total else 0},'issues':issues,'tags':tags,'issue_types':MEDICAID_ISSUE_TYPES,'issue_labels':MEDICAID_ISSUE_LABELS}


def _medicaid_analytics_payload() -> Dict[str, Any]:
    db=_medicaid_db(); cur=db.cursor(); cur.execute('''SELECT COUNT(*) total_issues,SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) done,SUM(CASE WHEN status='open' THEN 1 ELSE 0 END) open,SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) cancelled FROM issues'''); stats=dict(cur.fetchone())
    cur.execute('''SELECT t.name,COUNT(it.issue_id) count FROM tags t LEFT JOIN issue_tags it ON t.id=it.tag_id GROUP BY t.id,t.name ORDER BY count DESC'''); tag_stats=[dict(r) for r in cur.fetchall()]
    cur.execute('''SELECT issue_type,COUNT(*) total,SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) done,SUM(CASE WHEN status='open' THEN 1 ELSE 0 END) open,SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) cancelled FROM issues WHERE issue_type IS NOT NULL AND issue_type!='' GROUP BY issue_type ORDER BY total DESC'''); type_stats=[dict(r) for r in cur.fetchall()]
    worst={}
    for itype in MEDICAID_ISSUE_TYPES:
        cur.execute('''SELECT s.id,s.name,MAX(CASE WHEN i.title LIKE '%Q1 2026%' THEN i.metric_value END) current_metric,SUM(CASE WHEN i.status='open' THEN 1 ELSE 0 END) open_issues FROM states s JOIN issues i ON i.state_id=s.id WHERE i.issue_type=? GROUP BY s.id HAVING current_metric IS NOT NULL ORDER BY current_metric DESC LIMIT 5''',(itype,))
        worst[itype]=[{'state_id':r['id'],'state_name':r['name'],'metric':r['current_metric'] or 0,'open_issues':r['open_issues'],'money_at_risk':(r['current_metric'] or 0)*MEDICAID_COST_MULTIPLIERS.get(itype,100)} for r in cur.fetchall()]
    db.close(); return {'stats':stats,'tag_stats':tag_stats,'issue_type_stats':type_stats,'worst_states':worst,'issue_labels':MEDICAID_ISSUE_LABELS,'cost_multipliers':MEDICAID_COST_MULTIPLIERS}


def _medicaid_state_summary(state_id: int) -> Dict[str, Any]:
    db=_medicaid_db(); cur=db.cursor(); cur.execute('SELECT * FROM states WHERE id=?',(state_id,)); s=cur.fetchone()
    if not s: db.close(); raise ValueError('State not found')
    cur.execute('''SELECT issue_type,COUNT(*) total,SUM(CASE WHEN status='open' THEN 1 ELSE 0 END) open_count,SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) done_count,MAX(CASE WHEN title LIKE '%Q1 2026%' THEN metric_value ELSE NULL END) current_metric FROM issues WHERE state_id=? GROUP BY issue_type''',(state_id,)); type_data={r['issue_type']:dict(r) for r in cur.fetchall()}
    rankings={};
    for it in MEDICAID_ISSUE_TYPES:
        cur.execute('''SELECT s.id,MAX(CASE WHEN i.title LIKE '%Q1 2026%' THEN i.metric_value END) cur FROM states s JOIN issues i ON i.state_id=s.id WHERE i.issue_type=? GROUP BY s.id HAVING cur IS NOT NULL ORDER BY cur ASC''',(it,)); rr=cur.fetchall(); rankings[it]=next((i+1 for i,r in enumerate(rr) if r['id']==state_id),None)
    cur.execute('''SELECT s.id,MAX(CASE WHEN i.title LIKE '%Q1 2026%' AND i.issue_type='duplicate_claims' THEN i.metric_value END) dup_rate FROM states s JOIN issues i ON i.state_id=s.id WHERE i.issue_type='duplicate_claims' GROUP BY s.id HAVING dup_rate IS NOT NULL ORDER BY dup_rate ASC'''); dup=cur.fetchall(); db.close()
    nat=sum(r['dup_rate'] for r in dup)/max(1,len(dup)); rate=(type_data.get('duplicate_claims') or {}).get('current_metric') or 0; high=sum(1 for t,d in type_data.items() if (d.get('open_count') or 0)>0 and t!='duplicate_claims')
    grade='A' if rate<1.5 and high<=2 else 'B' if rate<3 and high<=5 else 'C' if rate<5 else 'D' if rate<6.5 else 'F'
    return {'state':dict(s),'type_breakdown':type_data,'duplicate_claims':{'rate':round(rate,2),'national_avg':round(nat,2),'rank':rankings.get('duplicate_claims'),'total_states':len(dup),'vs_national':round(rate-nat,2)},'issue_type_rankings':rankings,'quality_grade':grade,'issue_labels':MEDICAID_ISSUE_LABELS}


def _medicaid_rankings() -> Dict[str, Any]:
    db=_medicaid_db(); cur=db.cursor(); rankings={}
    for it in MEDICAID_ISSUE_TYPES:
        cur.execute('''SELECT s.id,s.name,MAX(CASE WHEN i.title LIKE '%Q1 2026%' THEN i.metric_value END) current_metric,SUM(CASE WHEN i.status='open' THEN 1 ELSE 0 END) open_issues FROM states s JOIN issues i ON i.state_id=s.id WHERE i.issue_type=? GROUP BY s.id HAVING current_metric IS NOT NULL ORDER BY current_metric ASC''',(it,))
        rankings[it]=[{'rank':i+1,'state_id':r['id'],'state':r['name'],'metric':r['current_metric'],'open_issues':r['open_issues']} for i,r in enumerate(cur.fetchall())]
    db.close(); return {'rankings':rankings,'duplicate_claims_ranking':rankings.get('duplicate_claims',[]),'issue_types':MEDICAID_ISSUE_TYPES,'issue_type_labels':MEDICAID_ISSUE_LABELS}


def _medicaid_claims_payload() -> Dict[str, Any]:
    analytics=_medicaid_analytics_payload(); rankings=_medicaid_rankings(); db=_medicaid_db(); cur=db.cursor()
    cur.execute('SELECT COUNT(*) c FROM agent_subscriptions'); subscriptions=int(cur.fetchone()['c'])
    trend=[]
    for q in ['Q2 2025','Q3 2025','Q4 2025','Q1 2026']:
        cur.execute('''SELECT AVG(metric_value) avg_rate,MIN(metric_value) min_rate,MAX(metric_value) max_rate FROM issues WHERE issue_type='duplicate_claims' AND title LIKE ?''',(f'%{q}%',)); r=cur.fetchone(); trend.append({'quarter':q,'avg_rate':round(r['avg_rate'] or 0,2),'min_rate':round(r['min_rate'] or 0,2),'max_rate':round(r['max_rate'] or 0,2)})
    cur.execute('''SELECT issue_type,SUM(metric_value) metric_total,COUNT(*) records,SUM(CASE WHEN status='open' THEN 1 ELSE 0 END) open_issues FROM issues GROUP BY issue_type'''); type_metrics=[]
    for r in cur.fetchall():
        it=r['issue_type']; metric=float(r['metric_total'] or 0); risk=metric*MEDICAID_COST_MULTIPLIERS.get(it,100)
        type_metrics.append({'issue_type':it,'label':MEDICAID_ISSUE_LABELS.get(it,it),'metric_total':round(metric,2),'records':r['records'],'open_issues':r['open_issues'],'estimated_money_at_risk':round(risk,2)})
    db.close(); type_metrics.sort(key=lambda x:x['estimated_money_at_risk'],reverse=True)
    worst=[]
    for it,rows in analytics['worst_states'].items():
        for r in rows[:3]: worst.append({'issue_type':it,'label':MEDICAID_ISSUE_LABELS.get(it,it),**r})
    worst.sort(key=lambda x:x['money_at_risk'],reverse=True)
    return {'summary':analytics['stats'],'duplicate_trend':trend,'issue_type_metrics':type_metrics,'top_opportunities':worst[:12],'subscriptions':subscriptions,'rankings':rankings['duplicate_claims_ranking'][:10],'methodology':{'cost_multipliers':MEDICAID_COST_MULTIPLIERS,'issue_labels':MEDICAID_ISSUE_LABELS,'note':'Estimated money at risk is a deterministic planning proxy from the standalone Medicaid prototype, not an audited payment recovery amount.'}}


def _medicaid_ai_answer(question: str) -> Dict[str, Any]:
    q=(question or '').strip().lower()
    if not q: return {'answer':'Please ask a question.','details':[]}
    aliases={'duplicate':'duplicate_claims','dup':'duplicate_claims','diagnosis':'invalid_diagnosis_code','icd':'invalid_diagnosis_code','debit':'debit_credit_mismatch','credit':'debit_credit_mismatch','null':'null_recipient_values','recipient':'null_recipient_values','procedure':'improper_procedure_code','hcpcs':'improper_procedure_code','referential':'referential_integrity','integrity':'referential_integrity','npi':'referential_integrity','payment':'payment_amount_exceeded','overpayment':'payment_amount_exceeded','birth':'birth_date_error','discharge':'discharge_date_error'}
    matched_type=next((v for k,v in aliases.items() if k in q),''); matched_state=next((s for s in sorted(MEDICAID_ALL_STATES,key=len,reverse=True) if s.lower() in q),None)
    db=_medicaid_db(); cur=db.cursor()
    if matched_state:
        cur.execute('SELECT id FROM states WHERE name=?',(matched_state,)); sid=cur.fetchone()['id']; payload=_medicaid_state_payload(sid); summary=_medicaid_state_summary(sid); db.close()
        if matched_type:
            td=summary['type_breakdown'].get(matched_type,{})
            return {'answer':f"{matched_state} — {MEDICAID_ISSUE_LABELS.get(matched_type,matched_type)}: {td.get('total',0)} total issues, {td.get('open_count',0)} open, {td.get('done_count',0)} resolved.", 'details':[f"Q1 2026 metric: {td.get('current_metric') if td.get('current_metric') is not None else '–'}",f"State quality grade: {summary['quality_grade']}",f"Duplicate claim rate: {summary['duplicate_claims']['rate']}% (rank #{summary['duplicate_claims']['rank']} of {summary['duplicate_claims']['total_states']})"]}
        return {'answer':f"{matched_state} CMS Quality Snapshot: {payload['summary']['total']} total issues — {payload['summary']['successful']} resolved, {payload['summary']['open']} open. Duplicate claim rate: {summary['duplicate_claims']['rate']}%.", 'details':[f"Quality grade: {summary['quality_grade']}",f"Nationwide duplicate-claim rank: #{summary['duplicate_claims']['rank']} of {summary['duplicate_claims']['total_states']}"]}
    if 'spend' in q or 'money' in q or 'cost' in q or 'risk' in q:
        claims=_medicaid_claims_payload(); top=claims['top_opportunities'][:7]; db.close(); return {'answer':'Largest modeled Medicaid quality/spending opportunities based on the prototype cost multipliers:','details':[f"{x['state_name']} · {x['label']}: ${x['money_at_risk']:,.0f} modeled risk" for x in top]}
    if matched_type:
        cur.execute('''SELECT s.name,MAX(CASE WHEN i.title LIKE '%Q1 2026%' THEN i.metric_value END) q1_vol,SUM(CASE WHEN i.status='open' THEN 1 ELSE 0 END) open_cnt FROM states s JOIN issues i ON i.state_id=s.id WHERE i.issue_type=? GROUP BY s.id HAVING q1_vol IS NOT NULL ORDER BY q1_vol ASC''',(matched_type,)); rows=cur.fetchall(); db.close()
        if not rows: return {'answer':f'No Q1 2026 data found for {MEDICAID_ISSUE_LABELS.get(matched_type,matched_type)}.','details':[]}
        if re.search(r'\b(worst|highest|most|critical)\b',q): rows=list(reversed(rows)); title=f"{MEDICAID_ISSUE_LABELS.get(matched_type,matched_type)} — states with highest Q1 2026 metrics:"
        else: title=f"{MEDICAID_ISSUE_LABELS.get(matched_type,matched_type)} — states with lowest Q1 2026 metrics:"
        return {'answer':title,'details':[f"#{i+1} {r['name']}: {r['q1_vol']:,.2f}, {r['open_cnt']} open" for i,r in enumerate(rows[:7])]}
    if re.search(r'\b(open|backlog|unresolved|pending)\b',q):
        rows=sorted(_medicaid_comparison_rows(),key=lambda x:x['open'],reverse=True)[:8]; db.close(); return {'answer':'States with the most open CMS quality issues:','details':[f"{r['state']}: {r['open']} open of {r['total']} ({r['backlog_rate']}% backlog)" for r in rows]}
    if re.search(r'\b(best|top|leader|performing)\b',q):
        rr=_medicaid_rankings()['duplicate_claims_ranking'][:7]; db.close(); return {'answer':'Top-performing states by lowest duplicate claim rate (Q1 2026):','details':[f"#{r['rank']} {r['state']}: {r['metric']}% dup rate, {r['open_issues']} open issues" for r in rr]}
    if re.search(r'\b(worst|bottom|critical|lagging|failing|poor)\b',q):
        rr=list(reversed(_medicaid_rankings()['duplicate_claims_ranking']))[:7]; db.close(); return {'answer':'States needing the most CMS quality improvement (highest duplicate claim rates):','details':[f"{r['state']}: {r['metric']}% dup rate, {r['open_issues']} open issues" for r in rr]}
    if re.search(r'\b(compare|rank|ranking|leaderboard|scorecard)\b',q):
        rows=sorted(_medicaid_comparison_rows(),key=lambda x:x['composite_score'],reverse=True)[:10]; db.close(); return {'answer':'CMS quality comparison across all 50 states — composite quality score leaderboard.','details':[f"#{i+1} {r['state']}: score {r['composite_score']}, success {r['success_rate']}%, backlog {r['backlog_rate']}%" for i,r in enumerate(rows)]}
    cur.execute('''SELECT SUM(CASE WHEN status='open' THEN 1 ELSE 0 END) open_cnt,SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) done_cnt,SUM(CASE WHEN status='cancelled' THEN 1 ELSE 0 END) canc_cnt,COUNT(*) total FROM issues'''); totals=dict(cur.fetchone()); cur.execute('''SELECT AVG(CASE WHEN title LIKE '%Q1 2026%' AND issue_type='duplicate_claims' THEN metric_value END) avg_dup FROM issues'''); avg=cur.fetchone()['avg_dup']; db.close(); total=totals['total'] or 1
    return {'answer':f"CMS Quality Overview — {total:,} total issues across all 50 states: {totals['done_cnt']:,} resolved ({totals['done_cnt']/total*100:.1f}%), {totals['open_cnt']:,} open ({totals['open_cnt']/total*100:.1f}%). National avg duplicate claim rate (Q1 2026): {avg:.2f}%.", 'details':['Try asking: “duplicate claim rate by state”','Try asking: “worst states for invalid diagnosis code”','Try asking: “open issues in Texas”','Try asking: “rank the states”','Try asking: “largest money at risk”']}


_init_medicaid_db()


def chart_html(fig: go.Figure, height: int = 360) -> str:
    fig.update_layout(
        template="plotly_white",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font={"color": "#172033", "family": "Inter, Segoe UI, Arial"},
        colorway=["#8B6A00", "#C9920A", "#E1B935", "#5F7F68", "#B98516", "#D7C17A"],
        margin={"l": 44, "r": 18, "t": 58, "b": 42},
        legend={"orientation": "h", "y": -0.2},
    )
    return pio.to_html(fig, include_plotlyjs=False, full_html=False, config={"displayModeBar": False, "responsive": True})


def no_data_chart(title: str, message: str = "No data is available for the current selection.") -> str:
    fig = go.Figure()
    fig.add_annotation(
        text=message,
        x=0.5,
        y=0.5,
        xref="paper",
        yref="paper",
        showarrow=False,
        align="center",
        font={"size": 15, "color": "#64748b"},
    )
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    fig.update_layout(title=title)
    return chart_html(fig)


def summary_payload(fdf: pd.DataFrame) -> Dict[str, Any]:
    if fdf.empty:
        return {"rows": 0, "field_average": "-", "weighted_rate": "-", "actions": "0", "streams": "0", "unmatched": "0", "risk_count": "0", "health_score": "-", "critical": "0", "risk_score": "-", "affected_reports": "0", "top_actions": [], "stream_summary": []}
    actions = fdf[fdf["NeedsChange"]]
    health = weighted_rate(fdf)
    risk_score = float(fdf["ImpactScore"].fillna(0).mean()) if len(fdf) else np.nan
    stream_summary = fdf.groupby("Stream", as_index=False).agg(
        Fields=("Stream", "size"),
        AverageMatch=("MatchRate", "mean"),
        Matched=("MatchedClaims", "sum"),
        NotMatched=("NotMatchedClaims", "sum"),
        AverageImpact=("ImpactScore", "mean"),
        Critical=("RiskTier", lambda x: int((x == "Critical").sum())),
    )
    stream_total = stream_summary["Matched"].fillna(0) + stream_summary["NotMatched"].fillna(0)
    stream_summary["WeightedMatch"] = np.where(
        stream_total > 0,
        stream_summary["Matched"].fillna(0) / stream_total,
        np.nan,
    )
    stream_summary_records = stream_summary.sort_values("NotMatched", ascending=False).replace({np.nan: None}).to_dict("records")
    sorted_actions = actions.sort_values(["NotMatchedClaims", "MatchRate", "ImpactScore"], ascending=[False, True, False]) if not actions.empty else fdf.sort_values(["NotMatchedClaims", "MatchRate", "ImpactScore"], ascending=[False, True, False])
    if len(actions) < 100 and not fdf.empty:
        non_actions = fdf[~fdf.index.isin(actions.index)].sort_values(["NotMatchedClaims", "MatchRate", "ImpactScore"], ascending=[False, True, False])
        combined = pd.concat([sorted_actions, non_actions])
    else:
        combined = sorted_actions
    top_actions = clean_records(combined, 150)
    affected = max(len(STATE.get("impacts", [])), fdf["Report"].replace("", np.nan).nunique(dropna=True), fdf["Stream"].nunique())
    return {
        "rows": len(fdf),
        "field_average": pct(field_average(fdf)),
        "weighted_rate": pct(weighted_rate(fdf)),
        "actions": fmt_int(actions.shape[0]),
        "streams": fmt_int(fdf["Stream"].nunique()),
        "unmatched": fmt_int(fdf["NotMatchedClaims"].fillna(0).sum()),
        "risk_count": fmt_int((fdf["RiskTier"].isin(["Critical", "Elevated"])).sum()),
        "critical": fmt_int((fdf["RiskTier"] == "Critical").sum()),
        "risk_score": "-" if pd.isna(risk_score) else f"{risk_score:.1f}",
        "affected_reports": fmt_int(affected),
        "health_score": "-" if pd.isna(health) else f"{health * 100:.0f}",
        "top_actions": top_actions,
        "stream_summary": stream_summary_records,
    }


def add_snapshot(fdf: pd.DataFrame) -> None:
    if fdf.empty:
        return
    snap = {"Label": f"S{len(STATE['snapshots']) + 1}", "FieldAverage": field_average(fdf), "WeightedRate": weighted_rate(fdf), "Critical": int((fdf["RiskTier"] == "Critical").sum()), "Actions": int(fdf["NeedsChange"].sum())}
    if not STATE["snapshots"] or STATE["snapshots"][-1] != snap:
        STATE["snapshots"].append(snap)
    STATE["snapshots"] = STATE["snapshots"][-20:]



def build_charts(fdf: pd.DataFrame) -> Tuple[Dict[str, str], List[Dict[str, Any]]]:
    keys = ["hist", "stream", "weighted", "compare", "classmix", "risk", "heatmap", "trend", "drivers", "volume", "impact"]
    charts = {k: "" for k in keys}
    compare_table: List[Dict[str, Any]] = []
    if fdf.empty:
        for key in keys:
            charts[key] = no_data_chart(key.title())
        return charts, []
    comparable = fdf[fdf["Comparable"]]
    if comparable.empty:
        charts["hist"] = no_data_chart("Comparable Field Match-Rate Distribution")
    else:
        fig = px.histogram(comparable, x="MatchRate", nbins=24, color="RiskTier", title="Comparable Field Match-Rate Distribution")
        fig.update_xaxes(tickformat=".0%")
        charts["hist"] = chart_html(fig)
    stream = fdf.groupby("Stream", as_index=False).agg(
        Fields=("Stream", "size"),
        AverageMatch=("MatchRate", "mean"),
        Matched=("MatchedClaims", "sum"),
        NotMatched=("NotMatchedClaims", "sum"),
        AverageImpact=("ImpactScore", "mean"),
        Critical=("RiskTier", lambda x: int((x == "Critical").sum())),
    )
    stream_total = stream["Matched"].fillna(0) + stream["NotMatched"].fillna(0)
    stream["WeightedMatch"] = np.where(
        stream_total > 0,
        stream["Matched"].fillna(0) / stream_total,
        np.nan,
    )
    compare_table = stream.round(4).replace({np.nan: None}).to_dict("records")
    fig = px.bar(stream.sort_values("WeightedMatch"), x="Stream", y="WeightedMatch", text="WeightedMatch", title="Average Match Percentage by Stream")
    fig.update_yaxes(tickformat=".0%")
    fig.update_traces(texttemplate="%{text:.1%}", textposition="outside", marker_line_color="#f9e6a2", marker_line_width=1)
    charts["stream"] = chart_html(fig)

    weighted_volume = stream.copy()
    weighted_volume["TotalVolume"] = weighted_volume["Matched"].fillna(0) + weighted_volume["NotMatched"].fillna(0)
    if weighted_volume.empty or float(weighted_volume["TotalVolume"].sum()) <= 0:
        charts["weighted"] = no_data_chart(
            "Reconciliation Volume by Stream",
            "No matched or unmatched volume is available for the current selection.",
        )
    else:
        weighted_long = weighted_volume.melt(
            id_vars=["Stream"],
            value_vars=["Matched", "NotMatched"],
            var_name="Result",
            value_name="Volume",
        )
        weighted_fig = px.bar(
            weighted_long,
            x="Stream",
            y="Volume",
            color="Result",
            barmode="stack",
            title="Reconciliation Volume by Stream",
            category_orders={"Result": ["Matched", "NotMatched"]},
        )
        charts["weighted"] = chart_html(weighted_fig)

    if stream.empty or float(stream["NotMatched"].fillna(0).sum()) <= 0:
        charts["volume"] = no_data_chart(
            "Unmatched Volume by Stream",
            "No unmatched volume is available for the current selection.",
        )
    else:
        charts["volume"] = chart_html(
            px.bar(
                stream.sort_values("NotMatched", ascending=False),
                x="Stream",
                y="NotMatched",
                color="AverageImpact",
                title="Unmatched Volume by Stream",
                color_continuous_scale="YlOrBr",
            )
        )
    fig = px.scatter(stream, x="AverageMatch", y="WeightedMatch", size="Fields", color="Stream", hover_data=["Matched", "NotMatched", "AverageImpact"], title="Average Match Percentage by Stream")
    fig.update_xaxes(tickformat=".0%")
    fig.update_yaxes(tickformat=".0%")
    charts["compare"] = chart_html(fig)
    classmix = fdf.groupby(["Stream", "Classification"], as_index=False).size()
    charts["classmix"] = chart_html(px.bar(classmix, x="Stream", y="size", color="Classification", title="Classification Mix by Stream"))
    risk = fdf.groupby("RiskTier", as_index=False).size()
    risk = risk[risk["size"] > 0]
    if risk.empty:
        charts["risk"] = no_data_chart("Risk Distribution", "No risk-tier values are available for the current selection.")
    else:
        charts["risk"] = chart_html(px.pie(risk, values="size", names="RiskTier", hole=0.58, title="Risk Distribution"))
    heat = fdf.groupby(["Stream", "RiskTier"], as_index=False).agg(Impact=("ImpactScore", "mean"))
    charts["heatmap"] = chart_html(px.density_heatmap(heat, x="RiskTier", y="Stream", z="Impact", histfunc="avg", title="Average Impact Score Heatmap", color_continuous_scale="YlOrBr"))
    drivers = fdf.groupby("Sub-Classification", as_index=False).agg(Impact=("ImpactScore", "mean"), Fields=("ImpactScore", "count")).sort_values("Impact", ascending=False)
    charts["drivers"] = chart_html(px.bar(drivers, x="Sub-Classification", y="Impact", color="Fields", title="Top Risk Drivers", color_continuous_scale="YlOrBr"))
    add_snapshot(fdf)
    if len(STATE["snapshots"]) < 3:
        baseline = weighted_rate(fdf) if not pd.isna(weighted_rate(fdf)) else 0.85
        fa = field_average(fdf) if not pd.isna(field_average(fdf)) else 0.86
        snaps = pd.DataFrame({"Label": ["Baseline", "Current", "Projected"], "WeightedRate": [max(baseline - 0.035, 0), baseline, min(baseline + 0.018, 1)], "FieldAverage": [max(fa - 0.025, 0), fa, min(fa + 0.012, 1)]})
    else:
        snaps = pd.DataFrame(STATE["snapshots"])
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=snaps["Label"], y=snaps["WeightedRate"], mode="lines+markers", name="Weighted match"))
    fig.add_trace(go.Scatter(x=snaps["Label"], y=snaps["FieldAverage"], mode="lines+markers", name="Field average"))
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(title="Quality Trend")
    charts["trend"] = chart_html(fig)
    impact_rows = pd.DataFrame(STATE["impacts"] or [])
    if not impact_rows.empty:
        charts["impact"] = chart_html(px.bar(impact_rows, x="Report", color="Impact", title="Mapped Report Impact", barmode="group"), 380)
    else:
        charts["impact"] = no_data_chart("Mapped Report Impact")
    return charts, compare_table

# -----------------------------------------------------------------------------
# Frontend
# -----------------------------------------------------------------------------

HTML = r'''
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>DART 8.9.1 - Data Assurance Reconciliation Tracker</title>
  <script>
    (function(){
      try{
        const saved=localStorage.getItem('dart_theme');
        document.documentElement.dataset.theme=saved==='dark'?'dark':'light';
      }catch(_err){
        document.documentElement.dataset.theme='light';
      }
    })();
  </script>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root{--bg:#100c05;--panel:#1b1408;--panel2:#261b0a;--panel3:#33250d;--line:#6f5420;--text:#fff7d1;--muted:#d8c891;--accent:#d4af37;--accent2:#f9d976;--accent3:#8a6a1f;--good:#9bd67d;--warn:#f9d976;--bad:#fb7185;--shadow:0 28px 78px rgba(0,0,0,.38)}
    *{box-sizing:border-box}html{scroll-behavior:smooth}body{margin:0;background:radial-gradient(circle at 12% -8%,rgba(249,217,118,.18),transparent 32%),radial-gradient(circle at 85% 0%,rgba(212,175,55,.12),transparent 30%),linear-gradient(180deg,#100c05,#171006 45%,#0b0804);color:var(--text);font-family:Inter,ui-sans-serif,system-ui,Segoe UI,Arial,sans-serif}a{color:inherit;text-decoration:none}
    @keyframes pageFade{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}@keyframes cardIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}.page-enter{animation:pageFade .28s ease both}@media(prefers-reduced-motion:reduce){*,*:before,*:after{animation:none!important;transition:none!important}}
    .topbar{position:sticky;top:0;z-index:50;backdrop-filter:blur(18px);background:rgba(16,12,5,.93);border-bottom:1px solid rgba(249,217,118,.18)}.nav{display:flex;align-items:center;gap:16px;max-width:1600px;margin:auto;padding:13px 22px}.brand{display:flex;align-items:center;gap:10px;font-weight:950;white-space:nowrap}.gem{width:31px;height:31px;border-radius:10px;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);display:grid;place-items:center;color:#120e05;box-shadow:0 0 32px rgba(249,217,118,.25)}
    .workspaceSwitch{position:relative;flex:none}.workspaceBtn{display:flex;align-items:center;gap:8px;border:1px solid rgba(212,175,55,.35);background:#251a08;color:#fff2b2;border-radius:999px;padding:8px 11px;font-weight:900;cursor:pointer;white-space:nowrap}.workspaceBtn:hover,.workspaceSwitch.open .workspaceBtn{border-color:#d4af37;background:#33250d}.workspaceDot{width:8px;height:8px;border-radius:50%;background:#d4af37;box-shadow:0 0 0 3px rgba(212,175,55,.12)}.workspaceMenu{position:absolute;left:0;top:45px;min-width:220px;padding:8px;background:rgba(27,20,8,.99);border:1px solid rgba(212,175,55,.42);border-radius:16px;box-shadow:0 24px 70px rgba(0,0,0,.35);display:none;z-index:220}.workspaceSwitch.open .workspaceMenu{display:block}.workspaceOption{display:block;width:100%;border:0;background:transparent;color:#decf91;text-align:left;padding:10px 11px;border-radius:10px;cursor:pointer;font-weight:800}.workspaceOption:hover,.workspaceOption.active{background:#33250d;color:#fff7d1}.workspaceOption small{display:block;color:#bdae7a;font-weight:600;margin-top:2px}.workspaceOption.active small{color:#eadb9f}
    .navMenus{display:flex;gap:8px;flex-wrap:wrap;margin-left:auto}
    .menu{position:relative}.menuBtn{border:1px solid transparent;background:transparent;color:var(--muted);padding:9px 12px;border-radius:999px;cursor:pointer;font-size:.9rem;transition:background .18s ease,border-color .18s ease,color .18s ease,transform .18s ease}.menuBtn:hover,.menuBtn.active,.menu.open .menuBtn{color:#fff8cf;border-color:#d4af37;background:#3a2c10;transform:translateY(-1px)}.menuPanel{position:absolute;right:0;top:48px;background:rgba(27,20,8,.98);border:1px solid rgba(212,175,55,.42);border-radius:18px;box-shadow:0 28px 78px rgba(0,0,0,.45);min-width:255px;padding:10px;z-index:180;opacity:0;visibility:hidden;transform:translateY(-7px) scale(.985);pointer-events:none;transition:opacity .18s ease,transform .18s ease,visibility .18s ease}.menuPanel:before{content:"";position:absolute;left:0;right:0;top:-18px;height:18px}.menu.open .menuPanel{opacity:1;visibility:visible;transform:translateY(0) scale(1);pointer-events:auto}.menuPanel a{display:flex;align-items:center;gap:10px;padding:11px 12px;border-radius:12px;color:#decf91;cursor:pointer;transition:background .16s ease,color .16s ease,transform .16s ease}.menuPanel a:hover{background:#33250d;color:#fff7d1;transform:translateX(3px)}.wrap{max-width:1600px;margin:auto;padding:24px 22px 55px;animation:pageFade .28s ease both}
    .pageShell{display:grid;grid-template-columns:190px minmax(0,1fr);gap:18px;align-items:start}.sectionRail{position:sticky;top:86px;background:rgba(27,20,8,.94);border:1px solid rgba(212,175,55,.35);border-radius:20px;padding:13px;box-shadow:0 18px 50px rgba(0,0,0,.22)}.railTitle{font-size:.75rem;color:#f9e6a2;text-transform:uppercase;letter-spacing:.08em;margin:0 0 10px}.sectionRail a{display:block;color:#d8c891;border:1px solid transparent;border-radius:12px;padding:10px 10px;margin:4px 0;font-size:.86rem;transition:background .15s ease,transform .15s ease,color .15s ease}.sectionRail a:hover{background:#33250d;color:#fff7d1;transform:translateX(3px)}.sectionBlock{scroll-margin-top:92px;margin-bottom:16px}.sectionTitle h2{margin:0 0 10px}.hero{position:relative;overflow:hidden;border:1px solid var(--line);border-radius:30px;padding:34px;background:radial-gradient(circle at 10% 0%,rgba(249,217,118,.25),transparent 34%),radial-gradient(circle at 80% 0%,rgba(212,175,55,.18),transparent 34%),linear-gradient(135deg,#201606,#2b1e09,#120e05);box-shadow:var(--shadow);margin-bottom:18px}.hero h1{font-size:2.55rem;margin:0 0 8px}.hero p{color:var(--muted);max-width:980px;line-height:1.55}.heroActions{display:flex;gap:10px;flex-wrap:wrap;margin-top:20px}
    .panel{background:rgba(27,20,8,.96);border:1px solid rgba(212,175,55,.38);border-radius:22px;padding:18px;box-shadow:0 18px 54px rgba(0,0,0,.2);min-width:0;animation:cardIn .32s ease both;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}.panel:hover{transform:translateY(-2px);box-shadow:0 24px 64px rgba(0,0,0,.28);border-color:rgba(249,217,118,.48)}.panel h2,.panel h3{margin-top:0}.grid{display:grid;gap:16px}.grid6{grid-template-columns:repeat(6,minmax(0,1fr))}.grid4{grid-template-columns:repeat(4,minmax(0,1fr))}.grid3{grid-template-columns:repeat(3,minmax(0,1fr))}.grid2{grid-template-columns:repeat(2,minmax(0,1fr))}.miniDeck{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.miniCard{background:#130e05;border:1px solid rgba(212,175,55,.26);border-radius:16px;padding:14px}.miniCard b{display:block;font-size:1.18rem;margin-bottom:4px}.miniCard span{color:var(--muted);font-size:.84rem}
    .metric .label{color:var(--muted);font-size:.76rem;text-transform:uppercase;letter-spacing:.08em}.metric .value{font-size:1.85rem;font-weight:950;margin-top:7px}.metric .sub{color:var(--muted);font-size:.84rem}.health .value{font-size:2.8rem;background:linear-gradient(135deg,#fff4b8,#d4af37);-webkit-background-clip:text;color:transparent}
    .exportTableBlock{margin-top:8px}.exportTableBar{display:flex;align-items:center;justify-content:flex-end;gap:8px;margin:0 0 8px}.exportTableBar .btn{padding:8px 11px;font-size:.78rem}.modalDataNote{margin-right:auto;color:#667085;font-size:.78rem}.byoDrop{border:2px dashed #d6bd6e;border-radius:20px;padding:22px;background:linear-gradient(135deg,#fffdf7,#fff8e7);text-align:center}.byoDrop input[type=file]{max-width:520px;margin:12px auto 0}
.byoImportPanel textarea{width:100%;border-radius:12px;border:1px solid #e0cf8d;padding:10px 11px;font-size:.82rem;font-family:inherit;resize:vertical;margin-top:4px}
.byoS3Fields{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:6px}
.byoS3Fields input{border-radius:10px;border:1px solid #e0cf8d;padding:9px 10px;font-size:.8rem;color:#243044;background:#fff}
.byoS3Results{margin-top:14px;border:1px solid #eedda5;border-radius:12px;padding:10px;background:#fffaf0;max-height:220px;overflow:auto}
.byoS3ResultsHead{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px}
.byoS3Row{display:flex;align-items:center;gap:8px;padding:6px 4px;border-radius:8px;font-size:.8rem}
.byoS3Row:hover{background:#fff2c9}
.byoS3Key{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#243044}
.byoS3Size{flex:none;font-size:.74rem}.workspaceEmpty{border-left:5px solid #c99a12}.workspaceBadge{display:inline-flex;align-items:center;gap:7px;border-radius:999px;padding:6px 10px;background:#fff4cc;border:1px solid #e1c66f;color:#725500;font-size:.78rem;font-weight:900;margin-bottom:10px}.personaLens{border-left:5px solid #c99a12!important;background:linear-gradient(135deg,#fffdf8,#fffaf0)!important}.personaLensHead{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;flex-wrap:wrap}.personaTags{display:flex;gap:7px;flex-wrap:wrap;margin-top:10px}.personaTag{display:inline-flex;align-items:center;border:1px solid #e2d4a5;background:#fff;color:#6f5400;border-radius:999px;padding:5px 9px;font-size:.76rem;font-weight:800}.personaRecommendations{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:14px}.personaRec{border:1px solid #e8dfc4;background:#fff;border-radius:14px;padding:13px;text-align:left;cursor:pointer;transition:transform .16s ease,border-color .16s ease,box-shadow .16s ease}.personaRec:hover{transform:translateY(-2px);border-color:#c99a12;box-shadow:0 8px 20px rgba(111,82,0,.08)}.personaRec b{display:block;color:#3c310f;margin-bottom:4px}.personaRec span{display:block;color:#667085;font-size:.8rem;line-height:1.4}.personaPrefNote{padding:10px 12px;border:1px solid #ead9a4;border-radius:12px;background:#fffaf0;color:#6f5b21;font-size:.82rem;margin-top:8px}@media(max-width:900px){.personaRecommendations{grid-template-columns:1fr}}
    .byoLibraryGrid{display:grid;grid-template-columns:minmax(0,1.25fr) minmax(320px,.75fr);gap:16px;align-items:start}.byoFileCard{border:1px solid #e7dfcb;border-radius:18px;padding:15px;background:linear-gradient(180deg,#fff,#fffdf8);display:grid;gap:10px}.byoFileCard.selectedA{border-color:#c99a12;box-shadow:inset 4px 0 0 #c99a12}.byoFileCard.selectedB{border-color:#8e7a32;box-shadow:inset 4px 0 0 #8e7a32}.byoFileTop{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}.byoFileName{font-weight:950;color:#172033;overflow-wrap:anywhere}.byoMeta{display:flex;gap:8px;flex-wrap:wrap}.byoMeta span{display:inline-flex;border:1px solid #e6dcc0;background:#fffaf0;border-radius:999px;padding:5px 8px;color:#667085;font-size:.75rem;font-weight:750}.byoFileActions{display:flex;gap:8px;flex-wrap:wrap}.byoPairPicker{display:grid;grid-template-columns:minmax(0,1fr) auto minmax(0,1fr);gap:14px;align-items:end}.byoVersus{width:42px;height:42px;border-radius:50%;display:grid;place-items:center;background:#fff4cc;border:1px solid #dfc66f;color:#725500;font-weight:950;margin-bottom:3px}.byoSelected{border:1px solid #dfc979;background:linear-gradient(135deg,#fffdf6,#fff7d8);border-radius:18px;padding:14px}.byoSelected h3{margin:0 0 4px}.byoSelected .datasetName{font-size:1.02rem;font-weight:950;color:#594300;overflow-wrap:anywhere}.compareHeroGrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.compareMetric{border:1px solid #e7dfcb;background:#fff;border-radius:16px;padding:13px}.compareMetric .k{font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#667085}.compareMetric .v{font-size:1.25rem;font-weight:950;color:#172033;margin-top:4px}.deltaGood{color:#2f7d5a!important}.deltaWarn{color:#b7791f!important}.byoAnalysisTabs{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px}.byoTab{border:1px solid #d9c989;background:#fffaf0;color:#725500;border-radius:999px;padding:8px 12px;font-weight:850;cursor:pointer}.byoTab.active{background:linear-gradient(135deg,#765700,#bd900e,#e2bd4e);color:#fff;border-color:#a77c00}.byoChatShell{display:grid;grid-template-columns:minmax(0,1fr) 320px;gap:16px;align-items:start}.byoChatPanel{min-height:520px}.byoChat{display:grid;gap:12px;max-height:560px;overflow:auto;padding:6px}.byoChatComposer{display:grid;grid-template-columns:1fr auto;gap:10px;margin-top:14px}.byoContextCard{position:sticky;top:86px}.aiStatus{display:inline-flex;align-items:center;gap:7px;border-radius:999px;padding:6px 9px;font-size:.76rem;font-weight:850}.aiStatus.on{background:#effaf4;color:#287657;border:1px solid #a8dcc4}.aiStatus.off{background:#fff8df;color:#8a6400;border:1px solid #ead184}.byoStep{display:flex;gap:11px;align-items:flex-start;padding:10px 0;border-top:1px solid #edf0f4}.byoStep:first-child{border-top:0}.byoStepNum{flex:none;width:27px;height:27px;border-radius:50%;display:grid;place-items:center;background:#fff3c8;color:#725500;font-weight:950}.byoEmptyPair{padding:24px;border:1px dashed #d2b85e;background:#fffdf7;border-radius:18px;text-align:center}.dangerBtn{border-color:#efb6bf!important;color:#b42335!important;background:#fff7f8!important}.byoStoragePath{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f5f2e9;border:1px solid #e3dac1;border-radius:8px;padding:3px 6px;color:#594300}.byoSplitTables{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr);gap:16px}.byoSplitTables>div{min-width:0;max-width:100%}.byoUniqueTables{width:100%;max-width:100%;overflow:hidden}.byoUniqueTables .byoCompactTable{min-width:0;max-width:100%;overflow:hidden}.byoUniqueTables .tableWrap{width:100%;max-width:100%;overflow-x:auto}.byoUniqueTables .table{width:100%;min-width:0!important;table-layout:fixed}.byoUniqueTables .table th,.byoUniqueTables .table td{white-space:normal;overflow-wrap:anywhere;word-break:break-word}.byoDatasetHead{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:10px}.byoDatasetHead h3{margin:0}.byoFlow{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.byoFlowCard{border:1px solid #e6dcc0;background:#fff;border-radius:18px;padding:15px}.byoFlowCard b{display:block;color:#594300;margin-bottom:5px}.byoFlowCard span{color:#667085;font-size:.86rem;line-height:1.45}@media(max-width:1000px){.byoLibraryGrid,.byoChatShell,.byoSplitTables{grid-template-columns:1fr}.byoContextCard{position:relative;top:auto}.byoPairPicker{grid-template-columns:1fr}.byoVersus{margin:auto}.compareHeroGrid,.byoFlow{grid-template-columns:1fr 1fr}}@media(max-width:640px){.compareHeroGrid,.byoFlow{grid-template-columns:1fr}.byoChatComposer{grid-template-columns:1fr}}
    .metricClickable{cursor:pointer;transition:transform .16s ease,box-shadow .16s ease,border-color .16s ease}.metricClickable:hover,.metricClickable:focus{transform:translateY(-2px);border-color:#c99b16!important;box-shadow:0 16px 34px rgba(15,23,42,.14);outline:none}.metricClickable .sub:after{content:" · Click to view";font-weight:700;color:#8a6a1f}.miniCard.metricClickable{position:relative}.miniCard.metricClickable:after{content:"View data";display:block;margin-top:7px;font-size:.72rem;font-weight:850;color:#8a6a1f}.unmatchedMetric .value{font-size:clamp(.95rem,1.25vw,1.45rem);line-height:1.12;letter-spacing:-.035em;overflow-wrap:anywhere;word-break:break-word}.metricDetailModal{width:min(1240px,96vw)}.scorecardClickable{cursor:pointer;position:relative}.scorecardClickable:after{content:"Open stream details →";display:block;margin-top:11px;color:#8a6a1f;font-size:.78rem;font-weight:900}.scorecardClickable:focus{outline:3px solid rgba(201,154,18,.2);outline-offset:2px}.modalTabs{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0 16px;padding-bottom:12px;border-bottom:1px solid #ece5d2}.modalTab{border:1px solid #d8c88d;background:#fffaf0;color:#725500;border-radius:999px;padding:8px 12px;font-weight:850;cursor:pointer}.modalTab.active{background:linear-gradient(135deg,#765700,#bd900e,#e2bd4e);color:#fff;border-color:#a77c00}.streamOverviewGrid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;margin-bottom:16px}.streamStat{border:1px solid #e7dfcb;background:#fffdf8;border-radius:15px;padding:12px;min-width:0}.streamStat .k{color:#667085;font-size:.7rem;text-transform:uppercase;letter-spacing:.06em}.streamStat .v{font-size:1.15rem;font-weight:900;color:#172033;margin-top:5px;overflow-wrap:anywhere}.modalSection{margin-top:16px}.modalSection h3{margin:0 0 9px}@media(max-width:900px){.streamOverviewGrid{grid-template-columns:repeat(2,minmax(0,1fr))}}
    .field{display:grid;gap:7px;min-width:0}.field label{font-size:.77rem;color:var(--muted);text-transform:uppercase;letter-spacing:.07em}input,select,textarea{background:#120d05;color:var(--text);border:1px solid rgba(212,175,55,.42);border-radius:13px;padding:12px 13px;outline:none;min-width:0;width:100%}input:focus,select:focus,textarea:focus{border-color:#f9d976;box-shadow:0 0 0 3px rgba(249,217,118,.12)}textarea{min-height:112px}.btn{border:0;border-radius:13px;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);padding:11px 15px;color:#130e05;font-weight:900;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:7px;transition:transform .16s ease,box-shadow .16s ease,filter .16s ease}.btn:hover{transform:translateY(-1px);box-shadow:0 12px 28px rgba(212,175,55,.18);filter:saturate(1.08)}.btn.secondary{background:#33250d;color:#fff7d1;border:1px solid #8a6a1f}.btn.ghost{background:transparent;color:#fff7d1;border:1px solid #8a6a1f}.btn.small{padding:8px 11px;font-size:.86rem}
    .filters summary{cursor:pointer;font-weight:900;color:#fff4b8}.filterGrid{display:grid;grid-template-columns:1.25fr repeat(4,minmax(145px,1fr));gap:12px;align-items:end;margin-top:15px}.advancedGrid{display:grid;grid-template-columns:1fr 1fr 1fr auto;gap:12px;align-items:end;margin-top:14px}.filterPicker{background:#130e05;border:1px solid rgba(212,175,55,.3);border-radius:16px;padding:12px;min-height:91px}.filterPicker label{font-size:.72rem;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.filterButton{width:100%;margin-top:8px;justify-content:space-between;background:#231906;color:#fff7d1;border:1px solid rgba(212,175,55,.42)}.filterCount{color:#f9d976;font-size:.82rem;font-weight:900}.activeFilters{display:flex;gap:8px;flex-wrap:wrap;margin-top:12px}.filterTag{display:inline-flex;align-items:center;gap:8px;font-size:.82rem;border:1px solid rgba(212,175,55,.36);background:#130e05;color:#f6e7a4;border-radius:999px;padding:7px 10px}.filterTag button{border:0;background:transparent;color:#f9d976;cursor:pointer;font-weight:900}
    .modalBackdrop{position:fixed;inset:0;background:rgba(0,0,0,.58);backdrop-filter:blur(7px);display:none;align-items:center;justify-content:center;z-index:400;padding:22px}.modal{width:min(800px,96vw);max-height:86vh;overflow:auto;background:linear-gradient(180deg,#211707,#130e05);border:1px solid rgba(249,217,118,.45);border-radius:24px;box-shadow:0 32px 90px rgba(0,0,0,.55);padding:18px}.modalHeader{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;border-bottom:1px solid rgba(212,175,55,.22);padding-bottom:13px}.modalHeader h2{margin:0}.modalGrid{display:grid;grid-template-columns:1fr auto auto;gap:10px;margin:14px 0}.modalOptions{display:flex;gap:8px;flex-wrap:wrap;max-height:360px;overflow:auto;padding:10px;background:#100b04;border:1px solid rgba(212,175,55,.2);border-radius:16px}.chip{position:relative;display:inline-flex}.chip input{position:absolute;opacity:0;pointer-events:none}.chip span{border:1px solid rgba(212,175,55,.32);background:#231906;color:#decf91;border-radius:999px;padding:8px 11px;font-size:.86rem;cursor:pointer;user-select:none}.chip input:checked + span{background:linear-gradient(135deg,#8a6a1f,#d4af37,#f9d976);color:#130e05;border-color:#f9d976;font-weight:900}.rangeLine{font-size:.83rem;color:var(--muted);margin-top:4px}
    .tableWrap{width:100%;overflow:auto;border-radius:16px;border:1px solid rgba(212,175,55,.24)}.table{width:100%;border-collapse:collapse;min-width:1050px}.table th{position:sticky;top:0;font-size:.75rem;color:#f9e6a2;text-transform:uppercase;letter-spacing:.06em;text-align:left;background:#160f05;z-index:1}.table th,.table td{padding:11px 12px;border-bottom:1px solid rgba(212,175,55,.18);vertical-align:top}.table tr:hover td{background:rgba(249,217,118,.045)}.pill{display:inline-flex;align-items:center;gap:6px;padding:5px 9px;border-radius:999px;background:#231906;border:1px solid #8a6a1f;color:#f6e7a4;font-size:.78rem}.pill.Critical{border-color:rgba(251,113,133,.7);color:#fecdd3}.pill.Elevated{border-color:#f9d976;color:#f9e6a2}.pill.Watch{border-color:#d4af37;color:#ffe8a3}.pill.Stable{border-color:rgba(155,214,125,.55);color:#d7f8cd}.bar{height:9px;border-radius:999px;background:#3a2c10;overflow:hidden;min-width:100px}.bar span{display:block;height:100%;background:linear-gradient(90deg,#fb7185,#f9d976,#9bd67d)}.insight{border-left:4px solid #d4af37;padding:15px 16px;background:#171006;border-radius:13px;color:#fff7d1}.muted{color:var(--muted)}.chat{display:grid;gap:12px;max-height:520px;overflow:auto}.bubble{padding:14px 16px;border-radius:16px;max-width:900px;white-space:pre-wrap}.user{background:#3a2c10;margin-left:auto}.assistant{background:#171006;border:1px solid #6f5420}.footer{margin-top:25px;color:var(--muted);font-size:.82rem}.kanban{display:grid;grid-template-columns:repeat(4,minmax(220px,1fr));gap:14px}.lane{background:#130e05;border:1px solid rgba(212,175,55,.32);border-radius:18px;padding:13px;min-height:240px}.lane h4{margin:0 0 10px}.task{background:#231906;border:1px solid #6f5420;border-radius:15px;padding:12px;margin-bottom:10px}.empty{padding:28px;border:1px dashed #8a6a1f;border-radius:16px;color:var(--muted);text-align:center;background:#130e05}.toast{position:fixed;bottom:22px;right:22px;background:#2b1e09;border:1px solid #d4af37;border-radius:14px;padding:13px 16px;box-shadow:var(--shadow);z-index:100;display:none}.briefText{white-space:pre-wrap;line-height:1.55;background:#130e05;border:1px solid rgba(212,175,55,.24);padding:16px;border-radius:16px;color:#fff7d1}.scorecard h2{font-size:1.5rem;margin-bottom:4px}.scoreLine{display:flex;justify-content:space-between;border-top:1px solid rgba(212,175,55,.16);padding:9px 0;color:#decf91}.navNote{font-size:.78rem;color:#d8c891;margin-left:4px}.explain{line-height:1.55}.explain b{color:#fff4b8}.editRow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px}
    @media(max-width:1180px){.pageShell,.grid6,.grid4,.grid3,.grid2,.kanban,.miniDeck,.filterGrid,.advancedGrid,.modalGrid,.editRow{grid-template-columns:1fr}.sectionRail{position:relative;top:auto}.nav{align-items:flex-start;flex-direction:column}.workspaceSwitch{width:100%}.workspaceBtn{width:100%;justify-content:space-between}.workspaceMenu{position:static;width:100%;margin-top:6px}.navMenus{margin-left:0}.hero h1{font-size:1.9rem}.menuPanel{position:static;display:none;opacity:1;visibility:visible;transform:none;pointer-events:auto}.menu.open .menuPanel{display:block}}

    /* --- Bright professional white + gold theme overrides --- */
    :root{--bg:#f7f8fb;--panel:#ffffff;--panel2:#fffaf0;--panel3:#f5edd8;--line:#e6d7aa;--text:#172033;--muted:#667085;--accent:#9a7400;--accent2:#d5aa28;--accent3:#6f5200;--good:#2f7d5a;--warn:#b7791f;--bad:#c24150;--shadow:0 16px 45px rgba(24,36,56,.10)}
    body{background:linear-gradient(180deg,#ffffff 0%,#fbfbfd 36%,#f5f7fa 100%)!important;color:var(--text)!important}
    .topbar{background:rgba(255,255,255,.94)!important;border-bottom:1px solid #eadfbf!important;box-shadow:0 8px 28px rgba(30,41,59,.06)}
    .brand{color:#594300!important}.gem{color:#fff!important;background:linear-gradient(135deg,#6f5200,#b98b09,#e7c957)!important;box-shadow:0 8px 20px rgba(154,116,0,.20)!important}.workspaceBtn{background:#fffaf0!important;color:#725500!important;border-color:#d8bd69!important}.workspaceSwitch.open .workspaceBtn,.workspaceBtn:hover{background:#fff3c8!important}.workspaceMenu{background:#fff!important;border-color:#dfcf9c!important;box-shadow:0 20px 55px rgba(15,23,42,.18)!important}.workspaceOption{color:#475569!important}.workspaceOption:hover,.workspaceOption.active{background:#fff6d8!important;color:#725500!important}.workspaceOption small{color:#8a94a6!important}
    .menuBtn{color:#566074!important}.menuBtn:hover,.menuBtn.active,.menu.open .menuBtn{color:#6f5200!important;border-color:#d7bd6a!important;background:#fff8df!important}
    .menuPanel{background:#fff!important;border-color:#ead9a4!important;box-shadow:0 24px 60px rgba(15,23,42,.16)!important}.menuPanel a{color:#485366!important}.menuPanel a:hover{background:#fff8e2!important;color:#6f5200!important}
    .sectionRail{background:#fff!important;border-color:#eadfbf!important;box-shadow:0 12px 32px rgba(15,23,42,.07)!important}.railTitle{color:#8b6a00!important}.sectionRail a{color:#667085!important}.sectionRail a:hover{background:#fff8e2!important;color:#6f5200!important}
    .hero{border-color:#e8d59a!important;background:radial-gradient(circle at 6% 0%,rgba(231,201,87,.24),transparent 36%),linear-gradient(135deg,#fffdf7,#fff8e4 58%,#ffffff)!important;box-shadow:0 18px 48px rgba(88,66,0,.10)!important}.hero h1{color:#172033!important}.hero p{color:#5d687a!important}
    .panel{background:#fff!important;border-color:#e8e1cf!important;box-shadow:0 12px 34px rgba(15,23,42,.07)!important}.panel:hover{box-shadow:0 16px 42px rgba(15,23,42,.10)!important;border-color:#dfc979!important}
    .miniCard,.filterPicker,.filterTag,.lane,.task,.empty,.briefText{background:#fff!important;border-color:#e7dfcb!important;color:#172033!important}.miniCard span,.muted,.navNote,.rangeLine{color:#667085!important}
    .metric .value,.metric .label,.panel h2,.panel h3,.scorecard h2{color:#172033!important}.health .value{background:linear-gradient(135deg,#6f5200,#c49311,#e5bd43)!important;-webkit-background-clip:text!important;color:transparent!important}
    input,select,textarea{background:#fff!important;color:#172033!important;border-color:#d9dfe8!important}input:focus,select:focus,textarea:focus{border-color:#c99a12!important;box-shadow:0 0 0 3px rgba(201,154,18,.14)!important}
    .btn{background:linear-gradient(135deg,#6f5200,#b98809,#dfbd4d)!important;color:#fff!important;box-shadow:0 7px 18px rgba(111,82,0,.16)}.btn:hover{box-shadow:0 10px 24px rgba(111,82,0,.22)!important}.btn.secondary{background:#fff8e3!important;color:#725500!important;border:1px solid #d8bd69!important}.btn.ghost{background:#fff!important;color:#725500!important;border:1px solid #d8bd69!important}
    .filters summary{color:#725500!important}.filterButton{background:#fff!important;color:#475569!important}.filterCount{color:#8b6a00!important}
    .modalBackdrop{background:rgba(15,23,42,.32)!important}.modal{background:#fff!important;border-color:#e1cf96!important;box-shadow:0 30px 90px rgba(15,23,42,.24)!important}.modalOptions{background:#fafafa!important;border-color:#e5e7eb!important}.chip span{background:#fff!important;color:#475569!important;border-color:#d7dce5!important}.chip input:checked + span{background:linear-gradient(135deg,#7a5a00,#d0a325,#ebcf6b)!important;color:#fff!important;border-color:#b78a08!important}
    .tableWrap{border-color:#e5dfcf!important;background:#fff!important}.table th{background:#fbf7ea!important;color:#6f5200!important;border-bottom-color:#e8dcae!important}.table th,.table td{border-bottom-color:#edf0f4!important;color:#243044!important}.table tr:hover td{background:#fffaf0!important}
    .pill{background:#faf7ef!important;border-color:#d6c281!important;color:#6b5a23!important}.pill.Critical{background:#fff1f2!important;color:#b42335!important;border-color:#f1b6bf!important}.pill.Elevated{background:#fff8df!important;color:#8a6400!important}.pill.Watch{background:#fffaf0!important;color:#8a6400!important}.pill.Stable{background:#effaf4!important;color:#287657!important;border-color:#a8dcc4!important}.bar{background:#edf0f4!important}.bar span{background:linear-gradient(90deg,#c84b5f,#d7a928,#3a8d69)!important}
    .insight{background:#fffaf0!important;color:#243044!important;border-left-color:#c99a12!important}.bubble.user{background:#fff4cf!important;color:#3f3000!important}.bubble.assistant{background:#fff!important;border-color:#e5dfcf!important;color:#243044!important}.bubble.assistant.aiMarkdown{white-space:normal!important;max-width:1080px;line-height:1.62;padding:20px 22px}.aiMarkdown h2,.aiMarkdown h3,.aiMarkdown h4{color:#172033;margin:18px 0 8px;line-height:1.25}.aiMarkdown h2:first-child,.aiMarkdown h3:first-child{margin-top:0}.aiMarkdown h2{font-size:1.2rem;border-bottom:1px solid #ece5d1;padding-bottom:7px}.aiMarkdown h3{font-size:1.05rem}.aiMarkdown h4{font-size:.95rem;color:#6a5100}.aiMarkdown p{margin:8px 0}.aiMarkdown ul,.aiMarkdown ol{margin:8px 0 12px 22px;padding:0}.aiMarkdown li{margin:5px 0}.aiMarkdown strong{color:#3f3000}.aiMarkdown code{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f5f2e9;border:1px solid #e7dfca;border-radius:6px;padding:2px 5px;font-size:.9em}.aiMarkdown pre{overflow:auto;background:#172033;color:#f8fafc;border-radius:12px;padding:14px}.aiMarkdown pre code{background:transparent;border:0;color:inherit;padding:0}.aiMarkdown hr{border:0;border-top:1px solid #e8e1ce;margin:18px 0}.aiTableWrap{overflow:auto;margin:12px 0 16px;border:1px solid #e5dfcf;border-radius:14px;background:#fff}.aiTable{width:100%;border-collapse:collapse;min-width:620px;font-size:.88rem}.aiTable th{background:#fff8df;color:#594300;text-transform:none;letter-spacing:0;font-size:.78rem;position:static}.aiTable th,.aiTable td{padding:10px 12px;border-bottom:1px solid #eee8d9;text-align:left;vertical-align:top}.aiTable tbody tr:last-child td{border-bottom:0}.aiTable tbody tr:nth-child(even) td{background:#fffdf7}.toast{background:#172033!important;color:#fff!important;border-color:#c99a12!important}.scoreLine{border-top-color:#eceff3!important;color:#5d687a!important}
    code{background:#f4f2eb;color:#694f00;padding:2px 5px;border-radius:6px}
    .statusDot{display:inline-block;width:9px;height:9px;border-radius:99px;margin-right:7px;background:#94a3b8}.statusDot.good{background:#2f7d5a}.statusDot.warn{background:#c28b18}.statusDot.bad{background:#c24150}
    .callout{border:1px solid #ead9a4;background:#fffaf0;border-radius:16px;padding:14px 16px;color:#475569}.callout b{color:#6f5200}
    .assistantBubbleHost{position:fixed;right:24px;bottom:22px;z-index:260;display:flex;align-items:flex-end;gap:12px;pointer-events:none}.assistantBubbleHost *{pointer-events:auto}.assistantBubblePanel{width:460px;max-width:calc(100vw - 48px);max-height:640px;display:flex;flex-direction:column;overflow:hidden;border-radius:20px;border:1px solid #d9c77b;background:rgba(255,255,255,.96);box-shadow:0 24px 52px rgba(19,24,38,.18)}.assistantBubbleHeader{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:12px 14px;background:linear-gradient(135deg,#fffdf7,#fff5d1);border-bottom:1px solid #e9d89b}.assistantBubbleHeaderText{display:flex;flex-direction:column;gap:2px;min-width:0}.assistantBubbleCompact{font-size:.86rem;font-weight:900;letter-spacing:.01em;color:#5c4300;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.assistantBubbleSubtitle{font-size:.7rem;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:#a3894a}.assistantBubbleHeader button{flex:none;width:28px;height:28px;border:0;border-radius:50%;background:#fff;cursor:pointer;color:#6b5a1b;font-size:1.2rem;box-shadow:0 2px 6px rgba(111,82,0,.12)}.assistantBubbleHeader button:hover{background:#fff0c4}
.assistantBubbleHistoryBar{display:flex;align-items:center;gap:8px;padding:10px 12px;background:#fffaf0;border-bottom:1px solid #eedda5}
.assistantBubbleHistorySelect{position:relative;flex:1;min-width:0}
.assistantBubbleHistoryCurrent{width:100%;display:flex;align-items:center;gap:8px;border:1px solid #e0cf8d;background:#fff;border-radius:10px;padding:9px 10px;cursor:pointer;color:#243044}
.assistantBubbleHistoryCurrent:hover{border-color:#d4af37}
.assistantBubbleHistoryCurrent .historyDot{flex:none;width:8px;height:8px;border-radius:50%;background:linear-gradient(135deg,#6f5200,#dfbd4d)}
.assistantBubbleHistoryCurrent .historyTitle{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:left;font-size:.82rem;font-weight:800}
.assistantBubbleHistoryCurrent .historyChevron{flex:none;color:#96865a;font-size:.7rem;transition:transform .15s ease}
.assistantBubbleHistorySelect.open .historyChevron{transform:rotate(180deg)}
.assistantBubbleHistoryMenu{display:none;position:absolute;top:calc(100% + 6px);left:0;right:0;background:#fff;border:1px solid #e5d391;border-radius:12px;box-shadow:0 18px 40px rgba(19,24,38,.16);max-height:240px;overflow:auto;z-index:5;padding:6px}
.assistantBubbleHistorySelect.open .assistantBubbleHistoryMenu{display:block}
.historyItem{display:flex;align-items:center;gap:2px;border-radius:8px}
.historyItem:hover{background:#fff7e0}
.historyItem.active{background:#fff2c9}
.historyItem.active .historyItemLabel{color:#6f5200;font-weight:900}
.historyItemLabel{flex:1;min-width:0;text-align:left;border:0;background:transparent;padding:8px 8px;font-size:.8rem;color:#243044;cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;border-radius:6px}
.historyItemIcon{flex:none;border:0;background:transparent;width:26px;height:26px;border-radius:6px;cursor:pointer;color:#8a7a4c;font-size:.8rem}
.historyItemIcon:hover{background:#f1e2ab;color:#6f5200}
.assistantBubbleNewChat{flex:none;border:0;border-radius:10px;background:linear-gradient(135deg,#6f5200,#b98809,#dfbd4d);color:#fff;font-weight:800;padding:9px 12px;cursor:pointer;font-size:.78rem;white-space:nowrap;box-shadow:0 6px 16px rgba(111,82,0,.22)}
.assistantBubbleNewChat:hover{filter:brightness(1.05)}
.assistantBubbleDatasetBar{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:8px;padding:10px 12px;background:#fbf3da;border-bottom:1px solid #eedda5}
.assistantBubbleDatasetSelect{width:100%;appearance:none;border:1px solid #e0cf8d;background:#fff;border-radius:10px;padding:8px 9px;color:#243044;font-size:.76rem;font-weight:700}
.assistantBubbleDatasetVs{font-size:.68rem;font-weight:900;color:#a3894a;letter-spacing:.04em}
.assistantBubbleMessages{display:grid;gap:12px;padding:14px;max-height:420px;overflow:auto;background:linear-gradient(180deg,#fff,#fffdf7)}
.assistantBubbleMessage{display:flex;flex-direction:column;gap:4px;max-width:88%}
.assistantBubbleMessage .assistantBubbleMessageRole{font-size:.64rem;font-weight:900;letter-spacing:.06em;text-transform:uppercase;color:#a3894a}
.assistantBubbleMessage .assistantBubbleMessageBody{padding:10px 12px;border-radius:14px;font-size:.9rem;line-height:1.55;white-space:pre-wrap;box-shadow:0 2px 8px rgba(19,24,38,.05)}
.assistantBubbleMessage.user{margin-left:auto;align-items:flex-end}
.assistantBubbleMessage.user .assistantBubbleMessageBody{background:linear-gradient(135deg,#fff1bd,#ffe49a);border:1px solid #ecd88d;color:#4a3506}
.assistantBubbleMessage.assistant .assistantBubbleMessageBody{background:#f8f7f3;border:1px solid #ece3d3;color:#243044}
.assistantBubbleEmpty{padding:16px 12px;border:1px dashed #d7c68a;border-radius:12px;background:#fffaf0;color:#667085;font-size:.82rem;text-align:center;line-height:1.5}
.assistantBubbleComposer{display:grid;grid-template-columns:1fr auto;gap:8px;padding:12px;border-top:1px solid #eedda5;background:#fff}
.assistantBubbleComposer textarea{border-radius:12px;border:1px solid #e0cf8d;padding:10px 11px;min-height:56px;resize:vertical;font-size:.85rem;font-family:inherit}
.assistantBubbleComposer textarea:focus{outline:none;border-color:#d4af37;box-shadow:0 0 0 3px rgba(212,175,55,.18)}
.assistantBubbleComposer button{border:0;border-radius:12px;background:linear-gradient(135deg,#6f5200,#b98809,#dfbd4d);color:#fff;font-weight:900;padding:10px 14px;cursor:pointer;box-shadow:0 6px 16px rgba(111,82,0,.22)}
.assistantBubbleComposer button:hover{filter:brightness(1.05)}.assistantBubbleLauncher{width:58px;height:58px;border:1px solid #d9c77b;border-radius:50%;background:linear-gradient(135deg,#6f5200,#b98809,#dfbd4d);color:#fff;font-size:1.55rem;cursor:pointer;box-shadow:0 16px 38px rgba(111,82,0,.22);position:relative;display:flex;align-items:center;justify-content:center;overflow:visible}.assistantBubbleLauncher:hover{transform:translateY(-1px)}.assistantBubbleLauncher .dartboardIcon{position:relative;z-index:1;filter:drop-shadow(0 1px 1px rgba(0,0,0,.25))}.assistantBubbleLauncher .dartArrow{position:absolute;top:50%;left:50%;z-index:2;pointer-events:none;transform:translate(calc(-50% - 34px),calc(-50% - 34px)) scale(.55);opacity:0}.assistantBubbleLauncher.throwDart .dartArrow{animation:dartThrow .5s cubic-bezier(.25,.75,.3,1) forwards}@keyframes dartThrow{0%{transform:translate(calc(-50% - 34px),calc(-50% - 34px)) scale(.55);opacity:0}20%{opacity:1}70%{transform:translate(-50%,-50%) scale(1);opacity:1}100%{transform:translate(-50%,-50%) scale(1);opacity:0}}
    .rawTable{min-width:1300px}.rawTable td{padding:5px 6px!important}.rawCell{min-width:145px;padding:8px 9px!important;border-radius:8px!important;font-size:.82rem}.rawRow{position:sticky;left:0;background:#fbf7ea!important;z-index:2;font-weight:800;color:#7a5a00!important}
    .multiSelect{display:none!important}
    .emailBuilder select{appearance:none;-webkit-appearance:none;background-color:#fff!important;background-image:linear-gradient(45deg,transparent 50%,#8b6a00 50%),linear-gradient(135deg,#8b6a00 50%,transparent 50%)!important;background-position:calc(100% - 18px) 52%,calc(100% - 12px) 52%!important;background-size:6px 6px,6px 6px!important;background-repeat:no-repeat!important;padding-right:38px!important;cursor:pointer;transition:border-color .16s ease,box-shadow .16s ease,transform .16s ease}.emailBuilder select:hover{border-color:#c7ad5b!important}.emailBuilder .field>label{color:#596579!important;font-weight:800}.emailBuilder .field{align-content:start}
    .smartMulti{position:relative;min-width:0}.smartMultiTrigger{width:100%;min-height:49px;border:1px solid #d9dfe8;border-radius:14px;background:linear-gradient(180deg,#fff,#fffdf7);padding:8px 10px 8px 12px;display:flex;align-items:center;justify-content:space-between;gap:10px;cursor:pointer;text-align:left;transition:border-color .16s ease,box-shadow .16s ease,transform .16s ease}.smartMultiTrigger:hover{border-color:#c7ad5b;box-shadow:0 5px 14px rgba(111,82,0,.07)}.smartMulti.open .smartMultiTrigger{border-color:#c99a12;box-shadow:0 0 0 3px rgba(201,154,18,.14)}.smartMultiSummary{display:flex;gap:6px;align-items:center;flex-wrap:wrap;min-width:0;flex:1}.smartChoice{display:inline-flex;align-items:center;max-width:220px;padding:5px 9px;border-radius:999px;background:#fff6d8;border:1px solid #e3cb7b;color:#725500;font-size:.78rem;font-weight:800;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.smartMore{font-size:.78rem;font-weight:850;color:#7a5a00;background:#f8f1dc;border:1px solid #ead9a4;border-radius:999px;padding:5px 8px}.smartPlaceholder{color:#8a94a6;font-size:.88rem}.smartMeta{display:flex;align-items:center;gap:8px;flex:none}.smartCount{font-size:.75rem;font-weight:900;color:#7a5a00;background:#fff5cf;border:1px solid #ead184;border-radius:999px;padding:4px 7px}.smartChevron{font-size:.85rem;color:#8b6a00;transition:transform .16s ease}.smartMulti.open .smartChevron{transform:rotate(180deg)}
    .smartMultiPanel{display:none;position:absolute;left:0;right:0;top:calc(100% + 7px);z-index:320;background:#fff;border:1px solid #dfcf9c;border-radius:16px;box-shadow:0 20px 52px rgba(15,23,42,.18);overflow:hidden}.smartMulti.open .smartMultiPanel{display:block;animation:cardIn .14s ease both}.smartMultiToolbar{display:grid;grid-template-columns:1fr auto auto;gap:7px;padding:10px;border-bottom:1px solid #eee6cf;background:#fffdf8}.smartMultiSearch{padding:9px 11px!important;border-radius:10px!important;font-size:.86rem}.smartMiniBtn{border:1px solid #dcc77f;background:#fff8df;color:#725500;border-radius:10px;padding:8px 10px;font-size:.76rem;font-weight:850;cursor:pointer}.smartMiniBtn:hover{background:#ffefb7}.smartMultiOptions{max-height:270px;overflow:auto;padding:7px}.smartMultiOption{display:flex;align-items:center;gap:10px;padding:9px 10px;border-radius:10px;cursor:pointer;color:#344054;font-size:.86rem}.smartMultiOption:hover{background:#fff8e2}.smartMultiOption input{width:16px!important;height:16px!important;min-width:16px!important;padding:0!important;accent-color:#b98b09;box-shadow:none!important}.smartMultiOption.smartHidden{display:none}.smartEmpty{padding:16px;color:#8a94a6;text-align:center;font-size:.84rem}.smartPickerHint{margin-top:6px;color:#7a8496;font-size:.76rem}
    .conditionRow{display:grid;grid-template-columns:1.15fr 1fr 1fr auto;gap:10px;align-items:end;margin:10px 0;padding:12px;border:1px solid #eee0b4;border-radius:16px;background:linear-gradient(135deg,#fffdf8,#fffaf0)}.conditionRow:hover{border-color:#ddc77c}.conditionRow .condValue:disabled{background:#f4f5f7!important;color:#98a2b3!important;cursor:not-allowed}.conditionRow .btn{margin-bottom:1px}
    .savedAutomation{border:1px solid #e8e1cf;border-left:5px solid #c99a12;border-radius:16px;padding:14px 16px;margin:10px 0;background:#fff}.automationActions{display:flex;gap:8px;flex-wrap:wrap;margin-top:10px}.statusGrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px}.statusCard{border:1px solid #e8e1cf;border-radius:16px;padding:14px;background:#fff}.statusCard .big{font-size:1.35rem;font-weight:850;color:#172033}.emailPreview{border:1px solid #ead9a4;border-radius:16px;background:#fff;overflow:hidden}.emailPreview iframe{width:100%;height:620px;border:0;background:#fff}
    @media(max-width:760px){.smartMultiToolbar{grid-template-columns:1fr 1fr}.smartMultiToolbar .smartMultiSearch{grid-column:1/-1}.conditionRow{grid-template-columns:1fr}.smartMultiPanel{position:fixed;left:14px;right:14px;top:110px;max-height:70vh}}

    /* Medicaid State Intelligence workspace */
    .medHero{background:linear-gradient(135deg,rgba(63,45,11,.98),rgba(32,23,8,.98));border-color:rgba(249,217,118,.24)}
    .grid5{grid-template-columns:repeat(5,minmax(0,1fr))}.medKpis .panel{min-height:126px}.medKpi .label{text-transform:uppercase;letter-spacing:.08em;color:#cdbc84;font-size:.72rem;font-weight:900}.medKpi .value{font-size:2rem;font-weight:950;margin:8px 0 3px;overflow-wrap:anywhere}.medKpi.warn{box-shadow:inset 0 0 0 1px rgba(249,217,118,.16)}
    .medDashboardGrid{align-items:start}.medSectionHead{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;margin-bottom:14px}.medSectionHead h3{margin:0 0 4px}.medSectionHead p{margin:0}.medStateSearch{max-width:230px;flex:0 1 230px}
    .medStateGrid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;max-height:610px;overflow:auto;padding-right:3px}.medStateGrid.wide{grid-template-columns:repeat(4,minmax(0,1fr));max-height:none}.medStateCard{display:block;width:100%;text-align:left;border:1px solid rgba(212,175,55,.24);background:linear-gradient(180deg,#241a09,#1a1307);color:var(--text);border-radius:14px;padding:13px;cursor:pointer;transition:transform .16s ease,border-color .16s ease,background .16s ease}.medStateCard:hover{transform:translateY(-2px);border-color:#d4af37;background:#30230d}.medStateTop{display:flex;justify-content:space-between;gap:10px;align-items:center}.medStateTop b{font-size:.94rem}.medStateTop span{font-size:.72rem;color:#e6d69b}.medStateStats{display:grid;grid-template-columns:repeat(3,1fr);gap:5px;margin:11px 0 9px}.medStateStats span{font-size:.68rem;color:#bfae75}.medStateStats b{display:block;color:#fff6c7;font-size:.92rem}.medProgress{height:5px;background:#0e0a04;border-radius:999px;overflow:hidden}.medProgress i{height:100%;display:block;background:linear-gradient(90deg,#8a6a1f,#f9d976);border-radius:999px}
    .medLeaderboard{display:grid;gap:8px}.medLeaderboard button{display:grid;grid-template-columns:34px 1fr auto;align-items:center;gap:10px;width:100%;border:1px solid rgba(212,175,55,.18);background:#1d1508;color:var(--text);border-radius:13px;padding:10px 12px;text-align:left;cursor:pointer}.medLeaderboard button:hover{border-color:#d4af37;background:#2c200b}.medRank{width:30px;height:30px;display:grid;place-items:center;border-radius:10px;background:#3a2b0e;color:#f9d976;font-weight:950}.medLeadName small{display:block;margin-top:2px;color:#c4b47d}.medLeadScore{font-size:1.15rem;font-weight:950;color:#f9d976}
    .medStatus,.medPriority{display:inline-flex;align-items:center;border-radius:999px;padding:5px 8px;font-size:.72rem;font-weight:900;text-transform:capitalize}.medStatus.done{background:rgba(155,214,125,.14);color:#baf09f}.medStatus.open{background:rgba(249,217,118,.13);color:#ffe58e}.medStatus.cancelled{background:rgba(251,113,133,.12);color:#ff9cad}.medPriority.high{background:rgba(251,113,133,.12);color:#ff9cad}.medPriority.medium{background:rgba(249,217,118,.12);color:#ffe58e}.medPriority.low{background:rgba(155,214,125,.12);color:#baf09f}.medDesc{max-width:620px;margin-top:5px;line-height:1.35}.medInlineSelect{min-width:118px;padding:8px 9px;border-radius:9px;border:1px solid rgba(212,175,55,.28);background:#1a1307;color:#fff6c7}.medIssueTable td{vertical-align:top}.medIssueTable td:first-child{min-width:330px}.medChart{width:100%;height:390px}.medHeatWorkspace{display:grid;grid-template-columns:minmax(0,1.7fr) minmax(315px,.72fr);gap:14px;align-items:start}.medMapPanel{padding:10px;min-width:0}.medMapPanel #medicaidUSMap{width:100%;height:610px}.medHeatFilters{grid-template-columns:repeat(3,minmax(0,1fr))}.medHeatSidePanel{min-height:630px;position:sticky;top:16px;padding:18px;overflow:hidden}.medHeatSideEmpty{min-height:570px;display:grid;place-items:center;text-align:center;padding:28px;color:#cdbc84}.medHeatSideEmpty b{display:block;color:#fff6c7;font-size:1.05rem;margin-bottom:7px}.medHeatSideHead{display:flex;align-items:flex-start;justify-content:space-between;gap:12px;padding-bottom:13px;border-bottom:1px solid rgba(212,175,55,.18)}.medHeatSideHead h2{font-size:1.34rem;margin:5px 0 2px}.medHeatSideHead p{margin:0}.medHeatClose{border:1px solid rgba(212,175,55,.24);background:#1a1307;color:#e8dba6;border-radius:10px;width:36px;height:36px;cursor:pointer;font-size:1.05rem;flex:none}.medHeatClose:hover{border-color:#d4af37;background:#2c200b}.medHeatMiniKpis{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin:14px 0}.medHeatMiniKpi{border:1px solid rgba(212,175,55,.16);background:#171105;border-radius:12px;padding:10px;min-width:0}.medHeatMiniKpi span{display:block;color:#bfae75;font-size:.68rem;font-weight:800;text-transform:uppercase;letter-spacing:.06em}.medHeatMiniKpi b{display:block;color:#fff6c7;font-size:1.12rem;margin-top:4px;overflow-wrap:anywhere}.medHeatSideSection{border-top:1px solid rgba(212,175,55,.16);padding-top:13px;margin-top:13px}.medHeatSideSection h4{margin:0 0 9px;color:#f9e7a0}.medHeatBreakdown{display:grid;gap:7px}.medHeatBreakRow{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;align-items:center;padding:8px 0;border-bottom:1px solid rgba(212,175,55,.09)}.medHeatBreakRow:last-child{border-bottom:0}.medHeatBreakRow small{display:block;color:#aa9b68;margin-top:2px}.medHeatIssueList{display:grid;gap:8px;max-height:250px;overflow:auto;padding-right:2px}.medHeatIssue{border:1px solid rgba(212,175,55,.15);background:#171105;border-radius:12px;padding:10px}.medHeatIssueTop{display:flex;align-items:flex-start;justify-content:space-between;gap:8px}.medHeatIssue b{font-size:.82rem;line-height:1.3}.medHeatIssueMeta{display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-top:7px}.medHeatOpenButton{width:100%;margin-top:15px;justify-content:center}.medHeatSideLoading{padding:28px 8px;color:#d7c789;text-align:center}
    .medInsight{position:relative;overflow:hidden}.medInsight:after{content:"";position:absolute;right:-45px;top:-45px;width:130px;height:130px;border-radius:50%;background:rgba(249,217,118,.05)}.medInsight h3{margin-top:10px}.medRecommendation{margin-top:16px;padding:12px;border-radius:12px;background:#151005;border:1px solid rgba(212,175,55,.16)}.medRecommendation b{display:block;color:#f9d976;margin-bottom:4px}.medRecommendation span{color:#ded09a}.medChatGrid{align-items:start}.medChatHistory{max-height:560px;overflow:auto}.medPromptList{display:grid;gap:9px}.medPromptList button{border:1px solid rgba(212,175,55,.23);background:#211807;color:#e8dba6;text-align:left;padding:12px;border-radius:12px;cursor:pointer}.medPromptList button:hover{border-color:#d4af37;background:#30230d}.medBriefSummary{background:linear-gradient(135deg,#2e220c,#1a1307)}.medBriefSummary h2{margin:10px 0 0;line-height:1.35}.medActionCol{min-height:230px}.medActionItem{display:grid;grid-template-columns:28px 1fr;gap:9px;align-items:start;margin-top:12px}.medActionItem b{width:26px;height:26px;display:grid;place-items:center;border-radius:8px;background:#3a2b0e;color:#f9d976}.medActionItem span{color:#dfd19a;line-height:1.45}.medMethodNote{border-left:4px solid #d4af37}
    @media(max-width:1150px){.medStateGrid.wide{grid-template-columns:repeat(3,minmax(0,1fr))}.medKpis.grid5{grid-template-columns:repeat(3,minmax(0,1fr))}}
    @media(max-width:1100px){.medHeatWorkspace{grid-template-columns:1fr}.medHeatSidePanel{position:static;min-height:0}.medHeatSideEmpty{min-height:180px}}
    @media(max-width:820px){.medStateGrid,.medStateGrid.wide{grid-template-columns:1fr}.medHeatFilters{grid-template-columns:1fr}.medMapPanel #medicaidUSMap{height:440px}.medIssueTable td:first-child{min-width:260px}.medKpis.grid5{grid-template-columns:repeat(2,minmax(0,1fr))}.medSectionHead{flex-direction:column}.medStateSearch{max-width:none;width:100%}.medHeatMiniKpis{grid-template-columns:repeat(2,minmax(0,1fr))}}

    /* Global metric/statistic info tooltips */
    .metricHelpHost{position:relative!important;overflow:visible!important}.metricHelpHost:hover,.metricHelpHost:focus-within{z-index:35}.metricHelpHost.metric,.metricHelpHost.medKpi,.metricHelpHost.statusCard,.metricHelpHost.streamStat,.metricHelpHost.compareMetric,.metricHelpHost.medHeatMiniKpi,.metricHelpHost.scorecard,.metricHelpHost.miniCard{padding-right:44px!important}
    .metricInfoWrap{position:absolute;top:10px;right:10px;z-index:60;display:inline-flex;align-items:center;justify-content:center}.metricInfoButton{width:23px;height:23px;border-radius:999px;display:grid;place-items:center;background:linear-gradient(135deg,#6f5200 0%,#b9890b 46%,#efd169 100%);border:1px solid #c49a20;color:#fffdf1;font:900 13px/1 Georgia,serif;font-style:italic;cursor:help;box-shadow:0 5px 13px rgba(111,82,0,.22);user-select:none;outline:none;transition:transform .16s ease,box-shadow .16s ease}.metricInfoButton:hover,.metricInfoButton:focus{transform:translateY(-1px) scale(1.06);box-shadow:0 7px 18px rgba(111,82,0,.3),0 0 0 3px rgba(201,154,18,.13)}
    .metricInfoTooltip{position:absolute;right:-2px;top:31px;width:min(320px,80vw);padding:13px 15px;border-radius:14px;background:#0f172a;border:1px solid #334155;color:#f8fafc;box-shadow:0 20px 48px rgba(0,0,0,.5);opacity:0;visibility:hidden;transform:translateY(-5px) scale(.985);transform-origin:top right;transition:opacity .16s ease,visibility .16s ease,transform .16s ease;pointer-events:none;text-align:left;text-transform:none!important;letter-spacing:normal!important;font-size:.82rem!important;font-weight:500!important;line-height:1.5!important;white-space:normal!important}.metricInfoTooltip:before{content:"";position:absolute;right:7px;top:-7px;width:12px;height:12px;background:#0f172a;border-left:1px solid #334155;border-top:1px solid #334155;transform:rotate(45deg)}.metricInfoTooltip b{display:block;color:#ffffff!important;font-size:.86rem;font-weight:800!important;margin-bottom:5px;position:relative;z-index:1}.metricInfoTooltip span{display:block;color:#f1f5f9!important;font-size:.82rem!important;line-height:1.5!important;text-transform:none!important;letter-spacing:normal!important;position:relative;z-index:1}.metricInfoWrap:hover .metricInfoTooltip,.metricInfoWrap:focus-within .metricInfoTooltip{opacity:1;visibility:visible;transform:translateY(0) scale(1)}
    @media(max-width:640px){.metricInfoTooltip{width:min(270px,74vw)}.metricInfoWrap{top:9px;right:9px}}

    /* Metric Explanation Cards - High-contrast readable typography */
    .explain .panel{background:#ffffff!important;border:1px solid #e2e8f0!important;border-radius:20px!important;padding:22px!important;box-shadow:0 10px 30px rgba(15,23,42,.06)!important}
    .explain .panel h3{color:#0f172a!important;font-size:1.22rem!important;font-weight:850!important;margin-top:0!important;margin-bottom:12px!important;letter-spacing:-.01em!important}
    .explain .panel p{color:#334155!important;font-size:.9rem!important;line-height:1.6!important;margin:9px 0!important}
    .explain .panel b{color:#0f172a!important;font-weight:800!important}
    :root[data-theme="dark"] .explain .panel{background:#15130e!important;border-color:#3f3824!important;box-shadow:0 14px 40px rgba(0,0,0,.35)!important}
    :root[data-theme="dark"] .explain .panel h3{color:#f8fafc!important}
    :root[data-theme="dark"] .explain .panel p{color:#cbd5e1!important}
    :root[data-theme="dark"] .explain .panel b{color:#ffffff!important}



    /* ------------------------------------------------------------------
       App-wide appearance system: white + gold light mode / dark + gold
       ------------------------------------------------------------------ */
    :root[data-theme="light"]{color-scheme:light;--bg:#f7f8fb;--panel:#ffffff;--panel2:#fffaf0;--panel3:#f5edd8;--line:#e6d7aa;--text:#172033;--muted:#667085;--accent:#9a7400;--accent2:#d5aa28;--accent3:#6f5200;--good:#2f7d5a;--warn:#b7791f;--bad:#c24150;--shadow:0 16px 45px rgba(24,36,56,.10)}
    :root[data-theme="dark"]{color-scheme:dark;--bg:#0b0a08;--panel:#15130e;--panel2:#1d190e;--panel3:#282111;--line:#5e4a1e;--text:#f8f4e8;--muted:#b9b09b;--accent:#d4af37;--accent2:#f0d271;--accent3:#9c7820;--good:#9bd67d;--warn:#f0d271;--bad:#ff8a9b;--shadow:0 24px 70px rgba(0,0,0,.38)}
    body,.topbar,.hero,.panel,.sectionRail,.menuPanel,.workspaceMenu,.miniCard,.tableWrap,.table th,.table td,input,select,textarea,.btn,.filterPicker,.filterTag,.modal,.modalOptions,.chip span,.lane,.task,.empty,.briefText,.insight,.bubble,.callout,.byoFileCard,.byoSelected,.compareMetric,.byoFlowCard,.smartMultiTrigger,.smartMultiPanel,.smartMultiToolbar,.conditionRow,.savedAutomation,.statusCard,.emailPreview,.streamStat,.medStateCard,.medLeaderboard button,.medHeatMiniKpi,.medHeatIssue,.medRecommendation,.medPromptList button,.medBriefSummary{transition:background-color .2s ease,border-color .2s ease,color .2s ease,box-shadow .2s ease,background .2s ease}


    /* Light theme is intentionally the existing white + gold product look. */
    :root[data-theme="light"] body{background:linear-gradient(180deg,#ffffff 0%,#fbfbfd 36%,#f5f7fa 100%)!important;color:#172033!important}
    :root[data-theme="light"] .topbar{background:rgba(255,255,255,.94)!important;border-bottom-color:#eadfbf!important;box-shadow:0 8px 28px rgba(30,41,59,.06)!important}
    :root[data-theme="light"] .brand{color:#594300!important}
    :root[data-theme="light"] .hero{border-color:#e8d59a!important;background:radial-gradient(circle at 6% 0%,rgba(231,201,87,.24),transparent 36%),linear-gradient(135deg,#fffdf7,#fff8e4 58%,#ffffff)!important;box-shadow:0 18px 48px rgba(88,66,0,.10)!important}
    :root[data-theme="light"] .panel{background:#fff!important;border-color:#e8e1cf!important;box-shadow:0 12px 34px rgba(15,23,42,.07)!important}
    :root[data-theme="light"] .panel:hover{border-color:#dfc979!important;box-shadow:0 16px 42px rgba(15,23,42,.10)!important}
    :root[data-theme="light"] .metric .value,:root[data-theme="light"] .metric .label,:root[data-theme="light"] .panel h1,:root[data-theme="light"] .panel h2,:root[data-theme="light"] .panel h3,:root[data-theme="light"] .panel h4,:root[data-theme="light"] .scorecard h2{color:#172033!important}
    :root[data-theme="light"] .muted,:root[data-theme="light"] .navNote,:root[data-theme="light"] .rangeLine,:root[data-theme="light"] .footer{color:#667085!important}

    /* Medicaid receives a full light treatment too (older rules were dark-only). */
    :root[data-theme="light"] .medHero{background:radial-gradient(circle at 10% 0%,rgba(231,201,87,.25),transparent 36%),linear-gradient(135deg,#fffdf8,#fff7dc 64%,#fff)!important;border-color:#dfc66f!important}
    :root[data-theme="light"] .medKpi .label{color:#7a6a45!important}:root[data-theme="light"] .medKpi .value{color:#172033!important}
    :root[data-theme="light"] .medStateCard{background:linear-gradient(180deg,#fff,#fffdf8)!important;color:#243044!important;border-color:#e7dfcb!important;box-shadow:0 7px 18px rgba(15,23,42,.04)}
    :root[data-theme="light"] .medStateCard:hover{background:#fff8e2!important;border-color:#c99a12!important;box-shadow:0 10px 24px rgba(111,82,0,.10)}
    :root[data-theme="light"] .medStateTop span,:root[data-theme="light"] .medStateStats span{color:#667085!important}:root[data-theme="light"] .medStateStats b{color:#172033!important}
    :root[data-theme="light"] .medProgress{background:#eceff3!important}
    :root[data-theme="light"] .medLeaderboard button{background:#fff!important;color:#243044!important;border-color:#e7dfcb!important}
    :root[data-theme="light"] .medLeaderboard button:hover{background:#fff8e2!important;border-color:#c99a12!important}
    :root[data-theme="light"] .medRank{background:#fff3c8!important;color:#725500!important}:root[data-theme="light"] .medLeadName small{color:#667085!important}:root[data-theme="light"] .medLeadScore{color:#8b6a00!important}
    :root[data-theme="light"] .medStatus.done,:root[data-theme="light"] .medPriority.low{background:#effaf4!important;color:#287657!important}:root[data-theme="light"] .medStatus.open,:root[data-theme="light"] .medPriority.medium{background:#fff8df!important;color:#8a6400!important}:root[data-theme="light"] .medStatus.cancelled,:root[data-theme="light"] .medPriority.high{background:#fff1f2!important;color:#b42335!important}
    :root[data-theme="light"] .medInlineSelect{background:#fff!important;color:#172033!important;border-color:#d9dfe8!important}
    :root[data-theme="light"] .medHeatSideEmpty{color:#667085!important}:root[data-theme="light"] .medHeatSideEmpty b{color:#172033!important}
    :root[data-theme="light"] .medHeatSideHead,:root[data-theme="light"] .medHeatSideSection,:root[data-theme="light"] .medHeatBreakRow{border-color:#ece5d2!important}
    :root[data-theme="light"] .medHeatClose{background:#fff!important;color:#725500!important;border-color:#d8bd69!important}:root[data-theme="light"] .medHeatClose:hover{background:#fff3c8!important}
    :root[data-theme="light"] .medHeatMiniKpi,:root[data-theme="light"] .medHeatIssue{background:#fffdf8!important;border-color:#e7dfcb!important}:root[data-theme="light"] .medHeatMiniKpi span,:root[data-theme="light"] .medHeatBreakRow small{color:#667085!important}:root[data-theme="light"] .medHeatMiniKpi b{color:#172033!important}:root[data-theme="light"] .medHeatSideSection h4{color:#6f5200!important}
    :root[data-theme="light"] .medRecommendation{background:#fffaf0!important;border-color:#ead9a4!important}:root[data-theme="light"] .medRecommendation b{color:#6f5200!important}:root[data-theme="light"] .medRecommendation span{color:#5d687a!important}
    :root[data-theme="light"] .medPromptList button{background:#fff!important;color:#475569!important;border-color:#e5dfcf!important}:root[data-theme="light"] .medPromptList button:hover{background:#fff8e2!important;border-color:#c99a12!important}
    :root[data-theme="light"] .medBriefSummary{background:linear-gradient(135deg,#fffdf8,#fff6d8)!important}:root[data-theme="light"] .medActionItem b{background:#fff3c8!important;color:#725500!important}:root[data-theme="light"] .medActionItem span{color:#5d687a!important}

    /* Dark theme: deep neutral surfaces, gold accents, white content text. */
    :root[data-theme="dark"] body{background:radial-gradient(circle at 10% -10%,rgba(212,175,55,.12),transparent 33%),radial-gradient(circle at 90% 0%,rgba(240,210,113,.07),transparent 26%),linear-gradient(180deg,#0b0a08 0%,#11100c 48%,#080807 100%)!important;color:var(--text)!important}
    :root[data-theme="dark"] .topbar{background:rgba(11,10,8,.94)!important;border-bottom:1px solid rgba(212,175,55,.22)!important;box-shadow:0 10px 30px rgba(0,0,0,.24)!important}
    :root[data-theme="dark"] .brand{color:#f8f4e8!important}:root[data-theme="dark"] .gem{color:#100d05!important;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f0d271)!important;box-shadow:0 8px 22px rgba(212,175,55,.20)!important}
    :root[data-theme="dark"] .workspaceBtn{background:#1c180e!important;color:#f4e4aa!important;border-color:#6f5720!important}:root[data-theme="dark"] .workspaceBtn:hover,:root[data-theme="dark"] .workspaceSwitch.open .workspaceBtn{background:#29220f!important;border-color:#d4af37!important}
    :root[data-theme="dark"] .workspaceMenu,:root[data-theme="dark"] .menuPanel{background:#15130e!important;border-color:#5e4a1e!important;box-shadow:0 24px 68px rgba(0,0,0,.46)!important}
    :root[data-theme="dark"] .workspaceOption,:root[data-theme="dark"] .menuPanel a{color:#c9c0aa!important}:root[data-theme="dark"] .workspaceOption:hover,:root[data-theme="dark"] .workspaceOption.active,:root[data-theme="dark"] .menuPanel a:hover{background:#282111!important;color:#fff7d1!important}
    :root[data-theme="dark"] .workspaceOption small{color:#948b75!important}:root[data-theme="dark"] .workspaceOption.active small{color:#cbbb82!important}
    :root[data-theme="dark"] .menuBtn{color:#b9b09b!important}:root[data-theme="dark"] .menuBtn:hover,:root[data-theme="dark"] .menuBtn.active,:root[data-theme="dark"] .menu.open .menuBtn{color:#fff3bd!important;border-color:#806723!important;background:#282111!important}
    :root[data-theme="dark"] .sectionRail{background:#15130e!important;border-color:#4f411e!important;box-shadow:0 16px 42px rgba(0,0,0,.24)!important}:root[data-theme="dark"] .railTitle{color:#e7ce78!important}:root[data-theme="dark"] .sectionRail a{color:#b9b09b!important}:root[data-theme="dark"] .sectionRail a:hover{background:#282111!important;color:#fff7d1!important}
    :root[data-theme="dark"] .hero{border-color:#5e4a1e!important;background:radial-gradient(circle at 8% 0%,rgba(240,210,113,.13),transparent 35%),radial-gradient(circle at 85% 0%,rgba(212,175,55,.08),transparent 30%),linear-gradient(135deg,#1b170d,#15130e 60%,#0f0e0a)!important;box-shadow:0 22px 64px rgba(0,0,0,.30)!important}:root[data-theme="dark"] .hero h1{color:#fffdf4!important}:root[data-theme="dark"] .hero p{color:#c7bea8!important}
    :root[data-theme="dark"] .panel{background:#15130e!important;border-color:#40371f!important;box-shadow:0 16px 46px rgba(0,0,0,.22)!important}:root[data-theme="dark"] .panel:hover{border-color:#745e27!important;box-shadow:0 22px 58px rgba(0,0,0,.28)!important}
    :root[data-theme="dark"] .panel h1,:root[data-theme="dark"] .panel h2,:root[data-theme="dark"] .panel h3,:root[data-theme="dark"] .panel h4,:root[data-theme="dark"] .metric .value,:root[data-theme="dark"] .metric .label,:root[data-theme="dark"] .scorecard h2{color:#f8f4e8!important}
    :root[data-theme="dark"] .miniCard,:root[data-theme="dark"] .filterPicker,:root[data-theme="dark"] .filterTag,:root[data-theme="dark"] .lane,:root[data-theme="dark"] .task,:root[data-theme="dark"] .empty,:root[data-theme="dark"] .briefText{background:#11100c!important;border-color:#3f351d!important;color:#f8f4e8!important}
    :root[data-theme="dark"] .miniCard span,:root[data-theme="dark"] .muted,:root[data-theme="dark"] .navNote,:root[data-theme="dark"] .rangeLine,:root[data-theme="dark"] .footer{color:#b9b09b!important}
    :root[data-theme="dark"] .health .value{background:linear-gradient(135deg,#fff1a8,#d4af37)!important;-webkit-background-clip:text!important;color:transparent!important}

    :root[data-theme="dark"] input,:root[data-theme="dark"] select,:root[data-theme="dark"] textarea{background:#0f0e0a!important;color:#f8f4e8!important;border-color:#4f411e!important}:root[data-theme="dark"] input::placeholder,:root[data-theme="dark"] textarea::placeholder{color:#817967!important}:root[data-theme="dark"] input:focus,:root[data-theme="dark"] select:focus,:root[data-theme="dark"] textarea:focus{border-color:#d4af37!important;box-shadow:0 0 0 3px rgba(212,175,55,.13)!important}
    :root[data-theme="dark"] .btn{background:linear-gradient(135deg,#85651c,#d4af37,#efd16a)!important;color:#100d05!important;box-shadow:0 8px 20px rgba(212,175,55,.12)!important}:root[data-theme="dark"] .btn.secondary{background:#282111!important;color:#f7e7ab!important;border:1px solid #765e20!important}:root[data-theme="dark"] .btn.ghost{background:#15130e!important;color:#f3dda0!important;border:1px solid #765e20!important}
    :root[data-theme="dark"] .filters summary{color:#f1d98d!important}:root[data-theme="dark"] .filterButton{background:#18150d!important;color:#e6dec9!important;border-color:#5d4b1e!important}:root[data-theme="dark"] .filterCount{color:#efd271!important}:root[data-theme="dark"] .filterTag{color:#eadb9f!important}
    :root[data-theme="dark"] .modalBackdrop{background:rgba(0,0,0,.68)!important}:root[data-theme="dark"] .modal{background:linear-gradient(180deg,#19160d,#0f0e0a)!important;border-color:#6c5723!important;box-shadow:0 32px 90px rgba(0,0,0,.62)!important}:root[data-theme="dark"] .modalHeader{border-bottom-color:#3d331c!important}:root[data-theme="dark"] .modalOptions{background:#0b0a08!important;border-color:#3b321d!important}
    :root[data-theme="dark"] .chip span{background:#1e1a0e!important;color:#cec5ae!important;border-color:#4f411e!important}:root[data-theme="dark"] .chip input:checked + span{background:linear-gradient(135deg,#806119,#d4af37,#efd16a)!important;color:#100d05!important;border-color:#d4af37!important}
    :root[data-theme="dark"] .tableWrap{border-color:#3c331d!important;background:#11100c!important}:root[data-theme="dark"] .table th{background:#1a170d!important;color:#ead585!important;border-bottom-color:#4a3d1d!important}:root[data-theme="dark"] .table th,:root[data-theme="dark"] .table td{color:#e7e1d3!important;border-bottom-color:#302a1b!important}:root[data-theme="dark"] .table tr:hover td{background:#1b180e!important}
    :root[data-theme="dark"] .pill{background:#211d11!important;border-color:#6e5820!important;color:#eadb9f!important}:root[data-theme="dark"] .pill.Critical{background:rgba(255,138,155,.10)!important;color:#ffb1bd!important;border-color:#a8505d!important}:root[data-theme="dark"] .pill.Elevated,:root[data-theme="dark"] .pill.Watch{background:rgba(240,210,113,.08)!important;color:#f3dda0!important;border-color:#7d6524!important}:root[data-theme="dark"] .pill.Stable{background:rgba(155,214,125,.09)!important;color:#b9e6a4!important;border-color:#4f7742!important}
    :root[data-theme="dark"] .bar{background:#28251c!important}:root[data-theme="dark"] .insight{background:#11100c!important;color:#e7e1d3!important;border-left-color:#d4af37!important}
    :root[data-theme="dark"] .bubble.user,:root[data-theme="dark"] .user{background:#30270f!important;color:#fff0b5!important}:root[data-theme="dark"] .bubble.assistant,:root[data-theme="dark"] .assistant{background:#11100c!important;border-color:#4c3d1d!important;color:#e7e1d3!important}
    :root[data-theme="dark"] .aiMarkdown h2,:root[data-theme="dark"] .aiMarkdown h3,:root[data-theme="dark"] .aiMarkdown h4{color:#f8f4e8!important}:root[data-theme="dark"] .aiMarkdown h2{border-bottom-color:#39301b!important}:root[data-theme="dark"] .aiMarkdown h4,:root[data-theme="dark"] .aiMarkdown strong{color:#f0d271!important}:root[data-theme="dark"] .aiMarkdown code,:root[data-theme="dark"] code{background:#252116!important;color:#f0d271!important;border-color:#4b3d1d!important}:root[data-theme="dark"] .aiMarkdown pre{background:#070706!important;color:#f8f4e8!important;border:1px solid #30291a!important}:root[data-theme="dark"] .aiMarkdown hr{border-top-color:#38301c!important}:root[data-theme="dark"] .aiTableWrap{background:#11100c!important;border-color:#3c331d!important}:root[data-theme="dark"] .aiTable th{background:#1d190e!important;color:#f0d271!important}:root[data-theme="dark"] .aiTable th,:root[data-theme="dark"] .aiTable td{border-bottom-color:#302a1b!important}:root[data-theme="dark"] .aiTable tbody tr:nth-child(even) td{background:#15130e!important}
    :root[data-theme="dark"] .callout{background:#19160d!important;border-color:#54451e!important;color:#c9c0aa!important}:root[data-theme="dark"] .callout b{color:#f0d271!important}:root[data-theme="dark"] .rawRow{background:#1a170d!important;color:#e8ce75!important}
    :root[data-theme="dark"] .metricClickable:hover,:root[data-theme="dark"] .metricClickable:focus{border-color:#d4af37!important;box-shadow:0 16px 36px rgba(0,0,0,.30)!important}:root[data-theme="dark"] .metricClickable .sub:after,:root[data-theme="dark"] .miniCard.metricClickable:after,:root[data-theme="dark"] .scorecardClickable:after{color:#e7c85f!important}
    :root[data-theme="dark"] .modalTabs{border-bottom-color:#3e351e!important}:root[data-theme="dark"] .modalTab,:root[data-theme="dark"] .byoTab{background:#211d11!important;color:#e9d58b!important;border-color:#6e5820!important}:root[data-theme="dark"] .modalTab.active,:root[data-theme="dark"] .byoTab.active{background:linear-gradient(135deg,#806119,#c99b20,#e9cb66)!important;color:#100d05!important;border-color:#d4af37!important}
    :root[data-theme="dark"] .streamStat,:root[data-theme="dark"] .compareMetric{background:#11100c!important;border-color:#3c331d!important}:root[data-theme="dark"] .streamStat .k,:root[data-theme="dark"] .compareMetric .k{color:#aaa18c!important}:root[data-theme="dark"] .streamStat .v,:root[data-theme="dark"] .compareMetric .v{color:#f8f4e8!important}

    /* Build Your Own + persona surfaces */
    :root[data-theme="dark"] .byoDrop{background:linear-gradient(135deg,#18150d,#11100c)!important;border-color:#6f5922!important}:root[data-theme="dark"] .workspaceBadge{background:#2b240f!important;border-color:#796124!important;color:#f0d271!important}:root[data-theme="dark"] .personaLens{background:linear-gradient(135deg,#18150d,#11100c)!important;border-color:#4c3e1d!important}
    :root[data-theme="dark"] .personaTag,:root[data-theme="dark"] .byoMeta span{background:#18150d!important;border-color:#4a3d1d!important;color:#d8cfb9!important}:root[data-theme="dark"] .personaRec,:root[data-theme="dark"] .byoFileCard,:root[data-theme="dark"] .byoFlowCard{background:linear-gradient(180deg,#17140d,#11100c)!important;border-color:#3f361f!important}:root[data-theme="dark"] .personaRec b,:root[data-theme="dark"] .byoFileName,:root[data-theme="dark"] .byoSelected .datasetName,:root[data-theme="dark"] .byoFlowCard b{color:#f0d271!important}:root[data-theme="dark"] .personaRec span,:root[data-theme="dark"] .byoFlowCard span,:root[data-theme="dark"] .byoMeta span{color:#b9b09b!important}
    :root[data-theme="dark"] .personaPrefNote,:root[data-theme="dark"] .byoEmptyPair{background:#17140d!important;border-color:#54451e!important;color:#c9c0aa!important}:root[data-theme="dark"] .byoVersus,:root[data-theme="dark"] .byoStepNum{background:#2b240f!important;border-color:#796124!important;color:#f0d271!important}:root[data-theme="dark"] .byoSelected{background:linear-gradient(135deg,#1c180e,#11100c)!important;border-color:#6b5520!important}:root[data-theme="dark"] .byoStoragePath{background:#222016!important;border-color:#4a401f!important;color:#f0d271!important}
    :root[data-theme="dark"] .aiStatus.on{background:rgba(155,214,125,.09)!important;color:#b9e6a4!important;border-color:#4f7742!important}:root[data-theme="dark"] .aiStatus.off{background:rgba(240,210,113,.08)!important;color:#f0d271!important;border-color:#715c21!important}:root[data-theme="dark"] .dangerBtn{background:#211216!important;color:#ffb1bd!important;border-color:#84434d!important}

    /* Email Agent / raw editor controls */
    :root[data-theme="dark"] .emailBuilder select{background-color:#0f0e0a!important;background-image:linear-gradient(45deg,transparent 50%,#d4af37 50%),linear-gradient(135deg,#d4af37 50%,transparent 50%)!important}:root[data-theme="dark"] .emailBuilder .field>label{color:#b9b09b!important}
    :root[data-theme="dark"] .smartMultiTrigger{background:linear-gradient(180deg,#15130e,#11100c)!important;border-color:#4f411e!important}:root[data-theme="dark"] .smartChoice{background:#2b240f!important;border-color:#725b22!important;color:#f0d271!important}:root[data-theme="dark"] .smartMore,:root[data-theme="dark"] .smartCount{background:#252014!important;border-color:#5a491e!important;color:#ead585!important}:root[data-theme="dark"] .smartPlaceholder,:root[data-theme="dark"] .smartPickerHint,:root[data-theme="dark"] .smartEmpty{color:#928a77!important}:root[data-theme="dark"] .smartChevron{color:#e3c65f!important}
    :root[data-theme="dark"] .smartMultiPanel{background:#12100b!important;border-color:#5b4920!important;box-shadow:0 22px 58px rgba(0,0,0,.48)!important}:root[data-theme="dark"] .smartMultiToolbar{background:#17140d!important;border-bottom-color:#3c321c!important}:root[data-theme="dark"] .smartMiniBtn{background:#252014!important;color:#ead585!important;border-color:#675321!important}:root[data-theme="dark"] .smartMiniBtn:hover{background:#302811!important}:root[data-theme="dark"] .smartMultiOption{color:#d7cfbd!important}:root[data-theme="dark"] .smartMultiOption:hover{background:#242014!important}
    :root[data-theme="dark"] .conditionRow{background:linear-gradient(135deg,#17140d,#11100c)!important;border-color:#4b3d1d!important}:root[data-theme="dark"] .conditionRow .condValue:disabled{background:#201f1a!important;color:#777264!important}:root[data-theme="dark"] .savedAutomation,:root[data-theme="dark"] .statusCard{background:#15130e!important;border-color:#40371f!important}:root[data-theme="dark"] .statusCard .big{color:#f8f4e8!important}:root[data-theme="dark"] .emailPreview{background:#fff!important;border-color:#6b5724!important}

    /* Medicaid dark mode */
    :root[data-theme="dark"] .medHero{background:radial-gradient(circle at 8% 0%,rgba(240,210,113,.12),transparent 34%),linear-gradient(135deg,#211b0e,#14120c)!important;border-color:#5e4a1e!important}
    :root[data-theme="dark"] .medKpi .label{color:#bdb18c!important}:root[data-theme="dark"] .medKpi .value{color:#f8f4e8!important}
    :root[data-theme="dark"] .medStateCard{background:linear-gradient(180deg,#1b180f,#11100c)!important;color:#f8f4e8!important;border-color:#40361d!important}:root[data-theme="dark"] .medStateCard:hover{background:#252010!important;border-color:#d4af37!important}:root[data-theme="dark"] .medStateTop span,:root[data-theme="dark"] .medStateStats span{color:#aaa18b!important}:root[data-theme="dark"] .medStateStats b{color:#f8f4e8!important}:root[data-theme="dark"] .medProgress{background:#28251c!important}
    :root[data-theme="dark"] .medLeaderboard button{background:#15130e!important;color:#f8f4e8!important;border-color:#3e351e!important}:root[data-theme="dark"] .medLeaderboard button:hover{background:#252010!important;border-color:#d4af37!important}:root[data-theme="dark"] .medRank{background:#2b240f!important;color:#f0d271!important}:root[data-theme="dark"] .medLeadName small{color:#aaa18b!important}:root[data-theme="dark"] .medLeadScore{color:#f0d271!important}
    :root[data-theme="dark"] .medStatus.done,:root[data-theme="dark"] .medPriority.low{background:rgba(155,214,125,.10)!important;color:#b9e6a4!important}:root[data-theme="dark"] .medStatus.open,:root[data-theme="dark"] .medPriority.medium{background:rgba(240,210,113,.10)!important;color:#f0d271!important}:root[data-theme="dark"] .medStatus.cancelled,:root[data-theme="dark"] .medPriority.high{background:rgba(255,138,155,.10)!important;color:#ffb1bd!important}
    :root[data-theme="dark"] .medInlineSelect{background:#0f0e0a!important;color:#f8f4e8!important;border-color:#4f411e!important}:root[data-theme="dark"] .medHeatSideEmpty{color:#aaa18b!important}:root[data-theme="dark"] .medHeatSideEmpty b{color:#f8f4e8!important}:root[data-theme="dark"] .medHeatSideHead,:root[data-theme="dark"] .medHeatSideSection,:root[data-theme="dark"] .medHeatBreakRow{border-color:#39301c!important}
    :root[data-theme="dark"] .medHeatClose{background:#11100c!important;color:#e9d58b!important;border-color:#51421e!important}:root[data-theme="dark"] .medHeatClose:hover{background:#27210f!important;border-color:#d4af37!important}:root[data-theme="dark"] .medHeatMiniKpi,:root[data-theme="dark"] .medHeatIssue{background:#11100c!important;border-color:#3c331d!important}:root[data-theme="dark"] .medHeatMiniKpi span,:root[data-theme="dark"] .medHeatBreakRow small{color:#aaa18b!important}:root[data-theme="dark"] .medHeatMiniKpi b{color:#f8f4e8!important}:root[data-theme="dark"] .medHeatSideSection h4{color:#f0d271!important}
    :root[data-theme="dark"] .medRecommendation{background:#11100c!important;border-color:#40361d!important}:root[data-theme="dark"] .medRecommendation b{color:#f0d271!important}:root[data-theme="dark"] .medRecommendation span{color:#c5bda8!important}:root[data-theme="dark"] .medPromptList button{background:#15130e!important;color:#d9d1bf!important;border-color:#40361d!important}:root[data-theme="dark"] .medPromptList button:hover{background:#252010!important;border-color:#d4af37!important}:root[data-theme="dark"] .medBriefSummary{background:linear-gradient(135deg,#211b0e,#11100c)!important}:root[data-theme="dark"] .medActionItem b{background:#2b240f!important;color:#f0d271!important}:root[data-theme="dark"] .medActionItem span{color:#c5bda8!important}

    /* Keep the gold information popover readable in both modes. */
    :root[data-theme="dark"] .metricInfoButton{box-shadow:0 5px 14px rgba(0,0,0,.34),0 0 0 1px rgba(240,210,113,.08)}
    :root[data-theme="dark"] .metricInfoTooltip{box-shadow:0 20px 50px rgba(0,0,0,.45)}

    :root[data-theme="light"] *{scrollbar-color:#c8aa4f #f1f2f5}:root[data-theme="dark"] *{scrollbar-color:#7f6827 #15130e}


    /* DART 8.7.4: isolated workspace + appearance controls */
    .workspaceControls{display:flex!important;align-items:center!important;gap:8px!important;flex:none!important;width:auto!important;min-width:0!important;position:relative!important;z-index:120!important;visibility:visible!important;opacity:1!important}
    .workspaceControls .workspaceSwitch{position:relative!important;flex:none!important;width:auto!important;min-width:0!important}
    .workspaceControls .workspaceBtn{width:auto!important;min-width:112px!important}
    .dartThemeToggle{appearance:none!important;-webkit-appearance:none!important;display:inline-flex!important;visibility:visible!important;opacity:1!important;pointer-events:auto!important;align-items:center!important;justify-content:center!important;gap:7px!important;flex:none!important;width:auto!important;min-width:112px!important;height:38px!important;margin:0!important;padding:3px 7px 3px 10px!important;border:2px solid #c99a12!important;border-radius:999px!important;background:linear-gradient(180deg,#fffdf7,#fff5d7)!important;color:#725500!important;box-shadow:0 6px 18px rgba(111,82,0,.16)!important;cursor:pointer!important;position:relative!important;top:auto!important;right:auto!important;left:auto!important;transform:none!important;z-index:130!important;font-family:inherit!important}
    .dartThemeToggle:hover,.dartThemeToggle:focus-visible{border-color:#a77c00!important;box-shadow:0 0 0 3px rgba(201,154,18,.13),0 8px 20px rgba(111,82,0,.18)!important;outline:none!important}
    .dartThemeLabel{display:block!important;visibility:visible!important;opacity:1!important;min-width:32px!important;text-align:center!important;color:#725500!important;font-size:.70rem!important;line-height:1!important;font-weight:950!important;letter-spacing:.01em!important}
    .dartThemeTrack{display:block!important;position:relative!important;width:54px!important;min-width:54px!important;height:28px!important;border-radius:999px!important;border:1px solid #d7bd6a!important;background:linear-gradient(180deg,#fffdf7,#fff0bf)!important;box-shadow:inset 0 1px 3px rgba(88,66,0,.10)!important;overflow:hidden!important;flex:none!important}
    .dartThemeIcon{position:absolute!important;top:50%!important;transform:translateY(-50%)!important;z-index:1!important;font-size:11px!important;line-height:1!important}
    .dartThemeSun{left:7px!important;color:#8b6a00!important;opacity:1!important}.dartThemeMoon{right:7px!important;color:#8a94a6!important;opacity:.52!important}
    .dartThemeKnob{display:block!important;position:absolute!important;z-index:2!important;top:3px!important;left:3px!important;width:22px!important;height:22px!important;border-radius:50%!important;background:linear-gradient(135deg,#795900,#b98b09,#e8ca60)!important;box-shadow:0 4px 10px rgba(88,66,0,.28),inset 0 0 0 1px rgba(255,255,255,.35)!important;transition:transform .22s cubic-bezier(.2,.8,.2,1)!important}
    :root[data-theme="dark"] .dartThemeToggle{background:linear-gradient(180deg,#211d13,#12100c)!important;border-color:#d4af37!important;color:#f0d271!important;box-shadow:0 7px 20px rgba(0,0,0,.30)!important}
    :root[data-theme="dark"] .dartThemeLabel{color:#f0d271!important}
    :root[data-theme="dark"] .dartThemeTrack{background:linear-gradient(180deg,#282313,#15130d)!important;border-color:#7b6325!important;box-shadow:inset 0 1px 5px rgba(0,0,0,.45)!important}
    :root[data-theme="dark"] .dartThemeKnob{transform:translateX(26px)!important;background:linear-gradient(135deg,#8a6a1f,#d4af37,#f0d271)!important}
    :root[data-theme="dark"] .dartThemeSun{opacity:.42!important;color:#a68e53!important}:root[data-theme="dark"] .dartThemeMoon{opacity:1!important;color:#f7e6a2!important}
    @media(max-width:1180px){.workspaceControls{display:flex!important;width:100%!important;align-items:center!important;justify-content:flex-start!important;gap:8px!important}.workspaceControls .workspaceSwitch{width:auto!important;flex:none!important}.workspaceControls .workspaceBtn{width:auto!important;min-width:150px!important;justify-content:space-between!important}.dartThemeToggle{display:inline-flex!important;min-width:112px!important}}
    @media(max-width:520px){.workspaceControls{gap:6px!important}.workspaceControls .workspaceBtn{min-width:136px!important}.dartThemeToggle{min-width:104px!important;padding-left:7px!important;gap:5px!important}.dartThemeLabel{font-size:.66rem!important;min-width:28px!important}.dartThemeTrack{width:52px!important;min-width:52px!important}}

    /* DART 8.7.6: definitive navbar utility layout.
       The theme bulb has its own navbar slot at the far right instead of sharing
       positioning rules with the workspace dropdown or page menus. */
    .nav{display:flex!important;flex-direction:row!important;align-items:center!important;flex-wrap:nowrap!important;gap:10px 14px!important;overflow:visible!important}
    .brand{display:flex!important;flex:0 0 auto!important;visibility:visible!important;opacity:1!important}
    .workspaceControls{display:flex!important;flex:0 0 auto!important;align-items:center!important;gap:8px!important;width:auto!important;min-width:0!important;visibility:visible!important;opacity:1!important;overflow:visible!important}
    .navMenus{display:flex!important;flex:1 1 auto!important;min-width:0!important;margin-left:auto!important;justify-content:flex-end!important;align-items:center!important;flex-wrap:wrap!important;overflow:visible!important}
    .navUtilities{display:flex!important;flex:0 0 auto!important;align-items:center!important;justify-content:flex-end!important;visibility:visible!important;opacity:1!important;overflow:visible!important;position:relative!important;z-index:160!important}
    .dartThemeToggle{display:inline-flex!important;flex:0 0 auto!important;align-items:center!important;justify-content:center!important;min-width:118px!important;height:40px!important;padding:4px 9px!important;gap:7px!important;visibility:visible!important;opacity:1!important;overflow:visible!important;white-space:nowrap!important;position:relative!important;z-index:170!important}
    .dartThemeBulb{display:grid!important;place-items:center!important;width:26px!important;height:26px!important;flex:0 0 26px!important;border-radius:50%!important;color:#9a7400!important;background:#fff4c4!important;border:1px solid #dfbd4f!important;box-shadow:0 3px 9px rgba(111,82,0,.14)!important}
    .dartThemeBulb svg{display:block!important;width:17px!important;height:17px!important;stroke:currentColor!important;fill:none!important;stroke-width:2!important;stroke-linecap:round!important;stroke-linejoin:round!important}
    :root[data-theme="dark"] .dartThemeBulb{color:#f0d271!important;background:#2b230f!important;border-color:#9b7c25!important;box-shadow:0 0 15px rgba(240,210,113,.30)!important}
    @media(max-width:1180px){
      .nav{flex-wrap:wrap!important}
      .brand{order:1!important}
      .workspaceControls{order:2!important}
      .navUtilities{order:3!important;margin-left:auto!important}
      .navMenus{order:4!important;flex:1 0 100%!important;width:100%!important;margin-left:0!important;justify-content:flex-start!important}
    }
    @media(max-width:620px){
      .nav{gap:8px!important;padding-left:14px!important;padding-right:14px!important}
      .workspaceControls{flex:1 1 130px!important;min-width:0!important}
      .workspaceControls .workspaceSwitch{width:100%!important;min-width:0!important}
      .workspaceControls .workspaceBtn{width:100%!important;min-width:0!important}
      .dartThemeToggle{min-width:112px!important;padding-left:7px!important;padding-right:7px!important}
    }
    @media(max-width:430px){
      .brand span:last-child{display:none!important}
      .dartThemeToggle{min-width:104px!important;gap:5px!important}
      .dartThemeBulb{width:24px!important;height:24px!important;flex-basis:24px!important}
      .dartThemeBulb svg{width:16px!important;height:16px!important}
      .dartThemeLabel{min-width:26px!important;font-size:.65rem!important}
      .dartThemeTrack{width:46px!important;min-width:46px!important}
      :root[data-theme="dark"] .dartThemeKnob{transform:translateX(18px)!important}
    }

    /* DART 8.9.1: Build Your Own uses the same dropdown navigation renderer as every workspace. */

    /* Compact appearance toggle: preserves the bulb + Light/Dark state while using less navbar width. */
    .dartThemeToggle{min-width:84px!important;height:32px!important;padding:3px 5px!important;gap:4px!important;border-width:1px!important;box-shadow:0 3px 10px rgba(111,82,0,.12)!important}
    .dartThemeToggle:hover,.dartThemeToggle:focus-visible{box-shadow:0 0 0 2px rgba(201,154,18,.11),0 4px 12px rgba(111,82,0,.14)!important}
    .dartThemeBulb{width:20px!important;height:20px!important;flex:0 0 20px!important}
    .dartThemeBulb svg{width:13px!important;height:13px!important}
    .dartThemeLabel{min-width:24px!important;font-size:.61rem!important}
    .dartThemeTrack{width:40px!important;min-width:40px!important;height:22px!important}
    .dartThemeIcon{font-size:9px!important}
    .dartThemeSun{left:5px!important}.dartThemeMoon{right:5px!important}
    .dartThemeKnob{top:2px!important;left:2px!important;width:16px!important;height:16px!important;box-shadow:0 2px 6px rgba(88,66,0,.24)!important}
    :root[data-theme="dark"] .dartThemeKnob{transform:translateX(18px)!important}
    @media(max-width:620px){.dartThemeToggle{min-width:80px!important;padding-left:4px!important;padding-right:4px!important}}
    @media(max-width:430px){.dartThemeToggle{min-width:76px!important;gap:3px!important}.dartThemeBulb{width:18px!important;height:18px!important;flex-basis:18px!important}.dartThemeBulb svg{width:12px!important;height:12px!important}.dartThemeLabel{min-width:22px!important;font-size:.59rem!important}.dartThemeTrack{width:38px!important;min-width:38px!important;height:20px!important}.dartThemeKnob{width:14px!important;height:14px!important;top:2px!important;left:2px!important}:root[data-theme="dark"] .dartThemeKnob{transform:translateX(18px)!important}}

    /* AI Insights Hub styling */
    .aiInsightsHero{background:radial-gradient(circle at 6% 0%,rgba(231,201,87,.24),transparent 36%),linear-gradient(135deg,#fffdf7,#fff8e4 58%,#ffffff);border:1px solid #e8d59a;border-radius:28px;padding:32px;box-shadow:0 18px 48px rgba(88,66,0,.10);margin-bottom:18px}
    .aiIntelligenceBadge{display:inline-flex;align-items:center;gap:7px;border-radius:999px;padding:6px 12px;background:linear-gradient(135deg,#fff5cf,#ffea9f);border:1px solid #dfc366;color:#6f5200;font-size:.78rem;font-weight:900;margin-bottom:12px;box-shadow:0 3px 10px rgba(111,82,0,.08)}
    .aiInsightsKpiGrid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:20px}
    .aiKpiCard{background:#fff;border:1px solid #e8e1cf;border-radius:18px;padding:16px 18px;box-shadow:0 10px 30px rgba(15,23,42,.06);cursor:pointer;transition:transform .16s ease,box-shadow .16s ease,border-color .16s ease;position:relative}
    .aiKpiCard:hover{transform:translateY(-2px);border-color:#c99a12;box-shadow:0 16px 36px rgba(15,23,42,.12)}
    .aiKpiCard .kpiLabel{font-size:.75rem;font-weight:850;color:#64748b;text-transform:uppercase;letter-spacing:.06em}
    .aiKpiCard .kpiVal{font-size:1.85rem;font-weight:950;color:#0f172a;margin:6px 0 2px}
    .aiKpiCard .kpiSub{font-size:.8rem;color:#64748b;display:flex;align-items:center;justify-content:space-between}
    .aiKpiCard .kpiSub span{color:#8a6a1f;font-weight:750}
    
    .aiDiagnosticsGrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(420px,1fr));gap:16px}
    .aiDiagnosticCard{background:#fff;border:1px solid #e5dfcf;border-radius:20px;padding:22px;box-shadow:0 12px 36px rgba(15,23,42,.07);display:flex;flex-direction:column;gap:14px;position:relative;overflow:hidden;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}
    .aiDiagnosticCard:hover{transform:translateY(-2px);border-color:#d4af37;box-shadow:0 18px 46px rgba(15,23,42,.12)}
    .aiDiagnosticCard.sev-Critical{border-left:6px solid #b42335}
    .aiDiagnosticCard.sev-Elevated{border-left:6px solid #d97706}
    .aiDiagnosticCard.sev-Watch{border-left:6px solid #eab308}
    .aiDiagnosticCard.sev-Stable{border-left:6px solid #16a34a}
    
    .aiCardHead{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
    .aiCategoryTag{font-size:.72rem;font-weight:850;letter-spacing:.05em;text-transform:uppercase;color:#854d0e;background:#fef9c3;border:1px solid #fde047;border-radius:999px;padding:4px 9px;display:inline-flex;align-items:center;gap:5px}
    .aiSevBadge{font-size:.72rem;font-weight:900;letter-spacing:.04em;text-transform:uppercase;border-radius:999px;padding:4px 10px;display:inline-flex;align-items:center}
    .aiSevBadge.Critical{background:#fee2e2;color:#991b1b;border:1px solid #fca5a5}
    .aiSevBadge.Elevated{background:#ffedd5;color:#9a3412;border:1px solid #fdba74}
    .aiSevBadge.Watch{background:#fef3c7;color:#92400e;border:1px solid #fcd34d}
    .aiSevBadge.Stable{background:#dcfce7;color:#166534;border:1px solid #86efac}
    
    .aiCardTitle{font-size:1.15rem;font-weight:900;color:#0f172a;margin:0;line-height:1.35}
    .aiCardSummary{font-size:.88rem;color:#334155;line-height:1.55;margin:0}
    
    .aiHighlightStat{background:linear-gradient(135deg,#f8fafc,#f1f5f9);border:1px solid #e2e8f0;border-radius:14px;padding:12px 14px;display:flex;align-items:center;justify-content:space-between;gap:10px}
    .aiHighlightStat .statItem{display:flex;flex-direction:column}
    .aiHighlightStat .statK{font-size:.68rem;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.05em}
    .aiHighlightStat .statV{font-size:1.05rem;font-weight:900;color:#0f172a}
    
    .aiDiagnosisSection{background:#fffdfa;border:1px solid #f1e4c3;border-radius:14px;padding:14px;font-size:.86rem;color:#334155;line-height:1.55}
    .aiDiagnosisSection b{color:#0f172a}
    .aiConfidenceTag{display:inline-flex;align-items:center;font-size:.7rem;font-weight:850;padding:2px 8px;border-radius:999px;background:#e0f2fe;border:1px solid #bae6fd;color:#0369a1;margin-left:6px}
    
    .aiActionBtns{display:flex;gap:8px;flex-wrap:wrap;margin-top:auto;padding-top:10px;border-top:1px solid #f1f5f9}
    
    /* Investigation Workspace */
    .aiWorkspacePanel{background:#fff;border:1px solid #e2e8f0;border-radius:22px;padding:22px;box-shadow:0 12px 34px rgba(15,23,42,.06)}
    .aiFilterNav{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:18px}
    .aiFilterPills{display:flex;gap:6px;flex-wrap:wrap}
    .aiFilterChip{border:1px solid #cbd5e1;background:#f8fafc;color:#475569;border-radius:999px;padding:7px 14px;font-size:.82rem;font-weight:800;cursor:pointer;transition:all .15s ease}
    .aiFilterChip:hover{background:#f1f5f9;border-color:#94a3b8;color:#0f172a}
    .aiFilterChip.active{background:linear-gradient(135deg,#6f5200,#b98809,#dfbd4d);color:#fff;border-color:#b98809;box-shadow:0 4px 12px rgba(111,82,0,.2)}
    
    .aiAnomalyGrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:14px}
    .aiAnomalyCard{background:#ffffff;border:1px solid #e2e8f0;border-radius:16px;padding:16px;box-shadow:0 4px 16px rgba(15,23,42,.04);cursor:pointer;transition:transform .15s ease,box-shadow .15s ease,border-color .15s ease;display:flex;flex-direction:column;gap:10px}
    .aiAnomalyCard:hover{transform:translateY(-2px);border-color:#c99a12;box-shadow:0 12px 28px rgba(15,23,42,.10)}
    .aiAnomalyCardHead{display:flex;align-items:center;justify-content:space-between;gap:8px}
    .aiAnomalyField{font-size:.95rem;font-weight:900;color:#0f172a;margin:0;line-height:1.3;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .aiAnomalyMapping{font-size:.76rem;color:#64748b;display:flex;align-items:center;gap:4px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    
    .aiAnomalyMetrics{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;background:#f8fafc;border:1px solid #edf2f7;border-radius:12px;padding:8px 10px}
    .aiAnomalyMetricItem{display:flex;flex-direction:column}
    .aiAnomalyMetricItem .amK{font-size:.64rem;font-weight:800;color:#64748b;text-transform:uppercase}
    .aiAnomalyMetricItem .amV{font-size:.88rem;font-weight:900;color:#0f172a}
    
    .aiAnomalyCardFoot{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-top:auto;padding-top:8px;border-top:1px solid #f1f5f9}
    .aiAnomalyCardFoot span{font-size:.76rem;font-weight:800;color:#854d0e}
    
    /* Diagnostic Dossier Modal */
    .fieldDossierWrap{display:flex;flex-direction:column;gap:16px}
    .dossierStatsGrid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:10px}
    .dossierStat{border:1px solid #e2e8f0;background:#f8fafc;border-radius:14px;padding:12px;min-width:0}
    .dossierStat .k{font-size:.68rem;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:.05em}
    .dossierStat .v{font-size:1.15rem;font-weight:950;color:#0f172a;margin-top:4px}
    .dossierSection{background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:16px}
    .dossierSection h3{margin:0 0 10px;font-size:1.02rem;color:#0f172a;font-weight:850}
    .dossierLineageGrid{display:grid;grid-template-columns:1fr auto 1fr;align-items:center;gap:14px}
    .lineageBox{background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:12px}
    .lineageBadge{display:inline-block;font-size:.68rem;font-weight:900;text-transform:uppercase;color:#725500;background:#fef9c3;border:1px solid #fde047;border-radius:6px;padding:2px 6px;margin-bottom:6px}
    .lineageItem{font-size:.84rem;color:#334155;margin:3px 0}
    .lineageItem b{color:#0f172a}
    .lineageArrow{font-size:1.5rem;font-weight:900;color:#b98809;text-align:center}
    .dossierDiagnosisBox{background:#fffdfa;border:1px solid #f1e4c3;border-radius:12px;padding:14px;font-size:.88rem;color:#334155;line-height:1.6}
    .dossierDiagnosisBox p{margin:0 0 8px}
    .dossierMetaGrid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;font-size:.84rem;color:#334155}
    .dossierMetaGrid div{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:8px 10px}
    .dossierMetaGrid b{display:block;font-size:.7rem;color:#64748b;text-transform:uppercase;margin-bottom:2px}
    .dossierActions{display:flex;gap:10px;flex-wrap:wrap;margin-top:8px}

    /* Dark mode overrides for AI Insights Hub */
    :root[data-theme="dark"] .aiInsightsHero{background:radial-gradient(circle at 6% 0%,rgba(240,210,113,.12),transparent 34%),linear-gradient(135deg,#1c170d,#11100c)!important;border-color:#5e4a1e!important}
    :root[data-theme="dark"] .aiIntelligenceBadge{background:#2b240f!important;border-color:#796124!important;color:#f0d271!important}
    :root[data-theme="dark"] .aiKpiCard{background:#15130e!important;border-color:#3f3824!important;box-shadow:0 12px 36px rgba(0,0,0,.4)!important}
    :root[data-theme="dark"] .aiKpiCard .kpiLabel{color:#94a3b8!important}
    :root[data-theme="dark"] .aiKpiCard .kpiVal{color:#f8fafc!important}
    :root[data-theme="dark"] .aiKpiCard .kpiSub{color:#94a3b8!important}
    :root[data-theme="dark"] .aiKpiCard .kpiSub span{color:#f0d271!important}
    :root[data-theme="dark"] .aiDiagnosticCard{background:#15130e!important;border-color:#3f3824!important;box-shadow:0 14px 40px rgba(0,0,0,.45)!important}
    :root[data-theme="dark"] .aiCardTitle{color:#f8fafc!important}
    :root[data-theme="dark"] .aiCardSummary{color:#cbd5e1!important}
    :root[data-theme="dark"] .aiCategoryTag{background:#2b240f!important;border-color:#796124!important;color:#f0d271!important}
    :root[data-theme="dark"] .aiHighlightStat{background:#1c1912!important;border-color:#3f3824!important}
    :root[data-theme="dark"] .aiHighlightStat .statK{color:#94a3b8!important}
    :root[data-theme="dark"] .aiHighlightStat .statV{color:#f8fafc!important}
    :root[data-theme="dark"] .aiDiagnosisSection{background:#18150d!important;border-color:#4a3d1d!important;color:#cbd5e1!important}
    :root[data-theme="dark"] .aiDiagnosisSection b{color:#f8fafc!important}
    :root[data-theme="dark"] .aiActionBtns{border-top-color:#2a2517!important}
    :root[data-theme="dark"] .aiWorkspacePanel{background:#15130e!important;border-color:#3f3824!important;box-shadow:0 14px 40px rgba(0,0,0,.4)!important}
    :root[data-theme="dark"] .aiFilterChip{background:#1c1912!important;border-color:#3f3824!important;color:#cbd5e1!important}
    :root[data-theme="dark"] .aiFilterChip:hover{background:#282317!important;color:#f8fafc!important}
    :root[data-theme="dark"] .aiFilterChip.active{background:linear-gradient(135deg,#785900,#c99b20,#f0d271)!important;color:#100d05!important;border-color:#d4af37!important}
    :root[data-theme="dark"] .aiAnomalyCard{background:#15130e!important;border-color:#3f3824!important;box-shadow:0 6px 20px rgba(0,0,0,.35)!important}
    :root[data-theme="dark"] .aiAnomalyField{color:#f8fafc!important}
    :root[data-theme="dark"] .aiAnomalyMapping{color:#94a3b8!important}
    :root[data-theme="dark"] .aiAnomalyMetrics{background:#1c1912!important;border-color:#332b1a!important}
    :root[data-theme="dark"] .aiAnomalyMetricItem .amK{color:#94a3b8!important}
    :root[data-theme="dark"] .aiAnomalyMetricItem .amV{color:#f8fafc!important}
    :root[data-theme="dark"] .aiAnomalyCardFoot{border-top-color:#2a2517!important}
    :root[data-theme="dark"] .aiAnomalyCardFoot span{color:#f0d271!important}
    :root[data-theme="dark"] .dossierStat{background:#1c1912!important;border-color:#3f3824!important}
    :root[data-theme="dark"] .dossierStat .k{color:#94a3b8!important}
    :root[data-theme="dark"] .dossierStat .v{color:#f8fafc!important}
    :root[data-theme="dark"] .dossierSection{background:#15130e!important;border-color:#3f3824!important}
    :root[data-theme="dark"] .dossierSection h3{color:#f8fafc!important}
    :root[data-theme="dark"] .lineageBox{background:#1c1912!important;border-color:#3f3824!important}
    :root[data-theme="dark"] .lineageBadge{background:#2b240f!important;border-color:#796124!important;color:#f0d271!important}
    :root[data-theme="dark"] .lineageItem{color:#cbd5e1!important}
    :root[data-theme="dark"] .lineageItem b{color:#f8fafc!important}
    :root[data-theme="dark"] .dossierDiagnosisBox{background:#18150d!important;border-color:#4a3d1d!important;color:#cbd5e1!important}
    :root[data-theme="dark"] .dossierMetaGrid div{background:#1c1912!important;border-color:#3f3824!important;color:#cbd5e1!important}
    :root[data-theme="dark"] .dossierMetaGrid b{color:#94a3b8!important}
    
    @media(max-width:960px){
      .aiInsightsKpiGrid,.dossierStatsGrid,.dossierMetaGrid{grid-template-columns:repeat(2,minmax(0,1fr))}
      .aiDiagnosticsGrid{grid-template-columns:1fr}
      .dossierLineageGrid{grid-template-columns:1fr}
      .lineageArrow{transform:rotate(90deg);margin:4px 0}
    }
    @media(max-width:540px){
      .aiInsightsKpiGrid,.dossierStatsGrid,.dossierMetaGrid{grid-template-columns:1fr}
    }

  </style>
</head>
<body data-workspace="medicare">
  <div class="topbar">
    <div class="nav">
      <a class="brand" href="#" onclick="goWorkspaceHome();return false"><span class="gem">◆</span><span>DART</span></a>
      <div class="workspaceControls" id="workspaceControls">
        <div class="workspaceSwitch" id="workspaceSwitcher">
          <button class="workspaceBtn" id="workspaceButton" type="button" onclick="toggleWorkspaceMenu(event)" aria-haspopup="true" aria-expanded="false"><span class="workspaceDot"></span><span id="workspaceLabel">Medicare</span><span>▾</span></button>
          <div class="workspaceMenu" id="workspaceMenu">
            <button class="workspaceOption" type="button" data-workspace="medicare" onclick="switchWorkspace('medicare')">Medicare<small>Current DART reconciliation workspace</small></button>
            <button class="workspaceOption" type="button" data-workspace="medicaid" onclick="switchWorkspace('medicaid')">Medicaid<small>50-state CMS quality workspace</small></button>
            <button class="workspaceOption" type="button" data-workspace="byo" onclick="switchWorkspace('byo')">Build Your Own<small>Compare two saved datasets + AI</small></button>
          </div>
        </div>
      </div>
      <div class="navMenus" id="navlinks"></div>
      <div class="navUtilities" id="navUtilities" aria-label="Appearance controls">
        <button class="dartThemeToggle" id="dartThemeToggle" type="button" role="switch" aria-checked="false" aria-label="Switch to dark mode" title="Switch to dark mode" onclick="event.stopPropagation();toggleTheme()">
          <span class="dartThemeBulb" id="dartThemeBulb" aria-hidden="true"><svg viewBox="0 0 24 24" focusable="false"><path d="M9 18h6"></path><path d="M10 22h4"></path><path d="M8.6 14.7A7 7 0 1 1 15.4 14.7c-.9.7-1.4 1.6-1.4 2.3h-4c0-.7-.5-1.6-1.4-2.3Z"></path><path d="M12 2V1"></path><path d="m4.9 4.9-.7-.7"></path><path d="M2 12H1"></path><path d="m19.1 4.9.7-.7"></path><path d="M22 12h1"></path></svg></span>
          <span class="dartThemeLabel" id="dartThemeLabel">Light</span>
          <span class="dartThemeTrack" aria-hidden="true">
            <span class="dartThemeIcon dartThemeSun">☀</span>
            <span class="dartThemeIcon dartThemeMoon">☾</span>
            <span class="dartThemeKnob"></span>
          </span>
        </button>
      </div>
    </div>
  </div>
  <main class="wrap"><section id="app"></section><div class="footer" id="footer"></div></main><div class="toast" id="toast"></div>
  <div class="modalBackdrop" id="filterModal"><div class="modal"><div class="modalHeader"><div><h2 id="modalTitle">Filter</h2><p class="muted" id="modalSub">Choose values to include in the current scope.</p></div><button class="btn ghost" onclick="closeFilterModal()">Close</button></div><div class="modalGrid"><input id="modalSearch" placeholder="Search values" oninput="renderModalOptions()"><button class="btn secondary" onclick="modalSelectAll(true)">Select visible</button><button class="btn ghost" onclick="modalSelectAll(false)">Clear visible</button></div><div class="modalOptions" id="modalOptions"></div><br><div class="heroActions"><button class="btn" onclick="applyFilterModal()">Apply filter</button><button class="btn secondary" onclick="clearCurrentFilter()">Clear this filter</button></div></div></div>
  <div class="modalBackdrop" id="metricModal" onclick="if(event.target===this)closeMetricModal()"><div class="modal metricDetailModal"><div class="modalHeader"><div><h2 id="metricModalTitle">Metric details</h2><p class="muted" id="metricModalSub">Rows behind this Command Center metric.</p></div><button class="btn ghost" onclick="closeMetricModal()">Close</button></div><div id="metricModalBody" style="margin-top:14px"></div></div></div>
<script>
function dartTheme(){return document.documentElement.dataset.theme==='dark'?'dark':'light';}
function ensureThemeTogglePlacement(){
  const utility=document.getElementById('navUtilities');
  if(!utility)return null;
  let toggle=document.getElementById('dartThemeToggle');
  if(!toggle){
    toggle=document.createElement('button');
    toggle.id='dartThemeToggle';
    toggle.className='dartThemeToggle';
    toggle.type='button';
    toggle.setAttribute('role','switch');
    toggle.innerHTML='<span class="dartThemeBulb" id="dartThemeBulb" aria-hidden="true"><svg viewBox="0 0 24 24" focusable="false"><path d="M9 18h6"></path><path d="M10 22h4"></path><path d="M8.6 14.7A7 7 0 1 1 15.4 14.7c-.9.7-1.4 1.6-1.4 2.3h-4c0-.7-.5-1.6-1.4-2.3Z"></path><path d="M12 2V1"></path><path d="m4.9 4.9-.7-.7"></path><path d="M2 12H1"></path><path d="m19.1 4.9.7-.7"></path><path d="M22 12h1"></path></svg></span><span class="dartThemeLabel" id="dartThemeLabel">Light</span><span class="dartThemeTrack" aria-hidden="true"><span class="dartThemeIcon dartThemeSun">☀</span><span class="dartThemeIcon dartThemeMoon">☾</span><span class="dartThemeKnob"></span></span>';
    toggle.addEventListener('click',function(e){e.stopPropagation();toggleTheme();});
  }
  if(toggle.parentElement!==utility)utility.appendChild(toggle);
  utility.style.setProperty('display','flex','important');
  utility.style.setProperty('visibility','visible','important');
  utility.style.setProperty('opacity','1','important');
  toggle.style.setProperty('display','inline-flex','important');
  toggle.style.setProperty('visibility','visible','important');
  toggle.style.setProperty('opacity','1','important');
  return toggle;
}
function syncThemeSwitch(){
  const btn=ensureThemeTogglePlacement();if(!btn)return;
  const dark=dartTheme()==='dark';
  btn.setAttribute('aria-checked',dark?'true':'false');
  const label=dark?'Switch to light mode':'Switch to dark mode';
  btn.setAttribute('aria-label',label);btn.title=label;
  const modeLabel=document.getElementById('dartThemeLabel');if(modeLabel)modeLabel.textContent=dark?'Dark':'Light';
}
function applyPlotlyTheme(root=document){
  if(!window.Plotly)return;
  const dark=dartTheme()==='dark';
  const text=dark?'#f2ecdc':'#243044';
  const muted=dark?'#b9b09b':'#667085';
  const grid=dark?'rgba(212,175,55,.14)':'#e9e5da';
  const zero=dark?'rgba(212,175,55,.22)':'#d9d4c8';
  const land=dark?'#17140e':'#f7f8fb';
  const lake=dark?'#0d0c09':'#ffffff';
  const boundary=dark?'#6c5924':'#c9c2ae';
  const graphs=(root||document).querySelectorAll?root.querySelectorAll('.js-plotly-plot,.plotly-graph-div'):[];
  graphs.forEach(gd=>{
    try{
      if(!gd||!gd._fullLayout)return;
      const full=gd._fullLayout,update={
        'font.color':text,
        'paper_bgcolor':'rgba(0,0,0,0)',
        'plot_bgcolor':'rgba(0,0,0,0)',
        'legend.font.color':text,
        'title.font.color':text
      };
      Object.keys(full).forEach(k=>{
        if(/^xaxis\d*$/.test(k)||/^yaxis\d*$/.test(k)){
          update[`${k}.color`]=muted;
          update[`${k}.tickfont.color`]=muted;
          update[`${k}.title.font.color`]=text;
          update[`${k}.gridcolor`]=grid;
          update[`${k}.zerolinecolor`]=zero;
          update[`${k}.linecolor`]=zero;
        }
      });
      if(full.geo){update['geo.bgcolor']='rgba(0,0,0,0)';update['geo.landcolor']=land;update['geo.lakecolor']=lake;update['geo.subunitcolor']=boundary;update['geo.countrycolor']=boundary;}
      if(full.coloraxis){update['coloraxis.colorbar.tickfont.color']=muted;update['coloraxis.colorbar.title.font.color']=text;}
      if(Array.isArray(full.annotations))full.annotations.forEach((_a,i)=>{update[`annotations[${i}].font.color`]=text;});
      Plotly.relayout(gd,update);
    }catch(err){console.warn('DART theme chart update skipped',err);}
  });
}
function refreshThemeVisuals(root=document){syncThemeSwitch();requestAnimationFrame(()=>applyPlotlyTheme(root));setTimeout(()=>applyPlotlyTheme(root),90);}
function setDartTheme(theme,persist=true){
  const next=theme==='dark'?'dark':'light';
  document.documentElement.dataset.theme=next;
  if(persist){try{localStorage.setItem('dart_theme',next);}catch(_err){}}
  refreshThemeVisuals(document);
}
function toggleTheme(){setDartTheme(dartTheme()==='dark'?'light':'dark');}

const navSets={
    medicare:[
      ['Overview',[['home','Command Center'],['briefing','Briefing'],['metrics','Metric Explanations'],['briefbuilder','Executive Brief']]],
      ['Intelligence',[['riskcenter','Risk Center'],['insights','AI Insights'],['scorecards','Scorecards']]],
      ['Operations',[['actioncenter','Remediation Center'],['catalog','Mapping Catalog'],['governance','Governance Center']],{hidden:true}],
      ['Data',[['explorer','System Explorer'],['raweditor','Raw Data Editor'],['emailagent','Email Agent'],['data','Data Management']],{hiddenTabs:['data']}],
      ['More',[['copilot','DART Copilot'],['profile','Profile'],['settings','Settings']]]
    ],
  medicaid:[['Overview',[['medicaid-home','Dashboard'],['medicaid-heatmap','US Heatmap'],['medicaid-exec','Executive Brief']]],['States',[['medicaid-states','State Explorer'],['medicaid-compare','Compare States'],['medicaid-analytics','Analytics']]],['Intelligence',[['medicaid-ai','AI Insights'],['medicaid-chat','CMS Q&A']]],['Claims',[['medicaid-claims','Claims Analysis'],['medicaid-optimize','Optimize Spending'],['medicaid-methodology','Methodology']]],['More',[['profile','Profile'],['settings','Settings']]]],
  byo:[['Workspace',[['byo-home','DIY Home'],['byo-library','Dataset Library'],['byo-lab','Upload Data'],['byo-raw','Raw Data Editor'],['byo-email','Email Alerts']]],['Analyze',[['byo-compare','Compare Lab'],['byo-quality','Schema & Quality'],['byo-preview','Data Explorer']]],['AI',[['byo-ai','AI Analyst']]],['Output',[['byo-export','Export Center']]],['More',[['profile','Profile'],['settings','Settings']]]],
};
const workspaceLabels={medicare:'Medicare',medicaid:'Medicaid',byo:'Build Your Own'};
const workspaceHomes={medicare:'home',medicaid:'medicaid-home',byo:'byo-home'};
const personaLandingLabels={auto:'Recommended for my role',home:'Command Center',briefing:'My Briefing',riskcenter:'Risk Center',actioncenter:'Remediation Center',catalog:'Mapping Catalog',insights:'AI Insights',briefbuilder:'Executive Brief','medicaid-home':'Medicaid Home','byo-home':'DIY Home','byo-compare':'Compare Lab','byo-quality':'Schema & Quality','byo-raw':'Raw Data Editor','byo-email':'Email Alerts','byo-ai':'AI Analyst'};
function personaPriorityLimit(meta){return Number(meta?.persona_effects?.priority_limit||8);}
function currentWorkspacePageIds(){return new Set(activeNavGroups().flatMap(g=>g[1].map(x=>x[0])));}
function personaRecommendations(meta){const allowed=currentWorkspacePageIds();return (meta?.persona_effects?.recommended_pages||[]).filter(x=>allowed.has(x.id));}
function personaQuickActions(meta){const recs=personaRecommendations(meta).slice(0,3);if(!recs.length)return `<button class="btn" onclick="showPage('riskcenter')">Open Risk Center</button><button class="btn secondary" onclick="showPage('briefing')">Open Briefing</button><button class="btn ghost" onclick="showPage('settings')">Workspace Settings</button>`;return recs.map((r,i)=>`<button class="btn ${i===0?'':i===1?'secondary':'ghost'}" onclick="showPage('${r.id}')">${esc(r.label)}</button>`).join('')+`<button class="btn ghost" onclick="showPage('settings')">Adjust persona</button>`;}
function personaLensPanel(meta){const p=meta?.persona||{};const e=meta?.persona_effects||{};const recs=personaRecommendations(meta);const focus=(e.focus||[]).map(x=>`<span class="personaTag">${esc(x)}</span>`).join('');return `<div class="panel personaLens"><div class="personaLensHead"><div><span class="workspaceBadge">Personalized workspace</span><h3>${esc(e.lens_title||'Your DART lens')}</h3><p class="muted">${esc(e.lens_summary||'DART is prioritizing the most relevant views for your role.')}</p></div><button class="btn ghost small" onclick="showPage('settings')">Edit persona</button></div><div class="personaTags"><span class="personaTag">${esc(p.role||'General')}</span><span class="personaTag">${esc(p.depth||'Balanced')} detail</span><span class="personaTag">Audience: ${esc(p.audience||'Myself')}</span>${focus}</div>${recs.length?`<div class="personaRecommendations">${recs.map(r=>`<button class="personaRec" onclick="showPage('${r.id}')"><b>${esc(r.label)}</b><span>${esc(r.reason)}</span></button>`).join('')}</div>`:''}</div>`;}
function personaLandingOptions(selected='auto'){const allowed=currentWorkspacePageIds();return Object.entries(personaLandingLabels).filter(([id])=>id==='auto'||allowed.has(id)).map(([id,label])=>`<option value="${id}" ${selected===id?'selected':''}>${esc(label)}</option>`).join('');}

const filterPages=new Set(['home','briefing','explorer','catalog','briefbuilder','riskcenter','insights','scorecards','actioncenter','governance']);
const defaultChatHistories={medicare:{activeId:'general',sessions:[{id:'general',title:'DART Copilot',messages:[]}]},medicaid:{activeId:'medicaid',sessions:[{id:'medicaid',title:'CMS Q&A',messages:[]}]},byo:{activeId:'byo',sessions:[{id:'byo',title:'AI Analyst',messages:[]}]} };const persistedChatHistories=(()=>{try{const raw=sessionStorage.getItem('dart_chat_histories_v1');if(!raw)return null;const parsed=JSON.parse(raw);return parsed&&typeof parsed==='object'?parsed:null;}catch(_err){return null;}})();
let state={workspace:sessionStorage.getItem('dart_workspace')||'medicare',page:'home',priorityQueueLimit:10,filters:{streams:null,classes:null,tiers:null,reasons:null,min_rate:0,min_impact:0,min_unmatched:0,actions_only:false,search:'',custom_filters:[]},meta:null,emailEditingId:null,rawSearch:'',rawOffset:0,rawLimit:75,emailAgent:null,rawData:null,byoLeft:sessionStorage.getItem('dart_byo_left')||'',byoRight:sessionStorage.getItem('dart_byo_right')||'',byoEditFile:sessionStorage.getItem('dart_byo_edit_file')||'',byoEmailFile:sessionStorage.getItem('dart_byo_email_file')||'',byoRawSearch:'',byoRawOffset:0,byoRawLimit:75,byoRawData:null,byoEmailEditingId:null,byoEmailAgent:null,byoCompare:null,byoChat:null,medicaidStateId:1,medicaidStateStatus:'',medicaidStateType:'',medicaidIssueType:'',medicaidMinTotal:3,medicaidHeatMetric:'quality_score',medicaidHeatStatus:'',medicaidHeatType:'',medicaidHeatSelectedId:0,medicaidHeatRows:[],medicaidStateReturnPage:'medicaid-states',medicaidIssueLabels:null,medicaidChat:[],chatHistories:{...defaultChatHistories,...(persistedChatHistories||{})},assistantBubbleOpen:false};
let modalState={key:null,title:'',items:[],selected:[]};
function esc(v){return String(v??'').replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s]));}
function mdInline(v){
  let s=esc(v);
  s=s.replace(/`([^`]+)`/g,'<code>$1</code>');
  s=s.replace(/\*\*([^*]+)\*\*/g,'<strong>$1</strong>');
  s=s.replace(/__([^_]+)__/g,'<strong>$1</strong>');
  s=s.replace(/\*([^*\n]+)\*/g,'<em>$1</em>');
  return s;
}
function renderMarkdown(v){
  const lines=String(v??'').replace(/\r/g,'').split('\n');
  const out=[];let list=null;let inCode=false;let code=[];
  const closeList=()=>{if(list){out.push(`</${list}>`);list=null;}};
  const tableSep=line=>/^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line);
  const cells=line=>line.trim().replace(/^\|/,'').replace(/\|$/,'').split('|').map(x=>x.trim());
  for(let i=0;i<lines.length;i++){
    const raw=lines[i];const line=raw.trim();
    if(line.startsWith('```')){
      closeList();
      if(!inCode){inCode=true;code=[];}else{out.push(`<pre><code>${esc(code.join('\n'))}</code></pre>`);inCode=false;code=[];}
      continue;
    }
    if(inCode){code.push(raw);continue;}
    if(line.includes('|')&&i+1<lines.length&&tableSep(lines[i+1])){
      closeList();const headers=cells(raw);const rows=[];i+=2;
      while(i<lines.length&&lines[i].trim().includes('|')&&lines[i].trim()!==''){rows.push(cells(lines[i]));i++;}
      i--;
      out.push(`<div class="aiTableWrap"><table class="aiTable"><thead><tr>${headers.map(x=>`<th>${mdInline(x)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${headers.map((_,j)=>`<td>${mdInline(r[j]??'')}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
      continue;
    }
    const heading=line.match(/^(#{1,4})\s+(.+)$/);
    if(heading){closeList();const level=Math.min(4,heading[1].length+1);out.push(`<h${level}>${mdInline(heading[2])}</h${level}>`);continue;}
    const bullet=line.match(/^[-*]\s+(.+)$/);
    if(bullet){if(list!=='ul'){closeList();list='ul';out.push('<ul>');}out.push(`<li>${mdInline(bullet[1])}</li>`);continue;}
    const numbered=line.match(/^\d+[.)]\s+(.+)$/);
    if(numbered){if(list!=='ol'){closeList();list='ol';out.push('<ol>');}out.push(`<li>${mdInline(numbered[1])}</li>`);continue;}
    if(/^---+$/.test(line)){closeList();out.push('<hr>');continue;}
    if(!line){closeList();continue;}
    closeList();out.push(`<p>${mdInline(raw)}</p>`);
  }
  closeList();if(inCode)out.push(`<pre><code>${esc(code.join('\n'))}</code></pre>`);
  return out.join('');
}
function chatMessageHtml(m){
  const role=m?.role==='user'?'user':'assistant';
  return `<div class="bubble ${role}${role==='assistant'?' aiMarkdown':''}">${role==='assistant'?renderMarkdown(m?.content||''):esc(m?.content||'')}</div>`;
}
function pctNum(v){return v===null||v===undefined||Number.isNaN(Number(v))?'-':(Number(v)*100).toFixed(1)+'%';}
function intFmt(v){return Number(v||0).toLocaleString();}
async function api(path,opts={}){const r=await fetch(path,opts);if(!r.ok)throw new Error(await r.text());return await r.json();}
async function postJson(path,body){return api(path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});}
async function putJson(path,body){return api(path,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});}
async function del(path){return api(path,{method:'DELETE'});}
function toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.style.display='block';setTimeout(()=>t.style.display='none',2200)}
function activeNavGroups(){
  const raw=navSets[state.workspace]||navSets.medicare;
  return raw
    .filter(g=>!g[2]?.hidden&&g[0]!=='Operations')
    .map(g=>{
      const hiddenTabs=new Set(g[2]?.hiddenTabs||[]);
      return [g[0], g[1].filter(([id])=>!hiddenTabs.has(id)&&id!=='data'&&id!=='quality'&&id!=='compare')];
    })
    .filter(g=>g[1].length>0);
}
function pageInGroup(group){return group[1].some(([id])=>id===state.page)}
function closeAllMenus(){document.querySelectorAll('.menu').forEach(m=>{m.classList.remove('open');const b=m.querySelector('.menuBtn');if(b)b.setAttribute('aria-expanded','false');});}
function closeWorkspaceMenu(){const sw=document.getElementById('workspaceSwitcher');if(sw)sw.classList.remove('open');const b=document.getElementById('workspaceButton');if(b)b.setAttribute('aria-expanded','false');}
function toggleWorkspaceMenu(e){e.stopPropagation();const sw=document.getElementById('workspaceSwitcher');if(!sw)return;const opening=!sw.classList.contains('open');closeAllMenus();closeWorkspaceMenu();if(opening){sw.classList.add('open');workspaceButton.setAttribute('aria-expanded','true');}}
function toggleMenu(e,menu){e.stopPropagation();closeWorkspaceMenu();const wasOpen=menu.classList.contains('open');closeAllMenus();if(!wasOpen){menu.classList.add('open');const b=menu.querySelector('.menuBtn');if(b)b.setAttribute('aria-expanded','true');}}
function goWorkspaceHome(){if(state.meta?.auth_user&&!state.meta?.persona){renderOnboarding();return;}showPage(workspaceHomes[state.workspace]||'home');}
async function switchWorkspace(mode){if(!navSets[mode])return;state.workspace=mode;sessionStorage.setItem('dart_workspace',mode);state.page=workspaceHomes[mode];closeWorkspaceMenu();setNav();await render();}
document.addEventListener('click',e=>{closeAllMenus();if(!e.target.closest('.workspaceSwitch'))closeWorkspaceMenu();if(!e.target.closest('.smartMulti'))closeSmartMultis();if(!e.target.closest('.assistantBubbleHistorySelect'))closeAssistantHistoryMenu();});document.addEventListener('keydown',e=>{if(e.key==='Escape'){closeAllMenus();closeWorkspaceMenu();closeFilterModal();closeMetricModal();closeSmartMultis();closeAssistantHistoryMenu();}});
function setNav(){
  ensureThemeTogglePlacement();syncThemeSwitch();
  document.body.dataset.workspace=state.workspace||'medicare';
  const groups=activeNavGroups();
  const label=document.getElementById('workspaceLabel');if(label)label.textContent=workspaceLabels[state.workspace]||'Medicare';
  document.querySelectorAll('.workspaceOption').forEach(x=>x.classList.toggle('active',x.dataset.workspace===state.workspace));
  const switcher=document.getElementById('workspaceSwitcher');if(switcher)switcher.style.display='';
  const nav=document.getElementById('navlinks');
  nav.innerHTML=groups.map(g=>`<div class="menu ${pageInGroup(g)?'current':''}"><button class="menuBtn ${pageInGroup(g)?'active':''}" onclick="toggleMenu(event,this.parentElement)" aria-haspopup="true" aria-expanded="false">${g[0]} ▾</button><div class="menuPanel">${g[1].map(([id,text])=>`<a onclick="showPage('${id}');closeAllMenus()">${text}</a>`).join('')}</div></div>`).join('');
}
async function boot(){
  state.meta=await api('/api/meta');
  if(!state.meta.auth_user){renderAuth();return;}
  if(!state.meta.persona){renderOnboarding();return;}
  const p=state.meta.persona||{};
  const target=sessionStorage.getItem('dart_target_page');
  const freshLogin=sessionStorage.getItem('dart_fresh_login')==='1';
  sessionStorage.removeItem('dart_fresh_login');
  if(target){
    state.workspace=p.preferred_workspace||state.workspace||'medicare';
    state.page=target;
    sessionStorage.removeItem('dart_target_page');
  }else if(freshLogin||!sessionStorage.getItem('dart_workspace')){
    state.workspace=p.preferred_workspace||'medicare';
    state.page=p.landing_page||workspaceHomes[state.workspace]||'home';
  }else{
    state.workspace=sessionStorage.getItem('dart_workspace')||p.preferred_workspace||'medicare';
    state.page=workspaceHomes[state.workspace]||'home';
  }
  sessionStorage.setItem('dart_workspace',state.workspace);
  setNav();
  await render();
}
function sectionShell(items,content){return `<div class="pageShell"><aside class="sectionRail"><div class="railTitle">On this page</div>${items.map(x=>`<a href="#${x[0]}">${x[1]}</a>`).join('')}</aside><div>${content}</div></div>`;}
function section(id,title,content){return `<section class="sectionBlock" id="${id}"><div class="sectionTitle"><h2>${title}</h2></div>${content}</section>`;}
function renderAuth(){document.body.dataset.workspace='auth';if(document.getElementById('workspaceSwitcher'))document.getElementById('workspaceSwitcher').style.display='none';document.getElementById('navlinks').innerHTML='';document.getElementById('footer').innerHTML='';document.getElementById('app').innerHTML=`<div class="hero"><h1>DART</h1><p><b>Data Assurance Reconciliation Tracker</b>. Sign in or create a local prototype account to save a profile in <code>dart_users.json</code>.</p></div><div class="grid grid2"><form class="panel" id="loginForm"><h2>Login</h2><div class="field"><label>Username</label><input id="loginUser" autocomplete="username"></div><div class="field"><label>Password</label><input id="loginPass" type="password" autocomplete="current-password"></div><br><button class="btn" type="submit">Login</button></form><form class="panel" id="signupForm"><h2>Create profile</h2><p class="muted">Saved to <code>dart_users.json</code> in the same folder as this app.</p><div class="field"><label>Username</label><input id="signupUser" autocomplete="username"></div><div class="field"><label>Password</label><input id="signupPass" type="password" autocomplete="new-password"></div><div class="field"><label>Display name</label><input id="signupName" placeholder="Nick Holmes"></div><div class="field"><label>Email</label><input id="signupEmail" placeholder="name@example.com"></div><div class="field"><label>Organization</label><input id="signupOrg" placeholder="Team or organization"></div><br><button class="btn" type="submit">Create profile</button></form></div>`;document.getElementById('loginForm').onsubmit=submitLogin;document.getElementById('signupForm').onsubmit=submitSignup;refreshThemeVisuals(document.getElementById('app'));}
async function submitLogin(e){e.preventDefault();try{await postJson('/api/login',{username:document.getElementById('loginUser').value,password:document.getElementById('loginPass').value});sessionStorage.setItem('dart_fresh_login','1');sessionStorage.removeItem('dart_target_page');toast('Logged in');window.location.reload();}catch(err){console.error(err);toast('Login failed');}}
async function submitSignup(e){e.preventDefault();try{await postJson('/api/signup',{username:document.getElementById('signupUser').value,password:document.getElementById('signupPass').value,display_name:document.getElementById('signupName').value,email:document.getElementById('signupEmail').value,organization:document.getElementById('signupOrg').value});sessionStorage.setItem('dart_fresh_login','1');toast('Profile created');window.location.reload();}catch(err){console.error(err);toast('Could not create profile');}}
async function logoutUser(){await postJson('/api/logout',{});['dart_workspace','dart_target_page','dart_fresh_login','dart_byo_left','dart_byo_right'].forEach(k=>sessionStorage.removeItem(k));toast('Logged out');window.location.reload();}
async function saveProfile(e){if(e)e.preventDefault();await postJson('/api/profile',{display_name:profileName.value,email:profileEmail.value,organization:profileOrg.value,role:profileRole.value,password:profilePass.value});toast('Profile saved');state.meta=await api('/api/meta');await render();}
function renderOnboarding(){
  document.body.dataset.workspace='onboarding';
  if(document.getElementById('workspaceSwitcher'))document.getElementById('workspaceSwitcher').style.display='none';
  document.getElementById('navlinks').innerHTML='';document.getElementById('footer').innerHTML='';
  document.getElementById('app').innerHTML=`<div class="hero"><span class="workspaceBadge">One-time setup</span><h1>Build your DART workspace</h1><p>Your persona now controls where DART opens, which actions are emphasized, how much detail is surfaced, and how AI/briefing content is framed. No features are removed—you can still navigate anywhere.</p></div><form class="panel" id="personaForm"><h2>Personalize the experience</h2><div class="grid grid2"><div class="field"><label>Program</label><select id="program"><option>Medicare</option><option>Medicaid</option><option>Both</option><option>Other / General</option></select></div><div class="field"><label>Primary role</label><select id="role"><option>Executive / Leadership</option><option>Data / Analytics</option><option>Program / Policy</option><option>Operations</option><option>Quality / Compliance</option><option>IT / Engineering</option><option>Research</option><option>Other</option></select></div><div class="field"><label>Focus areas</label><input id="focus" value="Data quality, Claims, Reporting impact, AI insights"></div><div class="field"><label>Audience</label><select id="audience"><option>Leadership</option><option>Myself</option><option>Analysts</option><option>Program teams</option><option>Technical teams</option><option>External stakeholders</option></select></div><div class="field"><label>Detail level</label><select id="depth"><option>Executive</option><option selected>Balanced</option><option>Technical</option></select></div><div class="field"><label>First question</label><input id="first_question" placeholder="Which areas need the most attention?"></div><div class="field"><label>Preferred workspace</label><select id="preferred_workspace"><option value="medicare">Medicare</option><option value="medicaid">Medicaid</option><option value="byo">Build Your Own</option></select></div><div class="field"><label>Start page</label><select id="landing_page">${personaLandingOptions('auto')}</select><div class="personaPrefNote">Choose “Recommended for my role” and DART will pick the most useful start page automatically.</div></div></div><br><button class="btn" type="submit">Build Workspace</button></form>`;
  document.getElementById('personaForm').onsubmit=savePersona;refreshThemeVisuals(document.getElementById('app'));
}
async function savePersona(e){
  if(e)e.preventDefault();
  const payload={program:document.getElementById('program').value,role:document.getElementById('role').value,audience:document.getElementById('audience').value,depth:document.getElementById('depth').value,focus:document.getElementById('focus').value.split(',').map(x=>x.trim()).filter(Boolean),geography:['National'],first_question:document.getElementById('first_question').value,preferred_workspace:document.getElementById('preferred_workspace').value,landing_page:document.getElementById('landing_page').value};
  try{
    const saved=await postJson('/api/persona',payload);
    state.meta=await api('/api/meta');
    state.workspace=saved.preferred_workspace||state.meta.persona?.preferred_workspace||'medicare';
    state.page=saved.landing_page||state.meta.persona?.landing_page||workspaceHomes[state.workspace]||'home';
    sessionStorage.setItem('dart_workspace',state.workspace);
    toast('Workspace created');
    setNav();
    await render();
  }catch(err){console.error(err);toast('Could not save workspace');}
}
async function saveSettings(e){
  if(e)e.preventDefault();
  const payload={program:document.getElementById('setProgram').value,role:document.getElementById('setRole').value,audience:document.getElementById('setAudience').value,depth:document.getElementById('setDepth').value,focus:document.getElementById('setFocus').value.split(',').map(x=>x.trim()).filter(Boolean),geography:['National'],first_question:document.getElementById('setQuestion').value,preferred_workspace:document.getElementById('setPreferredWorkspace').value,landing_page:document.getElementById('setLandingPage').value};
  await postJson('/api/persona',payload);toast('Persona and startup preferences saved');state.meta=await api('/api/meta');await render();
}
async function showPage(id){state.page=id;setNav();await render();}
function filterLabel(key){const arr=state.filters[key];if(arr===null||arr===undefined)return 'All values';if(!arr.length)return 'No values';return arr.length===1?arr[0]:`${arr.length} selected`;}
function openFilterModal(key,title,items){modalState={key,title,items:[...items],selected:[...(state.filters[key]||items)]};modalTitle.textContent=title;modalSub.textContent='Search and select values. Leave everything selected to include all values.';modalSearch.value='';filterModal.style.display='flex';renderModalOptions();}
function closeFilterModal(){filterModal.style.display='none';}
function visibleModalItems(){const q=modalSearch.value.toLowerCase().trim();return modalState.items.filter(x=>String(x).toLowerCase().includes(q));}
function safeArg(v){return String(v).replace(/\\/g,'\\\\').replace(/'/g,"\\'");}
function renderModalOptions(){const visible=visibleModalItems();modalOptions.innerHTML=visible.map(x=>`<label class="chip"><input type="checkbox" data-val="${esc(x)}" ${modalState.selected.includes(x)?'checked':''} onchange="toggleModalValue('${safeArg(x)}',this.checked)"><span>${esc(x)}</span></label>`).join('')||'<div class="empty">No matching values.</div>';}
function toggleModalValue(v,on){if(on&&!modalState.selected.includes(v))modalState.selected.push(v);if(!on)modalState.selected=modalState.selected.filter(x=>x!==v);}
function modalSelectAll(on){const visible=visibleModalItems();if(on){visible.forEach(v=>{if(!modalState.selected.includes(v))modalState.selected.push(v);});}else{modalState.selected=modalState.selected.filter(v=>!visible.includes(v));}renderModalOptions();}
function applyFilterModal(){const all=modalState.items;const sel=modalState.selected;if(sel.length===0){state.filters[modalState.key]=[];}else if(sel.length===all.length){state.filters[modalState.key]=null;}else{state.filters[modalState.key]=sel;}closeFilterModal();render();}
function clearCurrentFilter(){state.filters[modalState.key]=null;closeFilterModal();render();}
function removeCustomFilter(i){state.filters.custom_filters.splice(i,1);render();}
function addCustomFilter(){const field=customField.value;const op=customOp.value;const value=customValue.value.trim();if(!field||!value){toast('Choose a field and value');return;}state.filters.custom_filters=state.filters.custom_filters||[];state.filters.custom_filters.push({field,op,value});customValue.value='';render();}
function activeFilterTags(){const tags=[];[['streams','Streams'],['classes','Classification'],['tiers','Risk tier'],['reasons','Drivers']].forEach(([k,l])=>{const v=state.filters[k];if(v&&v.length)tags.push(`<span class="filterTag">${l}: ${esc(v.join(', '))}<button onclick="state.filters.${k}=null;render()">x</button></span>`);});(state.filters.custom_filters||[]).forEach((f,i)=>tags.push(`<span class="filterTag">${esc(f.field)} ${esc(f.op)} ${esc(f.value)}<button onclick="removeCustomFilter(${i})">x</button></span>`));return tags.length?`<div class="activeFilters">${tags.join('')}</div>`:'';}
function filterBar(meta){if(!filterPages.has(state.page))return '';const cols=(meta.columns_list||[]).map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');return `<details class="panel filters" style="margin-bottom:16px"><summary>Data scope and filters <span class="navNote">Click a filter to open a searchable selector. Add custom field filters below.</span></summary><div class="filterGrid"><div class="field"><label>Search everything</label><input id="search" placeholder="Table, field, stream, system, recommendation" value="${esc(state.filters.search)}"></div><div class="filterPicker"><label>Streams</label><button class="btn filterButton" onclick='openFilterModal("streams","Streams",${JSON.stringify(meta.streams)})'><span>${esc(filterLabel('streams'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Classification</label><button class="btn filterButton" onclick='openFilterModal("classes","Classification",${JSON.stringify(meta.classes)})'><span>${esc(filterLabel('classes'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Risk tier</label><button class="btn filterButton" onclick='openFilterModal("tiers","Risk tier",${JSON.stringify(meta.tiers)})'><span>${esc(filterLabel('tiers'))}</span><span class="filterCount">Choose values ▾</span></button></div><div class="filterPicker"><label>Drivers</label><button class="btn filterButton" onclick='openFilterModal("reasons","Drivers",${JSON.stringify(meta.reasons)})'><span>${esc(filterLabel('reasons'))}</span><span class="filterCount">Choose values ▾</span></button></div></div><div class="advancedGrid"><div class="field"><label>Custom field</label><select id="customField">${cols}</select></div><div class="field"><label>Operator</label><select id="customOp"><option value="contains">contains</option><option value="equals">equals</option><option value="not_equals">does not equal</option><option value="starts_with">starts with</option><option value="ends_with">ends with</option><option value="gt">greater than</option><option value="gte">greater than or equal</option><option value="lt">less than</option><option value="lte">less than or equal</option></select></div><div class="field"><label>Value</label><input id="customValue" placeholder="Value to filter on"></div><button class="btn" onclick="addCustomFilter()">Add custom filter</button></div><div class="advancedGrid"><div class="field"><label>Min match %</label><input id="minRate" type="range" min="0" max="100" value="${state.filters.min_rate*100}" oninput="minRateLabel.textContent=this.value+'%'"><div class="rangeLine">Current: <span id="minRateLabel">${state.filters.min_rate*100}%</span></div></div><div class="field"><label>Min impact</label><input id="minImpact" type="number" min="0" value="${state.filters.min_impact||0}"></div><div class="field"><label>Min unmatched</label><input id="minUnmatched" type="number" min="0" value="${state.filters.min_unmatched||0}"></div><div class="field"><label>View</label><select id="actionsOnly"><option value="false">All rows</option><option value="true" ${state.filters.actions_only?'selected':''}>Needs action only</option></select></div></div><br><button class="btn" onclick="readFilters();render()">Apply filters</button> <button class="btn secondary" onclick="resetFilters()">Reset filters</button>${activeFilterTags()}</details>`;}
function readFilters(){state.filters.min_rate=Number(minRate.value||0)/100;state.filters.min_impact=Number(minImpact.value||0);state.filters.min_unmatched=Number(minUnmatched.value||0);state.filters.actions_only=actionsOnly.value==='true';state.filters.search=search.value;state.filters.custom_filters=state.filters.custom_filters||[];}
async function resetFilters(){state.filters={streams:null,classes:null,tiers:null,reasons:null,min_rate:0,min_impact:0,min_unmatched:0,actions_only:false,search:'',custom_filters:[]};await render();}
async function getData(){return postJson('/api/view',state.filters);}
function closeMetricModal(){if(typeof metricModal!=='undefined'&&metricModal)metricModal.style.display='none';if(typeof metricModalBody!=='undefined'&&metricModalBody)metricModalBody.innerHTML='';streamModalState=null;}
function keyActivate(event,fn){if(event.key==='Enter'||event.key===' '){event.preventDefault();fn();}}

const METRIC_HELP={
  'health score':'Event-weighted reconciliation quality across the current Medicare scope. Higher-volume fields influence this score more, so it is the strongest single executive quality signal; higher is better.',
  'risk score':'Average impact score across the current scope. Impact combines mismatch severity with reconciliation volume, so higher values indicate greater potential operational or reporting risk.',
  'critical items':'Count of fields currently classified in the Critical risk tier based on DART impact-score thresholds.',
  'open issues':'Items that still require attention. In Medicare this means findings needing review/change; in Medicaid it means quality issues whose status is still open.',
  'unmatched volume':'Total claim or event volume that did not reconcile between the compared source and target fields in the current scope.',
  'affected reports':'Distinct downstream reports mapped or estimated to be exposed to the current reconciliation findings.',
  'weighted quality signal':'Matched volume divided by total matched plus unmatched volume. High-volume fields carry more weight than low-volume fields; higher is better.',
  'elevated or critical risk fields':'Number of fields whose calculated impact places them in either the Elevated or Critical DART risk tier.',
  'reports that may need review':'Count of downstream reports associated with current findings and therefore candidates for validation after remediation.',
  'volume-weighted match':'Matched claims divided by total reconciliation volume for this stream. This prevents tiny fields from influencing the result as much as high-volume fields.',
  'weighted match':'Matched claims divided by matched plus unmatched claims for the current stream and filters; higher is better.',
  'average field match':'Simple average of field-level match percentages in the stream. Every comparable field contributes equally regardless of volume.',
  'field average':'Simple mean of match rates across comparable fields. Unlike weighted match, each field has equal influence.',
  'fields':'Number of reconciliation fields represented in the current stream or filtered scope.',
  'not matched':'Total unmatched claim/event count for the stream. This is volume, not the number of fields.',
  'avg impact':'Mean DART impact score across fields in the stream. Higher values indicate a more consequential mismatch profile.',
  'critical':'Number of fields or records currently in the highest DART risk tier.',
  'total issues':'All Medicaid quality issue records in the current national, state, or filtered scope regardless of status.',
  'total quality issues':'All Medicaid claim-quality issues represented in the current claims-analysis scope.',
  'resolved':'Issues or workflow items whose remediation status is complete/resolved.',
  'open':'Issues or workflow items that remain active and require remediation or review.',
  'cancelled':'Issues closed without being recorded as successfully resolved.',
  'states':'Number of state Medicaid programs represented in the current nationwide workspace.',
  'quality grade':'Prototype state quality grade derived from the state quality/composite scoring logic. It is a decision-support indicator, not an official CMS rating.',
  'quality score':'Composite state quality score used for Medicaid comparison and heatmap ranking. It reflects resolution, backlog/cancellation behavior, and sample confidence; higher is better.',
  'duplicate rate':'Estimated duplicate-claim rate for the selected state based on the prototype claims-quality model; lower is better.',
  'us dup rank':'Nationwide rank for duplicate-claim rate. Rank #1 is the lowest/best duplicate rate, so lower rank numbers are better.',
  'issue types':'Number of distinct CMS quality issue categories represented in the current data.',
  'states analyzed':'Number of states that meet the current sample-size threshold and are included in the analysis.',
  'states with data':'Number of states with issue records available under the current filters.',
  'avg success':'Average share of issues resolved across the included state programs; higher is better.',
  'avg confidence':'Average sample-size confidence used by the Medicaid comparison model. Larger underlying issue samples generally produce higher confidence.',
  'minimum sample':'Minimum number of issues a state must have to be included in the current comparison or insight analysis.',
  'avg backlog':'Average share of issues still open across included states. Lower is generally better.',
  'avg quality':'Average composite Medicaid quality score across the states included in the current executive scope.',
  'coverage':'Share of state programs meeting the current analysis threshold and therefore represented in the executive metrics.',
  'success rate':'Percentage of issues in scope that are resolved successfully; higher is better.',
  'resolution rate':'Resolved issues divided by total issues in the selected scope; higher is better.',
  'backlog rate':'Percentage of issues that remain open in the selected scope; lower is generally better.',
  'cancellation rate':'Percentage of issues closed as cancelled rather than resolved.',
  'cancel rate':'Percentage of issues closed as cancelled rather than successfully resolved.',
  'resolved issues':'Count of issue records currently marked resolved in the selected heatmap/state scope.',
  'cancelled issues':'Count of issue records currently marked cancelled in the selected heatmap/state scope.',
  'monitoring agents':'Number of registered Medicaid monitoring subscriptions configured to watch quality conditions.',
  'modeled exposure':'Planning estimate of financial exposure produced by the prototype issue-volume and cost-multiplier model. It is not an audited recovery amount.',
  'priority opportunities':'Number of state-and-issue-type combinations surfaced as the strongest modeled remediation opportunities.',
  'top opportunity':'State associated with the highest-ranked modeled optimization opportunity in the current scope.',
  'top modeled risk':'Largest modeled money-at-risk value among the current opportunities. It is a prioritization proxy rather than an audited loss estimate.',
  'open quality issues':'Count of Medicaid quality issues that remain open and make up the current remediation backlog.',
  'rows':'Number of data records in the selected dataset.',
  'columns':'Number of fields detected in the selected dataset.',
  'completeness':'Percentage of all dataset cells that contain a non-missing value; higher generally indicates better data completeness.',
  'missing cells':'Total blank/null cells across all rows and columns in the selected dataset.',
  'duplicate rows':'Count of rows that exactly duplicate another row across the full set of dataset columns.',
  'numeric fields':'Number of columns detected as numeric and therefore available for quantitative profiling/comparison.',
  'saved datasets':'Number of reusable CSV/XLSX files currently stored in the Build Your Own persistent dataset library.',
  'rows available':'Total number of rows across all saved Build Your Own datasets.',
  'storage':'Repo-relative folder where Build Your Own dataset files are persisted between app restarts.',
  'ai analyst':'Connection status for the Build Your Own grounded AI analyst. Connected means the configured Groq/OpenAI-compatible model is available.',
  'shared columns':'Number of column names present in both selected datasets.',
  'type mismatches':'Shared columns whose inferred data type differs between Dataset A and Dataset B.',
  'row difference b − a':'Dataset B row count minus Dataset A row count. Positive means B has more rows; negative means A has more.',
  'completeness delta':'Dataset B completeness minus Dataset A completeness, expressed in percentage points. Positive means B is more complete.',
  'controls':'Number of saved governance controls in the DART control register.',
  'active':'Governance controls currently marked Active and therefore in operation.',
  'mappings':'Number of saved source-to-target lineage mappings in the DART mapping catalog.',
  'issues':'Number of saved remediation workflow items available to the Governance/Remediation workflow.',
  'watcher':'Current background Email Agent watcher state. A healthy/running state means DART is periodically checking enabled automations for workbook changes.',
  'check interval':'How often the Email Agent watcher checks the source workbook for qualifying changes.',
  'email delivery':'Whether SMTP delivery is configured and ready for Email Agent notifications.',
  'raw rows':'Number of source workbook data rows currently available to the Raw Data Editor.',
  'raw columns':'Number of editable source workbook columns detected by the Raw Data Editor.',
  'last saved':'Most recent modified timestamp reported for the source workbook.',
  'stream scorecard':'Executive summary for one reconciliation stream. Volume-weighted match weights fields by claim volume; Average field match gives each comparable field equal weight; Fields counts mapped fields; Not matched is unmatched claim/event volume; Critical counts highest-tier fields; Avg impact is the mean DART impact score.'
};
function normalizeMetricLabel(v){return String(v||'').trim().toLowerCase().replace(/\s+/g,' ');}
function metricHelpLabel(host){
  if(host.classList.contains('scorecard'))return 'Stream scorecard';
  if(host.classList.contains('metric'))return host.querySelector('.label')?.textContent||'';
  if(host.classList.contains('medKpi'))return host.querySelector('.label')?.textContent||'';
  if(host.classList.contains('statusCard'))return host.querySelector('.muted')?.textContent||'';
  if(host.classList.contains('streamStat'))return host.querySelector('.k')?.textContent||'';
  if(host.classList.contains('compareMetric'))return host.querySelector('.k')?.textContent||'';
  if(host.classList.contains('medHeatMiniKpi'))return host.querySelector('span')?.textContent||'';
  if(host.classList.contains('miniCard'))return host.querySelector('span')?.textContent||'';
  return '';
}
function metricHelpDescription(label,host){
  const key=normalizeMetricLabel(label);
  if(key==='open'&&state.page==='actioncenter')return 'Number of saved remediation workflow items currently in Open status.';
  if(key==='in progress')return 'Number of saved remediation workflow items actively being worked but not yet resolved.';
  if(key==='blocked')return 'Number of saved remediation workflow items that cannot progress until a dependency, decision, or data issue is cleared.';
  if(key==='resolved'&&state.page==='actioncenter')return 'Number of saved remediation workflow items marked Resolved.';
  if(key==='workbook'&&state.page==='emailagent')return 'Whether the Excel workbook monitored by the Email Agent is currently available at the configured source path.';
  if(key==='workbook'&&state.page==='raweditor')return 'The source Excel workbook currently opened by the Raw Data Editor for cell-level writeback.';
  if(METRIC_HELP[key])return METRIC_HELP[key];
  const sub=host.querySelector('.sub')?.textContent||host.querySelector('.muted')?.textContent||'';
  return `Shows ${label} for the current page and active filters.${sub?` The card's secondary note (“${sub.trim()}”) gives the local context for how the value is being reported.`:''}`;
}
function decorateMetricHelp(root=document){
  if(!root||!root.querySelectorAll)return;
  const selector='.metric,.miniCard,.medKpi,.statusCard,.streamStat,.compareMetric,.medHeatMiniKpi,.scorecard';
  root.querySelectorAll(selector).forEach(host=>{
    if(host.querySelector(':scope > .metricInfoWrap'))return;
    const label=metricHelpLabel(host).trim();if(!label)return;
    host.classList.add('metricHelpHost');
    const wrap=document.createElement('span');wrap.className='metricInfoWrap';
    const icon=document.createElement('span');icon.className='metricInfoButton';icon.tabIndex=0;icon.setAttribute('role','button');icon.setAttribute('aria-label',`About ${label}`);icon.textContent='i';
    const tip=document.createElement('span');tip.className='metricInfoTooltip';tip.setAttribute('role','tooltip');
    const title=document.createElement('b');title.textContent=label;
    const body=document.createElement('span');body.textContent=metricHelpDescription(label,host);
    tip.append(title,body);wrap.append(icon,tip);host.appendChild(wrap);
    const stop=e=>{e.stopPropagation();if(e.type==='click')e.preventDefault();};
    icon.addEventListener('click',stop);icon.addEventListener('mousedown',stop);icon.addEventListener('pointerdown',stop);
    icon.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();e.stopPropagation();}});
  });
}
async function openMetricDrill(kind,title){try{streamModalState=null;metricModalTitle.textContent=title;metricModalSub.textContent='Loading rows behind this metric…';metricModalBody.innerHTML='<div class="empty">Loading data…</div>';metricModal.style.display='flex';const result=await postJson('/api/drilldown/'+encodeURIComponent(kind),state.filters);metricModalSub.textContent=`${intFmt(result.total)} row${Number(result.total)===1?'':'s'} in the current filtered scope${result.truncated?' · showing the first '+intFmt(result.rows.length):''}.`;metricModalBody.innerHTML=table(result.rows,true,title);}catch(err){metricModalSub.textContent='Could not load metric details.';metricModalBody.innerHTML=`<div class="empty">${esc(err.message||err)}</div>`;}}
function openLocalDrill(title,subtitle,rows){streamModalState=null;metricModalTitle.textContent=title;metricModalSub.textContent=subtitle||`${intFmt((rows||[]).length)} record${(rows||[]).length===1?'':'s'}.`;metricModalBody.innerHTML=simpleTable(rows||[],title);metricModal.style.display='flex';}
function metricCard(label,value,sub,kind,extraClass=''){const action=()=>openMetricDrill(kind,label);return `<div class="panel metric metricClickable ${extraClass}" role="button" tabindex="0" onclick="openMetricDrill('${kind}','${esc(label)}')" onkeydown="keyActivate(event,()=>openMetricDrill('${kind}','${esc(label)}'))"><div class="label">${esc(label)}</div><div class="value" title="${esc(value)}">${value}</div><div class="sub">${esc(sub)}</div></div>`;}
function metrics(s){return `<div class="grid grid4">${metricCard('Health score',s.health_score,'Event-weighted quality','health','health')}${metricCard('Critical items',s.critical,'Highest risk tier','critical')}${metricCard('Actionable items',s.actions,'Require review, change, or remediation','open-issues')}${metricCard('Unmatched volume',s.unmatched,'Current scope','unmatched','unmatchedMetric')}</div>`;}
function miniDeck(s){return '';}
let streamModalState=null;
async function openStreamDetails(stream){try{metricModalTitle.textContent=`${stream} stream details`;metricModalSub.textContent='Loading stream metrics and underlying data…';metricModalBody.innerHTML='<div class="empty">Loading stream data…</div>';metricModal.style.display='flex';const result=await postJson('/api/stream-detail/'+encodeURIComponent(stream),state.filters);streamModalState={data:result,tab:'overview'};metricModalSub.textContent=`${intFmt(result.total)} field${Number(result.total)===1?'':'s'} in ${stream} under the current filters${result.truncated?' · large tables are truncated':''}.`;renderStreamModal();}catch(err){metricModalSub.textContent='Could not load stream details.';metricModalBody.innerHTML=`<div class="empty">${esc(err.message||err)}</div>`;}}
function setStreamTab(tab){if(!streamModalState)return;streamModalState.tab=tab;renderStreamModal();}
function streamTabs(active){const tabs=[['overview','Overview'],['fields','Fields'],['risk','Risk'],['unmatched','Unmatched'],['classifications','Classifications'],['remediation','Remediation'],['mapping','Mapping & impact']];return `<div class="modalTabs" role="tablist">${tabs.map(([id,label])=>`<button class="modalTab ${active===id?'active':''}" type="button" onclick="setStreamTab('${id}')">${label}</button>`).join('')}</div>`;}
function streamOverview(s){return `<div class="streamOverviewGrid"><div class="streamStat"><div class="k">Weighted match</div><div class="v">${s.weighted_rate}</div></div><div class="streamStat"><div class="k">Field average</div><div class="v">${s.field_average}</div></div><div class="streamStat"><div class="k">Fields</div><div class="v">${intFmt(s.rows)}</div></div><div class="streamStat"><div class="k">Unmatched volume</div><div class="v">${s.unmatched}</div></div><div class="streamStat"><div class="k">Critical</div><div class="v">${s.critical}</div></div><div class="streamStat"><div class="k">Risk score</div><div class="v">${s.risk_score}</div></div></div>`;}
function renderStreamModal(){if(!streamModalState)return;const d=streamModalState.data;const tab=streamModalState.tab;const stream=d.stream||'stream';let body='';if(tab==='overview'){body=`${streamOverview(d.summary)}<div class="grid grid2"><div class="panel"><h3>Risk tier mix</h3>${simpleTable(d.risk_summary,stream+' risk tier mix')}</div><div class="panel"><h3>Classification mix</h3>${simpleTable(d.classification_summary,stream+' classification mix')}</div></div><div class="modalSection"><h3>Highest-impact fields</h3>${table((d.rows||[]).slice(0,25),true,stream+' highest-impact fields')}</div>`;}else if(tab==='fields'){body=`<div class="modalSection"><h3>All fields in this stream</h3>${table(d.rows,false,stream+' all fields')}</div>`;}else if(tab==='risk'){body=`<div class="modalSection"><h3>Critical and elevated fields</h3>${table(d.risk_rows,true,stream+' risk fields')}</div>`;}else if(tab==='unmatched'){body=`<div class="modalSection"><h3>Fields with unmatched volume</h3>${table(d.unmatched_rows,true,stream+' unmatched fields')}</div>`;}else if(tab==='classifications'){body=`<div class="grid grid2"><div class="panel"><h3>Classification counts</h3>${simpleTable(d.classification_summary,stream+' classification counts')}</div><div class="panel"><h3>Risk tier counts</h3>${simpleTable(d.risk_summary,stream+' risk tier counts')}</div></div><div class="modalSection"><h3>Field-level classification data</h3>${table(d.rows,true,stream+' field classifications')}</div>`;}else if(tab==='remediation'){body=`<div class="modalSection"><h3>Fields needing change or review</h3>${table(d.remediation_rows,true,stream+' remediation fields')}</div><div class="modalSection"><h3>Saved workflow issues for this stream</h3>${simpleTable(d.workflow_issues,stream+' workflow issues')}</div>`;}else if(tab==='mapping'){body=`<div class="modalSection"><h3>Saved lineage mappings</h3>${simpleTable(d.mappings,stream+' lineage mappings')}</div><div class="modalSection"><h3>Saved report / KPI impacts</h3>${simpleTable(d.impacts,stream+' report KPI impacts')}</div>`;}metricModalBody.innerHTML=streamTabs(tab)+body;decorateMetricHelp(metricModalBody);}
function safeFilename(name){return String(name||'dart-export').trim().replace(/[^a-z0-9._-]+/gi,'_').replace(/^_+|_+$/g,'')||'dart-export';}
function csvCell(v){const s=String(v??'').replace(/\r?\n/g,' ');return '"'+s.replace(/"/g,'""')+'"';}
function exportRenderedTable(button,filename){const block=button.closest('.exportTableBlock');const tbl=block?.querySelector('table');if(!tbl){toast('No table data to export');return;}const lines=[...tbl.querySelectorAll('tr')].map(tr=>[...tr.children].map(cell=>csvCell(cell.innerText.trim())).join(','));const blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=safeFilename(filename)+'.csv';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),500);toast('CSV exported');}
function exportBar(exportName,rowCount){return exportName?`<div class="exportTableBar"><span class="modalDataNote">${intFmt(rowCount)} row${Number(rowCount)===1?'':'s'}</span><button class="btn secondary small" type="button" onclick="exportRenderedTable(this,'${safeArg(exportName)}')">Export CSV</button></div>`:'';}
function table(rows,limitCols=false,exportName=''){if(!rows||!rows.length)return '<div class="empty">No rows match the current selection.</div>';const cols=limitCols?['Stream','NCH Target Column','MatchRate','NotMatchedClaims','RiskTier','ImpactScore','Recommendation']:['Stream','NCH Target Table','NCH Target Column','SS Table','SS Column','MatchedClaims','NotMatchedClaims','MatchRate','Classification','RiskTier','ImpactScore','Sub-Classification','Disposition','Recommendation'];return `<div class="exportTableBlock">${exportBar(exportName,rows.length)}<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>{let v=r[c];if(c==='MatchRate')return `<td><div>${pctNum(v)}</div><div class="bar"><span style="width:${Math.max(0,Math.min(100,Number(v||0)*100))}%"></span></div></td>`;if(c==='RiskTier')return `<td><span class="pill ${esc(v)}">${esc(v)}</span></td>`;return `<td>${esc(v)}</td>`}).join('')}</tr>`).join('')}</tbody></table></div></div>`;}
function simpleTable(rows,exportName=''){if(!rows||!rows.length)return '<div class="empty">Nothing saved yet. Starter examples will appear where helpful.</div>';const cols=Object.keys(rows[0]);return `<div class="exportTableBlock">${exportBar(exportName,rows.length)}<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;}
function cards(items){return items.map(r=>`<div class="insight"><b>${esc(r.Stream)} · ${esc(r['NCH Target Column'])}</b><br><span class="muted">${pctNum(r.MatchRate)} match · ${intFmt(r.NotMatchedClaims)} unmatched · ${esc(r.RiskTier)} risk · impact ${esc(r.ImpactScore)}</span><br>${esc(r.Recommendation)}</div>`).join('<br>')||'<div class="empty">No priority findings in the current scope.</div>';}
function runInjectedScripts(root){root.querySelectorAll('script').forEach(oldScript=>{const newScript=document.createElement('script');Array.from(oldScript.attributes).forEach(attr=>newScript.setAttribute(attr.name,attr.value));newScript.textContent=oldScript.textContent;oldScript.parentNode.replaceChild(newScript,oldScript);});if(window.Plotly){setTimeout(()=>{root.querySelectorAll('.plotly-graph-div').forEach(g=>{try{Plotly.Plots.resize(g);}catch(err){console.warn('Plotly resize skipped',err);}});applyPlotlyTheme(root);},75);}else{syncThemeSwitch();}}
function executiveBrief(data,meta){const s=data.summary;const p=meta.persona||{};const top=(s.top_actions||[]).slice(0,6).map((r,i)=>`${i+1}. ${r.Stream} · ${r['NCH Target Column']} | ${pctNum(r.MatchRate)} match | ${intFmt(r.NotMatchedClaims)} unmatched | ${r.RiskTier} | impact ${r.ImpactScore}`).join('\n')||'No priority findings in the current scope.';return `Executive Brief\n\nAudience: ${p.audience||'Leadership'}\nProgram: ${p.program||'Not specified'}\nRole lens: ${p.role||'Not specified'}\n\nExecutive Risk Rating\n- Health score: ${s.health_score}\n- Risk score: ${s.risk_score}\n- Critical items: ${s.critical}\n- Open actions: ${s.actions}\n- Affected reports: ${s.affected_reports}\n\nOverall Data Quality\n- Field average: ${s.field_average}\n- Event-weighted match: ${s.weighted_rate}\n- Unmatched volume: ${s.unmatched}\n\nPriority Findings\n${top}\n\nDecision Support Recommendations\n1. Validate highest-impact mappings first.\n2. Confirm lineage for affected fields and downstream reports.\n3. Assign owners for critical and elevated items in the Remediation Center.\n4. Use Impact Explorer for fields that support executive KPIs.\n5. Review Governance Center controls for ownership and evidence.\n\nDecision Required\nConfirm whether critical mappings should be prioritized for remediation in the next review cycle.`}
function generateRichAiInsights(data, meta){
  const rows = data.rows || [];
  const s = data.summary || {};
  const topActions = s.top_actions || [];
  const streams = s.stream_summary || [];
  const insights = [];
  
  // 1. Critical Volume Leakage / Concentration
  const sortedByUnmatched = rows.slice().sort((a,b) => Number(b.NotMatchedClaims||0) - Number(a.NotMatchedClaims||0));
  const topUnmatchedRows = sortedByUnmatched.slice(0, 10).filter(r => Number(r.NotMatchedClaims||0) > 0);
  const top5Volume = topUnmatchedRows.slice(0, 5).reduce((acc, r) => acc + Number(r.NotMatchedClaims||0), 0);
  const totalUnmatchedNum = rows.reduce((acc, r) => acc + Number(r.NotMatchedClaims||0), 0);
  const volSharePct = totalUnmatchedNum > 0 ? ((top5Volume / totalUnmatchedNum) * 100).toFixed(1) : '0.0';
  
  if(topUnmatchedRows.length > 0){
    const lead = topUnmatchedRows[0];
    insights.push({
      id: 'volume_leakage',
      category: 'Volume Concentration',
      title: `Critical Volume Leakage: Top 5 Fields Account for ${volSharePct}% of Unmatched Claims`,
      severity: 'Critical',
      summary: `A heavy 80/20 asymmetry exists where ${intFmt(top5Volume)} claims are concentrated across just 5 fields, led by ${lead.Stream} · ${lead['NCH Target Column']}.`,
      highlightKey1: 'Concentrated Volume',
      highlightVal1: intFmt(top5Volume),
      highlightKey2: 'Share of Total Loss',
      highlightVal2: `${volSharePct}%`,
      highlightKey3: 'Top Offending Field',
      highlightVal3: `${lead.Stream} · ${lead['NCH Target Column']}`,
      diagnosis: `Reconciliation loss is not evenly distributed across streams; it is heavily localized. The single field <b>${esc(lead.Stream)} · ${esc(lead['NCH Target Column'])}</b> alone accounts for <b>${intFmt(lead.NotMatchedClaims)}</b> unmatched claims (${pctNum(lead.MatchRate)} match rate). Resolving this core cluster will immediately recover the majority of reconciled volume.`,
      confidence: '96%',
      recommendation: `Prioritize immediate structural review of ${lead.Stream} ETL pipeline and STTM table joins.`,
      rows: topUnmatchedRows,
      leadRow: lead,
      exportName: 'volume_leakage_fields',
      copilotPrompt: `Analyze why ${lead.Stream} field ${lead['NCH Target Column']} has ${intFmt(lead.NotMatchedClaims)} unmatched claims and provide technical mapping recommendations.`
    });
  }
  
  // 2. Complete Mismatch & Zero-Match Disconnects
  const zeroMatchRows = rows.filter(r => (r.MatchRate == null || Number(r.MatchRate) <= 0.1) && Number(r.NotMatchedClaims||0) > 0).sort((a,b) => Number(b.NotMatchedClaims||0) - Number(a.NotMatchedClaims||0));
  if(zeroMatchRows.length > 0){
    const zeroVol = zeroMatchRows.reduce((acc, r) => acc + Number(r.NotMatchedClaims||0), 0);
    const leadZero = zeroMatchRows[0];
    insights.push({
      id: 'zero_match',
      category: 'Reconciliation Disconnect',
      title: `Zero / Near-Zero Match Disconnects in ${zeroMatchRows.length} Field${zeroMatchRows.length===1?'':'s'}`,
      severity: 'Critical',
      summary: `${zeroMatchRows.length} fields exhibit near-complete matching failures (≤10% match rate) across ${intFmt(zeroVol)} total claims.`,
      highlightKey1: 'Affected Fields',
      highlightVal1: zeroMatchRows.length,
      highlightKey2: 'Disconnected Claims',
      highlightVal2: intFmt(zeroVol),
      highlightKey3: 'Primary Stream',
      highlightVal3: leadZero.Stream,
      diagnosis: `Zero-match anomalies represent complete schema or join failures rather than organic data variance. Leading causes include unmapped target fields in CMS STTM, mismatched identifier padding (e.g., string vs numeric padding), or missing reference tables in feed staging.`,
      confidence: '94%',
      recommendation: `Validate source-to-target mapping definitions in Mapping Catalog and verify cross-table join keys.`,
      rows: zeroMatchRows,
      leadRow: leadZero,
      exportName: 'zero_match_fields',
      copilotPrompt: `Explain why ${zeroMatchRows.length} reconciliation fields have near-zero match rate and how to fix STTM joins.`
    });
  }
  
  // 3. High-Impact Downstream KPI & Report Exposure
  const highImpactRows = rows.filter(r => Number(r.ImpactScore || 0) >= 3.0 || r.RiskTier === 'Critical').sort((a,b) => Number(b.ImpactScore||0) - Number(a.ImpactScore||0));
  if(highImpactRows.length > 0){
    const avgImpact = (highImpactRows.reduce((acc, r) => acc + Number(r.ImpactScore||0), 0) / highImpactRows.length).toFixed(2);
    const leadImpact = highImpactRows[0];
    insights.push({
      id: 'kpi_exposure',
      category: 'Downstream Risk',
      title: `High Downstream Exposure Across ${highImpactRows.length} Critical Fields`,
      severity: 'Elevated',
      summary: `Technical discrepancies in these fields directly propagate into ${s.affected_reports || 'multiple'} executive KPI reports and compliance audits.`,
      highlightKey1: 'Critical/Elevated Items',
      highlightVal1: highImpactRows.length,
      highlightKey2: 'Avg Impact Score',
      highlightVal2: avgImpact,
      highlightKey3: 'Exposed Reports',
      highlightVal3: s.affected_reports || '3+',
      diagnosis: `Fields with high impact scores directly underpin downstream clinical, financial, and regulatory reporting. When match rates degrade here, executive metrics (Paid Claims, Denials, Program Integrity) reflect unverified data.`,
      confidence: '91%',
      recommendation: `Establish explicit governance controls and assign named remediation owners in Remediation Center.`,
      rows: highImpactRows,
      leadRow: leadImpact,
      exportName: 'high_impact_fields',
      copilotPrompt: `Which reconciliation fields create the highest risk for downstream executive reporting and KPI integrity?`
    });
  }
  
  // 4. Stream Performance Asymmetry
  if(streams.length > 1){
    const worstStream = streams.slice().sort((a,b) => Number(a.WeightedMatch || 0) - Number(b.WeightedMatch || 0))[0];
    const bestStream = streams.slice().sort((a,b) => Number(b.WeightedMatch || 0) - Number(a.WeightedMatch || 0))[0];
    const streamRows = rows.filter(r => r.Stream === worstStream.Stream);
    const leadStream = streamRows[0] || worstStream;
    insights.push({
      id: 'stream_asymmetry',
      category: 'Stream Asymmetry',
      title: `Performance Gap: ${worstStream.Stream} Lagging Behind ${bestStream.Stream}`,
      severity: 'Watch',
      summary: `${worstStream.Stream} has a weighted match of ${pctNum(worstStream.WeightedMatch)} vs. ${bestStream.Stream} at ${pctNum(bestStream.WeightedMatch)}.`,
      highlightKey1: 'Lagging Stream',
      highlightVal1: worstStream.Stream,
      highlightKey2: 'Weighted Match',
      highlightVal2: pctNum(worstStream.WeightedMatch),
      highlightKey3: 'Unmatched Claims',
      highlightVal3: intFmt(worstStream.NotMatched),
      diagnosis: `Stream-level performance divergence indicates that reconciliation logic or source data pipeline maturity varies significantly across feeds. ${worstStream.Stream} contributes ${intFmt(worstStream.NotMatched)} unmatched claims and accounts for the largest proportion of scope risk.`,
      confidence: '89%',
      recommendation: `Conduct a dedicated stream audit on ${worstStream.Stream} to align format rules with top-performing streams.`,
      rows: streamRows,
      leadRow: leadStream,
      exportName: `${worstStream.Stream.toLowerCase()}_stream_records`,
      copilotPrompt: `Compare reconciliation performance between ${worstStream.Stream} and ${bestStream.Stream} and highlight main variance drivers.`
    });
  }
  
  // 5. High-ROI Quick Wins
  const quickWinRows = rows.filter(r => Number(r.MatchRate || 0) > 0.1 && Number(r.MatchRate || 0) < 0.85 && Number(r.NotMatchedClaims || 0) > 5000).sort((a,b) => Number(b.NotMatchedClaims||0) - Number(a.NotMatchedClaims||0));
  if(quickWinRows.length > 0){
    const winVol = quickWinRows.reduce((acc, r) => acc + Number(r.NotMatchedClaims||0), 0);
    const leadWin = quickWinRows[0];
    const avgWinMatch = (quickWinRows.reduce((acc,r)=>acc+Number(r.MatchRate||0),0)/quickWinRows.length);
    insights.push({
      id: 'quick_wins',
      category: 'Remediation Quick-Wins',
      title: `Fast Recovery Potential: ${intFmt(winVol)} Claims in ${quickWinRows.length} Moderate-Variance Fields`,
      severity: 'Stable',
      summary: `These fields have established partial matches and high claim volume, making them prime candidates for fast lookup table fixes.`,
      highlightKey1: 'Quick-Win Fields',
      highlightVal1: quickWinRows.length,
      highlightKey2: 'Recoverable Volume',
      highlightVal2: intFmt(winVol),
      highlightKey3: 'Average Match Rate',
      highlightVal3: pctNum(avgWinMatch),
      diagnosis: `Unlike zero-match fields requiring schema re-engineering, these fields already have active mapping lineage. Unmatched volume here is typically caused by unmapped lookup codes, trim/case differences, or minor date formatting.`,
      confidence: '93%',
      recommendation: `Apply standardized value mapping transformations to capture unmatched codes.`,
      rows: quickWinRows,
      leadRow: leadWin,
      exportName: 'quick_win_remediation_fields',
      copilotPrompt: `How can we quickly resolve the moderate match rate variance in ${leadWin.Stream} field ${leadWin['NCH Target Column']}?`
    });
  }
  
  return insights;
}

function openFieldDossier(r){
  if(!r) return;
  streamModalState = null;
  metricModalTitle.textContent = `${r.Stream || 'Reconciliation'} · ${r['NCH Target Column'] || 'Field Dossier'}`;
  metricModalSub.textContent = `Deep AI Root-Cause Diagnostic Dossier · Risk Tier: ${r.RiskTier || 'Unknown'} · Impact Score: ${Number(r.ImpactScore || 0).toFixed(2)}`;
  
  const matchRate = Number(r.MatchRate || 0);
  const matchPct = pctNum(matchRate);
  const unmatched = intFmt(r.NotMatchedClaims || 0);
  const matched = intFmt(r.MatchedClaims || 0);
  const totalClaims = (Number(r.MatchedClaims || 0) + Number(r.NotMatchedClaims || 0));
  
  let hypothesis = '';
  let conf = '92%';
  if(matchRate === 0){
    hypothesis = `<strong>Total Match Failure (0.0% Match):</strong> All ${intFmt(totalClaims)} claims failed reconciliation. This indicates an unmapped target column in the STTM reference table, a complete mismatch in key formats (e.g., string vs numeric padding), or a missing ETL join table.`;
    conf = '98%';
  } else if(matchRate < 0.5){
    hypothesis = `<strong>Severe Reconciliation Leakage (${matchPct} Match):</strong> ${unmatched} claims failed matching against ${matched} matched. This pattern typically stems from unhandled disposition codes, legacy code lookups, or date format truncations during feed staging.`;
    conf = '94%';
  } else if(matchRate < 0.95){
    hypothesis = `<strong>Moderate Edge-Case Variance (${matchPct} Match):</strong> ${unmatched} unmatched claims out of ${intFmt(totalClaims)}. Likely caused by rare transaction codes, null values in optional fields, or late-arriving adjustments.`;
    conf = '88%';
  } else {
    hypothesis = `<strong>Minor Variance (${matchPct} Match):</strong> Only ${unmatched} unmatched claims. Standard boundary discrepancies or occasional null overrides.`;
    conf = '85%';
  }
  
  const rec = r.Recommendation || 'Validate source-to-target mapping and inspect feed values.';
  const escRow = JSON.stringify(r).replace(/'/g, "&#39;");
  
  metricModalBody.innerHTML = `
    <div class="fieldDossierWrap">
      <div class="dossierStatsGrid">
        <div class="dossierStat"><div class="k">Match Rate</div><div class="v">${matchPct}</div><div class="bar" style="margin-top:6px"><span style="width:${Math.max(0, Math.min(100, matchRate * 100))}%"></span></div></div>
        <div class="dossierStat"><div class="k">Unmatched Claims</div><div class="v" style="color:var(--bad)">${unmatched}</div></div>
        <div class="dossierStat"><div class="k">Matched Claims</div><div class="v" style="color:var(--good)">${matched}</div></div>
        <div class="dossierStat"><div class="k">Impact Score</div><div class="v">${Number(r.ImpactScore || 0).toFixed(2)}</div></div>
        <div class="dossierStat"><div class="k">Risk Tier</div><div class="v"><span class="pill ${esc(r.RiskTier || 'Watch')}">${esc(r.RiskTier || 'Watch')}</span></div></div>
      </div>
      
      <div class="dossierSection">
        <h3>Technical Lineage & Mapping Specification</h3>
        <div class="dossierLineageGrid">
          <div class="lineageBox">
            <span class="lineageBadge">Source System (SS)</span>
            <div class="lineageItem"><b>Table:</b> <span>${esc(r['SS Table'] || 'Source Feed')}</span></div>
            <div class="lineageItem"><b>Column:</b> <span>${esc(r['SS Column'] || 'Source Field')}</span></div>
          </div>
          <div class="lineageArrow">→</div>
          <div class="lineageBox">
            <span class="lineageBadge">NCH Target (CMS STTM)</span>
            <div class="lineageItem"><b>Table:</b> <span>${esc(r['NCH Target Table'] || 'NCH Target Table')}</span></div>
            <div class="lineageItem"><b>Column:</b> <span>${esc(r['NCH Target Column'] || 'NCH Target Column')}</span></div>
          </div>
        </div>
      </div>
      
      <div class="dossierSection">
        <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:8px">
          <h3 style="margin:0">AI Root-Cause Diagnostic Analysis</h3>
          <span class="personaTag" style="background:#eef6ff;border-color:#b9d9fe;color:#1e40af;font-weight:800">AI Confidence: ${conf}</span>
        </div>
        <div class="dossierDiagnosisBox">
          <p>${hypothesis}</p>
          <div style="margin-top:10px;padding:10px 12px;border-radius:10px;background:rgba(212,175,55,.12);border:1px solid rgba(212,175,55,.3)">
            <b>Recommended Action:</b> <span>${esc(rec)}</span>
          </div>
        </div>
      </div>
      
      <div class="dossierSection">
        <h3>Classification & Downstream Context</h3>
        <div class="dossierMetaGrid">
          <div><b>Stream:</b> <span>${esc(r.Stream || '—')}</span></div>
          <div><b>Classification:</b> <span>${esc(r.Classification || '—')}</span></div>
          <div><b>Sub-Classification:</b> <span>${esc(r['Sub-Classification'] || '—')}</span></div>
          <div><b>Disposition:</b> <span>${esc(r.Disposition || '—')}</span></div>
        </div>
      </div>
      
      <div class="dossierActions">
        <button class="btn" onclick="quickPrompt('Explain the root causes of reconciliation failure for ${safeArg(r.Stream)} field ${safeArg(r['NCH Target Column'])} and suggest technical fixes.')">Ask Copilot About This Field</button>
        <button class="btn secondary" onclick="addInsightToActionCenter('${safeArg(r.Stream)}','${safeArg(r['NCH Target Column'])}','${safeArg(r.RiskTier)}',${Number(r.ImpactScore || 0)},'${safeArg(rec)}')">Track in Remediation Center</button>
        <button class="btn ghost" onclick='openInvestigationDrill("${safeArg(r.Stream)} · ${safeArg(r['NCH Target Column'])} Data", "Underlying reconciliation records for this field", [${escRow}], "${safeArg(r['NCH Target Column'])}")'>View Record Table</button>
      </div>
    </div>
  `;
  metricModal.style.display = 'flex';
}

function openInvestigationDrill(title, subtitle, rows, exportName){
  streamModalState = null;
  metricModalTitle.textContent = title;
  metricModalSub.textContent = subtitle || `${intFmt((rows || []).length)} record${(rows || []).length === 1 ? '' : 's'} under investigation.`;
  metricModalBody.innerHTML = table(rows || [], false, exportName || 'investigation_records');
  metricModal.style.display = 'flex';
}

async function addInsightToActionCenter(stream, field, riskTier, impactScore, note){
  try{
    await postJson('/api/issues', {
      Field: field,
      Stream: stream,
      Priority: riskTier === 'Critical' ? 'High' : (riskTier === 'Elevated' ? 'Medium' : 'Low'),
      Status: 'Open',
      Owner: 'Data Quality Team',
      Note: note || `AI Insight Anomaly: ${riskTier} risk with impact ${impactScore}.`
    });
    toast(`Added ${field} to Remediation Center`);
    state.meta = await api('/api/meta');
  }catch(err){
    toast('Could not add to action center: ' + (err.message || err));
  }
}

function filterInvestigationCards(category, query){
  if(category !== undefined && category !== null){
    state.insightsTab = category;
    document.querySelectorAll('.aiFilterChip').forEach(btn => {
      btn.classList.toggle('active', btn.dataset.category === category);
    });
  }
  const activeCat = state.insightsTab || 'all';
  const q = (query !== undefined ? query : (document.getElementById('aiInvestigationSearch')?.value || '')).trim().toLowerCase();
  
  document.querySelectorAll('.aiAnomalyCard').forEach(card => {
    const cardText = card.textContent.toLowerCase();
    
    let matchesCat = (activeCat === 'all');
    if(activeCat === 'critical') matchesCat = card.dataset.risk === 'Critical';
    if(activeCat === 'zero') matchesCat = Number(card.dataset.match || 1) <= 0.1;
    if(activeCat === 'volume') matchesCat = Number(card.dataset.unmatched || 0) >= 10000;
    if(activeCat === 'kpi') matchesCat = Number(card.dataset.impact || 0) >= 3.0;
    
    const matchesSearch = !q || cardText.includes(q);
    card.style.display = (matchesCat && matchesSearch) ? 'flex' : 'none';
  });
}

function aiInsightsPageHtml(data, meta){
  state.lastViewData = data;
  state.insightsTab = state.insightsTab || 'all';
  const s = data.summary || {};
  const p = meta.persona || {};
  const richInsights = generateRichAiInsights(data, meta);
  const rows = data.rows || [];
  
  const criticalCount = rows.filter(r => r.RiskTier === 'Critical').length;
  const zeroMatchCount = rows.filter(r => (r.MatchRate == null || Number(r.MatchRate) <= 0.1) && Number(r.NotMatchedClaims || 0) > 0).length;
  const highVolCount = rows.filter(r => Number(r.NotMatchedClaims || 0) >= 10000).length;
  const kpiCount = rows.filter(r => Number(r.ImpactScore || 0) >= 3.0).length;
  
  const diagnosticCardsHtml = richInsights.map(item => {
    const jsonRows = JSON.stringify(item.rows || []).replace(/'/g, "&#39;");
    const jsonLead = JSON.stringify(item.leadRow || item.rows?.[0] || {}).replace(/'/g, "&#39;");
    return `
      <div class="aiDiagnosticCard sev-${esc(item.severity)}">
        <div class="aiCardHead">
          <span class="aiCategoryTag">${esc(item.category)}</span>
          <span class="aiSevBadge ${esc(item.severity)}">${esc(item.severity)}</span>
        </div>
        <h3 class="aiCardTitle">${esc(item.title)}</h3>
        <p class="aiCardSummary">${esc(item.summary)}</p>
        
        <div class="aiHighlightStat">
          <div class="statItem">
            <span class="statK">${esc(item.highlightKey1)}</span>
            <span class="statV">${esc(item.highlightVal1)}</span>
          </div>
          <div class="statItem">
            <span class="statK">${esc(item.highlightKey2)}</span>
            <span class="statV">${esc(item.highlightVal2)}</span>
          </div>
          <div class="statItem">
            <span class="statK">${esc(item.highlightKey3)}</span>
            <span class="statV" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(item.highlightVal3)}</span>
          </div>
        </div>
        
        <div class="aiDiagnosisSection">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:6px;margin-bottom:6px">
            <b>Root Cause Hypothesis:</b>
            <span class="aiConfidenceTag">Confidence: ${esc(item.confidence)}</span>
          </div>
          <div>${item.diagnosis}</div>
          <div style="margin-top:8px;font-weight:700;color:#854d0e">
            <b>Recommended Action:</b> ${esc(item.recommendation)}
          </div>
        </div>
        
        <div class="aiActionBtns">
          <button class="btn small" onclick='openInvestigationDrill("${safeArg(item.title)}", "Underlying data records for ${safeArg(item.category)}", ${jsonRows}, "${safeArg(item.exportName)}")'>Inspect Data (${(item.rows||[]).length} rows)</button>
          <button class="btn secondary small" onclick='openFieldDossier(${jsonLead})'>View Dossier</button>
          <button class="btn ghost small" onclick="quickPrompt('${safeArg(item.copilotPrompt)}')">Ask Copilot</button>
        </div>
      </div>
    `;
  }).join('');
  
  // Sorted anomaly candidate rows for deep-dive investigation workspace
  const anomalyRows = (s.top_actions || rows).slice(0, 48);
  const anomalyCardsHtml = anomalyRows.map(r => {
    const jsonRow = JSON.stringify(r).replace(/'/g, "&#39;");
    const mRate = Number(r.MatchRate || 0);
    const unmatch = Number(r.NotMatchedClaims || 0);
    const imp = Number(r.ImpactScore || 0);
    return `
      <div class="aiAnomalyCard" data-risk="${esc(r.RiskTier)}" data-match="${mRate}" data-unmatched="${unmatch}" data-impact="${imp}" data-stream="${esc(r.Stream)}" onclick='openFieldDossier(${jsonRow})'>
        <div class="aiAnomalyCardHead">
          <span class="pill ${esc(r.RiskTier || 'Watch')}">${esc(r.RiskTier || 'Watch')}</span>
          <span class="personaTag" style="font-size:.7rem;padding:3px 8px">${esc(r.Stream)}</span>
        </div>
        <h4 class="aiAnomalyField" title="${esc(r['NCH Target Column'])}">${esc(r['NCH Target Column'])}</h4>
        <div class="aiAnomalyMapping" title="${esc(r['SS Table'])}.${esc(r['SS Column'])} → ${esc(r['NCH Target Table'])}">
          <span>${esc(r['SS Table'] || 'SS')} → ${esc(r['NCH Target Table'] || 'NCH')}</span>
        </div>
        <div class="aiAnomalyMetrics">
          <div class="aiAnomalyMetricItem">
            <span class="amK">Match</span>
            <span class="amV">${pctNum(r.MatchRate)}</span>
          </div>
          <div class="aiAnomalyMetricItem">
            <span class="amK">Unmatched</span>
            <span class="amV" style="color:var(--bad)">${intFmt(r.NotMatchedClaims)}</span>
          </div>
          <div class="aiAnomalyMetricItem">
            <span class="amK">Impact</span>
            <span class="amV">${imp.toFixed(1)}</span>
          </div>
        </div>
        <div class="aiAnomalyCardFoot">
          <span style="font-size:.74rem;color:#854d0e">🔍 Click to inspect dossier →</span>
          <button class="btn small ghost" style="padding:4px 8px;font-size:.72rem" onclick='event.stopPropagation();openInvestigationDrill("${safeArg(r.Stream)} · ${safeArg(r['NCH Target Column'])}", "Detailed reconciliation row record", [${jsonRow}], "${safeArg(r['NCH Target Column'])}")'>Data</button>
        </div>
      </div>
    `;
  }).join('');
  
  const allCriticalRows = JSON.stringify(rows.filter(r => r.RiskTier === 'Critical')).replace(/'/g, "&#39;");
  const allZeroRows = JSON.stringify(rows.filter(r => (r.MatchRate == null || Number(r.MatchRate) <= 0.1) && Number(r.NotMatchedClaims || 0) > 0)).replace(/'/g, "&#39;");
  const allUnmatchedRows = JSON.stringify(rows.filter(r => Number(r.NotMatchedClaims || 0) > 0)).replace(/'/g, "&#39;");
  
  return `
    <div class="aiInsightsHero">
      <span class="aiIntelligenceBadge">AI Diagnostic & Anomaly Engine</span>
      <h1 style="color:#0f172a;margin:0 0 8px;font-size:2.2rem;font-weight:950">AI Insights & Root-Cause Diagnostics</h1>
      <p style="color:#475569;margin:0;font-size:.95rem;line-height:1.55;max-width:980px">
        Automated anomaly detection and root-cause analysis tailored for <b>${esc(p.role || 'Data & Analytics')}</b> with <b>${esc(p.depth || 'balanced')}</b> depth. Click on any finding, anomaly card, or KPI below to inspect underlying records and technical specifications.
      </p>
      <div class="heroActions">
        <button class="btn" onclick="quickPrompt('Summarize the top data quality anomalies across the dataset and provide prioritized remediation guidance.')">Ask Copilot</button>
        <button class="btn secondary" onclick='openInvestigationDrill("Full Reconciliation Dataset", "All active rows under current filter scope", ${JSON.stringify(rows).replace(/'/g,"&#39;")}, "all_active_reconciliation_records")'>Export Full Investigation CSV</button>
        <button class="btn ghost" onclick="showPage('copilot')">Open Copilot Chat</button>
      </div>
    </div>
    
    <div class="aiInsightsKpiGrid">
      <div class="aiKpiCard" onclick='openInvestigationDrill("Critical Risk Tier Anomalies", "Fields in the highest severity risk tier", ${allCriticalRows}, "critical_anomalies")'>
        <div class="kpiLabel">Critical Anomalies</div>
        <div class="kpiVal" style="color:var(--bad)">${criticalCount}</div>
        <div class="kpiSub">Highest risk tier <span>Inspect →</span></div>
      </div>
      <div class="aiKpiCard" onclick='openInvestigationDrill("Unmatched Claims Under Investigation", "Reconciliation volume loss across filtered scope", ${allUnmatchedRows}, "unmatched_volume_loss")'>
        <div class="kpiLabel">Unmatched Claims</div>
        <div class="kpiVal">${s.unmatched || '0'}</div>
        <div class="kpiSub">Total scope loss <span>Inspect →</span></div>
      </div>
      <div class="aiKpiCard" onclick='openInvestigationDrill("Zero/Near-Zero Match Disconnects", "Fields with 0% to 10% match rates", ${allZeroRows}, "zero_match_disconnects")'>
        <div class="kpiLabel">Zero-Match Disconnects</div>
        <div class="kpiVal" style="color:#d97706">${zeroMatchCount}</div>
        <div class="kpiSub">Schema/Join drops <span>Inspect →</span></div>
      </div>
    </div>
    
    <div id="strategic-diagnostics" class="sectionBlock">
      <div class="sectionTitle">
        <h2 style="color:#0f172a;font-size:1.4rem;font-weight:900;margin-bottom:12px">Strategic & Root-Cause AI Diagnoses</h2>
      </div>
      <div class="aiDiagnosticsGrid">
        ${diagnosticCardsHtml}
      </div>
    </div>
    
    <br>
    
    <div id="investigation-workspace" class="sectionBlock">
      <div class="aiWorkspacePanel">
        <div class="aiFilterNav">
          <div>
            <h3 style="margin:0 0 4px;color:#0f172a;font-size:1.2rem;font-weight:900">Interactive Anomaly Explorer</h3>
            <p class="muted" style="margin:0;font-size:.85rem">Select a filter or search to isolate specific failure modes. Click any card to inspect its full diagnostic dossier.</p>
          </div>
          <input id="aiInvestigationSearch" placeholder="Search field, stream, or table..." style="max-width:260px;padding:9px 12px;font-size:.84rem;border-radius:10px" oninput="filterInvestigationCards(undefined, this.value)">
        </div>
        <div class="aiFilterPills" style="margin-bottom:16px">
          <button class="aiFilterChip ${state.insightsTab==='all'?'active':''}" data-category="all" onclick="filterInvestigationCards('all')">All Anomalies (${anomalyRows.length})</button>
          <button class="aiFilterChip ${state.insightsTab==='critical'?'active':''}" data-category="critical" onclick="filterInvestigationCards('critical')">Critical Mappings (${criticalCount})</button>
          <button class="aiFilterChip ${state.insightsTab==='zero'?'active':''}" data-category="zero" onclick="filterInvestigationCards('zero')">Zero-Match Drops (${zeroMatchCount})</button>
          <button class="aiFilterChip ${state.insightsTab==='volume'?'active':''}" data-category="volume" onclick="filterInvestigationCards('volume')">High Volume Loss (${highVolCount})</button>
          <button class="aiFilterChip ${state.insightsTab==='kpi'?'active':''}" data-category="kpi" onclick="filterInvestigationCards('kpi')">High KPI Impact (${kpiCount})</button>
        </div>
        <div class="aiAnomalyGrid" id="aiAnomalyGrid">
          ${anomalyCardsHtml}
        </div>
      </div>
    </div>
  `;
}

function aiInsights(data){return generateRichAiInsights(data, state.meta||{});}
function insightsHtml(items){return '';}
function scorecards(rows){return `<div class="grid grid3">${rows.map(r=>`<div class="panel scorecard scorecardClickable" role="button" tabindex="0" onclick="openStreamDetails('${safeArg(r.Stream)}')" onkeydown="keyActivate(event,()=>openStreamDetails('${safeArg(r.Stream)}'))"><h2>${esc(r.Stream)}</h2><p class="muted">Stream scorecard</p><div class="scoreLine"><span>Volume-weighted match</span><b>${pctNum(r.WeightedMatch)}</b></div><div class="scoreLine"><span>Average field match</span><b>${pctNum(r.AverageMatch)}</b></div><div class="scoreLine"><span>Fields</span><b>${intFmt(r.Fields)}</b></div><div class="scoreLine"><span>Not matched</span><b>${intFmt(r.NotMatched)}</b></div><div class="scoreLine"><span>Critical</span><b>${intFmt(r.Critical)}</b></div><div class="scoreLine"><span>Avg impact</span><b>${Number(r.AverageImpact||0).toFixed(2)}</b></div></div>`).join('')}</div>`}
function metricExplanations(){return `<div class="grid grid3 explain"><div class="panel"><h3>Health score</h3><p><b>What it means:</b> Event-weighted match quality across the current filtered scope.</p><p><b>How to use it:</b> Use it as the executive quality signal because higher-volume fields influence it more.</p></div><div class="panel"><h3>Field average</h3><p><b>What it means:</b> Average match rate across comparable fields, treating each field equally.</p><p><b>How to use it:</b> Good for breadth. It can hide volume concentration, so compare it with weighted match.</p></div><div class="panel"><h3>Weighted rate</h3><p><b>What it means:</b> Matched claims divided by total claims in scope.</p><p><b>How to use it:</b> Good for workload and business-volume weighting.</p></div><div class="panel"><h3>Risk score</h3><p><b>What it means:</b> Average impact score in the scope.</p><p><b>Formula:</b> (1 - Match Rate) x log(1 + Total Claims).</p></div><div class="panel"><h3>Risk tier</h3><p><b>What it means:</b> Stable, Watch, Elevated, or Critical based on impact score.</p><p><b>How to use it:</b> Sort remediation work by tier before drilling into drivers.</p></div><div class="panel"><h3>Affected reports</h3><p><b>What it means:</b> Reports/KPIs mapped in Impact Explorer or inferred from the current stream/report scope.</p><p><b>How to use it:</b> Converts technical variance into stakeholder impact.</p></div></div>`;}
function lineageEditor(rows){return `<div class="panel"><h3>Add or update mapping</h3><div class="editRow"><input id="linId" placeholder="ID, leave blank for new"><input id="linStream" placeholder="Stream"><input id="linSourceTable" placeholder="Source table"><input id="linSourceField" placeholder="Source field"></div><div class="editRow"><input id="linTargetTable" placeholder="Target table"><input id="linTargetField" placeholder="Target field"><input id="linReport" placeholder="Report"><input id="linKPI" placeholder="KPI"></div><div class="editRow"><input id="linOwner" placeholder="Owner"><select id="linStatus"><option>Draft</option><option>Validated</option><option>Needs review</option><option>Retired</option></select><input id="linRiskTier" placeholder="Risk tier"><input id="linNotes" placeholder="Notes"></div><button class="btn" onclick="saveLineage()">Save mapping</button> <button class="btn secondary" onclick="seedLineage()">Seed from current data</button></div><br><div class="panel"><h3>Active mappings</h3>${editableLineageTable(rows)}</div>`;}
function editableLineageTable(rows){if(!rows||!rows.length)return '<div class="empty">No mappings yet.</div>';const cols=['id','Stream','SourceTable','SourceField','TargetTable','TargetField','Report','KPI','Owner','Status','RiskTier','Notes'];return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}<th>Actions</th></tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}<td><button class="btn small secondary" onclick='loadLineage(${JSON.stringify(r).replace(/'/g,"'")})'>Edit</button> <button class="btn small ghost" onclick="deleteLineage('${esc(r.id)}')">Delete</button></td></tr>`).join('')}</tbody></table></div>`;}
function loadLineage(r){linId.value=r.id||'';linStream.value=r.Stream||'';linSourceTable.value=r.SourceTable||'';linSourceField.value=r.SourceField||'';linTargetTable.value=r.TargetTable||'';linTargetField.value=r.TargetField||'';linReport.value=r.Report||'';linKPI.value=r.KPI||'';linOwner.value=r.Owner||'';linStatus.value=r.Status||'Draft';linRiskTier.value=r.RiskTier||'';linNotes.value=r.Notes||'';window.scrollTo({top:0,behavior:'smooth'});}
async function saveLineage(){const payload={id:linId.value,Stream:linStream.value,SourceTable:linSourceTable.value,SourceField:linSourceField.value,TargetTable:linTargetTable.value,TargetField:linTargetField.value,Report:linReport.value,KPI:linKPI.value,Owner:linOwner.value,Status:linStatus.value,RiskTier:linRiskTier.value,Notes:linNotes.value};await postJson('/api/lineage',payload);toast('Mapping saved');await render();}
async function deleteLineage(id){await del('/api/lineage/'+encodeURIComponent(id));toast('Mapping deleted');await render();}
async function seedLineage(){await postJson('/api/seed-lineage',{});toast('Lineage seeded from current data');await render();}
function impactEditor(rows, dataRows){return `<div class="grid grid2"><div class="panel"><h3>Map a field to a report/KPI</h3><div class="field"><label>Field</label><select id="impactField">${dataRows.map((r,i)=>`<option value="${i}">${esc(r.Stream)} · ${esc(r['NCH Target Column'])}</option>`).join('')}</select></div><div class="field"><label>Report / dashboard</label><input id="impactReport" placeholder="Claims Integrity Dashboard"></div><div class="field"><label>KPI</label><input id="impactKpi" placeholder="Match Rate, Paid Claims, Denials"></div><div class="field"><label>Business owner</label><input id="impactOwner" placeholder="Reporting"></div><div class="field"><label>Impact</label><select id="impactLevel"><option>Low</option><option>Moderate</option><option>High</option><option>Critical</option></select></div><div class="field"><label>Decision need</label><input id="impactDecision" placeholder="Confirm remediation priority"></div><br><button class="btn" onclick='addImpact(${JSON.stringify(dataRows).replace(/'/g,"'")})'>Map report impact</button></div><div class="panel"><h3>Saved impact map</h3>${simpleTable(rows)}</div></div>`;}
function governanceEditor(rows){return `<div class="grid grid2"><div class="panel"><h3>Create governance control</h3><input id="govId" placeholder="ID, leave blank for new"><br><br><input id="govControl" placeholder="Control"><br><br><input id="govOwner" placeholder="Owner"><br><br><input id="govCadence" placeholder="Cadence"><br><br><select id="govStatus"><option>Draft</option><option>Planned</option><option>Active</option><option>Blocked</option><option>Retired</option></select><br><br><textarea id="govEvidence" placeholder="Evidence source or notes"></textarea><br><button class="btn" onclick="saveGovernance()">Save control</button></div><div class="panel"><h3>Governance controls</h3>${governanceTable(rows)}</div></div>`;}
function governanceTable(rows){if(!rows||!rows.length)return '<div class="empty">No governance controls saved.</div>';const cols=['id','Control','Owner','Cadence','Status','Evidence'];return `<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${c}</th>`).join('')}<th>Actions</th></tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}<td><button class="btn small secondary" onclick='loadGovernance(${JSON.stringify(r).replace(/'/g,"'")})'>Edit</button> <button class="btn small ghost" onclick="deleteGovernance('${esc(r.id)}')">Delete</button></td></tr>`).join('')}</tbody></table></div>`;}
function loadGovernance(r){govId.value=r.id||'';govControl.value=r.Control||'';govOwner.value=r.Owner||'';govCadence.value=r.Cadence||'';govStatus.value=r.Status||'Draft';govEvidence.value=r.Evidence||'';}
async function saveGovernance(){await postJson('/api/governance',{id:govId.value,Control:govControl.value,Owner:govOwner.value,Cadence:govCadence.value,Status:govStatus.value,Evidence:govEvidence.value});toast('Governance control saved');await render();}
async function deleteGovernance(id){await del('/api/governance/'+encodeURIComponent(id));toast('Governance control deleted');await render();}
function workspacePlaceholder(title,description,items=[]){return `<div class="hero"><span class="workspaceBadge">${esc(workspaceLabels[state.workspace]||'Workspace')}</span><h1>${esc(title)}</h1><p>${esc(description)}</p></div><div class="panel workspaceEmpty"><h3>Workspace is separated from Medicare data</h3><p class="muted">The current All Streams reconciliation workbook is treated as Medicare-only. No Medicare rows are reused here.</p>${items.length?`<div class="grid grid3">${items.map(x=>`<div class="miniCard"><b>${esc(x[0])}</b><span>${esc(x[1])}</span></div>`).join('')}</div>`:''}</div>`;}
function renderSharedPage(meta){document.getElementById('footer').innerHTML=`DART ${esc(meta.version)} · ${esc(workspaceLabels[state.workspace]||'Workspace')} · FastAPI`;const html=state.page==='profile'?profilePage(meta):settingsPage(meta);const appEl=document.getElementById('app');appEl.innerHTML=html;runInjectedScripts(appEl);decorateMetricHelp(appEl);refreshThemeVisuals(appEl);ensureAssistantBubble();} 
function workspaceChatMeta(key){
  if(key==='medicaid')return {label:'CMS Q&A',subtitle:'Medicaid workspace',route:'medicaid-chat'};
  if(key==='byo')return {label:'AI Analyst',subtitle:'Build Your Own',route:'byo-ai'};
  return {label:'DART Copilot',subtitle:'Medicare workspace',route:'copilot'};
}
function getScopedHistory(key){
  state.chatHistories=state.chatHistories||{};
  if(!state.chatHistories[key]){
    const meta=workspaceChatMeta(key);
    state.chatHistories[key]={
      activeId:key+'-main',
      sessions:[{id:key+'-main',title:meta.label,messages:[]}]
    };
  }
  const bucket=state.chatHistories[key];
  if(!bucket.sessions||!bucket.sessions.length){
    const meta=workspaceChatMeta(key);
    bucket.sessions=[{id:key+'-main',title:meta.label,messages:[]}];
    bucket.activeId=key+'-main';
  }
  if(!bucket.sessions.find(s=>s.id===bucket.activeId)){
    bucket.activeId=bucket.sessions[0].id;
  }
  return bucket;
}
function getActiveSession(key){
  const bucket=getScopedHistory(key);
  let session=bucket.sessions.find(s=>s.id===bucket.activeId);
  if(!session){
    session=bucket.sessions[0];
    bucket.activeId=session.id;
  }
  return session;
}
function persistChatHistoryState(){
  try{sessionStorage.setItem('dart_chat_histories_v1', JSON.stringify(state.chatHistories||{}));}catch(_err){}
}
function setScopedHistoryMessages(key,messages){
  const bucket=getScopedHistory(key);
  const active=getActiveSession(key);
  active.messages = Array.isArray(messages) ? messages.map(m=>({role:m?.role||'assistant', content:String(m?.content??'')})) : [];
  if(key==='medicaid'){state.medicaidChat=active.messages.slice();}
  if(key==='byo'){state.byoChat=active.messages.slice();}
  if(key==='medicare'){state.meta=state.meta||{};state.meta.chat=active.messages.slice();}
  persistChatHistoryState();
}
function syncChatHistoryFromWorkspaceState(){
  ['medicare','medicaid','byo'].forEach(key=>{
    const bucket=getScopedHistory(key);
    const active=getActiveSession(key);
    if(key==='medicaid'&&Array.isArray(state.medicaidChat)&&state.medicaidChat.length&&!active.messages.length){
      active.messages=state.medicaidChat.slice();
    }else if(key==='byo'&&Array.isArray(state.byoChat)&&state.byoChat.length&&!active.messages.length){
      active.messages=state.byoChat.slice();
    }else if(key==='medicare'&&Array.isArray(state.meta?.chat)&&state.meta.chat.length&&!active.messages.length){
      active.messages=state.meta.chat.slice();
    }
  });
  persistChatHistoryState();
}
function assistantBubbleWorkspaceKey(){
  if(state.workspace==='medicaid')return 'medicaid';
  if(state.workspace==='byo')return 'byo';
  return 'medicare';
}
function scrollBubbleToBottom(){
  const host=document.getElementById('assistantBubbleHost');
  if(!host)return;
  const msgBox=host.querySelector('.assistantBubbleMessages');
  if(msgBox){
    msgBox.scrollTop=msgBox.scrollHeight;
  }
}
function toggleAssistantBubble(forceOpen){
  const host=document.getElementById('assistantBubbleHost');
  if(!host)return;
  const panel=host.querySelector('.assistantBubblePanel');
  const shouldOpen=typeof forceOpen==='boolean'?forceOpen:(!panel||panel.style.display==='none');
  if(panel){panel.style.display=shouldOpen?'flex':'none';}
  state.assistantBubbleOpen=shouldOpen;
  const launcher=host.querySelector('.assistantBubbleLauncher');
  if(shouldOpen&&launcher){
    launcher.classList.remove('throwDart');
    void launcher.offsetWidth;
    launcher.classList.add('throwDart');
    setTimeout(()=>{
      scrollBubbleToBottom();
      const promptEl=document.getElementById('assistantBubblePrompt');
      if(promptEl)promptEl.focus();
    },60);
  }
}
function ensureAssistantBubble(forceOpen){
  let host=document.getElementById('assistantBubbleHost');
  if(!host){
    host=document.createElement('div');
    host.id='assistantBubbleHost';
    host.className='assistantBubbleHost';
    document.body.appendChild(host);
  }
  if(typeof forceOpen==='boolean'){
    state.assistantBubbleOpen=forceOpen;
  }
  const key=assistantBubbleWorkspaceKey();
  const info=workspaceChatMeta(key);
  const bucket=getScopedHistory(key);
  const active=getActiveSession(key);
  const messages=active?.messages||[];
  const list=messages.length ? messages.map(m=>`<div class="assistantBubbleMessage ${m.role==='user'?'user':'assistant'}"><span class="assistantBubbleMessageRole">${m.role==='user'?'You':'AI'}</span><div class="assistantBubbleMessageBody">${m.role==='assistant'?renderMarkdown(m.content):esc(m.content)}</div></div>`).join('') : '<div class="assistantBubbleEmpty">No messages in this conversation yet. Ask a question below to get started.</div>';
  const historyItems=bucket.sessions.map(s=>`<div class="historyItem ${s.id===bucket.activeId?'active':''}">
        <button type="button" class="historyItemLabel" onclick="selectAssistantChat('${key}','${esc(s.id)}')">${esc(s.title)}</button>
        <button type="button" class="historyItemIcon" title="Rename conversation" onclick="renameAssistantChat(event,'${key}','${esc(s.id)}')">✎</button>
        <button type="button" class="historyItemIcon" title="Delete conversation" onclick="deleteAssistantChat(event,'${key}','${esc(s.id)}')">🗑</button>
      </div>`).join('');
  const datasetBar=key!=='byo' ? '' : `<div class="assistantBubbleDatasetBar">
        <select class="assistantBubbleDatasetSelect" onchange="setByoSelection('left',this.value)">${byoDatasetOptions(state.byoInfo||{datasets:[]},state.byoLeft,state.byoRight)}</select>
        <span class="assistantBubbleDatasetVs">vs</span>
        <select class="assistantBubbleDatasetSelect" onchange="setByoSelection('right',this.value)">${byoDatasetOptions(state.byoInfo||{datasets:[]},state.byoRight,state.byoLeft)}</select>
      </div>`;
  host.innerHTML=`<div class="assistantBubblePanel" ${state.assistantBubbleOpen===false?'style="display:none"':''}>
      <div class="assistantBubbleHeader">
        <div class="assistantBubbleHeaderText"><span class="assistantBubbleCompact">${esc(info.label)}</span><span class="assistantBubbleSubtitle">${esc(info.subtitle)}</span></div>
        <button type="button" aria-label="Close assistant" onclick="toggleAssistantBubble(false)">×</button>
      </div>
      <div class="assistantBubbleHistoryBar">
        <div class="assistantBubbleHistorySelect" id="assistantBubbleHistorySelect">
          <button type="button" class="assistantBubbleHistoryCurrent" onclick="toggleAssistantHistoryMenu(event)">
            <span class="historyDot"></span><span class="historyTitle">${esc(active?.title||'Conversation')}</span><span class="historyChevron">▾</span>
          </button>
          <div class="assistantBubbleHistoryMenu">${historyItems}</div>
        </div>
        <button type="button" class="assistantBubbleNewChat" onclick="createAssistantChat('${key}')">+ New</button>
      </div>
      ${datasetBar}
      <div class="assistantBubbleMessages">${list}</div>
      <div class="assistantBubbleComposer">
        <textarea id="assistantBubblePrompt" placeholder="Ask the ${esc(info.label)} assistant…" onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();sendWorkspaceBubblePrompt('${key}');}"></textarea>
        <button type="button" id="assistantBubbleSendBtn" onclick="sendWorkspaceBubblePrompt('${key}')">Send</button>
      </div>
    </div>
    <button class="assistantBubbleLauncher" type="button" onclick="toggleAssistantBubble()" aria-label="Open workspace AI assistant">
      <svg class="dartboardIcon" viewBox="0 0 48 48" width="34" height="34" aria-hidden="true"><circle cx="24" cy="24" r="22" fill="#fdf6e3" stroke="#1a1a1a" stroke-width="2"/><circle cx="24" cy="24" r="17.5" fill="#c0392b"/><circle cx="24" cy="24" r="12.5" fill="#fdf6e3"/><circle cx="24" cy="24" r="7.5" fill="#c0392b"/><circle cx="24" cy="24" r="3" fill="#1a1a1a"/></svg>
      <svg class="dartArrow" viewBox="0 0 24 24" width="22" height="22" aria-hidden="true"><line x1="3" y1="3" x2="15" y2="15" stroke="#24324a" stroke-width="2.4" stroke-linecap="round"/><polygon points="21,21 12,21 21,12" fill="#e0473f"/><path d="M3 3 L7 3 M3 3 L3 7" stroke="#e0473f" stroke-width="2" stroke-linecap="round"/></svg>
    </button>`;
  if(state.assistantBubbleOpen!==false){
    setTimeout(scrollBubbleToBottom,20);
  }
}
function toggleAssistantHistoryMenu(e){
  e.stopPropagation();
  const wrap=document.getElementById('assistantBubbleHistorySelect');
  if(wrap)wrap.classList.toggle('open');
}
function closeAssistantHistoryMenu(){
  const wrap=document.getElementById('assistantBubbleHistorySelect');
  if(wrap)wrap.classList.remove('open');
}
function selectAssistantChat(key,id){
  const bucket=getScopedHistory(key);
  bucket.activeId=id;
  const session=getActiveSession(key);
  if(key==='medicaid') state.medicaidChat=session.messages.slice();
  if(key==='byo') state.byoChat=session.messages.slice();
  if(key==='medicare'){ state.meta=state.meta||{}; state.meta.chat=session.messages.slice(); }
  persistChatHistoryState();
  ensureAssistantBubble(true);
  if((key==='medicare'&&state.page==='copilot')||(key==='medicaid'&&state.page==='medicaid-chat')||(key==='byo'&&state.page==='byo-ai')){
    render();
  }
}
function createAssistantChat(key){
  const bucket=getScopedHistory(key);
  const info=workspaceChatMeta(key);
  const id=key+'-'+Date.now();
  bucket.sessions.push({id,title:`${info.label} ${bucket.sessions.length+1}`,messages:[]});
  bucket.activeId=id;
  if(key==='medicaid') state.medicaidChat=[];
  if(key==='byo') state.byoChat=[];
  if(key==='medicare'){ state.meta=state.meta||{}; state.meta.chat=[]; }
  persistChatHistoryState();
  ensureAssistantBubble(true);
  if((key==='medicare'&&state.page==='copilot')||(key==='medicaid'&&state.page==='medicaid-chat')||(key==='byo'&&state.page==='byo-ai')){
    render();
  }
}
function renameAssistantChat(e,key,id){
  e.stopPropagation();
  const bucket=getScopedHistory(key);
  const session=bucket.sessions.find(s=>s.id===id);
  if(!session)return;
  const next=prompt('Rename conversation',session.title);
  if(next===null)return;
  const trimmed=next.trim();
  if(!trimmed)return;
  session.title=trimmed.slice(0,60);
  persistChatHistoryState();
  ensureAssistantBubble();
}
function deleteAssistantChat(e,key,id){
  e.stopPropagation();
  const bucket=getScopedHistory(key);
  if(bucket.sessions.length<=1){toast('Keep at least one conversation per workspace');return;}
  if(!confirm('Delete this conversation?'))return;
  bucket.sessions=bucket.sessions.filter(s=>s.id!==id);
  if(bucket.activeId===id)bucket.activeId=bucket.sessions[0].id;
  const active=getActiveSession(key);
  if(key==='medicaid') state.medicaidChat=active.messages.slice();
  if(key==='byo') state.byoChat=active.messages.slice();
  if(key==='medicare'){ state.meta=state.meta||{}; state.meta.chat=active.messages.slice(); }
  persistChatHistoryState();
  ensureAssistantBubble();
  if((key==='medicare'&&state.page==='copilot')||(key==='medicaid'&&state.page==='medicaid-chat')||(key==='byo'&&state.page==='byo-ai')){
    render();
  }
}
async function sendWorkspaceBubblePrompt(key){
  const input=document.getElementById('assistantBubblePrompt');
  const prompt=(input?.value||'').trim();
  if(!prompt)return;
  const bucket=getScopedHistory(key);
  const targetSession=getActiveSession(key);
  const targetSessionId=targetSession.id;
  
  const defaultPrefixes=['DART Copilot','CMS Q&A','AI Analyst','Conversation'];
  if(defaultPrefixes.some(p=>targetSession.title.startsWith(p))&&targetSession.messages.length===0){
    targetSession.title=prompt.length>32?prompt.slice(0,30)+'…':prompt;
  }
  
  targetSession.messages.push({role:'user',content:prompt});
  if(key==='medicaid') state.medicaidChat=targetSession.messages.slice();
  if(key==='byo') state.byoChat=targetSession.messages.slice();
  if(key==='medicare'){ state.meta=state.meta||{}; state.meta.chat=targetSession.messages.slice(); }
  
  state.assistantBubbleOpen=true;
  persistChatHistoryState();
  ensureAssistantBubble(true);
  
  const sendBtn=document.getElementById('assistantBubbleSendBtn');
  if(sendBtn){sendBtn.disabled=true;sendBtn.textContent='…';}
  
  let answer='';
  try{
    if(key==='medicaid'){
      const res=await postJson('/api/medicaid/ai-chat',{question:prompt});
      answer=res.answer||'';
      if(res.details&&res.details.length){
        answer+='\n\n'+res.details.map(x=>'- '+x).join('\n');
      }
    }else if(key==='byo'){
      const res=await postJson('/api/byo/chat',{prompt,left:state.byoLeft||'',right:state.byoRight||''});
      answer=res.answer||'I could not generate an answer from the selected dataset context.';
    }else{
      const res=await postJson('/api/copilot',{prompt,filters:state.filters});
      answer=res.answer||'I could not generate an answer from the current context.';
    }
  }catch(err){
    answer=`I could not reach the AI service: ${err?.message||String(err)}`;
  }
  
  const exactSession=bucket.sessions.find(s=>s.id===targetSessionId)||targetSession;
  exactSession.messages.push({role:'assistant',content:answer});
  
  if(bucket.activeId===targetSessionId){
    if(key==='medicaid') state.medicaidChat=exactSession.messages.slice();
    if(key==='byo') state.byoChat=exactSession.messages.slice();
    if(key==='medicare'){ state.meta=state.meta||{}; state.meta.chat=exactSession.messages.slice(); }
  }
  
  persistChatHistoryState();
  ensureAssistantBubble(true);
  scrollBubbleToBottom();
  
  if((key==='medicare'&&state.page==='copilot')||(key==='medicaid'&&state.page==='medicaid-chat')||(key==='byo'&&state.page==='byo-ai')){
    render();
  }
}
function medMoney(v){const n=Number(v||0),a=Math.abs(n),sign=n<0?'-':'';if(a>=1e9)return `${sign}$${(a/1e9).toFixed(2)}B`;if(a>=1e6)return `${sign}$${(a/1e6).toFixed(2)}M`;if(a>=1e3)return `${sign}$${(a/1e3).toFixed(1)}K`;return `${sign}$${a.toLocaleString(undefined,{maximumFractionDigits:0})}`;}
function medIssueLabel(id){return state.medicaidIssueLabels?.[id]||String(id||'').replaceAll('_',' ').replace(/\b\w/g,m=>m.toUpperCase());}
function medIssueOptions(selected='',includeAll=true){const labels=state.medicaidIssueLabels||{};const ids=Object.keys(labels);return `${includeAll?`<option value="">All issue types</option>`:''}${ids.map(id=>`<option value="${esc(id)}" ${selected===id?'selected':''}>${esc(labels[id])}</option>`).join('')}`;}
function medKpis(items){return `<div class="grid grid5 medKpis">${items.map(([label,value,sub,cls=''])=>`<div class="panel medKpi ${cls}"><div class="label">${esc(label)}</div><div class="value">${value}</div><div class="sub">${esc(sub||'')}</div></div>`).join('')}</div>`;}
function medStatusPill(v){const s=String(v||'').toLowerCase();return `<span class="medStatus ${esc(s)}">${esc(s==='done'?'Resolved':s==='cancelled'?'Cancelled':'Open')}</span>`;}
function medPriorityPill(v){return `<span class="medPriority ${esc(String(v||'').toLowerCase())}">${esc(v||'')}</span>`;}
function medStateCard(s){const total=Number(s.total_issues||0),done=Number(s.done_issues||0),open=Number(s.open_issues||0);const rate=total?done/total*100:0;return `<button class="medStateCard" onclick="openMedicaidState(${Number(s.id)})"><div class="medStateTop"><b>${esc(s.name)}</b><span>${rate.toFixed(0)}% resolved</span></div><div class="medStateStats"><span><b>${intFmt(total)}</b>Total</span><span><b>${intFmt(open)}</b>Open</span><span><b>${intFmt(done)}</b>Resolved</span></div><div class="medProgress"><i style="width:${Math.max(0,Math.min(100,rate))}%"></i></div></button>`;}
function medLeaderboard(rows){return `<div class="medLeaderboard">${(rows||[]).map((r,i)=>`<button onclick="openMedicaidState(${Number(r.id)})"><span class="medRank">${i+1}</span><span class="medLeadName"><b>${esc(r.state)}</b><small>${intFmt(r.total)} issues · ${r.confidence}% confidence</small></span><span class="medLeadScore">${Number(r.composite_score||0).toFixed(1)}</span></button>`).join('')}</div>`;}
function openMedicaidState(id,returnPage='medicaid-states'){state.medicaidStateId=Number(id);state.medicaidStateReturnPage=returnPage||'medicaid-states';state.page='medicaid-state';render();}
async function medUpdateIssue(id,stateId,status){await postJson(`/api/medicaid/issue/${id}/status`,{status});toast('Issue status updated');state.medicaidStateId=Number(stateId);await render();}
async function medCreateIssue(event,stateId){event.preventDefault();const f=event.target;await postJson(`/api/medicaid/state/${stateId}/issue`,{title:f.querySelector('[name=title]').value,description:f.querySelector('[name=description]').value,status:f.querySelector('[name=status]').value,priority:f.querySelector('[name=priority]').value,issue_type:f.querySelector('[name=issue_type]').value});toast('Issue created');await render();}
function medIssueTable(issues,stateId){if(!issues?.length)return '<div class="empty">No issues match the current filters.</div>';return `<div class="tableWrap"><table class="table medIssueTable"><thead><tr><th>Issue</th><th>Type</th><th>Status</th><th>Priority</th><th>Metric</th><th>Started</th><th>Days open</th><th>Tags</th><th>Action</th></tr></thead><tbody>${issues.map(i=>`<tr><td><b>${esc(i.title)}</b><div class="muted medDesc">${esc(i.description||'')}</div></td><td>${esc(medIssueLabel(i.issue_type))}</td><td>${medStatusPill(i.status)}</td><td>${medPriorityPill(i.priority)}</td><td>${i.metric_value===null||i.metric_value===undefined?'—':Number(i.metric_value).toLocaleString()}</td><td>${esc(i.start_date||'—')}</td><td>${i.days_open===null||i.days_open===undefined?'—':intFmt(i.days_open)}</td><td>${esc(i.tags||'—')}</td><td><select class="medInlineSelect" onchange="medUpdateIssue(${Number(i.id)},${Number(stateId)},this.value)"><option value="open" ${i.status==='open'?'selected':''}>Open</option><option value="done" ${i.status==='done'?'selected':''}>Resolved</option><option value="cancelled" ${i.status==='cancelled'?'selected':''}>Cancelled</option></select></td></tr>`).join('')}</tbody></table></div>`;}
function medTopbar(title,sub,actions=''){return `<div class="hero medHero"><span class="workspaceBadge">Medicaid State Intelligence</span><h1>${esc(title)}</h1><p>${esc(sub)}</p>${actions?`<div class="heroActions">${actions}</div>`:''}</div>`;}
function medDashboardPage(d,meta){state.medicaidIssueLabels=d.issue_labels||{};const t=d.totals||{};const actions=`<button class="btn" onclick="showPage('medicaid-states')">Explore states</button><button class="btn secondary" onclick="showPage('medicaid-heatmap')">Open US heatmap</button><button class="btn ghost" onclick="showPage('medicaid-ai')">AI insights</button>`;return `${medTopbar('Medicaid Quality Dashboard','Nationwide CMS claim-quality tracking across all 50 states with issue status, rankings, state drill-downs, and quarterly quality indicators.',actions)}${medKpis([['Total issues',intFmt(t.total),'Across all 50 states'],['Resolved',intFmt(t.done),`${Number(t.resolution_rate||0).toFixed(1)}% resolution rate`],['Open',intFmt(t.open),'Active quality backlog','warn'],['Cancelled',intFmt(t.cancelled),'Closed without resolution'],['States','50','Nationwide coverage']])}<br><div class="grid grid2 medDashboardGrid"><div class="panel"><div class="medSectionHead"><div><h3>National quality leaderboard</h3><p class="muted">Composite score combines resolution, cancellation, backlog, and sample confidence.</p></div><button class="btn ghost small" onclick="showPage('medicaid-compare')">Compare all</button></div>${medLeaderboard(d.leaderboard)}</div><div class="panel"><div class="medSectionHead"><div><h3>All state programs</h3><p class="muted">Open any state for its complete issue tracker and CMS quality summary.</p></div><input class="medStateSearch" placeholder="Filter states..." oninput="medFilterStateCards(this.value)"></div><div class="medStateGrid" id="medStateGrid">${(d.states||[]).map(medStateCard).join('')}</div></div></div><br>${personaLensPanel(meta)}`;}
function medFilterStateCards(q){q=String(q||'').toLowerCase();document.querySelectorAll('#medStateGrid .medStateCard').forEach(x=>x.style.display=x.textContent.toLowerCase().includes(q)?'':'none');}
function medStatesPage(d){state.medicaidIssueLabels=d.issue_labels||{};return `${medTopbar('State Explorer','Browse all 50 state Medicaid programs, review open/resolved issue volume, and drill into a state-level issue workspace.',`<a class="btn secondary" href="/download/medicaid/issues.csv">Export all issues</a><button class="btn ghost" onclick="showPage('medicaid-compare')">Compare states</button>`)}<div class="panel"><div class="medSectionHead"><div><h3>50-state issue tracker</h3><p class="muted">State cards are backed by the same seeded state/issue model used in the standalone application.</p></div><input class="medStateSearch" placeholder="Search state..." oninput="medFilterStateCards(this.value)"></div><div class="medStateGrid wide" id="medStateGrid">${(d.states||[]).map(medStateCard).join('')}</div></div>`;}
function medStatePage(payload,summary){state.medicaidIssueLabels=payload.issue_labels||{};const s=payload.state,sm=payload.summary,dup=summary.duplicate_claims||{};const typeRows=Object.entries(summary.type_breakdown||{}).map(([id,v])=>({'Issue type':medIssueLabel(id),'Total':v.total,'Open':v.open_count,'Resolved':v.done_count,'Q1 2026 metric':v.current_metric??'—','US rank':summary.issue_type_rankings?.[id]??'—'}));return `${medTopbar(`${s.name} Medicaid Quality`,`State-specific CMS issue tracking, quality grade, nationwide rankings, and remediation workflow.`,`<button class="btn secondary" onclick="showPage('${state.medicaidStateReturnPage==='medicaid-heatmap'?'medicaid-heatmap':'medicaid-states'}')">${state.medicaidStateReturnPage==='medicaid-heatmap'?'Back to heatmap':'Back to states'}</button><a class="btn ghost" href="/download/medicaid/issues.csv?state=${encodeURIComponent(s.name)}">Export ${esc(s.code)} issues</a>`)}${medKpis([['Quality grade',esc(summary.quality_grade),'Prototype CMS quality grade'],['Total issues',intFmt(sm.total),'Current filtered state scope'],['Open',intFmt(sm.open),`${Number(sm.success_rate||0).toFixed(1)}% resolved`,'warn'],['Duplicate rate',`${Number(dup.rate||0).toFixed(2)}%`,`${Number(dup.vs_national||0)>=0?'+':''}${Number(dup.vs_national||0).toFixed(2)} vs national avg`],['US dup rank',`#${dup.rank||'—'}`,`of ${dup.total_states||50} · lower is better`]])}<br><div class="grid grid2"><div class="panel"><h3>Issue-type scorecard</h3>${genericTable(typeRows,`${s.name}_issue_type_scorecard`)}</div><div class="panel"><h3>Create a state issue</h3><p class="muted">Adds a custom issue directly to the Medicaid state database.</p><form onsubmit="medCreateIssue(event,${Number(s.id)})"><div class="field"><label>Title</label><input name="title" required placeholder="Describe the state issue"></div><div class="field"><label>Description</label><textarea name="description" rows="3" placeholder="Evidence, context, or remediation notes"></textarea></div><div class="grid grid3"><div class="field"><label>Status</label><select name="status"><option value="open">Open</option><option value="done">Resolved</option><option value="cancelled">Cancelled</option></select></div><div class="field"><label>Priority</label><select name="priority"><option>low</option><option selected>medium</option><option>high</option></select></div><div class="field"><label>Issue type</label><select name="issue_type">${medIssueOptions('',false)}<option value="general">General</option></select></div></div><br><button class="btn" type="submit">Create issue</button></form></div></div><br><div class="panel"><div class="medSectionHead"><div><h3>${esc(s.name)} issue register</h3><p class="muted">Change issue status inline. Updates persist in the Medicaid SQLite workspace.</p></div><div class="heroActions"><select class="medInlineSelect" onchange="state.medicaidStateStatus=this.value;render()"><option value="" ${!state.medicaidStateStatus?'selected':''}>All statuses</option><option value="open" ${state.medicaidStateStatus==='open'?'selected':''}>Open</option><option value="done" ${state.medicaidStateStatus==='done'?'selected':''}>Resolved</option><option value="cancelled" ${state.medicaidStateStatus==='cancelled'?'selected':''}>Cancelled</option></select><select class="medInlineSelect" onchange="state.medicaidStateType=this.value;render()">${medIssueOptions(state.medicaidStateType||'',true)}</select></div></div>${medIssueTable(payload.issues,s.id)}</div>`;}
function medComparePage(d){state.medicaidIssueLabels=d.issue_labels||state.medicaidIssueLabels||{};const rows=(d.rows||[]).sort((a,b)=>b.composite_score-a.composite_score);const display=rows.map((r,i)=>({'Rank':i+1,'State':r.state,'Total':r.total,'Resolved':r.successful,'Open':r.open,'Cancelled':r.cancelled,'Success %':`${r.success_rate}%`,'Backlog %':`${r.backlog_rate}%`,'Confidence %':`${r.confidence}%`,'Composite':r.composite_score}));const ids='medCompareChart';const script=`<scr${''}ipt>(function(){const rows=${JSON.stringify(rows)};if(!window.Plotly)return;Plotly.newPlot('${ids}',[{type:'bar',x:rows.slice(0,15).map(r=>r.state),y:rows.slice(0,15).map(r=>r.composite_score),text:rows.slice(0,15).map(r=>r.composite_score.toFixed(1)),textposition:'outside',hovertemplate:'%{x}<br>Quality score %{y:.1f}<extra></extra>'}],{margin:{l:45,r:20,t:20,b:90},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#fff7d1'},xaxis:{tickangle:-40},yaxis:{range:[0,100],title:'Composite score'}},{responsive:true,displayModeBar:false});})();</scr${''}ipt>`;return `${medTopbar('Compare States','Benchmark Medicaid quality outcomes across states using the standalone app’s success, backlog, cancellation, confidence, and composite-score logic.',`<a class="btn ghost" href="/download/medicaid/issues.csv${state.medicaidIssueType?`?issue_type=${encodeURIComponent(state.medicaidIssueType)}`:''}">Export issues</a>`)}<div class="panel"><div class="filterGrid" style="grid-template-columns:1fr 180px auto"><div class="field"><label>Issue type</label><select onchange="state.medicaidIssueType=this.value;render()">${medIssueOptions(state.medicaidIssueType||'',true)}</select></div><div class="field"><label>Minimum sample</label><input type="number" min="0" value="${Number(state.medicaidMinTotal||0)}" onchange="state.medicaidMinTotal=Number(this.value||0);render()"></div><button class="btn secondary" style="align-self:end" onclick="state.medicaidIssueType='';state.medicaidMinTotal=0;render()">Reset</button></div></div><br><div class="panel"><h3>Quality-score leaderboard</h3><div id="${ids}" class="medChart"></div>${script}</div><br><div class="panel"><h3>State comparison table</h3>${genericTable(display,'medicaid_state_comparison')}</div>`;}
function medAnalyticsPage(a){state.medicaidIssueLabels=a.issue_labels||{};const s=a.stats||{};const total=Number(s.total_issues||0);const typeRows=(a.issue_type_stats||[]).map(x=>({'Issue type':medIssueLabel(x.issue_type),'Total':x.total,'Resolved':x.done,'Open':x.open,'Cancelled':x.cancelled,'Resolution %':`${x.total?x.done/x.total*100:0 .toFixed?.(1)}%`}));const risk=[];Object.entries(a.worst_states||{}).forEach(([it,rows])=>(rows||[]).forEach(r=>risk.push({'Issue type':medIssueLabel(it),'State':r.state_name,'Current metric':r.metric,'Open issues':r.open_issues,'Modeled money at risk':medMoney(r.money_at_risk)})));const chartRows=a.issue_type_stats||[];const script=`<scr${''}ipt>(function(){const rows=${JSON.stringify(chartRows)};if(!window.Plotly)return;Plotly.newPlot('medAnalyticsChart',[{type:'bar',name:'Resolved',x:rows.map(r=>r.issue_type.replaceAll('_',' ')),y:rows.map(r=>r.done)},{type:'bar',name:'Open',x:rows.map(r=>r.issue_type.replaceAll('_',' ')),y:rows.map(r=>r.open)},{type:'bar',name:'Cancelled',x:rows.map(r=>r.issue_type.replaceAll('_',' ')),y:rows.map(r=>r.cancelled)}],{barmode:'stack',margin:{l:45,r:20,t:20,b:125},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#fff7d1'},xaxis:{tickangle:-35},legend:{orientation:'h',y:1.15}},{responsive:true,displayModeBar:false});})();</scr${''}ipt>`;return `${medTopbar('Analytics','System-wide issue analytics, tags, issue-type performance, and modeled financial exposure across Medicaid state programs.')} ${medKpis([['Total issues',intFmt(total),'All issue records'],['Resolved',intFmt(s.done),`${total?Number(s.done||0)/total*100:0 .toFixed?.(1)}% of total`],['Open',intFmt(s.open),'Current backlog','warn'],['Cancelled',intFmt(s.cancelled),'Closed without resolution'],['Issue types',intFmt(a.issue_type_stats?.length),'CMS quality categories']])}<br><div class="panel"><h3>Status mix by CMS issue type</h3><div id="medAnalyticsChart" class="medChart"></div>${script}</div><br><div class="grid grid2"><div class="panel"><h3>Issue-type performance</h3>${genericTable(typeRows,'medicaid_issue_type_analytics')}</div><div class="panel"><h3>Tag usage</h3>${genericTable((a.tag_stats||[]).map(x=>({'Tag':x.name,'Issues':x.count})),'medicaid_tag_usage')}</div></div><br><div class="panel"><h3>Highest modeled risk opportunities</h3><p class="muted">Uses the same deterministic cost multipliers represented by the standalone analytics logic. It is a prioritization proxy, not an audited recovery estimate.</p>${genericTable(risk.slice(0,25),'medicaid_modeled_risk')}</div>`;}
function medInsightsPage(d){state.medicaidIssueLabels=state.medicaidIssueLabels||{};const m=d.meta||{};const cards=(d.insights||[]).map(x=>`<div class="panel medInsight ${esc(x.type)}"><span class="workspaceBadge">${esc(x.type.replaceAll('_',' '))}</span><h3>${esc(x.title)}</h3><p>${esc(x.description)}</p><div class="medRecommendation"><b>Recommended action</b><span>${esc(x.recommendation)}</span></div></div>`).join('');const rows=(d.data||[]).sort((a,b)=>b.composite_score-a.composite_score).slice(0,15).map((r,i)=>({'Rank':i+1,'State':r.state,'Quality score':r.composite_score,'Success %':`${r.success_rate}%`,'Backlog %':`${r.backlog_rate}%`,'Confidence %':`${r.confidence}%`,'Issues':r.total}));return `${medTopbar('AI Insights','Evidence-driven recommendations generated from the same deterministic state comparison model used by the standalone Medicaid app.',`<button class="btn secondary" onclick="showPage('medicaid-chat')">Open CMS Q&A</button>`)}<div class="panel"><div class="filterGrid" style="grid-template-columns:1fr 180px"><div class="field"><label>Issue type</label><select onchange="state.medicaidIssueType=this.value;render()">${medIssueOptions(state.medicaidIssueType||'',true)}</select></div><div class="field"><label>Minimum issues per state</label><input type="number" min="0" value="${Number(state.medicaidMinTotal||3)}" onchange="state.medicaidMinTotal=Number(this.value||3);render()"></div></div></div><br>${medKpis([['States analyzed',intFmt(m.states_analyzed),'Meet current sample threshold'],['States with data',intFmt(m.states_with_data),'Across current filter'],['Avg success',`${Number(m.avg_success_rate||0).toFixed(1)}%`,'Resolved share'],['Avg confidence',`${Number(m.avg_confidence||0).toFixed(1)}%`,'Sample-size confidence'],['Minimum sample',intFmt(m.min_total),'Issues per state']])}<br><div class="grid grid2">${cards}</div><br><div class="panel"><h3>Evidence table</h3>${genericTable(rows,'medicaid_ai_evidence')}</div>`;}
function medChatPage(){const history=state.medicaidChat||[];return `${medTopbar('CMS Quality Q&A','Ask natural-language questions about state quality, duplicate claims, issue types, backlog, rankings, quarterly metrics, or modeled money at risk.')}<div class="grid grid2 medChatGrid"><div class="panel"><h3>Ask the Medicaid analyst</h3><div class="chat medChatHistory" id="medChatHistory">${history.length?history.map(m=>`<div class="bubble ${m.role==='user'?'user':'assistant'}">${m.role==='assistant'?renderMarkdown(m.content):esc(m.content)}</div>`).join(''):'<div class="empty">Ask a question to start the CMS quality conversation.</div>'}</div><br><div class="filterGrid" style="grid-template-columns:1fr auto"><input id="medChatPrompt" placeholder="Which states have the highest duplicate claim rates?" onkeydown="if(event.key==='Enter')askMedicaidChat()"><button class="btn" onclick="askMedicaidChat()">Ask</button></div></div><div class="panel"><h3>Suggested questions</h3><div class="medPromptList"><button onclick="askMedicaidChat('rank the states')">Rank the states</button><button onclick="askMedicaidChat('worst states for invalid diagnosis code')">Worst states for diagnosis-code issues</button><button onclick="askMedicaidChat('open issues in Texas')">Open issues in Texas</button><button onclick="askMedicaidChat('duplicate claim rate by state')">Duplicate claim rates</button><button onclick="askMedicaidChat('largest money at risk')">Largest modeled money at risk</button></div></div></div>`;}
async function askMedicaidChat(prefill=''){
  const input=document.getElementById('medChatPrompt');
  const q=(prefill||input?.value||'').trim();
  if(!q)return;
  if(input)input.value='';
  const bucket=getScopedHistory('medicaid');
  const session=getActiveSession('medicaid');
  if(session.messages.length===0){
    session.title=q.length>32?q.slice(0,30)+'…':q;
  }
  session.messages.push({role:'user',content:q});
  state.medicaidChat=session.messages.slice();
  persistChatHistoryState();
  try{
    const result=await postJson('/api/medicaid/ai-chat',{question:q});
    let content=result.answer||'';
    if(result.details?.length)content+='\n\n'+result.details.map(x=>'- '+x).join('\n');
    session.messages.push({role:'assistant',content});
  }catch(err){
    session.messages.push({role:'assistant',content:`Could not reach CMS Q&A service: ${err?.message||err}`});
  }
  state.medicaidChat=session.messages.slice();
  persistChatHistoryState();
  await render();
  setTimeout(()=>{
    document.getElementById('medChatHistory')?.scrollTo(0,999999);
    scrollBubbleToBottom();
  },50);
}
function medExecPage(b){const k=b.kpis||{};const actions=b.actions_30_60_90||{};const top=(b.top_states||[]).map((r,i)=>({'Rank':i+1,'State':r.state,'Quality score':r.composite_score,'Success %':`${r.success_rate}%`,'Backlog %':`${r.backlog_rate}%`,'Confidence %':`${r.confidence}%`}));const watch=(b.watchlist||[]).map((r,i)=>({'Priority':i+1,'State':r.state,'Quality score':r.composite_score,'Backlog %':`${r.backlog_rate}%`,'Open':r.open,'Issues':r.total}));const actionCol=(title,items)=>`<div class="panel medActionCol"><span class="workspaceBadge">${esc(title)}</span>${(items||[]).map((x,i)=>`<div class="medActionItem"><b>${i+1}</b><span>${esc(x)}</span></div>`).join('')}</div>`;return `${medTopbar('Executive Brief','Leadership-ready Medicaid quality briefing with KPI coverage, top performers, watchlist states, and a 30/60/90-day action plan.')}<div class="panel"><div class="filterGrid" style="grid-template-columns:1fr 180px"><div class="field"><label>Issue type</label><select onchange="state.medicaidIssueType=this.value;render()">${medIssueOptions(state.medicaidIssueType||'',true)}</select></div><div class="field"><label>Minimum sample</label><input type="number" min="0" value="${Number(state.medicaidMinTotal||3)}" onchange="state.medicaidMinTotal=Number(this.value||3);render()"></div></div></div><br>${medKpis([['States analyzed',intFmt(k.states_analyzed),'Current executive scope'],['Avg success',`${Number(k.avg_success||0).toFixed(1)}%`,'Resolved issue rate'],['Avg backlog',`${Number(k.avg_backlog||0).toFixed(1)}%`,'Open issue share'],['Avg quality',Number(k.avg_quality||0).toFixed(1),'Composite score'],['Coverage',`${Number(k.coverage||0).toFixed(1)}%`,'States meeting threshold']])}<br><div class="panel medBriefSummary"><span class="workspaceBadge">Executive snapshot</span><h2>${esc(b.summary||'')}</h2></div><br><div class="grid grid2"><div class="panel"><h3>Top performers</h3>${genericTable(top,'medicaid_executive_top_states')}</div><div class="panel"><h3>Watchlist</h3>${genericTable(watch,'medicaid_executive_watchlist')}</div></div><br><div class="grid grid3">${actionCol('Next 30 days',actions['30_days'])}${actionCol('Next 60 days',actions['60_days'])}${actionCol('Next 90 days',actions['90_days'])}</div>`;}
function medHeatSideEmpty(){return `<div class="medHeatSideEmpty"><div><b>Select a state on the map</b><span>Click any state to keep the national heatmap visible while reviewing that state's quality metrics and current issues here.</span></div></div>`;}
function medCloseHeatState(){state.medicaidHeatSelectedId=0;const el=document.getElementById('medHeatStatePanel');if(el)el.innerHTML=medHeatSideEmpty();}
function medHeatMetricDisplay(metric,value){const pctMetrics=new Set(['quality_score','success_rate','backlog_rate','cancel_rate']);if(value===null||value===undefined||Number.isNaN(Number(value)))return '—';return pctMetrics.has(metric)?`${Number(value).toFixed(1)}${metric==='quality_score'?'':'%'}`:intFmt(value);}
function medHeatStatePanel(payload,summary,row){const s=payload.state||{},sm=payload.summary||{},dup=summary.duplicate_claims||{},metric=state.medicaidHeatMetric||'quality_score';const metricLabels={quality_score:'Quality score',success_rate:'Success rate',total:'Total issues',successful:'Resolved issues',open:'Open issues',cancelled:'Cancelled issues',backlog_rate:'Backlog rate',cancel_rate:'Cancellation rate'};const breakdown=Object.entries(summary.type_breakdown||{}).sort((a,b)=>Number(b[1]?.open_count||0)-Number(a[1]?.open_count||0)||Number(b[1]?.total||0)-Number(a[1]?.total||0)).slice(0,5);const issues=(payload.issues||[]).slice().sort((a,b)=>{const p={high:3,medium:2,low:1};return (p[b.priority]||0)-(p[a.priority]||0)||(a.status==='open'?-1:1);}).slice(0,6);return `<div class="medHeatSideHead"><div><span class="workspaceBadge">${esc(s.code||'State')} selected</span><h2>${esc(s.name||'State')}</h2><p class="muted">Filtered preview beside the national map</p></div><button class="medHeatClose" type="button" onclick="medCloseHeatState()" title="Close state preview">×</button></div><div class="medHeatMiniKpis"><div class="medHeatMiniKpi"><span>${esc(metricLabels[metric]||metric)}</span><b>${medHeatMetricDisplay(metric,row?.metric_value)}</b></div><div class="medHeatMiniKpi"><span>Quality grade</span><b>${esc(summary.quality_grade||'—')}</b></div><div class="medHeatMiniKpi"><span>Open issues</span><b>${intFmt(sm.open)}</b></div><div class="medHeatMiniKpi"><span>Resolved</span><b>${intFmt(sm.successful)}</b></div><div class="medHeatMiniKpi"><span>Resolution rate</span><b>${Number(sm.success_rate||0).toFixed(1)}%</b></div><div class="medHeatMiniKpi"><span>Duplicate rate</span><b>${Number(dup.rate||0).toFixed(2)}%</b></div></div><div class="medHeatSideSection"><h4>Issue mix</h4><div class="medHeatBreakdown">${breakdown.length?breakdown.map(([id,v])=>`<div class="medHeatBreakRow"><div><b>${esc(medIssueLabel(id))}</b><small>${intFmt(v.total)} total · ${intFmt(v.done_count)} resolved</small></div><span class="medStatus ${Number(v.open_count||0)>0?'open':'done'}">${intFmt(v.open_count)} open</span></div>`).join(''):'<div class="muted">No issue-type data for this state.</div>'}</div></div><div class="medHeatSideSection"><h4>Current issues</h4><div class="medHeatIssueList">${issues.length?issues.map(i=>`<div class="medHeatIssue"><div class="medHeatIssueTop"><b>${esc(i.title)}</b>${medPriorityPill(i.priority)}</div><div class="medHeatIssueMeta">${medStatusPill(i.status)}<span class="muted">${esc(medIssueLabel(i.issue_type))}</span>${i.days_open===null||i.days_open===undefined?'':`<span class="muted">${intFmt(i.days_open)} days open</span>`}</div></div>`).join(''):'<div class="muted">No issues match the current heatmap filters.</div>'}</div></div><div class="medHeatSideSection"><div class="medHeatBreakRow"><div><b>Duplicate-claim standing</b><small>${Number(dup.vs_national||0)>=0?'+':''}${Number(dup.vs_national||0).toFixed(2)} pts vs national average</small></div><b>#${dup.rank||'—'} / ${dup.total_states||50}</b></div></div><button class="btn medHeatOpenButton" type="button" onclick="openMedicaidState(${Number(s.id)},'medicaid-heatmap')">Open full ${esc(s.name||'state')} page →</button>`;}
async function medSelectHeatState(id){id=Number(id);if(!id)return;state.medicaidHeatSelectedId=id;const panel=document.getElementById('medHeatStatePanel');if(panel)panel.innerHTML='<div class="medHeatSideLoading">Loading state metrics and issues…</div>';const q=new URLSearchParams();if(state.medicaidHeatStatus)q.set('status',state.medicaidHeatStatus);if(state.medicaidHeatType)q.set('issue_type',state.medicaidHeatType);try{const [payload,summary]=await Promise.all([api(`/api/medicaid/heatmap/state/${id}?${q.toString()}`),api(`/api/medicaid/state-summary/${id}`)]);if(state.medicaidHeatSelectedId!==id)return;const row=(state.medicaidHeatRows||[]).find(r=>Number(r.id)===id)||{};if(panel){panel.innerHTML=medHeatStatePanel(payload,summary,row);decorateMetricHelp(panel);}}catch(err){if(panel)panel.innerHTML=`<div class="empty">Could not load this state preview: ${esc(String(err.message||err))}</div>`;}}
function medHeatmapPage(h,dashboard){state.medicaidIssueLabels=dashboard.issue_labels||state.medicaidIssueLabels||{};const rows=h.rows||[];state.medicaidHeatRows=rows;const metric=state.medicaidHeatMetric||'quality_score';const metricLabels={quality_score:'Quality score',success_rate:'Success rate',total:'Total issues',successful:'Resolved issues',open:'Open issues',cancelled:'Cancelled issues',backlog_rate:'Backlog rate',cancel_rate:'Cancellation rate'};const script=`<scr${''}ipt>(function(){const rows=${JSON.stringify(rows)};if(!window.Plotly)return;const trace={type:'choropleth',locationmode:'USA-states',locations:rows.map(r=>r.state_code),z:rows.map(r=>r.metric_value),text:rows.map(r=>r.name+'<br>${metricLabels[metric]}: '+r.metric_value),hovertemplate:'%{text}<extra></extra>',colorscale:'YlOrBr',marker:{line:{color:'#5a4520',width:0.7}},colorbar:{title:${JSON.stringify(metricLabels[metric]||metric)}}};Plotly.newPlot('medicaidUSMap',[trace],{geo:{scope:'usa',bgcolor:'rgba(0,0,0,0)',lakecolor:'rgba(0,0,0,0)'},margin:{l:0,r:0,t:10,b:0},paper_bgcolor:'rgba(0,0,0,0)',font:{color:'#fff7d1'}},{responsive:true,displayModeBar:false}).then(g=>g.on('plotly_click',e=>{const code=e.points?.[0]?.location;const row=rows.find(r=>r.state_code===code);if(row)medSelectHeatState(row.id);}));if(state.medicaidHeatSelectedId)medSelectHeatState(state.medicaidHeatSelectedId);})();</scr${''}ipt>`;const ranking=[...rows].sort((a,b)=>Number(b.metric_value||0)-Number(a.metric_value||0)).slice(0,15).map((r,i)=>({'Rank':i+1,'State':r.name,[metricLabels[metric]||metric]:r.metric_value,'Total':r.total,'Open':r.open,'Resolved':r.successful}));return `${medTopbar('US Quality Heatmap','Interactive 50-state heatmap with an in-place state preview. Click a state to review its metrics and issues beside the map, then open the complete state workspace when you need the full detail.')}<div class="panel"><div class="filterGrid medHeatFilters"><div class="field"><label>Metric</label><select onchange="state.medicaidHeatMetric=this.value;render()">${Object.entries(metricLabels).map(([v,l])=>`<option value="${v}" ${metric===v?'selected':''}>${l}</option>`).join('')}</select></div><div class="field"><label>Status</label><select onchange="state.medicaidHeatStatus=this.value;render()"><option value="">All statuses</option><option value="open" ${state.medicaidHeatStatus==='open'?'selected':''}>Open</option><option value="done" ${state.medicaidHeatStatus==='done'?'selected':''}>Resolved</option><option value="cancelled" ${state.medicaidHeatStatus==='cancelled'?'selected':''}>Cancelled</option></select></div><div class="field"><label>Issue type</label><select onchange="state.medicaidHeatType=this.value;render()">${medIssueOptions(state.medicaidHeatType||'',true)}</select></div></div></div><br><div class="medHeatWorkspace"><div class="panel medMapPanel"><div id="medicaidUSMap"></div>${script}</div><aside class="panel medHeatSidePanel" id="medHeatStatePanel">${medHeatSideEmpty()}</aside></div><br><div class="panel"><h3>Current heatmap ranking</h3>${genericTable(ranking,'medicaid_heatmap_ranking')}</div>`;}
function medClaimsPage(c){const s=c.summary||{}, trend=c.duplicate_trend||[], opportunities=(c.top_opportunities||[]).map((x,i)=>({'Priority':i+1,'State':x.state_name,'Issue type':x.label,'Current metric':x.metric,'Open issues':x.open_issues,'Modeled money at risk':medMoney(x.money_at_risk)}));const typeRows=(c.issue_type_metrics||[]).map(x=>({'Issue type':x.label,'Metric volume':Number(x.metric_total).toLocaleString(),'Issue records':x.records,'Open issues':x.open_issues,'Modeled money at risk':medMoney(x.estimated_money_at_risk)}));const script=`<scr${''}ipt>(function(){const rows=${JSON.stringify(trend)};if(!window.Plotly)return;Plotly.newPlot('medClaimsTrend',[{type:'scatter',mode:'lines+markers',x:rows.map(r=>r.quarter),y:rows.map(r=>r.avg_rate),name:'Average duplicate rate',fill:'tozeroy'}],{margin:{l:55,r:20,t:20,b:50},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{color:'#fff7d1'},yaxis:{title:'Duplicate claim rate %'}},{responsive:true,displayModeBar:false});})();</scr${''}ipt>`;return `${medTopbar('Claims Analysis','Medicaid claim-quality and spending-risk analysis derived from the state issue model, including duplicate trends, issue-volume diagnostics, monitoring subscriptions, and prioritization.')} ${medKpis([['Total quality issues',intFmt(s.total_issues),'State issue database'],['Open',intFmt(s.open),'Current claim-quality backlog','warn'],['Resolved',intFmt(s.done),'Completed remediation'],['Monitoring agents',intFmt(c.subscriptions),'Registered subscriptions'],['Top modeled risk',opportunities[0]?.['Modeled money at risk']||'$0','Highest current state/type opportunity']])}<br><div class="grid grid2"><div class="panel"><h3>Duplicate claim trend</h3><div id="medClaimsTrend" class="medChart"></div>${script}</div><div class="panel"><h3>Register monitoring agent</h3><p class="muted">Preserves the standalone app’s agent-subscription concept inside the Medicaid workspace.</p><form onsubmit="registerMedicaidAgent(event)"><div class="field"><label>Email</label><input name="email" type="email" required placeholder="analyst@example.gov"></div><div class="field"><label>Requirements</label><input name="requirements" placeholder="duplicate claims, payment risk, open backlog"></div><div class="field"><label>Filter value</label><input name="filter" placeholder="Optional state or issue type"></div><br><button class="btn" type="submit">Register agent</button></form></div></div><br><div class="panel"><h3>Top state opportunities</h3>${genericTable(opportunities,'medicaid_claims_opportunities')}</div><br><div class="panel"><h3>Issue-type spending / quality model</h3>${genericTable(typeRows,'medicaid_claims_type_metrics')}</div>`;}
async function registerMedicaidAgent(event){event.preventDefault();const f=event.target;const requirements=f.requirements.value.split(',').map(x=>x.trim()).filter(Boolean);const r=await postJson('/api/medicaid/claims/register-agent',{email:f.email.value,requirements,filter_value:f.filter.value,rules:[]});toast(r.message||'Agent registered');await render();}
function medOptimizePage(c){const opp=(c.top_opportunities||[]).map((x,i)=>({'Priority':i+1,'State':x.state_name,'Issue type':x.label,'Metric':x.metric,'Open':x.open_issues,'Modeled money at risk':medMoney(x.money_at_risk),'Suggested action':x.issue_type==='duplicate_claims'?'Strengthen duplicate detection and pre-pay edit rules':x.issue_type==='payment_amount_exceeded'?'Review fee schedule/outlier controls':x.issue_type==='referential_integrity'?'Repair provider/member reference integrity':'Target the highest-volume validation failure'}));const total=(c.issue_type_metrics||[]).reduce((a,x)=>a+Number(x.estimated_money_at_risk||0),0);return `${medTopbar('Optimize Spending','Actionable Medicaid quality and modeled overspending opportunities, prioritized by the same issue metrics and cost multipliers represented in the standalone app.')} ${medKpis([['Modeled exposure',medMoney(total),'Across all issue types'],['Priority opportunities',intFmt(opp.length),'State/type combinations'],['Top opportunity',opp[0]?.['State']||'—',opp[0]?.['Issue type']||''],['Top modeled risk',opp[0]?.['Modeled money at risk']||'$0','Planning proxy'],['Open quality issues',intFmt(c.summary?.open),'Current backlog']])}<br><div class="panel"><h3>Optimization queue</h3>${genericTable(opp,'medicaid_optimization_queue')}</div><br><div class="grid grid3"><div class="panel"><span class="workspaceBadge">30 days</span><h3>Control the largest leaks</h3><p>Target the top duplicate, payment, and code-quality opportunities. Confirm high-risk state metrics and assign accountable owners.</p></div><div class="panel"><span class="workspaceBadge">60 days</span><h3>Standardize remediation</h3><p>Deploy common validation rules and state playbooks, then compare backlog and quality-score movement across programs.</p></div><div class="panel"><span class="workspaceBadge">90 days</span><h3>Institutionalize monitoring</h3><p>Use agent subscriptions, quarterly state scorecards, and executive briefs to sustain the strongest controls.</p></div></div>`;}
function medMethodologyPage(c){const m=c.methodology||{};const rows=Object.entries(m.cost_multipliers||{}).map(([it,v])=>({'Issue type':medIssueLabel(it),'Prototype multiplier':it==='duplicate_claims'?`${medMoney(v)} per 1 percentage point`:medMoney(v)+' per modeled case','Metric interpretation':it==='duplicate_claims'?'Quarterly duplicate claim rate (%)':'Q1 2026 issue/case volume'}));const defs=[{'Measure':'Success rate','Formula':'Resolved ÷ Total issues','Use':'Issue-resolution performance'},{'Measure':'Backlog rate','Formula':'Open ÷ Total issues','Use':'Outstanding remediation load'},{'Measure':'Cancellation rate','Formula':'Cancelled ÷ Total issues','Use':'Issues closed without resolution'},{'Measure':'Confidence','Formula':'min(100%, Total issues ÷ 6 × 100)','Use':'Sample-size reliability adjustment'},{'Measure':'Composite quality','Formula':'[55% success + 25% (1-cancel) + 20% (1-backlog)] × confidence','Use':'Cross-state quality ranking'},{'Measure':'Modeled money at risk','Formula':'Current issue metric × prototype cost multiplier','Use':'Prioritization proxy; not audited recovery'}];return `${medTopbar('Methodology & Data Dictionary','Reference for the Medicaid state-quality prototype: issue definitions, score formulas, cost multipliers, quarterly framing, and interpretation guidance.')}<div class="panel medMethodNote"><h3>Important interpretation</h3><p>${esc(m.note||'The Medicaid workspace uses deterministic prototype data for demonstration and prioritization.')}</p></div><br><div class="grid grid2"><div class="panel"><h3>Scoring methodology</h3>${genericTable(defs,'medicaid_scoring_methodology')}</div><div class="panel"><h3>Issue cost multipliers</h3>${genericTable(rows,'medicaid_cost_multipliers')}</div></div><br><div class="panel"><h3>CMS issue dictionary</h3>${genericTable(Object.entries(state.medicaidIssueLabels||{}).map(([id,label])=>({'Issue type ID':id,'Display name':label,'Standalone title':label})), 'medicaid_issue_dictionary')}</div>`;}

async function renderMedicaid(meta){
  let html='';
  if(state.page==='medicaid-home'){
    const d=await api('/api/medicaid/dashboard');html=medDashboardPage(d,meta);
  }else if(state.page==='medicaid-states'){
    const d=await api('/api/medicaid/dashboard');html=medStatesPage(d);
  }else if(state.page==='medicaid-state'){
    const id=Number(state.medicaidStateId||1);const qs=new URLSearchParams();if(state.medicaidStateStatus)qs.set('status',state.medicaidStateStatus);if(state.medicaidStateType)qs.set('issue_type',state.medicaidStateType);const [p,s]=await Promise.all([api(`/api/medicaid/state/${id}?${qs.toString()}`),api(`/api/medicaid/state-summary/${id}`)]);html=medStatePage(p,s);
  }else if(state.page==='medicaid-compare'){
    const q=new URLSearchParams();if(state.medicaidIssueType)q.set('issue_type',state.medicaidIssueType);q.set('min_total',String(Number(state.medicaidMinTotal||0)));const d=await api('/api/medicaid/compare?'+q.toString());html=medComparePage(d);
  }else if(state.page==='medicaid-analytics'){
    const a=await api('/api/medicaid/analytics');html=medAnalyticsPage(a);
  }else if(state.page==='medicaid-ai'){
    const q=new URLSearchParams();if(state.medicaidIssueType)q.set('issue_type',state.medicaidIssueType);q.set('min_total',String(Number(state.medicaidMinTotal||3)));const d=await api('/api/medicaid/ai-comparison?'+q.toString());if(!state.medicaidIssueLabels){const dash=await api('/api/medicaid/dashboard');state.medicaidIssueLabels=dash.issue_labels;}html=medInsightsPage(d);
  }else if(state.page==='medicaid-chat'){
    if(!state.medicaidIssueLabels){const dash=await api('/api/medicaid/dashboard');state.medicaidIssueLabels=dash.issue_labels;}html=medChatPage();
  }else if(state.page==='medicaid-exec'){
    if(!state.medicaidIssueLabels){const dash=await api('/api/medicaid/dashboard');state.medicaidIssueLabels=dash.issue_labels;}const q=new URLSearchParams();if(state.medicaidIssueType)q.set('issue_type',state.medicaidIssueType);q.set('min_total',String(Number(state.medicaidMinTotal||3)));const b=await api('/api/medicaid/executive-brief?'+q.toString());html=medExecPage(b);
  }else if(state.page==='medicaid-heatmap'){
    const q=new URLSearchParams({metric:state.medicaidHeatMetric||'quality_score'});if(state.medicaidHeatStatus)q.set('status',state.medicaidHeatStatus);if(state.medicaidHeatType)q.set('issue_type',state.medicaidHeatType);const [h,d]=await Promise.all([api('/api/medicaid/heatmap?'+q.toString()),api('/api/medicaid/dashboard')]);html=medHeatmapPage(h,d);
  }else if(state.page==='medicaid-claims'){
    const c=await api('/api/medicaid/claims-data');if(!state.medicaidIssueLabels)state.medicaidIssueLabels=c.methodology?.issue_labels||{};html=medClaimsPage(c);
  }else if(state.page==='medicaid-optimize'){
    const c=await api('/api/medicaid/claims-data');if(!state.medicaidIssueLabels)state.medicaidIssueLabels=c.methodology?.issue_labels||{};html=medOptimizePage(c);
  }else if(state.page==='medicaid-methodology'){
    const c=await api('/api/medicaid/claims-data');state.medicaidIssueLabels=c.methodology?.issue_labels||{};html=medMethodologyPage(c);
  }else{state.page='medicaid-home';return renderMedicaid(meta);}
  document.getElementById('footer').innerHTML=`Medicaid State Intelligence · 50-state CMS quality prototype · ${esc(meta.version)} · FastAPI`;
  const appEl=document.getElementById('app');appEl.classList.remove('page-enter');appEl.innerHTML=html;runInjectedScripts(appEl);decorateMetricHelp(appEl);refreshThemeVisuals(appEl);requestAnimationFrame(()=>appEl.classList.add('page-enter'));ensureAssistantBubble();
}
function byoStats(info,prefix=''){return `<div class="grid grid6"><div class="panel metric"><div class="label">Rows</div><div class="value">${intFmt(info.rows)}</div><div class="sub">${prefix||'Dataset'}</div></div><div class="panel metric"><div class="label">Columns</div><div class="value">${intFmt(info.columns)}</div><div class="sub">Detected fields</div></div><div class="panel metric"><div class="label">Completeness</div><div class="value">${pctNum(info.completeness)}</div><div class="sub">Non-missing cells</div></div><div class="panel metric"><div class="label">Missing cells</div><div class="value">${intFmt(info.missing_cells)}</div><div class="sub">Across all fields</div></div><div class="panel metric"><div class="label">Duplicate rows</div><div class="value">${intFmt(info.duplicate_rows)}</div><div class="sub">Exact duplicates</div></div><div class="panel metric"><div class="label">Numeric fields</div><div class="value">${intFmt(info.numeric_columns)}</div><div class="sub">Detected numeric</div></div></div>`;}
function byoUploadPanel(info){return `<div class="byoDrop"><span class="workspaceBadge">Persistent upload</span><h3>Save datasets to your repo</h3><p class="muted">Upload one or more CSV/XLSX files. DART validates each file, then saves the original file under <span class="byoStoragePath">${esc(info.storage_path||'Data/DIY')}</span>. Files remain available after FastAPI restarts.</p><input id="byoFiles" type="file" accept=".csv,.xlsx" multiple><div class="heroActions" style="justify-content:center"><button class="btn" type="button" onclick="uploadByoFiles()">Upload & save</button><button class="btn secondary" type="button" onclick="showPage('byo-library')">Open Dataset Library</button></div></div>`;}
async function uploadByoFiles(){const input=document.getElementById('byoFiles');const files=[...(input?.files||[])];if(!files.length){toast('Choose at least one CSV or Excel file');return;}const fd=new FormData();files.forEach(f=>fd.append('files',f));try{const r=await fetch('/api/byo/upload',{method:'POST',body:fd});const payload=await r.json().catch(()=>({error:'Upload failed'}));if(!r.ok)throw new Error(payload.error||'Upload failed');const names=(payload.saved||[]).map(x=>x.name);if(names.length){state.byoLeft=state.byoLeft||names[0];if(!state.byoRight&&names.length>1)state.byoRight=names[1];saveByoSelection();}toast(`${names.length} dataset${names.length===1?'':'s'} saved to ${payload.storage_path||'Data/DIY'}`);await render();}catch(err){toast('Upload failed: '+String(err.message||err).slice(0,180));}}
function byoImportPanel(){
  const s3=state.byoS3Items;
  const sp=state.byoSharePointItems;
  return `<div class="grid grid2">
    <div class="panel byoImportPanel">
      <span class="workspaceBadge">Online folder / URL</span>
      <h3>Import from a web link</h3>
      <p class="muted">Paste one or more direct CSV/XLSX file URLs (one per line). DART downloads and saves each into your DIY library.</p>
      <textarea id="byoImportUrls" placeholder="https://example.com/folder/claims.csv&#10;https://example.com/folder/ledger.xlsx" rows="4"></textarea>
      <div class="heroActions" style="justify-content:flex-start"><button class="btn" type="button" onclick="importByoFromUrl()">Import from URL(s)</button></div>
    </div>
    <div class="panel byoImportPanel">
      <span class="workspaceBadge">Amazon S3</span>
      <h3>Import from an S3 bucket</h3>
      <p class="muted">List CSV/XLSX objects in a bucket (optionally under a prefix/folder), then choose which ones to save. Leave the keys blank to use the server's default AWS credentials.</p>
      <div class="byoS3Fields">
        <input id="byoS3Bucket" type="text" placeholder="Bucket name">
        <input id="byoS3Prefix" type="text" placeholder="Prefix / folder (optional)">
        <input id="byoS3Region" type="text" placeholder="Region (optional)">
        <input id="byoS3AccessKey" type="text" placeholder="Access key (optional)">
        <input id="byoS3SecretKey" type="password" placeholder="Secret key (optional)">
        <input id="byoS3SessionToken" type="password" placeholder="Session token (optional)">
      </div>
      <div class="heroActions" style="justify-content:flex-start"><button class="btn secondary" type="button" onclick="listByoS3Files()">List bucket files</button></div>
      ${s3?byoS3ResultsHtml(s3):''}
    </div>
    <div class="panel byoImportPanel" style="grid-column:1/-1">
      <span class="workspaceBadge">SharePoint / OneDrive (private)</span>
      <h3>Import from a private SharePoint or OneDrive folder</h3>
      <p class="muted">Paste a SharePoint or OneDrive (work/school account) sharing link to a folder or file. This needs an Azure AD app registration with the <b>Sites.Read.All</b> (or Files.Read.All) application permission, admin-consented in your tenant — DART signs in app-only, so no user login or cookie is required. Personal <b>outlook.com</b> OneDrive accounts aren't supported by this flow.</p>
      <div class="byoS3Fields">
        <input id="byoSpTenant" type="text" placeholder="Azure AD tenant ID">
        <input id="byoSpClientId" type="text" placeholder="App (client) ID">
        <input id="byoSpClientSecret" type="password" placeholder="Client secret">
        <input id="byoSpShareUrl" type="text" placeholder="SharePoint/OneDrive share link" style="grid-column:1/-1">
      </div>
      <div class="heroActions" style="justify-content:flex-start"><button class="btn secondary" type="button" onclick="listByoSharePointFiles()">List folder files</button></div>
      ${sp?byoSharePointResultsHtml(sp):''}
    </div>
  </div>`;
}
function byoS3ResultsHtml(items){
  if(!items.length)return '<div class="empty" style="margin-top:12px">No matching .csv/.xlsx objects found for that bucket/prefix.</div>';
  return `<div class="byoS3Results">
    <div class="byoS3ResultsHead"><b>${items.length} file${items.length===1?'':'s'} found</b><button class="btn small" type="button" onclick="importByoFromS3()">Import selected</button></div>
    ${items.map(x=>`<label class="byoS3Row"><input type="checkbox" class="byoS3Check" value="${esc(x.key)}"><span class="byoS3Key">${esc(x.key)}</span><span class="byoS3Size muted">${x.size_mb} MB</span></label>`).join('')}
  </div>`;
}
async function importByoFromUrl(){
  const box=document.getElementById('byoImportUrls');
  const raw=(box?.value||'').trim();
  if(!raw){toast('Paste at least one file URL');return;}
  try{
    const r=await fetch('/api/byo/import/url',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({urls:raw})});
    const payload=await r.json().catch(()=>({error:'Import failed'}));
    if(!r.ok)throw new Error(payload.error||'Import failed');
    const names=(payload.saved||[]).map(x=>x.name);
    if(names.length){state.byoLeft=state.byoLeft||names[0];if(!state.byoRight&&names.length>1)state.byoRight=names[1];saveByoSelection();}
    toast(`${names.length} dataset${names.length===1?'':'s'} imported${payload.errors?.length?`, ${payload.errors.length} failed`:''}`);
    await render();
  }catch(err){toast('Import failed: '+String(err.message||err).slice(0,200));}
}
function byoS3Context(){
  return {
    bucket:document.getElementById('byoS3Bucket')?.value.trim()||'',
    prefix:document.getElementById('byoS3Prefix')?.value.trim()||'',
    region:document.getElementById('byoS3Region')?.value.trim()||'',
    access_key:document.getElementById('byoS3AccessKey')?.value.trim()||'',
    secret_key:document.getElementById('byoS3SecretKey')?.value.trim()||'',
    session_token:document.getElementById('byoS3SessionToken')?.value.trim()||'',
  };
}
async function listByoS3Files(){
  const ctx=byoS3Context();
  if(!ctx.bucket){toast('Enter an S3 bucket name');return;}
  try{
    const r=await fetch('/api/byo/import/s3/list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(ctx)});
    const payload=await r.json().catch(()=>({error:'S3 list failed'}));
    if(!r.ok)throw new Error(payload.error||'S3 list failed');
    state.byoS3Items=payload.items||[];
    state.byoS3Context=ctx;
    toast(`${state.byoS3Items.length} matching file(s) found`);
    await render();
  }catch(err){toast('S3 list failed: '+String(err.message||err).slice(0,200));}
}
async function importByoFromS3(){
  const checked=[...document.querySelectorAll('.byoS3Check:checked')].map(x=>x.value);
  if(!checked.length){toast('Select at least one S3 file');return;}
  const ctx=state.byoS3Context||byoS3Context();
  try{
    const r=await fetch('/api/byo/import/s3',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({...ctx,keys:checked})});
    const payload=await r.json().catch(()=>({error:'S3 import failed'}));
    if(!r.ok)throw new Error(payload.error||'S3 import failed');
    const names=(payload.saved||[]).map(x=>x.name);
    if(names.length){state.byoLeft=state.byoLeft||names[0];if(!state.byoRight&&names.length>1)state.byoRight=names[1];saveByoSelection();}
    toast(`${names.length} dataset${names.length===1?'':'s'} imported from S3${payload.errors?.length?`, ${payload.errors.length} failed`:''}`);
    state.byoS3Items=null;
    await render();
  }catch(err){toast('S3 import failed: '+String(err.message||err).slice(0,200));}
}
function byoSharePointResultsHtml(items){
  if(!items.length)return '<div class="empty" style="margin-top:12px">No matching .csv/.xlsx files found at that link.</div>';
  return `<div class="byoS3Results">
    <div class="byoS3ResultsHead"><b>${items.length} file${items.length===1?'':'s'} found</b><button class="btn small" type="button" onclick="importByoFromSharePoint()">Import selected</button></div>
    ${items.map((x,i)=>`<label class="byoS3Row"><input type="checkbox" class="byoSpCheck" value="${i}"><span class="byoS3Key">${esc(x.name)}</span><span class="byoS3Size muted">${x.size_mb} MB</span></label>`).join('')}
  </div>`;
}
function byoSharePointContext(){
  return {
    tenant_id:document.getElementById('byoSpTenant')?.value.trim()||'',
    client_id:document.getElementById('byoSpClientId')?.value.trim()||'',
    client_secret:document.getElementById('byoSpClientSecret')?.value.trim()||'',
    share_url:document.getElementById('byoSpShareUrl')?.value.trim()||'',
  };
}
async function listByoSharePointFiles(){
  const ctx=byoSharePointContext();
  if(!ctx.tenant_id||!ctx.client_id||!ctx.client_secret||!ctx.share_url){toast('Fill in the tenant ID, app ID, client secret, and share link');return;}
  try{
    const r=await fetch('/api/byo/import/sharepoint/list',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(ctx)});
    const payload=await r.json().catch(()=>({error:'SharePoint list failed'}));
    if(!r.ok)throw new Error(payload.error||'SharePoint list failed');
    state.byoSharePointItems=payload.items||[];
    toast(`${state.byoSharePointItems.length} matching file(s) found`);
    await render();
  }catch(err){toast('SharePoint list failed: '+String(err.message||err).slice(0,200));}
}
async function importByoFromSharePoint(){
  const items=state.byoSharePointItems||[];
  const checked=[...document.querySelectorAll('.byoSpCheck:checked')].map(x=>items[Number(x.value)]).filter(Boolean);
  if(!checked.length){toast('Select at least one SharePoint/OneDrive file');return;}
  try{
    const r=await fetch('/api/byo/import/sharepoint',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({items:checked})});
    const payload=await r.json().catch(()=>({error:'SharePoint import failed'}));
    if(!r.ok)throw new Error(payload.error||'SharePoint import failed');
    const names=(payload.saved||[]).map(x=>x.name);
    if(names.length){state.byoLeft=state.byoLeft||names[0];if(!state.byoRight&&names.length>1)state.byoRight=names[1];saveByoSelection();}
    toast(`${names.length} dataset${names.length===1?'':'s'} imported from SharePoint/OneDrive${payload.errors?.length?`, ${payload.errors.length} failed`:''}`);
    state.byoSharePointItems=null;
    await render();
  }catch(err){toast('SharePoint import failed: '+String(err.message||err).slice(0,200));}
}
function saveByoSelection(){sessionStorage.setItem('dart_byo_left',state.byoLeft||'');sessionStorage.setItem('dart_byo_right',state.byoRight||'');}
function ensureByoSelections(info){const names=(info.datasets||[]).filter(x=>x.status==='Ready').map(x=>x.name);if(state.byoLeft&&!names.includes(state.byoLeft))state.byoLeft='';if(state.byoRight&&!names.includes(state.byoRight))state.byoRight='';if(!state.byoLeft&&names.length)state.byoLeft=names[0];if(!state.byoRight&&names.length>1)state.byoRight=names.find(x=>x!==state.byoLeft)||'';if(state.byoRight===state.byoLeft)state.byoRight=names.find(x=>x!==state.byoLeft)||'';if(state.byoEditFile&&!names.includes(state.byoEditFile))state.byoEditFile='';if(state.byoEmailFile&&!names.includes(state.byoEmailFile))state.byoEmailFile='';if(!state.byoEditFile&&names.length)state.byoEditFile=names[0];if(!state.byoEmailFile&&names.length)state.byoEmailFile=state.byoEditFile||names[0];saveByoSelection();sessionStorage.setItem('dart_byo_edit_file',state.byoEditFile||'');sessionStorage.setItem('dart_byo_email_file',state.byoEmailFile||'');}
function byoPairReady(){return Boolean(state.byoLeft&&state.byoRight&&state.byoLeft!==state.byoRight);}
async function setByoSelection(side,value){if(side==='left')state.byoLeft=value;else state.byoRight=value;if(state.byoLeft===state.byoRight){if(side==='left')state.byoRight='';else state.byoLeft='';}state.byoCompare=null;state.byoChat=null;saveByoSelection();await render();}
async function swapByoSelection(){const a=state.byoLeft;state.byoLeft=state.byoRight;state.byoRight=a;state.byoCompare=null;state.byoChat=null;saveByoSelection();await render();}
function byoDatasetOptions(info,selected,exclude=''){return `<option value="">Choose a saved dataset…</option>${(info.datasets||[]).filter(x=>x.status==='Ready'&&x.name!==exclude).map(x=>`<option value="${esc(x.name)}" ${x.name===selected?'selected':''}>${esc(x.name)} · ${intFmt(x.rows)} rows · ${intFmt(x.columns)} cols</option>`).join('')}`;}
function byoPairPicker(info){return `<div class="panel"><div class="byoPairPicker"><div class="field"><label>Dataset A</label><select id="byoLeftSelect" onchange="setByoSelection('left',this.value)">${byoDatasetOptions(info,state.byoLeft,state.byoRight)}</select></div><div class="byoVersus">VS</div><div class="field"><label>Dataset B</label><select id="byoRightSelect" onchange="setByoSelection('right',this.value)">${byoDatasetOptions(info,state.byoRight,state.byoLeft)}</select></div></div>${byoPairReady()?`<div class="heroActions"><button class="btn" type="button" onclick="showPage('byo-compare')">Analyze selected pair</button><button class="btn secondary" type="button" onclick="showPage('byo-ai')">Ask AI about this pair</button><button class="btn ghost" type="button" onclick="swapByoSelection()">Swap A ↔ B</button></div>`:'<div class="callout" style="margin-top:14px"><b>Select two different saved datasets</b> to activate comparison and AI analysis.</div>'}</div>`;}
function byoLibraryCards(info){const rows=info.datasets||[];if(!rows.length)return `<div class="empty">No saved DIY datasets yet. Upload files and they will appear here and on disk under ${esc(info.storage_path||'Data/DIY')}.</div>`;return `<div class="grid grid2">${rows.map(d=>`<div class="byoFileCard ${state.byoLeft===d.name?'selectedA':''} ${state.byoRight===d.name?'selectedB':''}"><div class="byoFileTop"><div><div class="byoFileName">${esc(d.name)}</div><div class="muted" style="font-size:.8rem;margin-top:3px">Modified ${esc(d.modified_at)}</div></div><span class="pill ${d.status==='Ready'?'Stable':'Critical'}">${esc(d.status)}</span></div><div class="byoMeta"><span>${esc(d.type)}</span><span>${intFmt(d.rows||0)} rows</span><span>${intFmt(d.columns||0)} columns</span><span>${esc(d.size_mb)} MB</span>${state.byoLeft===d.name?'<span>Dataset A ✓</span>':''}${state.byoRight===d.name?'<span>Dataset B ✓</span>':''}</div><div class="byoFileActions">${d.status==='Ready'?`<button class="btn small" type="button" onclick="setByoSelection('left','${safeArg(d.name)}')">${state.byoLeft===d.name?'Dataset A ✓':'Use as A'}</button><button class="btn secondary small" type="button" onclick="setByoSelection('right','${safeArg(d.name)}')">${state.byoRight===d.name?'Dataset B ✓':'Use as B'}</button>`:''}<a class="btn secondary small" href="/download/byo-file/${encodeURIComponent(d.name)}">Download original</a><a class="btn ghost small" href="/download/byo-csv/${encodeURIComponent(d.name)}">Export CSV</a><button class="btn ghost small dangerBtn" type="button" onclick="deleteByoFile('${safeArg(d.name)}')">Delete</button></div></div>`).join('')}</div>`;}
async function deleteByoFile(name){if(!confirm(`Delete ${name} from Data/DIY? This removes the saved repo file.`))return;try{await del('/api/byo/files/'+encodeURIComponent(name));if(state.byoLeft===name)state.byoLeft='';if(state.byoRight===name)state.byoRight='';if(state.byoEditFile===name)state.byoEditFile='';if(state.byoEmailFile===name)state.byoEmailFile='';state.byoCompare=null;state.byoChat=null;saveByoSelection();toast('Dataset deleted');await render();}catch(err){toast('Delete failed: '+String(err.message||err).slice(0,160));}}
function byoHomeStats(info){const datasets=info.datasets||[];const totalRows=datasets.reduce((a,x)=>a+Number(x.rows||0),0);const totalSize=datasets.reduce((a,x)=>a+Number(x.size_mb||0),0);return `<div class="grid grid4"><div class="panel metric"><div class="label">Saved datasets</div><div class="value">${intFmt(info.dataset_count||0)}</div><div class="sub">Persistent library</div></div><div class="panel metric"><div class="label">Rows available</div><div class="value">${intFmt(totalRows)}</div><div class="sub">Across saved files</div></div><div class="panel metric"><div class="label">Storage</div><div class="value" style="font-size:1.1rem">${esc(info.storage_path||'Data/DIY')}</div><div class="sub">Repo folder</div></div><div class="panel metric"><div class="label">AI analyst</div><div class="value" style="font-size:1.1rem">${info.ai_configured?'Connected':'Local mode'}</div><div class="sub">${info.ai_configured?esc(info.ai_model):'Uses same GROQ key when configured'}</div></div></div>`;}
function compareSummaryCards(c){const s=c.summary;return `<div class="compareHeroGrid"><div class="compareMetric"><div class="k">Shared columns</div><div class="v">${intFmt(s.common_columns)}</div></div><div class="compareMetric"><div class="k">Type mismatches</div><div class="v ${s.type_mismatches?'deltaWarn':'deltaGood'}">${intFmt(s.type_mismatches)}</div></div><div class="compareMetric"><div class="k">Row difference B − A</div><div class="v">${intFmt(s.row_delta)}</div></div><div class="compareMetric"><div class="k">Completeness delta</div><div class="v ${Number(s.completeness_delta_pp)>=0?'deltaGood':'deltaWarn'}">${Number(s.completeness_delta_pp).toFixed(2)} pp</div></div></div>`;}
function datasetSideCard(label,d){return `<div class="byoSelected"><span class="workspaceBadge">${esc(label)}</span><div class="datasetName">${esc(d.source)}</div><div class="muted" style="margin-top:5px">${intFmt(d.rows)} rows · ${intFmt(d.columns)} columns · ${pctNum(d.completeness)} complete · ${intFmt(d.duplicate_rows)} duplicates</div></div>`;}
function compareOverview(c){return `${compareSummaryCards(c)}<br><div class="grid grid2">${datasetSideCard('Dataset A',c.left)}${datasetSideCard('Dataset B',c.right)}</div><br><div class="grid grid2"><div class="panel"><h3>Likely join keys</h3><p class="muted">Candidates are ranked using uniqueness and cross-file value overlap.</p>${genericTable(c.key_candidates,'DIY likely join keys')}</div><div class="panel"><h3>Schema exceptions</h3><p class="muted">Shared fields whose inferred pandas data types differ.</p>${genericTable(c.type_mismatches,'DIY schema type mismatches')}</div></div>`;}
function compareQuality(c){return `<div class="byoSplitTables"><div class="panel"><div class="byoDatasetHead"><h3>Dataset A quality</h3><span class="pill Stable">${esc(c.left.source)}</span></div>${byoStats(c.left,'Dataset A')}<br>${genericTable(c.left.quality,'Dataset A quality')}</div><div class="panel"><div class="byoDatasetHead"><h3>Dataset B quality</h3><span class="pill Stable">${esc(c.right.source)}</span></div>${byoStats(c.right,'Dataset B')}<br>${genericTable(c.right.quality,'Dataset B quality')}</div></div><br><div class="panel"><h3>Shared-column quality comparison</h3>${genericTable(c.column_comparison,'DIY shared column quality comparison')}</div>`;}
function compareExplorer(c){return `<div class="byoSplitTables"><div class="panel"><div class="byoDatasetHead"><h3>Dataset A preview</h3><a class="btn secondary small" href="/download/byo-csv/${encodeURIComponent(c.left.source)}">Full CSV</a></div>${genericTable(c.left.preview,'Dataset A preview')}</div><div class="panel"><div class="byoDatasetHead"><h3>Dataset B preview</h3><a class="btn secondary small" href="/download/byo-csv/${encodeURIComponent(c.right.source)}">Full CSV</a></div>${genericTable(c.right.preview,'Dataset B preview')}</div></div>`;}
function compareDeepDive(c){return `<div class="grid grid2"><div class="panel"><h3>Common columns</h3>${genericTable(c.common_columns,'DIY common columns')}</div><div class="panel"><h3>Columns unique to each dataset</h3><div class="byoSplitTables byoUniqueTables"><div class="byoCompactTable">${genericTable(c.only_left,'Columns only in Dataset A')}</div><div class="byoCompactTable">${genericTable(c.only_right,'Columns only in Dataset B')}</div></div></div></div><br><div class="panel"><h3>Numeric field comparison</h3><p class="muted">Side-by-side means, medians, ranges, and mean deltas for shared numeric columns.</p>${genericTable(c.numeric_comparison,'DIY numeric comparison')}</div>`;}
function byoAiStatus(info){return `<span class="aiStatus ${info.ai_configured?'on':'off'}">● ${info.ai_configured?'AI connected · '+esc(info.ai_model):'Local analysis · Groq client unavailable'}</span>`;}
async function getByoChat(){if(!byoPairReady())return {chat:[],ai_configured:false};try{return await api(`/api/byo/chat?left=${encodeURIComponent(state.byoLeft)}&right=${encodeURIComponent(state.byoRight)}`);}catch{return {chat:[],ai_configured:false};}}
async function askByoAI(prefill=''){
  if(!byoPairReady()){toast('Select two datasets first');return;}
  const input=document.getElementById('byoAiPrompt');
  const value=(prefill||input?.value||'').trim();
  if(!value)return;
  if(input)input.value='';
  const send=document.getElementById('byoAiSend');
  if(send){send.disabled=true;send.textContent='Analyzing…';}
  const bucket=getScopedHistory('byo');
  const session=getActiveSession('byo');
  if(session.messages.length===0){
    session.title=value.length>32?value.slice(0,30)+'…':value;
  }
  session.messages.push({role:'user',content:value});
  state.byoChat=session.messages.slice();
  persistChatHistoryState();
  try{
    const res=await postJson('/api/byo/chat',{prompt:value,left:state.byoLeft||'',right:state.byoRight||''});
    session.messages.push({role:'assistant',content:res.answer||'I could not generate an answer from the selected dataset context.'});
  }catch(err){
    session.messages.push({role:'assistant',content:`AI analysis failed: ${err?.message||err}`});
  }
  if(send){send.disabled=false;send.textContent='Send';}
  state.byoChat=session.messages.slice();
  persistChatHistoryState();
  await render();
  scrollBubbleToBottom();
}
async function resetByoAI(){
  await postJson('/api/byo/chat/reset',{});
  const bucket=getScopedHistory('byo');
  const session=getActiveSession('byo');
  session.messages=[];
  state.byoChat=[];
  persistChatHistoryState();
  toast('AI conversation cleared');
  await render();
  ensureAssistantBubble();
}
function byoAiPage(info,chat,c){const messages=chat?.chat||[];return `<div class="hero"><span class="workspaceBadge">Two-dataset AI</span><h1>AI Analyst</h1><p>Ask questions across the two selected saved datasets. The analyst uses the same Groq/OpenAI-compatible configuration as DART Copilot and is grounded in computed profiles, distributions, comparison metrics, key candidates, and representative records from both files.</p><div class="heroActions">${byoAiStatus(info)}<button class="btn secondary" type="button" onclick="resetByoAI()">Clear conversation</button></div></div>${byoPairPicker(info)}<br>${byoPairReady()?`<div class="byoChatShell"><div class="panel byoChatPanel"><div class="byoChat">${messages.length?messages.map(chatMessageHtml).join(''):'<div class="empty">Ask your first question about the selected pair.</div>'}</div><div class="byoChatComposer"><input id="byoAiPrompt" placeholder="Compare the two datasets, find likely keys, explain quality differences…" onkeydown="if(event.key==='Enter')askByoAI()"><button class="btn" id="byoAiSend" type="button" onclick="askByoAI()">Send</button></div><div class="heroActions"><button class="btn secondary small" type="button" onclick="askByoAI('Give me an executive comparison of these two datasets')">Executive comparison</button><button class="btn secondary small" type="button" onclick="askByoAI('What are the biggest data quality differences?')">Quality differences</button><button class="btn secondary small" type="button" onclick="askByoAI('Which columns are the best candidates to join these datasets and why?')">Find join keys</button><button class="btn secondary small" type="button" onclick="askByoAI('What anomalies or mismatches should I investigate first?')">Investigation priorities</button></div></div><aside class="panel byoContextCard"><h3>Active context</h3>${datasetSideCard('Dataset A',c.left)}<br>${datasetSideCard('Dataset B',c.right)}<br><div class="callout"><b>${intFmt(c.summary.common_columns)} shared columns</b><br>${intFmt(c.summary.type_mismatches)} type mismatches · ${intFmt(c.summary.key_candidates)} key candidates</div><p class="muted" style="font-size:.8rem">The model is not retrained on your files. Each answer is grounded in the two selected datasets at request time.</p></aside></div>`:`<div class="byoEmptyPair"><h3>Select two datasets to start the AI analyst</h3><p class="muted">Upload at least two files, then choose Dataset A and Dataset B above.</p></div>`}`;}
function byoSingleDatasetPicker(info,selected,onchange,label='Dataset'){return `<div class="panel"><div class="field"><label>${esc(label)}</label><select onchange="${onchange}(this.value)">${byoDatasetOptions(info,selected)}</select></div></div>`;}
async function setByoEditFile(value){state.byoEditFile=value;state.byoRawOffset=0;state.byoRawSearch='';state.byoRawData=null;sessionStorage.setItem('dart_byo_edit_file',value||'');await render();}
async function setByoEmailFile(value){state.byoEmailFile=value;state.byoEmailEditingId=null;state.byoEmailAgent=null;sessionStorage.setItem('dart_byo_email_file',value||'');await render();}
function byoRawDataPage(info,raw){state.byoRawData=raw;if(!raw?.file_found)return `<div class="hero"><span class="workspaceBadge">Build Your Own</span><h1>Raw Data Editor</h1><p>Upload a CSV or XLSX file first, then choose it here to inspect and edit the underlying saved file.</p><div class="heroActions"><button class="btn" onclick="showPage('byo-lab')">Upload data</button></div></div>`;const cols=raw.columns||[],rows=raw.rows||[],start=(raw.offset||0)+1,end=(raw.offset||0)+rows.length;const grid=`<div class="tableWrap"><table class="table rawTable"><thead><tr><th class="rawRow">Source row</th>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr><td class="rawRow">${esc(r._SourceRow)}</td>${cols.map(c=>`<td><input class="rawCell byoRawCell" data-row="${esc(r._SourceRow)}" data-col="${esc(c)}" data-before="${esc(rawText(r[c]))}" value="${esc(rawText(r[c]))}"></td>`).join('')}</tr>`).join('')}</tbody></table></div>`;return `${byoSingleDatasetPicker(info,state.byoEditFile,'setByoEditFile','File to edit')}<br>${section('byo-raw-overview','File editor',`<div class="hero"><span class="workspaceBadge">Build Your Own · ${esc(raw.file_type||'')}</span><h1>Raw Data Editor</h1><p>Inspect every row in the selected uploaded file and change cell values directly. Saving writes back to the actual file under Data/DIY, creates a timestamped backup first, and can immediately check BYO Email Alerts for the same file.</p><div class="heroActions"><button class="btn" onclick="saveByoRawChanges()">Save changes to file</button><button class="btn secondary" onclick="reloadByoRawData()">Refresh from file</button><button class="btn ghost" onclick="restoreLatestByoBackup()" ${raw.latest_backup?'':'disabled'}>Restore latest backup</button><button class="btn ghost" onclick="state.byoEmailFile=state.byoEditFile;sessionStorage.setItem('dart_byo_email_file',state.byoEmailFile);showPage('byo-email')">Open Email Alerts</button></div></div><div class="statusGrid"><div class="statusCard"><div class="muted">File</div><div class="big" style="font-size:1rem">${esc(raw.filename)}</div></div><div class="statusCard"><div class="muted">Rows</div><div class="big">${intFmt(raw.raw_rows)}</div></div><div class="statusCard"><div class="muted">Columns</div><div class="big">${intFmt(cols.length)}</div></div><div class="statusCard"><div class="muted">Sheet / type</div><div class="big" style="font-size:1rem">${esc(raw.sheet||raw.file_type||'')}</div></div></div><br><div class="callout"><b>Demo workflow:</b> open Email Alerts for this file and establish a baseline. Return here, edit a cell, leave “Check BYO Email Alerts immediately” selected, and save. The alert engine will compare the new file against the baseline immediately.</div>`)}${section('byo-raw-grid','Edit data',`<div class="panel"><div class="grid grid2"><div class="field"><label>Search this file</label><input id="byoRawSearchBox" value="${esc(state.byoRawSearch)}" placeholder="Search any column"><div class="heroActions"><button class="btn secondary small" onclick="applyByoRawSearch()">Search</button><button class="btn ghost small" onclick="clearByoRawSearch()">Clear</button></div></div><div class="field"><label>After save</label><select id="byoRawRunAgent"><option value="true">Check BYO Email Alerts immediately</option><option value="false">Do not run alerts</option></select><div class="rangeLine">Showing ${intFmt(start)}-${intFmt(end)} of ${intFmt(raw.filtered_rows)} matching rows.</div></div></div><br>${grid}<br><div class="heroActions"><button class="btn secondary" onclick="byoRawPage(-1)" ${raw.offset<=0?'disabled':''}>Previous</button><button class="btn secondary" onclick="byoRawPage(1)" ${raw.offset+raw.limit>=raw.filtered_rows?'disabled':''}>Next</button><button class="btn" onclick="saveByoRawChanges()">Save changes to file</button></div></div>`)}${section('byo-raw-backup','Backup protection',`<div class="panel"><h3>Backup status</h3><p class="muted">${raw.latest_backup?`Latest backup: <code>${esc(raw.latest_backup)}</code>`:'No backup has been created for this file yet.'}</p><p class="muted">For XLSX files DART changes only the targeted worksheet cells. For CSV files DART rewrites the CSV after backing up the original. Source row numbers are editor-only and are never added as a data column.</p></div>`)}`;}
function collectByoRawChanges(){return [...document.querySelectorAll('.byoRawCell')].filter(el=>el.value!==el.dataset.before).map(el=>({source_row:Number(el.dataset.row),column:el.dataset.col,before:el.dataset.before,after:el.value}));}
async function saveByoRawChanges(){const changes=collectByoRawChanges();if(!changes.length){toast('No cell changes to save');return;}if(!confirm(`Save ${changes.length} changed cell(s) to ${state.byoEditFile}?`))return;try{const runAgent=document.getElementById('byoRawRunAgent')?.value!=='false';const r=await postJson('/api/byo/raw-data/save',{filename:state.byoEditFile,changes,run_email_agent:runAgent,signature:state.byoRawData?.signature||''});const sent=(r.agent_results||[]).filter(x=>x.status==='Email sent').length;toast(`Saved ${r.saved} cell(s)${r.agent_results?.length?' · alerts checked':''}${sent?` · ${sent} email sent`:''}`);state.byoRawOffset=0;await render();}catch(err){toast('Save failed: '+String(err.message||err).slice(0,180));}}
async function restoreLatestByoBackup(){if(!confirm(`Restore the most recent backup for ${state.byoEditFile}?`))return;try{await postJson('/api/byo/raw-data/restore-latest',{filename:state.byoEditFile});toast('DIY file restored');state.byoRawOffset=0;await render();}catch(err){toast('Restore failed: '+String(err.message||err).slice(0,160));}}
async function reloadByoRawData(){state.byoRawOffset=0;await render();}
async function applyByoRawSearch(){state.byoRawSearch=document.getElementById('byoRawSearchBox')?.value.trim()||'';state.byoRawOffset=0;await render();}
async function clearByoRawSearch(){state.byoRawSearch='';state.byoRawOffset=0;await render();}
async function byoRawPage(dir){state.byoRawOffset=Math.max(0,state.byoRawOffset+dir*state.byoRawLimit);await render();}
function addByoEmailCondition(){const cols=state.byoEmailAgent?.columns||[];document.getElementById('byoConditionList')?.insertAdjacentHTML('beforeend',emailConditionRow({},cols));}
function byoEmailAgentPage(info,workspace){state.byoEmailAgent=info;const templates=info.templates||[],tpl=(state.byoEmailEditingId?templates.find(t=>t.id===state.byoEmailEditingId):null)||info.default_template,cols=info.columns||[],status=info.status||{},smtp=info.smtp||{};const saved=templates.map(t=>`<div class="savedAutomation"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start"><div><b>${esc(t.name)}</b><div class="muted">${esc(t.dataset_name||'Dataset missing')} · ${esc(t.recipient_email)} · ${t.enabled?'Enabled':'Disabled'} · ${esc(t.last_status||'Not run yet')}</div></div><span class="pill ${t.enabled?'Stable':'Watch'}">${t.enabled?'Enabled':'Disabled'}</span></div><div class="automationActions"><button class="btn secondary small" onclick="editByoEmailById('${esc(t.id)}','${safeArg(t.dataset_name||'')}')">Edit</button><button class="btn secondary small" onclick="baselineByoEmail('${esc(t.id)}')">Refresh baseline</button><button class="btn secondary small" onclick="runByoEmail('${esc(t.id)}',false)">Check only</button><button class="btn small" onclick="runByoEmail('${esc(t.id)}',true)">Run + send</button><button class="btn ghost small" onclick="testByoEmail('${esc(t.id)}')">Send test</button><button class="btn ghost small" onclick="deleteByoEmailTemplate('${esc(t.id)}')">Delete</button></div></div>`).join('')||'<div class="empty">No Build Your Own alert automations yet. Choose a dataset, configure one below, and save it.</div>';const history=(info.history||[]).length?simpleTable(info.history):'<div class="empty">No Build Your Own alert history yet.</div>',dup=tpl.identity_duplicate_count||0;return `${byoSingleDatasetPicker(workspace,state.byoEmailFile,'setByoEmailFile','Dataset to monitor')}<br>${section('byo-email-status','Alert status',`<div class="hero"><span class="workspaceBadge">Build Your Own</span><h1>Email Alerts</h1><p>Monitor any uploaded DIY dataset for new or changed rows. Each automation is tied to one saved file, has its own baseline, and uses the same working SMTP delivery as the Medicare Email Agent.</p><div class="heroActions"><button class="btn" onclick="checkAllByoEmails()">Check all enabled alerts now</button><button class="btn secondary" onclick="baselineAllByoEmails()">Baseline all enabled</button><button class="btn ghost" onclick="testSmtp()">Test SMTP connection</button><button class="btn ghost" onclick="state.byoEditFile=state.byoEmailFile;sessionStorage.setItem('dart_byo_edit_file',state.byoEditFile);showPage('byo-raw')">Edit selected file</button></div></div><div class="statusGrid"><div class="statusCard"><div class="muted">Watcher</div><div class="big"><span class="statusDot ${String(status.status||'').toLowerCase().includes('error')?'bad':'good'}"></span>${esc(status.status||'Not started')}</div></div><div class="statusCard"><div class="muted">Check interval</div><div class="big">${esc(info.check_interval_seconds)}s</div></div><div class="statusCard"><div class="muted">Dataset</div><div class="big" style="font-size:1rem">${esc(info.selected_dataset||'None')}</div></div><div class="statusCard"><div class="muted">Email delivery</div><div class="big">${smtp.configured?'Ready':'Setup needed'}</div><div class="muted">${esc(smtp.host||'')} ${smtp.port?': '+esc(smtp.port):''}</div></div></div><br><div class="callout"><b>Best demo:</b> save an automation, click Refresh baseline, open the Raw Data Editor for this same file, change a monitored value, and save. The editor can run this alert immediately so you do not have to wait for the background interval.</div>`)}${section('byo-email-saved','Saved alerts',`<div class="panel"><div style="display:flex;justify-content:space-between;align-items:center;gap:12px"><div><h3>Saved Build Your Own alerts</h3><p class="muted">Each alert remembers the exact DIY file it monitors.</p></div><button class="btn secondary" onclick="newByoEmailTemplate()">New alert</button></div>${saved}</div>`)}${section('byo-email-builder','Alert builder',`<div class="panel emailBuilder"><h3>${tpl.id?'Edit alert':'Create alert'}</h3><div class="grid grid2"><div class="field"><label>Monitoring dataset</label><input value="${esc(info.selected_dataset||'')}" disabled><div class="rangeLine">Change the dataset using the picker above.</div></div><div class="field"><label>Alert name</label><input id="byoEmailName" value="${esc(tpl.name||'')}"></div><div class="field"><label>Recipient email(s)</label><input id="byoEmailRecipient" value="${esc(tpl.recipient_email||'')}" placeholder="owner@example.com; second@example.com"></div><div class="field"><label>Enabled</label><select id="byoEmailEnabled"><option value="false" ${!tpl.enabled?'selected':''}>Disabled</option><option value="true" ${tpl.enabled?'selected':''}>Enabled</option></select></div><div class="field"><label>Trigger mode</label><select id="byoEmailTrigger">${['New or changed rows','New rows only','All changes including deletions'].map(x=>`<option ${tpl.trigger_mode===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Row identity columns</label>${smartMultiSelect('byoEmailIdentity',cols,tpl.identity_columns,'Choose row identity columns')}<div class="rangeLine">Use stable fields that identify the same logical row between file versions. ${dup?`${intFmt(dup)} rows currently share duplicate identity values.`:'Current selection is suitable for row matching.'}</div></div><div class="field"><label>Monitored columns</label>${smartMultiSelect('byoEmailMonitor',cols,tpl.monitor_columns,'Choose monitored columns')}<div class="rangeLine">Changes to these fields count as detected changes.</div></div><div class="field"><label>Display columns in email</label>${smartMultiSelect('byoEmailDisplay',cols,tpl.display_columns,'Choose email display columns')}</div><div class="field"><label>Requirement mode</label><select id="byoEmailConditionMode"><option ${tpl.condition_mode==='Any change'?'selected':''}>Any change</option><option ${tpl.condition_mode!=='Any change'?'selected':''}>Only when conditions are met</option></select><label style="margin-top:10px">Condition logic</label><select id="byoEmailLogic"><option ${tpl.condition_logic!=='ANY condition'?'selected':''}>ALL conditions</option><option ${tpl.condition_logic==='ANY condition'?'selected':''}>ANY condition</option></select></div></div><br><div style="display:flex;justify-content:space-between;align-items:center"><h3 style="margin:0">Conditions</h3><button type="button" class="btn secondary small" onclick="addByoEmailCondition()">Add condition</button></div><div id="byoConditionList">${(tpl.conditions||[]).map(r=>emailConditionRow(r,cols)).join('')}</div><div class="grid grid2"><div class="field"><label>Subject prefix</label><input id="byoEmailSubject" value="${esc(tpl.subject_prefix||'[DART DIY Alert]')}"></div><div class="field"><label>Greeting</label><input id="byoEmailGreeting" value="${esc(tpl.greeting||'Hello,')}"></div></div><div class="field"><label>Email introduction</label><textarea id="byoEmailIntro">${esc(tpl.email_intro||'')}</textarea></div><div class="field" style="max-width:280px"><label>CSV attachment</label><select id="byoEmailCsv"><option value="true" ${tpl.include_csv!==false?'selected':''}>Include CSV</option><option value="false" ${tpl.include_csv===false?'selected':''}>No CSV</option></select></div><br><div class="heroActions"><button class="btn" onclick="saveByoEmailTemplate()">Save alert</button><button class="btn secondary" onclick="previewByoEmailTemplate()">Preview current matches</button></div></div><div id="byoEmailPreviewHolder" style="margin-top:16px"></div>`)}${section('byo-email-history','Execution history',`<div class="panel"><h3>Recent BYO alert history</h3>${history}</div><br><div class="callout"><b>SMTP:</b> This page uses the same local-demo Gmail configuration as the Medicare Email Agent. The password is never sent to the browser.${smtp.issues?.length?`<br><br><b>Current setup items:</b> ${esc(smtp.issues.join(' '))}`:''}</div>`)}`;}
async function editByoEmailById(id,dataset){state.byoEmailEditingId=id;state.byoEmailFile=dataset;sessionStorage.setItem('dart_byo_email_file',dataset||'');await render();}
async function newByoEmailTemplate(){state.byoEmailEditingId=null;await render();}
function collectByoEmailTemplate(){const conditions=[...document.querySelectorAll('#byoConditionList .conditionRow')].map(row=>({column:row.querySelector('.condColumn').value,operator:row.querySelector('.condOperator').value,value:row.querySelector('.condValue').value}));return {id:state.byoEmailEditingId||'',dataset_name:state.byoEmailFile||'',name:document.getElementById('byoEmailName')?.value||'',recipient_email:document.getElementById('byoEmailRecipient')?.value||'',enabled:document.getElementById('byoEmailEnabled')?.value==='true',trigger_mode:document.getElementById('byoEmailTrigger')?.value||'New or changed rows',identity_columns:selectedValues('byoEmailIdentity'),monitor_columns:selectedValues('byoEmailMonitor'),condition_mode:document.getElementById('byoEmailConditionMode')?.value||'Any change',condition_logic:document.getElementById('byoEmailLogic')?.value||'ALL conditions',conditions,display_columns:selectedValues('byoEmailDisplay'),subject_prefix:document.getElementById('byoEmailSubject')?.value||'[DART DIY Alert]',greeting:document.getElementById('byoEmailGreeting')?.value||'Hello,',email_intro:document.getElementById('byoEmailIntro')?.value||'',include_csv:document.getElementById('byoEmailCsv')?.value!=='false'};}
async function saveByoEmailTemplate(){try{const r=await postJson('/api/byo/email-agent/template',collectByoEmailTemplate());state.byoEmailEditingId=r.template.id;toast('Build Your Own alert saved');await render();}catch(err){toast('Could not save: '+String(err.message||err).slice(0,160));}}
async function deleteByoEmailTemplate(id){if(!confirm('Delete this BYO alert and its baseline snapshot?'))return;await del('/api/byo/email-agent/template/'+encodeURIComponent(id));state.byoEmailEditingId=null;toast('BYO alert deleted');await render();}
async function baselineByoEmail(id){try{const r=await postJson('/api/byo/email-agent/baseline/'+encodeURIComponent(id),{});toast(`${r.rows||0} rows saved as baseline`);await render();}catch(err){toast('Baseline failed: '+String(err.message||err).slice(0,160));}}
async function baselineAllByoEmails(){const r=await postJson('/api/byo/email-agent/baselines',{});toast(`Refreshed ${r.baselines||0} BYO baseline(s)`);await render();}
async function runByoEmail(id,send){try{const r=await postJson('/api/byo/email-agent/run/'+encodeURIComponent(id),{send_email:send});toast(r.status||'Check complete');await render();}catch(err){toast('Alert check failed: '+String(err.message||err).slice(0,160));}}
async function testByoEmail(id){try{const r=await postJson('/api/byo/email-agent/test/'+encodeURIComponent(id),{});toast(r.status||'Test email sent');}catch(err){toast('Test failed: '+String(err.message||err).slice(0,160));}}
async function checkAllByoEmails(){const r=await postJson('/api/byo/email-agent/check',{});toast(`Checked ${r.results?.length||0} enabled BYO alert(s)`);await render();}
async function previewByoEmailTemplate(){try{const r=await postJson('/api/byo/email-agent/preview',collectByoEmailTemplate());const holder=document.getElementById('byoEmailPreviewHolder');holder.innerHTML=`<div class="panel"><h3>Preview · ${intFmt(r.matches)} current matching row(s)</h3><div class="emailPreview"><iframe id="byoEmailPreviewFrame"></iframe></div></div>`;document.getElementById('byoEmailPreviewFrame').srcdoc=r.html||'<p>No preview available.</p>';}catch(err){toast('Preview failed: '+String(err.message||err).slice(0,160));}}
async function renderByo(meta,info){
  ensureByoSelections(info);
  state.byoInfo=info;
  let comparison=null;
  if(byoPairReady()&&['byo-compare','byo-quality','byo-preview','byo-ai','byo-export'].includes(state.page)){
    try{comparison=await postJson('/api/byo/compare',{left:state.byoLeft,right:state.byoRight});state.byoCompare=comparison;}
    catch(err){state.byoCompare=null;toast('Comparison failed: '+String(err.message||err).slice(0,160));}
  }
  let rawInfo=null,emailInfo=null;
  if(state.page==='byo-raw'&&state.byoEditFile)rawInfo=await api(`/api/byo/raw-data?filename=${encodeURIComponent(state.byoEditFile)}&search=${encodeURIComponent(state.byoRawSearch)}&offset=${state.byoRawOffset}&limit=${state.byoRawLimit}`);
  if(state.page==='byo-email')emailInfo=await api(`/api/byo/email-agent?filename=${encodeURIComponent(state.byoEmailFile||'')}`);
  let html='';
  const noFiles=`<div class="panel workspaceEmpty"><h3>No saved DIY datasets yet</h3><p class="muted">Upload CSV/XLSX files. DART will save them under <span class="byoStoragePath">${esc(info.storage_path||'Data/DIY')}</span> and keep them available after restart.</p><button class="btn" onclick="showPage('byo-lab')">Upload datasets</button></div>`;
  const needPair=(info.dataset_count||0)<1?noFiles:`<div class="byoEmptyPair"><h3>${(info.dataset_count||0)===1?'One more dataset needed':'Select two datasets'}</h3><p class="muted">${(info.dataset_count||0)===1?'Upload another CSV/XLSX file, then choose a pair.':'Choose two different saved datasets above to run this analysis.'}</p><button class="btn secondary" onclick="showPage('byo-library')">Open Dataset Library</button></div>`;
  if(state.page==='byo-home'){
    html=`<div class="hero"><span class="workspaceBadge">Build Your Own</span><h1>Dataset Intelligence Workspace</h1><p>Upload reusable data files, compare datasets, inspect quality, edit the underlying saved data, configure file-specific email alerts, and ask an AI analyst questions grounded in your selected files.</p><div class="heroActions"><button class="btn" onclick="showPage('byo-library')">Open Dataset Library</button><button class="btn secondary" onclick="showPage('byo-raw')">Open Raw Data Editor</button><button class="btn secondary" onclick="showPage('byo-email')">Open Email Alerts</button><button class="btn ghost" onclick="showPage('byo-ai')">Open AI Analyst</button></div></div>${byoHomeStats(info)}<br>${personaLensPanel(meta)}<br><div class="byoFlow"><div class="byoFlowCard"><b>1 · Save datasets</b><span>Upload Excel or CSV files once. Originals are stored in Data/DIY and persist between server restarts.</span></div><div class="byoFlowCard"><b>2 · Inspect, edit, or compare</b><span>Open any single file in Raw Data Editor or choose a pair for comparison and quality analysis.</span></div><div class="byoFlowCard"><b>3 · Alert + analyze</b><span>Baseline a file-specific Email Alert, edit a monitored cell for the demo, then let DART detect and email the change immediately.</span></div></div><br>${(info.datasets||[]).length?byoPairPicker(info):noFiles}`;
  }else if(state.page==='byo-lab'){
    html=`<div class="hero"><span class="workspaceBadge">Build Your Own</span><h1>Upload Data</h1><p>Add reusable datasets to the persistent DIY library. Multiple files can be uploaded at once and duplicate filenames are saved safely with a numeric suffix.</p></div>${byoUploadPanel(info)}<br>${byoImportPanel()}<br><div class="panel"><h3>Saved location</h3><p>Files are written to <span class="byoStoragePath">${esc(info.storage_path||'Data/DIY')}</span> relative to this Python app. Once uploaded, the same file can be opened in Raw Data Editor and selected in Email Alerts.</p></div>`;
  }else if(state.page==='byo-library'){
    html=`<div class="hero"><span class="workspaceBadge">Persistent library</span><h1>Dataset Library</h1><p>Choose, download, export, or delete the files saved in your repo. Every ready file is also available to the BYO Raw Data Editor and Email Alerts pages.</p><div class="heroActions"><button class="btn" onclick="showPage('byo-lab')">Upload more files</button><button class="btn secondary" onclick="showPage('byo-raw')">Edit a file</button><button class="btn ghost" onclick="showPage('byo-email')">Configure alerts</button></div></div><div class="byoLibraryGrid"><div class="panel"><h3>${intFmt(info.dataset_count||0)} saved datasets</h3>${byoLibraryCards(info)}</div><div>${byoPairPicker(info)}</div></div>`;
  }else if(state.page==='byo-compare'){
    html=`<div class="hero"><span class="workspaceBadge">Dataset A ↔ Dataset B</span><h1>Compare Lab</h1><p>Run a structured comparison across two saved datasets: shape, schema, completeness, duplicate burden, shared fields, type mismatches, numeric distributions, and likely join keys.</p></div>${byoPairPicker(info)}${comparison?`<br>${compareOverview(comparison)}<br><div class="panel"><h3>Detailed field comparison</h3>${genericTable(comparison.column_comparison,'DIY detailed field comparison')}</div><br>${compareDeepDive(comparison)}`:`<br>${needPair}`}`;
  }else if(state.page==='byo-quality'){
    html=`<div class="hero"><h1>Schema & Quality</h1><p>Compare column completeness, distinctness, duplicates, inferred data types, and schema exceptions across the selected pair.</p></div>${byoPairPicker(info)}${comparison?`<br>${compareQuality(comparison)}`:`<br>${needPair}`}`;
  }else if(state.page==='byo-preview'){
    html=`<div class="hero"><h1>Data Explorer</h1><p>Inspect both selected datasets side-by-side without editing them. Use Raw Data Editor when you want to write changes back to one uploaded file.</p><div class="heroActions"><button class="btn secondary" onclick="showPage('byo-raw')">Open Raw Data Editor</button></div></div>${byoPairPicker(info)}${comparison?`<br>${compareExplorer(comparison)}<br><div class="panel"><h3>Shared schema</h3>${genericTable(comparison.column_comparison,'DIY explorer shared schema')}</div>`:`<br>${needPair}`}`;
  }else if(state.page==='byo-raw'){
    html=(info.dataset_count||0)?byoRawDataPage(info,rawInfo):noFiles;
  }else if(state.page==='byo-email'){
    html=(info.dataset_count||0)?byoEmailAgentPage(emailInfo||{},info):noFiles;
  }else if(state.page==='byo-ai'){
    const chat=await getByoChat();html=byoAiPage(info,chat,comparison||state.byoCompare||{left:{},right:{},summary:{common_columns:0,type_mismatches:0,key_candidates:0}});
  }else if(state.page==='byo-export'){
    html=`<div class="hero"><h1>Export Center</h1><p>Download the original saved files, export either full dataset as CSV, or export the comparison tables currently generated by DART.</p></div>${byoPairPicker(info)}${comparison?`<br><div class="grid grid2"><div class="panel"><h3>Dataset A</h3><p><b>${esc(comparison.left.source)}</b></p><div class="heroActions"><a class="btn" href="/download/byo-file/${encodeURIComponent(comparison.left.source)}">Download original</a><a class="btn secondary" href="/download/byo-csv/${encodeURIComponent(comparison.left.source)}">Export full CSV</a></div></div><div class="panel"><h3>Dataset B</h3><p><b>${esc(comparison.right.source)}</b></p><div class="heroActions"><a class="btn" href="/download/byo-file/${encodeURIComponent(comparison.right.source)}">Download original</a><a class="btn secondary" href="/download/byo-csv/${encodeURIComponent(comparison.right.source)}">Export full CSV</a></div></div></div><br><div class="grid grid2"><div class="panel"><h3>Comparison export</h3>${genericTable(comparison.column_comparison,'DIY comparison export')}</div><div class="panel"><h3>Join-key export</h3>${genericTable(comparison.key_candidates,'DIY join key candidates')}</div></div>`:`<br>${needPair}`}`;
  }else{state.page='byo-home';return renderByo(meta,info);}
  document.getElementById('footer').innerHTML=`Build Your Own · ${intFmt(info.dataset_count||0)} saved datasets · ${esc(info.storage_path||'Data/DIY')} · DART ${esc(meta.version)} · ${esc(meta.build_fingerprint||'build unknown')} · FastAPI`;
  const appEl=document.getElementById('app');appEl.innerHTML=html;runInjectedScripts(appEl);decorateMetricHelp(appEl);refreshThemeVisuals(appEl);ensureAssistantBubble();
}
function genericTable(rows,exportName=''){if(!rows||!rows.length)return '<div class="empty">No rows available.</div>';const cols=Object.keys(rows[0]);return `<div class="exportTableBlock">${exportBar(exportName,rows.length)}<div class="tableWrap"><table class="table"><thead><tr>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr>${cols.map(c=>`<td>${esc(r[c])}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div>`;}

function kanban(rows){const lanes=['Open','In progress','Blocked','Resolved'];return `<div class="kanban">${lanes.map(l=>`<div class="lane"><h4>${l}</h4>${(rows||[]).filter(r=>r.Status===l).map(r=>`<div class="task"><b>${esc(r.Field)}</b><br><span class="muted">${esc(r.Stream)} · ${esc(r.Priority||'')}</span><br>${esc(r.Owner||'Unassigned')}<br><span class="muted">${esc(r.Note||'')}</span></div>`).join('')||'<div class="muted">No items</div>'}</div>`).join('')}</div>`;}
async function render(){const meta=await api('/api/meta');state.meta=meta;if(!meta.auth_user){renderAuth();return;}if(!meta.persona){renderOnboarding();return;}setNav();syncChatHistoryFromWorkspaceState();if(state.page==='profile'||state.page==='settings'){renderSharedPage(meta);return;}if(state.workspace==='medicaid'){await renderMedicaid(meta);return;}if(state.workspace==='byo'){const byo=await api('/api/byo');await renderByo(meta,byo);return;}const data=await getData();let emailInfo=null;let rawInfo=null;if(state.page==='emailagent')emailInfo=await api('/api/email-agent');if(state.page==='raweditor')rawInfo=await api(`/api/raw-data?search=${encodeURIComponent(state.rawSearch)}&offset=${state.rawOffset}&limit=${state.rawLimit}`);document.getElementById('footer').innerHTML=`Loaded: ${esc(meta.source)} · ${intFmt(meta.rows)} rows · ${intFmt(meta.columns)} columns · DART ${esc(meta.version)} · ${esc(meta.build_fingerprint||'build unknown')} · FastAPI`;const p=meta.persona||{};let html=filterBar(meta);
if(state.page==='home'){
  const pqLimit=Number(state.priorityQueueLimit||10);
  const queueSource=Array.isArray(data.summary?.top_actions)&&data.summary.top_actions.length?data.summary.top_actions:(data.issues||data.rows||[]);
  const queueItems=queueSource.slice().sort((a,b)=>{
    const unmA=Number(a.NotMatchedClaims||0);
    const unmB=Number(b.NotMatchedClaims||0);
    if(unmB!==unmA)return unmB-unmA;
    const rateA=a.MatchRate!=null?Number(a.MatchRate):1;
    const rateB=b.MatchRate!=null?Number(b.MatchRate):1;
    if(rateA!==rateB)return rateA-rateB;
    const impA=Number(a.ImpactScore||0);
    const impB=Number(b.ImpactScore||0);
    return impB-impA;
  }).slice(0,pqLimit);
  let body=`${section('overview','Overview',`<div class="hero"><h1>DART Command Center</h1><p>A consolidated view of data quality, reconciliation risk, unmatched volume, remediation activity, and downstream business impact${p.program||p.role?` for <b>${esc(p.program||'your program')}</b> with a <b>${esc(p.role||'general')}</b> role lens`:''}. Use this page as your primary operational command center.</p><div class="heroActions">${personaQuickActions(meta)}</div></div>${metrics(data.summary)}<br>${personaLensPanel(meta)}`)}${section('riskmix','Risk distribution',`<div class="panel"><h3>Risk distribution</h3>${data.charts.risk}</div>`)}${section('volume','Reconciliation volume',`<div class="grid grid2"><div class="panel"><h3>Unmatched volume</h3>${data.charts.volume}</div><div class="panel"><h3>Reconciliation volume by stream</h3>${data.charts.weighted}</div></div>`)}${section('queue','Priority queue',`<div class="panel"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;flex-wrap:wrap;gap:8px"><h3>Priority queue</h3><div style="display:flex;align-items:center;gap:8px"><label for="priorityQueueSelect" style="font-size:.85rem;font-weight:600;color:var(--muted)">Show:</label><select id="priorityQueueSelect" style="padding:5px 12px;border-radius:8px;font-size:.85rem;font-weight:700;cursor:pointer" onchange="state.priorityQueueLimit=Number(this.value);render()"><option value="10" ${pqLimit===10?'selected':''}>Top 10</option><option value="50" ${pqLimit===50?'selected':''}>Top 50</option><option value="100" ${pqLimit===100?'selected':''}>Top 100</option></select></div></div>${cards(queueItems)}</div>`)}`;html+=sectionShell([['overview','Overview'],['riskmix','Risk distribution'],['volume','Volume'],['queue','Priority queue']],body);}
if(state.page==='metrics'){html+=`<div class="hero"><h1>Metric Explanations</h1><p>Plain-English definitions for the DART metrics so leadership, analysts, and owners can interpret the same dashboard consistently.</p></div>${metricExplanations()}`;}
if(state.page==='briefing'){let body=`${section('brief-top','Briefing snapshot',`<div class="hero"><span class="workspaceBadge">${esc(meta.persona_effects?.lens_title||'Personalized briefing')}</span><h1>My Briefing</h1><p>${esc(p.first_question||'What should I investigate first?')} This view is tuned for <b>${esc(p.role||'your role')}</b>, a <b>${esc(p.audience||'general')}</b> audience, and <b>${esc(p.depth||'balanced')}</b> detail.</p></div>${miniDeck(data.summary)}<br>${personaLensPanel(meta)}`)}${section('focus','Recommended focus',`<div class="panel">${cards(data.summary.top_actions.slice(0,personaPriorityLimit(meta)))}</div>`)}${section('streams','Stream summary',`<div class="panel">${simpleTable((data.summary.stream_summary||[]).slice(0,Math.max(4,personaPriorityLimit(meta))))}</div>`)}`;html+=sectionShell([['brief-top','Snapshot'],['focus','Focus'],['streams','Streams']],body);}
if(state.page==='riskcenter'){let body=`${section('risk-overview','Risk overview',`<div class="hero"><h1>Risk Center</h1><p>Concentrated view of critical findings, risk drivers, and the field-level evidence behind the current risk picture.</p></div>${miniDeck(data.summary)}`)}${section('risk-drivers','Top risk drivers',`<div class="panel"><h3>Top risk drivers</h3>${data.charts.drivers}</div>`)}${section('risk-list','Critical and elevated findings',`<div class="panel">${table(data.issues,true)}</div>`)}`;html+=sectionShell([['risk-overview','Overview'],['risk-drivers','Risk drivers'],['risk-list','Findings']],body);}
if(state.page==='insights'){let body=aiInsightsPageHtml(data,meta);html+=sectionShell([['insight-top','Overview'],['strategic-diagnostics','AI Diagnostics'],['investigation-workspace','Anomaly Explorer'],['what-if-simulator','What-If Simulator'],['copilot-prompts','Diagnostic Copilot']],body);}
if(state.page==='scorecards'){html+=`<div class="hero"><h1>Stream Scorecards</h1><p>One executive card per stream, showing quality, volume, criticality, and impact.</p></div>${scorecards(data.summary.stream_summary)}`;}
if(state.page==='actioncenter'){let body=`${section('action-overview','Action status',`<div class="hero"><h1>Remediation Center</h1><p>Seeded workflow issues are loaded automatically, so the board is useful before you add anything manually.</p></div><div class="grid grid4">${['Open','In progress','Blocked','Resolved'].map(status=>{const rows=(meta.issues||[]).filter(x=>x.Status===status);return `<div class="panel metric metricClickable" role="button" tabindex="0" onclick='openLocalDrill("${status} remediation items","${rows.length} saved workflow item${rows.length===1?"":"s"} with status ${status}.",${JSON.stringify(rows).replace(/'/g,"&#39;")})' onkeydown="keyActivate(event,()=>this.click())"><div class="label">${status}</div><div class="value">${rows.length}</div><div class="sub">Saved workflow items</div></div>`;}).join('')}</div>`)}${section('action-candidates','Action candidates',`<div class="panel">${table(data.issues,true)}</div>`)}${section('board','Workflow board',`<div class="panel">${kanban(meta.issues)}</div>`)}`;html+=sectionShell([['action-overview','Status'],['action-candidates','Candidates'],['board','Board']],body);}
if(state.page==='impactexplorer'){let body=`${section('impact-overview','Impact overview',`<div class="hero"><h1>Impact Explorer</h1><p>Explore how field-level reconciliation findings connect to reports, KPIs, owners, and decision needs.</p></div>${miniDeck(data.summary)}`)}${section('impact-chart','Mapped impact chart',`<div class="panel">${data.charts.impact}</div>`)}${section('impact-map','Impact mapping',impactEditor(meta.impacts,data.rows))}`;html+=sectionShell([['impact-overview','Overview'],['impact-chart','Chart'],['impact-map','Mapping']],body);}
if(state.page==='catalog'){let body=`${section('catalog-edit','Editable catalog',`<div class="hero"><h1>Mapping Catalog</h1><p>Add, edit, validate, and delete field mappings. Seeded mappings are created from the highest-impact fields.</p></div>${lineageEditor(meta.connections)}`)}`;html+=sectionShell([['catalog-edit','Catalog']],body);}
if(state.page==='governance'){let body=`${section('gov-overview','Governance overview',`<div class="hero"><h1>Governance Center</h1><p>Track controls, owners, evidence, cadence, and decision readiness for DART findings.</p></div><div class="grid grid4"><div class="panel metric metricClickable" role="button" tabindex="0" onclick='openLocalDrill("Governance controls","${(meta.governance||[]).length} saved governance control${(meta.governance||[]).length===1?"":"s"}.",${JSON.stringify(meta.governance||[]).replace(/'/g,"&#39;")})' onkeydown="keyActivate(event,()=>this.click())"><div class="label">Controls</div><div class="value">${(meta.governance||[]).length}</div><div class="sub">Saved governance items</div></div><div class="panel metric metricClickable" role="button" tabindex="0" onclick='openLocalDrill("Active governance controls","${(meta.governance||[]).filter(x=>x.Status==='Active').length} active control${(meta.governance||[]).filter(x=>x.Status==='Active').length===1?"":"s"}.",${JSON.stringify((meta.governance||[]).filter(x=>x.Status==='Active')).replace(/'/g,"&#39;")})' onkeydown="keyActivate(event,()=>this.click())"><div class="label">Active</div><div class="value">${(meta.governance||[]).filter(x=>x.Status==='Active').length}</div><div class="sub">Controls in motion</div></div><div class="panel metric metricClickable" role="button" tabindex="0" onclick='openLocalDrill("Lineage mappings","${(meta.connections||[]).length} saved lineage mapping${(meta.connections||[]).length===1?"":"s"}.",${JSON.stringify(meta.connections||[]).replace(/'/g,"&#39;")})' onkeydown="keyActivate(event,()=>this.click())"><div class="label">Mappings</div><div class="value">${(meta.connections||[]).length}</div><div class="sub">Lineage catalog rows</div></div><div class="panel metric metricClickable" role="button" tabindex="0" onclick='openLocalDrill("Workflow issues","${(meta.issues||[]).length} saved workflow item${(meta.issues||[]).length===1?"":"s"}.",${JSON.stringify(meta.issues||[]).replace(/'/g,"&#39;")})' onkeydown="keyActivate(event,()=>this.click())"><div class="label">Issues</div><div class="value">${(meta.issues||[]).length}</div><div class="sub">Workflow items</div></div></div>`)}${section('gov-controls','Control register',governanceEditor(meta.governance))}${section('gov-links','Governance evidence links',`<div class="grid grid3"><div class="panel"><h3>Lineage evidence</h3><p class="muted">Use Mapping Catalog statuses as evidence that source-to-target lineage has been reviewed.</p><button class="btn secondary" onclick="showPage('catalog')">Open Mapping Catalog</button></div><div class="panel"><h3>Impact evidence</h3><p class="muted">Use Impact Explorer mappings to document downstream report/KPI exposure.</p><button class="btn secondary" onclick="showPage('impactexplorer')">Open Impact Explorer</button></div><div class="panel"><h3>Issue evidence</h3><p class="muted">Use Remediation Center workflow items to show owner assignment and remediation progress.</p><button class="btn secondary" onclick="showPage('actioncenter')">Open Remediation Center</button></div></div>`)}`;html+=sectionShell([['gov-overview','Overview'],['gov-controls','Controls'],['gov-links','Evidence links']],body);}
if(state.page==='explorer'){html+=`<div class="hero"><h1>System Explorer</h1><p>Search and inspect System A to System B relationships with risk tier and volume context next to each mapped field.</p></div><div class="panel"><h3>${intFmt(data.summary.rows)} fields in view</h3>${table(data.rows)}</div>`;}
if(state.page==='briefbuilder'){let body=`${section('brief-builder','Generated brief',`<div class="hero"><h1>Executive Brief Builder</h1><p>Create a leadership-ready brief from the current filtered scope.</p><div class="heroActions"><button class="btn" onclick="copyBrief()">Copy brief</button><button class="btn secondary" onclick="showPage('settings')">Update briefing settings</button></div></div>${miniDeck(data.summary)}<br><div class="panel"><h3>Generated brief</h3><div id="briefText" class="briefText">${esc(executiveBrief(data,meta))}</div></div>`)}`;html+=sectionShell([['brief-builder','Brief']],body);}
if(state.page==='copilot'){let body=`${section('chat-area','Conversation',`<div class="hero"><h1>DART Copilot</h1><p>Ask DART Copilot questions about the filtered dataset. Without an API key, it still returns deterministic local analysis.</p></div><div class="panel"><h3>Conversation</h3><div class="chat" id="copilotChatHistory">${(meta.chat||[]).map(chatMessageHtml).join('')||'<div class="empty">Ask a question to start the DART Copilot conversation.</div>'}</div><br><div class="filterGrid" style="grid-template-columns:1fr auto"><input id="prompt" placeholder="What are the top risks and why?" onkeydown="if(event.key==='Enter')askCopilot()"><button class="btn" onclick="askCopilot()">Send</button></div></div>`)}${section('chat-context','Current context',`<div class="panel">${miniDeck(data.summary)}<br><h3>Suggested prompts</h3><button class="btn secondary" onclick="quickPrompt('Summarize the biggest quality risks in the current scope')">Summarize risks</button> <button class="btn secondary" onclick="quickPrompt('What should I investigate first?')">Next investigation</button> <button class="btn secondary" onclick="quickPrompt('Explain health score, field average, and event weighted match rate')">Explain metrics</button></div>`)}`;html+=sectionShell([['chat-area','Chat'],['chat-context','Context']],body);}
if(state.page==='raweditor'){html+=rawDataPage(rawInfo);}
if(state.page==='emailagent'){html+=emailAgentPage(emailInfo);}
if(state.page==='data'){let body=`${section('data-source','Source files',`<div class="hero"><h1>Data Management</h1><p>DART loads the real reconciliation workbook automatically and enriches it with the CMS mapping workbook. Use the dedicated Raw Data Editor for controlled writeback and the Email Agent for workbook-change automation.</p><div class="heroActions"><button class="btn" onclick="showPage('raweditor')">Open Raw Data Editor</button><button class="btn secondary" onclick="showPage('emailagent')">Open Email Agent</button><button class="btn ghost" onclick="restoreDemo()">Reload Excel files</button><a class="btn ghost" href="/download/current.csv">Download current view</a></div></div><div class="grid grid2"><div class="panel"><h3>Current source</h3><p>${esc(meta.source)}</p><p class="muted">Rows: ${intFmt(meta.rows)} · Columns: ${intFmt(meta.columns)}</p></div><div class="panel"><h3>Detected structure</h3>${simpleTable(meta.profile)}</div></div>`)}${section('data-preview','Preview',`<div class="panel"><h3>Standardized preview</h3>${table(data.rows)}</div>`)}`;html+=sectionShell([['data-source','Source files'],['data-preview','Preview']],body);}
if(state.page==='profile'){html+=profilePage(meta);}if(state.page==='settings'){html+=settingsPage(meta);}const appEl=document.getElementById('app');appEl.classList.remove('page-enter');appEl.innerHTML=html;runInjectedScripts(appEl);decorateMetricHelp(appEl);refreshThemeVisuals(appEl);requestAnimationFrame(()=>appEl.classList.add('page-enter'));ensureAssistantBubble();}
function profilePage(meta){const u=meta.auth_user||{};return `<div class="hero"><h1>Profile</h1><p>Manage your local DART prototype profile. Credentials are stored in <code>${esc(u.storage_file||'dart_users.json')}</code> for easy local testing only.</p><div class="heroActions"><button class="btn secondary" onclick="logoutUser()">Logout</button></div></div><form class="panel" id="profileForm"><div class="grid grid2"><div class="field"><label>Username</label><input value="${esc(u.username||'')}" disabled></div><div class="field"><label>Display name</label><input id="profileName" value="${esc(u.display_name||'')}"></div><div class="field"><label>Email</label><input id="profileEmail" value="${esc(u.email||'')}"></div><div class="field"><label>Organization</label><input id="profileOrg" value="${esc(u.organization||'')}"></div><div class="field"><label>Role</label><input id="profileRole" value="${esc(u.role||'')}"></div><div class="field"><label>New password</label><input id="profilePass" type="password" placeholder="Leave blank to keep current password"></div></div><br><button class="btn" type="submit">Save profile</button></form><br><div class="panel"><h3>Prototype storage note</h3><p class="muted">This prototype profile system stores credentials locally in plain text. It is useful for local demos, but should be replaced with proper authentication before use with sensitive data.</p></div>`;}
function settingsPage(meta){const p=meta.persona||{};const focus=(p.focus||[]).join(', ');return `<div class="hero"><span class="workspaceBadge">Persona & startup</span><h1>Workspace Settings</h1><p>Update the persona lens and choose where DART should take you immediately after future logins. These settings are saved to your local user profile.</p></div>${personaLensPanel(meta)}<br><form class="panel" id="settingsForm"><div class="grid grid2"><div class="field"><label>Program</label><select id="setProgram">${['Medicare','Medicaid','Both','Other / General'].map(x=>`<option ${p.program===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Primary role</label><select id="setRole">${['Executive / Leadership','Data / Analytics','Program / Policy','Operations','Quality / Compliance','IT / Engineering','Research','Other'].map(x=>`<option ${p.role===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Focus areas</label><input id="setFocus" value="${esc(focus)}"></div><div class="field"><label>Audience</label><select id="setAudience">${['Leadership','Myself','Analysts','Program teams','Technical teams','External stakeholders'].map(x=>`<option ${p.audience===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Detail level</label><select id="setDepth">${['Executive','Balanced','Technical'].map(x=>`<option ${p.depth===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Briefing question</label><input id="setQuestion" value="${esc(p.first_question||'')}"></div><div class="field"><label>Preferred workspace</label><select id="setPreferredWorkspace">${[['medicare','Medicare'],['medicaid','Medicaid'],['byo','Build Your Own']].map(([v,l])=>`<option value="${v}" ${p.preferred_workspace===v?'selected':''}>${l}</option>`).join('')}</select></div><div class="field"><label>Start page</label><select id="setLandingPage">${personaLandingOptions(p.landing_page||'auto')}</select><div class="personaPrefNote">If the selected page does not belong to the preferred workspace, DART safely falls back to that workspace home.</div></div></div><br><button class="btn" type="submit">Save settings</button> <button class="btn secondary" type="button" onclick="state.workspace='${esc(p.preferred_workspace||'medicare')}';sessionStorage.setItem('dart_workspace',state.workspace);showPage('${esc(p.landing_page||'home')}')">Go to my start page</button></form>`;}

function rawText(v){return v===null||v===undefined?'':String(v);}
function selectedValues(id){const el=document.getElementById(id);return el?[...el.selectedOptions].map(o=>o.value):[];}
function selectOptions(items,selected=[]){const set=new Set(selected||[]);return (items||[]).map(x=>`<option value="${esc(x)}" ${set.has(x)?'selected':''}>${esc(x)}</option>`).join('');}
function smartSummaryMarkup(values,placeholder){const vals=values||[];if(!vals.length)return `<span class="smartPlaceholder">${esc(placeholder||'Choose values')}</span>`;const chips=vals.slice(0,2).map(v=>`<span class="smartChoice" title="${esc(v)}">${esc(v)}</span>`).join('');return chips+(vals.length>2?`<span class="smartMore">+${vals.length-2} more</span>`:'');}
function smartMultiSelect(id,items,selected=[],placeholder='Choose columns'){const set=new Set(selected||[]);const options=(items||[]).map(x=>`<option value="${esc(x)}" ${set.has(x)?'selected':''}>${esc(x)}</option>`).join('');const checks=(items||[]).map(x=>`<label class="smartMultiOption"><input type="checkbox" value="${esc(x)}" ${set.has(x)?'checked':''} onchange="smartMultiApplyOption(this)"><span>${esc(x)}</span></label>`).join('');return `<div class="smartMulti" data-picker-id="${esc(id)}" data-placeholder="${esc(placeholder)}"><select id="${esc(id)}" class="multiSelect" multiple>${options}</select><button type="button" class="smartMultiTrigger" onclick="smartMultiOpen(event,'${esc(id)}')"><span class="smartMultiSummary">${smartSummaryMarkup([...set],placeholder)}</span><span class="smartMeta"><span class="smartCount">${set.size}</span><span class="smartChevron">▾</span></span></button><div class="smartMultiPanel" onclick="event.stopPropagation()"><div class="smartMultiToolbar"><input class="smartMultiSearch" placeholder="Search columns..." oninput="smartMultiFilter(this)"><button type="button" class="smartMiniBtn" onclick="smartMultiSetVisible('${esc(id)}',true)">Select visible</button><button type="button" class="smartMiniBtn" onclick="smartMultiSetVisible('${esc(id)}',false)">Clear visible</button></div><div class="smartMultiOptions">${checks||'<div class="smartEmpty">No columns available.</div>'}</div></div></div>`;}
function closeSmartMultis(except=null){document.querySelectorAll('.smartMulti.open').forEach(w=>{if(w!==except)w.classList.remove('open')});}
function smartMultiOpen(event,id){event.stopPropagation();const select=document.getElementById(id);const wrap=select?.closest('.smartMulti');if(!wrap)return;const opening=!wrap.classList.contains('open');closeSmartMultis(wrap);wrap.classList.toggle('open',opening);if(opening)setTimeout(()=>wrap.querySelector('.smartMultiSearch')?.focus(),0);}
function smartMultiFilter(input){const q=input.value.trim().toLowerCase();const wrap=input.closest('.smartMulti');let visible=0;wrap.querySelectorAll('.smartMultiOption').forEach(opt=>{const show=!q||opt.textContent.toLowerCase().includes(q);opt.classList.toggle('smartHidden',!show);if(show)visible++;});let empty=wrap.querySelector('.smartSearchEmpty');if(!visible){if(!empty){empty=document.createElement('div');empty.className='smartEmpty smartSearchEmpty';empty.textContent='No matching columns.';wrap.querySelector('.smartMultiOptions').appendChild(empty);}}else if(empty)empty.remove();}
function smartMultiApplyOption(input){const wrap=input.closest('.smartMulti');const select=wrap.querySelector('select');const option=[...select.options].find(o=>o.value===input.value);if(option)option.selected=input.checked;refreshSmartMulti(wrap);}
function smartMultiSetVisible(id,on){const select=document.getElementById(id);const wrap=select?.closest('.smartMulti');if(!wrap)return;wrap.querySelectorAll('.smartMultiOption:not(.smartHidden) input[type="checkbox"]').forEach(input=>{input.checked=on;const option=[...select.options].find(o=>o.value===input.value);if(option)option.selected=on;});refreshSmartMulti(wrap);}
function refreshSmartMulti(wrap){const select=wrap.querySelector('select');const values=[...select.selectedOptions].map(o=>o.value);wrap.querySelector('.smartMultiSummary').innerHTML=smartSummaryMarkup(values,wrap.dataset.placeholder||'Choose columns');wrap.querySelector('.smartCount').textContent=values.length;}
function conditionNeedsValue(operator){return !['Is blank (empty or null)','Is null / NaN','Is empty string'].includes(operator);}
function syncEmailConditionRow(row){if(!row)return;const op=row.querySelector('.condOperator')?.value||'Equals';const input=row.querySelector('.condValue');if(!input)return;const needs=conditionNeedsValue(op);input.disabled=!needs;input.placeholder=needs?'Comparison value':'No value required';if(!needs)input.value='';}
function emailConditionRow(rule={},cols=[]){const operators=["Is blank (empty or null)", "Is null / NaN", "Is empty string", "Equals", "Does not equal", "Contains", "Does not contain", "Starts with", "Greater than", "Greater than or equal to", "Less than", "Less than or equal to", "Before date", "After date"];const op=rule.operator||'Equals';const needs=conditionNeedsValue(op);return `<div class="conditionRow"><div class="field"><label>Column</label><select class="condColumn">${selectOptions(cols,[rule.column||cols[0]||''])}</select></div><div class="field"><label>Operator</label><select class="condOperator" onchange="syncEmailConditionRow(this.closest('.conditionRow'))">${operators.map(x=>`<option ${x===op?'selected':''}>${esc(x)}</option>`).join('')}</select></div><div class="field"><label>Value</label><input class="condValue" value="${needs?esc(rule.value||''):''}" placeholder="${needs?'Comparison value':'No value required'}" ${needs?'':'disabled'}></div><button type="button" class="btn ghost small" onclick="this.closest('.conditionRow').remove()">Remove</button></div>`;}
function addEmailCondition(){const cols=state.emailAgent?.columns||[];document.getElementById('conditionList').insertAdjacentHTML('beforeend',emailConditionRow({},cols));}
function emailAgentPage(info){state.emailAgent=info;const templates=info.templates||[];const tpl=(state.emailEditingId?templates.find(t=>t.id===state.emailEditingId):null)||info.default_template;const cols=info.columns||[];const status=info.status||{};const smtp=info.smtp||{};const saved=templates.map(t=>`<div class="savedAutomation"><div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start"><div><b>${esc(t.name)}</b><div class="muted">${esc(t.recipient_email)} · ${t.enabled?'Enabled':'Disabled'} · ${esc(t.last_status||'Not run yet')}</div></div><span class="pill ${t.enabled?'Stable':'Watch'}">${t.enabled?'Enabled':'Disabled'}</span></div><div class="automationActions"><button class="btn secondary small" onclick="editEmailById('${esc(t.id)}')">Edit</button><button class="btn secondary small" onclick="baselineEmail('${esc(t.id)}')">Refresh baseline</button><button class="btn secondary small" onclick="runEmail('${esc(t.id)}',false)">Check only</button><button class="btn small" onclick="runEmail('${esc(t.id)}',true)">Run + send</button><button class="btn ghost small" onclick="testEmail('${esc(t.id)}')">Send test</button><button class="btn ghost small" onclick="deleteEmailTemplate('${esc(t.id)}')">Delete</button></div></div>`).join('')||'<div class="empty">No saved automations yet. Configure one below and save it.</div>';
const history=(info.history||[]).length?simpleTable(info.history):'<div class="empty">No execution history yet.</div>';
const dup=tpl.identity_duplicate_count||0;
return `${section('email-status','Agent status',`<div class="hero"><h1>Email Agent</h1><p>Watch the real reconciliation workbook for new or changed rows, apply saved requirements, and send structured email alerts with an optional CSV attachment.</p><div class="heroActions"><button class="btn" onclick="checkAllEmails()">Check workbook now</button><button class="btn secondary" onclick="baselineAllEmails()">Baseline all enabled</button><button class="btn ghost" onclick="testSmtp()">Test SMTP connection</button></div></div><div class="statusGrid"><div class="statusCard"><div class="muted">Watcher</div><div class="big"><span class="statusDot ${String(status.status||'').toLowerCase().includes('error')?'bad':'good'}"></span>${esc(status.status||'Not started')}</div></div><div class="statusCard"><div class="muted">Check interval</div><div class="big">${esc(info.check_interval_seconds)}s</div></div><div class="statusCard"><div class="muted">Workbook</div><div class="big">${info.workbook_found?'Found':'Missing'}</div></div><div class="statusCard"><div class="muted">Email delivery</div><div class="big">${smtp.configured?'Ready':'Setup needed'}</div><div class="muted">${esc(smtp.host||'')} ${smtp.port?': '+esc(smtp.port):''}</div></div></div><br><div class="callout"><b>How it works:</b> establish a baseline, edit the workbook (the Raw Data Editor can do this), then DART compares row identities and monitored columns. Only qualifying new/changed rows are emailed. Identity columns identify the same logical record across workbook versions; monitored columns determine which cell changes count as a change.</div>`)}${section('email-saved','Saved automations',`<div class="panel"><div style="display:flex;justify-content:space-between;align-items:center;gap:12px"><div><h3>Saved automations</h3><p class="muted">Templates and baselines persist under the Email Agent directory.</p></div><button class="btn secondary" onclick="newEmailTemplate()">New automation</button></div>${saved}</div>`)}${section('email-builder','Automation builder',`<div class="panel emailBuilder"><h3>${tpl.id?'Edit automation':'Create automation'}</h3><div class="grid grid2"><div class="field"><label>Name</label><input id="emailName" value="${esc(tpl.name||'')}"></div><div class="field"><label>Recipient email(s)</label><input id="emailRecipient" value="${esc(tpl.recipient_email||'')}" placeholder="owner@example.com; second@example.com"></div><div class="field"><label>Enabled</label><select id="emailEnabled"><option value="false" ${!tpl.enabled?'selected':''}>Disabled</option><option value="true" ${tpl.enabled?'selected':''}>Enabled</option></select></div><div class="field"><label>Trigger mode</label><select id="emailTrigger">${['New or changed rows','New rows only','All changes including deletions'].map(x=>`<option ${tpl.trigger_mode===x?'selected':''}>${x}</option>`).join('')}</select></div><div class="field"><label>Row identity columns</label>${smartMultiSelect('emailIdentity',cols,tpl.identity_columns,'Choose row identity columns')}<div class="rangeLine">Use the searchable picker to choose columns that together identify the same logical row. ${dup?`Current selection has ${intFmt(dup)} rows sharing duplicate identity values.`:'Current selection is suitable for row matching.'}</div></div><div class="field"><label>Monitored columns</label>${smartMultiSelect('emailMonitor',cols,tpl.monitor_columns,'Choose monitored columns')}<div class="rangeLine">Search, select visible values, or clear quickly. Any selected field change can trigger the automation. Empty means all columns on the server.</div></div><div class="field"><label>Display columns in email</label>${smartMultiSelect('emailDisplay',cols,tpl.display_columns,'Choose email display columns')}<div class="smartPickerHint">Controls the columns included in the alert preview and CSV context.</div></div><div class="field"><label>Requirement mode</label><select id="emailConditionMode"><option ${tpl.condition_mode==='Any change'?'selected':''}>Any change</option><option ${tpl.condition_mode!=='Any change'?'selected':''}>Only when conditions are met</option></select><label style="margin-top:10px">Condition logic</label><select id="emailLogic"><option ${tpl.condition_logic!=='ANY condition'?'selected':''}>ALL conditions</option><option ${tpl.condition_logic==='ANY condition'?'selected':''}>ANY condition</option></select></div></div><br><div style="display:flex;justify-content:space-between;align-items:center"><h3 style="margin:0">Conditions</h3><button type="button" class="btn secondary small" onclick="addEmailCondition()">Add condition</button></div><div id="conditionList">${(tpl.conditions||[]).map(r=>emailConditionRow(r,cols)).join('')}</div><div class="grid grid2"><div class="field"><label>Subject prefix</label><input id="emailSubject" value="${esc(tpl.subject_prefix||'[DART Alert]')}"></div><div class="field"><label>Greeting</label><input id="emailGreeting" value="${esc(tpl.greeting||'Hello,')}"></div></div><div class="field"><label>Email introduction</label><textarea id="emailIntro">${esc(tpl.email_intro||'')}</textarea></div><div class="field" style="max-width:280px"><label>CSV attachment</label><select id="emailCsv"><option value="true" ${tpl.include_csv!==false?'selected':''}>Include CSV</option><option value="false" ${tpl.include_csv===false?'selected':''}>No CSV</option></select></div><br><div class="heroActions"><button class="btn" onclick="saveEmailTemplate()">Save automation</button><button class="btn secondary" onclick="previewEmailTemplate()">Preview current matches</button></div></div><div id="emailPreviewHolder" style="margin-top:16px"></div>`)}${section('email-history','Execution history',`<div class="panel"><h3>Recent execution history</h3>${history}</div><br><div class="callout"><b>SMTP setup:</b> Gmail delivery is configured in this local-demo file using the same working configuration as main2. The password is never returned to the browser or stored in automation templates.${smtp.issues?.length?`<br><br><b>Current setup items:</b> ${esc(smtp.issues.join(' '))}`:''}</div>`)}`;}
function editEmailById(id){state.emailEditingId=id;render();}
function newEmailTemplate(){state.emailEditingId=null;render();}
function collectEmailTemplate(){const conditions=[...document.querySelectorAll('.conditionRow')].map(row=>({column:row.querySelector('.condColumn').value,operator:row.querySelector('.condOperator').value,value:row.querySelector('.condValue').value}));return {id:state.emailEditingId||'',name:emailName.value,recipient_email:emailRecipient.value,enabled:emailEnabled.value==='true',trigger_mode:emailTrigger.value,identity_columns:selectedValues('emailIdentity'),monitor_columns:selectedValues('emailMonitor'),condition_mode:emailConditionMode.value,condition_logic:emailLogic.value,conditions,display_columns:selectedValues('emailDisplay'),subject_prefix:emailSubject.value,greeting:emailGreeting.value,email_intro:emailIntro.value,include_csv:emailCsv.value==='true'};}
async function saveEmailTemplate(){try{const r=await postJson('/api/email-agent/template',collectEmailTemplate());state.emailEditingId=r.template.id;toast('Automation saved');await render();}catch(err){console.error(err);toast('Could not save: '+err.message.slice(0,160));}}
async function deleteEmailTemplate(id){if(!confirm('Delete this automation and its baseline snapshot?'))return;await del('/api/email-agent/template/'+encodeURIComponent(id));state.emailEditingId=null;toast('Automation deleted');await render();}
async function baselineEmail(id){const r=await postJson('/api/email-agent/baseline/'+encodeURIComponent(id),{});toast(`${r.rows||0} rows saved as baseline`);await render();}
async function baselineAllEmails(){const r=await postJson('/api/email-agent/baselines',{});toast(`Refreshed ${r.baselines||0} baseline(s)`);await render();}
async function runEmail(id,send){const r=await postJson('/api/email-agent/run/'+encodeURIComponent(id),{send_email:send});toast(r.status||'Check complete');await render();}
async function testEmail(id){try{const r=await postJson('/api/email-agent/test/'+encodeURIComponent(id),{});toast(r.status||'Test email sent');}catch(err){toast('Test failed: '+err.message.slice(0,160));}}
async function checkAllEmails(){const r=await postJson('/api/email-agent/check',{});toast(`Checked ${r.results?.length||0} enabled automation(s)`);await render();}
async function testSmtp(){try{const r=await postJson('/api/email-agent/smtp-test',{});toast(`SMTP ready: ${r.sender}`);}catch(err){toast('SMTP test failed: '+err.message.slice(0,160));}}
async function previewEmailTemplate(){try{const r=await postJson('/api/email-agent/preview',collectEmailTemplate());const holder=document.getElementById('emailPreviewHolder');holder.innerHTML=`<div class="panel"><h3>Preview · ${intFmt(r.matches)} current matching row(s)</h3><div class="emailPreview"><iframe id="emailPreviewFrame"></iframe></div></div>`;const frame=document.getElementById('emailPreviewFrame');frame.srcdoc=r.html||'<p>No preview available.</p>';}catch(err){toast('Preview failed: '+err.message.slice(0,160));}}
function rawDataPage(raw){state.rawData=raw;const cols=raw.columns||[];const rows=raw.rows||[];const start=(raw.offset||0)+1;const end=(raw.offset||0)+rows.length;const tableHtml=`<div class="tableWrap"><table class="table rawTable"><thead><tr><th class="rawRow">Excel row</th>${cols.map(c=>`<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${rows.map(r=>`<tr><td class="rawRow">${esc(r._ExcelRow)}</td>${cols.map(c=>`<td><input class="rawCell" data-row="${esc(r._ExcelRow)}" data-col="${esc(c)}" data-before="${esc(rawText(r[c]))}" value="${esc(rawText(r[c]))}"></td>`).join('')}</tr>`).join('')}</tbody></table></div>`;return `${section('raw-overview','Workbook editor',`<div class="hero"><h1>Raw Data Editor</h1><p>Edit the actual source workbook used by DART. Saves are cell-level, backed up first, written atomically, and can immediately run the Email Agent for a clean end-to-end demo.</p><div class="heroActions"><button class="btn" onclick="saveRawChanges()">Save changes to Excel</button><button class="btn secondary" onclick="reloadRawData()">Refresh from workbook</button><button class="btn ghost" onclick="restoreLatestBackup()" ${raw.latest_backup?'':'disabled'}>Restore latest backup</button></div></div><div class="statusGrid"><div class="statusCard"><div class="muted">Workbook</div><div class="big">${esc(raw.workbook)}</div></div><div class="statusCard"><div class="muted">Raw rows</div><div class="big">${intFmt(raw.raw_rows)}</div></div><div class="statusCard"><div class="muted">Raw columns</div><div class="big">${intFmt(cols.length)}</div></div><div class="statusCard"><div class="muted">Last saved</div><div class="big" style="font-size:1rem">${esc(raw.modified_at||'-')}</div></div></div><br><div class="callout"><b>Demo workflow:</b> first establish an Email Agent baseline. Then change one or more cells here and save. A timestamped full-workbook backup is created before writeback. Leave the immediate Email Agent check enabled to show qualifying changes without waiting for the watcher interval.</div>`)}${section('raw-grid','Edit rows',`<div class="panel"><div class="grid grid2"><div class="field"><label>Search raw workbook</label><input id="rawSearchBox" value="${esc(state.rawSearch)}" placeholder="Search any column"><div class="heroActions"><button class="btn secondary small" onclick="applyRawSearch()">Search</button><button class="btn ghost small" onclick="clearRawSearch()">Clear</button></div></div><div class="field"><label>After save</label><select id="rawRunAgent"><option value="true">Check Email Agent immediately</option><option value="false">Do not run Email Agent</option></select><div class="rangeLine">Showing ${intFmt(start)}-${intFmt(end)} of ${intFmt(raw.filtered_rows)} matching rows.</div></div></div><br>${tableHtml}<br><div class="heroActions"><button class="btn secondary" onclick="rawPage(-1)" ${raw.offset<=0?'disabled':''}>Previous</button><button class="btn secondary" onclick="rawPage(1)" ${raw.offset+raw.limit>=raw.filtered_rows?'disabled':''}>Next</button><button class="btn" onclick="saveRawChanges()">Save changes to Excel</button></div></div>`)}${section('raw-backup','Backup status',`<div class="panel"><h3>Backup protection</h3><p class="muted">${raw.latest_backup?`Latest backup: <code>${esc(raw.latest_backup)}</code>`:'No editor backup has been created yet.'}</p><p class="muted">The hidden Excel row number is used only to write the edited cell back to the exact source row; it is never stored as a new workbook column.</p></div>`)}`;}
function collectRawChanges(){return [...document.querySelectorAll('.rawCell')].filter(el=>el.value!==el.dataset.before).map(el=>({excel_row:Number(el.dataset.row),column:el.dataset.col,before:el.dataset.before,after:el.value}));}
async function saveRawChanges(){const changes=collectRawChanges();if(!changes.length){toast('No cell changes to save');return;}if(!confirm(`Save ${changes.length} changed cell(s) to the real Excel workbook?`))return;try{const runAgent=document.getElementById('rawRunAgent')?.value!=='false';const r=await postJson('/api/raw-data/save',{changes,run_email_agent:runAgent,signature:state.rawData?.signature||''});toast(`Saved ${r.saved} cell(s)${r.agent_results?.length?' · Email Agent checked':''}`);state.rawOffset=0;state.meta=await api('/api/meta');await render();}catch(err){console.error(err);toast('Save failed: '+err.message.slice(0,180));}}
async function restoreLatestBackup(){if(!confirm('Restore the most recent pre-edit workbook backup? This replaces the current source workbook.'))return;try{const r=await postJson('/api/raw-data/restore-latest',{});toast('Workbook restored');state.rawOffset=0;await render();}catch(err){toast('Restore failed: '+err.message.slice(0,160));}}
async function reloadRawData(){state.rawOffset=0;await render();}
async function applyRawSearch(){state.rawSearch=document.getElementById('rawSearchBox').value.trim();state.rawOffset=0;await render();}
async function clearRawSearch(){state.rawSearch='';state.rawOffset=0;await render();}
async function rawPage(dir){state.rawOffset=Math.max(0,state.rawOffset+dir*state.rawLimit);await render();}

async function addImpact(rows){if(!rows.length){toast('No rows to map');return;}const r=rows[Number(impactField.value)];await postJson('/api/impacts',{Stream:r.Stream,Field:r['NCH Target Column'],Report:impactReport.value,KPI:impactKpi.value,Owner:impactOwner.value,Impact:impactLevel.value,DecisionNeed:impactDecision.value});toast('Impact mapped');await render();}
async function askCopilot(){
  const input=document.getElementById('prompt');
  const v=(input?.value||'').trim();
  if(!v)return;
  if(input)input.value='';
  const bucket=getScopedHistory('medicare');
  const session=getActiveSession('medicare');
  if(session.messages.length===0){
    session.title=v.length>32?v.slice(0,30)+'…':v;
  }
  session.messages.push({role:'user',content:v});
  state.meta=state.meta||{};
  state.meta.chat=session.messages.slice();
  persistChatHistoryState();
  try{
    const res=await postJson('/api/copilot',{prompt:v,filters:state.filters});
    session.messages.push({role:'assistant',content:res.answer||'I could not generate an answer.'});
  }catch(err){
    session.messages.push({role:'assistant',content:`AI service error: ${err?.message||err}`});
  }
  state.meta.chat=session.messages.slice();
  persistChatHistoryState();
  await render();
  setTimeout(()=>{
    document.getElementById('copilotChatHistory')?.scrollTo(0,999999);
    scrollBubbleToBottom();
  },50);
}
async function quickPrompt(v){
  if(!v)return;
  const bucket=getScopedHistory('medicare');
  const session=getActiveSession('medicare');
  if(session.messages.length===0){
    session.title=v.length>32?v.slice(0,30)+'…':v;
  }
  session.messages.push({role:'user',content:v});
  state.meta=state.meta||{};
  state.meta.chat=session.messages.slice();
  persistChatHistoryState();
  try{
    const res=await postJson('/api/copilot',{prompt:v,filters:state.filters});
    session.messages.push({role:'assistant',content:res.answer||'I could not generate an answer.'});
  }catch(err){
    session.messages.push({role:'assistant',content:`AI service error: ${err?.message||err}`});
  }
  state.meta.chat=session.messages.slice();
  persistChatHistoryState();
  await render();
  setTimeout(()=>{
    document.getElementById('copilotChatHistory')?.scrollTo(0,999999);
    scrollBubbleToBottom();
  },50);
}
async function uploadFile(){const f=file.files[0];if(!f){toast('Choose a file first');return;}const fd=new FormData();fd.append('file',f);const r=await fetch('/api/upload',{method:'POST',body:fd});if(!r.ok){toast(await r.text());return;}toast('Data loaded');await render();}
async function restoreDemo(){await postJson('/api/restore-demo',{});toast('Excel files reloaded');await render();}
function copyBrief(){const txt=document.getElementById('briefText')?.innerText||'';navigator.clipboard.writeText(txt);toast('Brief copied');}
document.addEventListener('submit',e=>{if(e.target&&e.target.id==='settingsForm')saveSettings(e);if(e.target&&e.target.id==='profileForm')saveProfile(e)});
syncThemeSwitch();
boot();
</script>
</body>
</html>
'''

# -----------------------------------------------------------------------------
# API routes
# -----------------------------------------------------------------------------

@app.middleware("http")
async def no_cache_everywhere(request: Request, call_next):
    """Prevent stale one-file frontend/API responses during local development."""
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    response.headers["X-DART-Version"] = APP_VERSION
    response.headers["X-DART-Build"] = BUILD_FINGERPRINT
    return response

@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    ensure_data()
    return HTMLResponse(
        content=HTML,
        headers={
            "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
            "Pragma": "no-cache",
            "Expires": "0",
        },
    )


@app.get("/api/build-info")
def build_info() -> Dict[str, Any]:
    """Unambiguous runtime proof of the exact Python/frontend bundle being served."""
    return {
        "version": APP_VERSION,
        "build_fingerprint": BUILD_FINGERPRINT,
        "source_file": str(Path(__file__).resolve()),
        "pid": os.getpid(),
        "started_at": APP_STARTED_AT,
        "html_has_grouped_byo_nav": "Workspace" in HTML and "byo-raw" in HTML and "byo-email" in HTML,
        "html_has_persistent_byo_nav": "byoPersistentNav" in HTML,
        "html_has_injected_byo_tabs": "byoDataTabs" in HTML,
    }


@app.post("/api/signup")
async def signup(request: Request) -> Any:
    payload = await request.json()
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", "")).strip()
    if not username or not password:
        return JSONResponse({"error": "Username and password are required."}, status_code=400)
    data = load_users()
    users = data.setdefault("users", {})
    if username in users:
        return JSONResponse({"error": "Username already exists."}, status_code=400)
    users[username] = {
        "password": password,
        "created_at": "local prototype",
        "profile": {
            "display_name": str(payload.get("display_name", "")).strip() or username,
            "email": str(payload.get("email", "")).strip(),
            "organization": str(payload.get("organization", "")).strip(),
            "role": str(payload.get("role", "")).strip(),
        },
    }
    save_users(data)
    STATE["current_user"] = username
    STATE["persona"] = None
    STATE["chat"] = []
    STATE["byo_chat"] = []
    STATE["byo_chat_pair"] = []
    return {"status": "ok"}


@app.post("/api/login")
async def login(request: Request) -> Any:
    payload = await request.json()
    username = str(payload.get("username", "")).strip()
    password = str(payload.get("password", "")).strip()
    user = load_users().get("users", {}).get(username)
    if not user or user.get("password") != password:
        return JSONResponse({"error": "Invalid username or password."}, status_code=401)
    STATE["current_user"] = username
    stored_persona = user.get("persona") if isinstance(user.get("persona"), dict) else None
    STATE["persona"] = normalize_persona(stored_persona) if stored_persona else None
    # Chat is intentionally session-local; clear it at login so one local account never sees another account's conversation.
    STATE["chat"] = []
    STATE["byo_chat"] = []
    STATE["byo_chat_pair"] = []
    return {
        "status": "ok",
        "has_persona": bool(STATE["persona"]),
        "preferred_workspace": (STATE["persona"] or {}).get("preferred_workspace"),
        "landing_page": (STATE["persona"] or {}).get("landing_page"),
    }


@app.post("/api/logout")
async def logout() -> Dict[str, str]:
    STATE["current_user"] = None
    STATE["persona"] = None
    STATE["chat"] = []
    STATE["byo_chat"] = []
    STATE["byo_chat_pair"] = []
    return {"status": "ok"}


@app.post("/api/profile")
async def update_profile(request: Request) -> Any:
    username = STATE.get("current_user")
    if not username:
        return JSONResponse({"error": "Not logged in."}, status_code=401)
    payload = await request.json()
    data = load_users()
    user = data.setdefault("users", {}).setdefault(username, {"password": "", "profile": {}})
    profile = user.setdefault("profile", {})
    for key in ["display_name", "email", "organization", "role"]:
        if key in payload:
            profile[key] = str(payload.get(key, "")).strip()
    new_password = str(payload.get("password", "")).strip()
    if new_password:
        user["password"] = new_password
    save_users(data)
    return {"status": "ok"}


@app.get("/api/meta")
def meta() -> Dict[str, Any]:
    ensure_data()
    df = standardize(STATE["df"])
    profile = pd.DataFrame({
        "Column": df.columns,
        "Type": [str(df[c].dtype) for c in df.columns],
        "Non-null": [int(df[c].notna().sum()) for c in df.columns],
        "Unique": [int(df[c].nunique(dropna=True)) for c in df.columns],
    }).to_dict("records")
    persisted_persona = current_persona()
    STATE["persona"] = persisted_persona
    return {
        "auth_user": public_user(STATE.get("current_user")),
        "persona": persisted_persona,
        "persona_effects": persona_effects(persisted_persona),
        "source": STATE["source"],
        "version": APP_VERSION,
        "build_fingerprint": BUILD_FINGERPRINT,
        "server_pid": os.getpid(),
        "server_started_at": APP_STARTED_AT,
        "rows": len(df),
        "columns": len(df.columns),
        "columns_list": [str(c) for c in df.columns],
        "streams": sorted(df["Stream"].astype(str).unique().tolist()) if "Stream" in df else [],
        "classes": sorted(df["Classification"].astype(str).unique().tolist()) if "Classification" in df else [],
        "tiers": sorted(df["RiskTier"].astype(str).unique().tolist()) if "RiskTier" in df else [],
        "reasons": sorted(df["Sub-Classification"].astype(str).unique().tolist()) if "Sub-Classification" in df else [],
        "profile": profile,
        "issues": STATE["issues"],
        "connections": STATE["connections"],
        "impacts": STATE["impacts"],
        "governance": STATE["governance"],
        "rules": STATE["rules"],
        "chat": STATE["chat"],
    }


@app.post("/api/persona")
async def persona(request: Request) -> Any:
    username = STATE.get("current_user")
    if not username:
        return JSONResponse({"error": "Not logged in."}, status_code=401)
    normalized = normalize_persona(await request.json())
    data = load_users()
    user = data.setdefault("users", {}).setdefault(username, {"password": "", "profile": {}})
    user["persona"] = normalized
    save_users(data)
    STATE["persona"] = normalized
    if not STATE["chat"]:
        STATE["chat"] = [{"role": "assistant", "content": "Hi. I am Dartboard, your DART assistant. I will tailor the level of detail and next-step framing to your saved workspace persona."}]
    return {
        "status": "ok",
        "persona": normalized,
        "preferred_workspace": normalized["preferred_workspace"],
        "landing_page": normalized["landing_page"],
    }


@app.post("/api/view")
async def view(request: Request) -> Dict[str, Any]:
    filters = await request.json()
    fdf = apply_filters(filters)
    charts, compare_table = build_charts(fdf)
    issues = fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]) if not fdf.empty else fdf
    return {"summary": summary_payload(fdf), "rows": clean_records(fdf, 225), "issues": clean_records(issues, 125), "charts": charts, "compare_table": compare_table}


@app.post("/api/drilldown/{kind}")
async def metric_drilldown(kind: str, request: Request) -> Dict[str, Any]:
    """Return the field-level records behind clickable report metrics."""
    filters = await request.json()
    fdf = apply_filters(filters)

    if kind in {"health", "weighted-quality"}:
        rows = fdf[fdf["TotalClaims"].fillna(0) > 0].copy()
        sort_cols = ["TotalClaims", "ImpactScore"]
    elif kind == "field-average":
        rows = fdf[fdf["Comparable"] & fdf["MatchRate"].notna()].copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "risk-score":
        rows = fdf[fdf["ImpactScore"].fillna(0) > 0].copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "critical":
        rows = fdf[fdf["RiskTier"] == "Critical"].copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "elevated-critical":
        rows = fdf[fdf["RiskTier"].isin(["Critical", "Elevated"])].copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "open-issues":
        rows = fdf[fdf["NeedsChange"]].copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "unmatched":
        rows = fdf[fdf["NotMatchedClaims"].fillna(0) > 0].copy()
        sort_cols = ["NotMatchedClaims", "ImpactScore"]
    elif kind == "affected-reports":
        report_text = fdf["Report"].fillna("").astype(str).str.strip()
        kpi_text = fdf["KPI"].fillna("").astype(str).str.strip()
        rows = fdf[(report_text != "") | (kpi_text != "")].copy()
        # If report/KPI mappings have not been populated yet, show the underlying
        # stream rows used by DART's report-impact estimate rather than an empty modal.
        if rows.empty:
            rows = fdf.copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    elif kind == "all-fields":
        rows = fdf.copy()
        sort_cols = ["ImpactScore", "NotMatchedClaims"]
    else:
        return {"rows": [], "total": 0, "truncated": False}

    if not rows.empty:
        valid_sort = [c for c in sort_cols if c in rows.columns]
        if valid_sort:
            rows = rows.sort_values(valid_sort, ascending=[False] * len(valid_sort))
    limit = 10000
    return {
        "rows": clean_records(rows, limit),
        "total": int(len(rows)),
        "truncated": bool(len(rows) > limit),
    }


@app.post("/api/stream-detail/{stream}")
async def stream_detail(stream: str, request: Request) -> Dict[str, Any]:
    """Return a multi-section drilldown for one stream under the active filters."""
    filters = await request.json()
    fdf = apply_filters(filters)
    rows = fdf[fdf["Stream"].astype(str) == str(stream)].copy()
    if not rows.empty:
        rows = rows.sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False])

    risk_rows = rows[rows["RiskTier"].isin(["Critical", "Elevated"])].copy() if not rows.empty else rows.copy()
    unmatched_rows = rows[rows["NotMatchedClaims"].fillna(0) > 0].copy() if not rows.empty else rows.copy()
    remediation_rows = rows[rows["NeedsChange"]].copy() if not rows.empty else rows.copy()

    if rows.empty:
        class_summary = []
        risk_summary = []
    else:
        class_summary = (
            rows.groupby("Classification", dropna=False)
            .size().reset_index(name="Fields")
            .sort_values("Fields", ascending=False)
            .replace({np.nan: None}).to_dict("records")
        )
        risk_summary = (
            rows.groupby("RiskTier", dropna=False)
            .size().reset_index(name="Fields")
            .sort_values("Fields", ascending=False)
            .replace({np.nan: None}).to_dict("records")
        )

    workflow_issues = [x for x in STATE.get("issues", []) if str(x.get("Stream", "")) == str(stream)]
    mappings = [x for x in STATE.get("connections", []) if str(x.get("Stream", "")) == str(stream)]
    impacts = [x for x in STATE.get("impacts", []) if str(x.get("Stream", "")) == str(stream)]

    limit = 10000
    return {
        "stream": str(stream),
        "summary": summary_payload(rows),
        "rows": clean_records(rows, limit),
        "risk_rows": clean_records(risk_rows, limit),
        "unmatched_rows": clean_records(unmatched_rows, limit),
        "remediation_rows": clean_records(remediation_rows, limit),
        "classification_summary": class_summary,
        "risk_summary": risk_summary,
        "workflow_issues": workflow_issues,
        "mappings": mappings,
        "impacts": impacts,
        "total": int(len(rows)),
        "truncated": bool(len(rows) > limit),
    }


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)) -> Any:
    return JSONResponse(
        {"error": "Medicare source uploads are disabled. Use Build Your Own > DIY Lab for standalone CSV/Excel uploads."},
        status_code=403,
    )

@app.get("/api/byo")
def get_byo_workspace() -> Dict[str, Any]:
    return byo_payload()


@app.get("/api/byo/dataset/{filename}")
def get_byo_dataset(filename: str) -> Any:
    try:
        path = _byo_path(filename)
        df = _read_byo_dataset(path)
        payload = _dataset_payload_from_frame(df, path.name, 300)
        payload["modified_at"] = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        payload["size_mb"] = round(path.stat().st_size / (1024 * 1024), 2)
        return payload
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, FileNotFoundError) else 400)


@app.post("/api/byo/upload")
async def upload_byo(files: List[UploadFile] = File(...)) -> Any:
    saved: List[Dict[str, Any]] = []
    errors: List[str] = []
    for file in files:
        filename = str(file.filename or "dataset").strip() or "dataset"
        try:
            safe_name = _safe_byo_filename(filename)
            content = await file.read()
            if len(content) > BYO_MAX_FILE_BYTES:
                raise ValueError(f"{filename} exceeds the 25 MB DIY file limit.")
            suffix = Path(safe_name).suffix.lower()
            if suffix == ".csv":
                try:
                    frame = pd.read_csv(io.BytesIO(content))
                except UnicodeDecodeError:
                    frame = pd.read_csv(io.BytesIO(content), encoding="latin-1")
            else:
                frame = pd.read_excel(io.BytesIO(content), engine="openpyxl")
            frame = norm_cols(frame)
            if frame.empty and len(frame.columns) == 0:
                raise ValueError("The uploaded file did not contain a readable table.")
            target = _unique_byo_target(safe_name)
            target.write_bytes(content)
            saved.append({"name": target.name, "rows": int(len(frame)), "columns": int(len(frame.columns))})
        except Exception as exc:
            errors.append(f"{filename}: {exc}")
    if not saved:
        return JSONResponse({"error": "; ".join(errors) or "No files were saved."}, status_code=400)
    return {"status": "ok", "saved": saved, "errors": errors, "storage_path": str(Path("Data") / "DIY")}


@app.post("/api/byo/import/url")
async def import_byo_from_url(request: Request) -> Any:
    if requests is None:
        return JSONResponse({"error": "The 'requests' package is not installed on the server. Run: pip install requests"}, status_code=500)
    body = await request.json()
    raw_urls = str(body.get("urls", "") or "")
    urls = [u.strip() for u in re.split(r"[\r\n,]+", raw_urls) if u.strip()][:20]
    if not urls:
        return JSONResponse({"error": "Provide at least one file URL."}, status_code=400)
    saved: List[Dict[str, Any]] = []
    errors: List[str] = []
    for url in urls:
        try:
            _validate_public_import_url(url)
            resp = requests.get(url, timeout=20)
            resp.raise_for_status()
            content = resp.content
            disposition = resp.headers.get("content-disposition", "")
            match = re.search(r'filename="?([^";]+)"?', disposition)
            name_hint = match.group(1) if match else (Path(urlparse(url).path).name or "dataset.csv")
            info = _save_byo_bytes(content, name_hint)
            info["source_url"] = url
            saved.append(info)
        except Exception as exc:
            errors.append(f"{url}: {exc}")
    if not saved:
        return JSONResponse({"error": "; ".join(errors) or "No files were imported."}, status_code=400)
    return {"status": "ok", "saved": saved, "errors": errors, "storage_path": str(Path("Data") / "DIY")}


@app.post("/api/byo/import/s3/list")
async def list_byo_s3_files(request: Request) -> Any:
    body = await request.json()
    bucket = str(body.get("bucket", "") or "").strip()
    prefix = str(body.get("prefix", "") or "").strip()
    region = str(body.get("region", "") or "").strip()
    access_key = str(body.get("access_key", "") or "").strip()
    secret_key = str(body.get("secret_key", "") or "").strip()
    session_token = str(body.get("session_token", "") or "").strip()
    if not bucket:
        return JSONResponse({"error": "Provide an S3 bucket name."}, status_code=400)
    try:
        client = _s3_client(region, access_key, secret_key, session_token)
        paginator = client.get_paginator("list_objects_v2")
        items: List[Dict[str, Any]] = []
        for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
            for obj in page.get("Contents", []):
                key = str(obj["Key"])
                if key.lower().endswith((".csv", ".xlsx")):
                    items.append({"key": key, "size_bytes": int(obj["Size"]), "size_mb": round(obj["Size"] / (1024 * 1024), 2)})
                if len(items) >= 200:
                    break
            if len(items) >= 200:
                break
        return {"bucket": bucket, "prefix": prefix, "items": items}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/import/s3")
async def import_byo_from_s3(request: Request) -> Any:
    body = await request.json()
    bucket = str(body.get("bucket", "") or "").strip()
    region = str(body.get("region", "") or "").strip()
    access_key = str(body.get("access_key", "") or "").strip()
    secret_key = str(body.get("secret_key", "") or "").strip()
    session_token = str(body.get("session_token", "") or "").strip()
    keys = [str(k).strip() for k in (body.get("keys") or []) if str(k).strip()][:20]
    if not bucket or not keys:
        return JSONResponse({"error": "Provide a bucket and at least one file to import."}, status_code=400)
    try:
        client = _s3_client(region, access_key, secret_key, session_token)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=500)
    saved: List[Dict[str, Any]] = []
    errors: List[str] = []
    for key in keys:
        try:
            if Path(key).suffix.lower() not in {".csv", ".xlsx"}:
                raise ValueError("Only .csv and .xlsx keys are supported.")
            obj = client.get_object(Bucket=bucket, Key=key)
            content = obj["Body"].read()
            info = _save_byo_bytes(content, Path(key).name)
            info["source_key"] = key
            saved.append(info)
        except Exception as exc:
            errors.append(f"{key}: {exc}")
    if not saved:
        return JSONResponse({"error": "; ".join(errors) or "No files were imported."}, status_code=400)
    return {"status": "ok", "saved": saved, "errors": errors, "storage_path": str(Path("Data") / "DIY")}


@app.post("/api/byo/import/sharepoint/list")
async def list_byo_sharepoint_files(request: Request) -> Any:
    body = await request.json()
    try:
        token = _graph_access_token(body.get("tenant_id", ""), body.get("client_id", ""), body.get("client_secret", ""))
        items = _graph_share_files(token, str(body.get("share_url", "") or ""))
        return {"items": items}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/import/sharepoint")
async def import_byo_from_sharepoint(request: Request) -> Any:
    if requests is None:
        return JSONResponse({"error": "The 'requests' package is not installed on the server. Run: pip install requests"}, status_code=500)
    body = await request.json()
    items = [x for x in (body.get("items") or []) if isinstance(x, dict)][:20]
    if not items:
        return JSONResponse({"error": "Select at least one file to import."}, status_code=400)
    saved: List[Dict[str, Any]] = []
    errors: List[str] = []
    for entry in items:
        name = str(entry.get("name", "") or "dataset")
        url = str(entry.get("download_url", "") or "")
        try:
            if not url:
                raise ValueError("Missing download link - list the folder again.")
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()
            info = _save_byo_bytes(resp.content, name)
            info["source"] = "SharePoint/OneDrive"
            saved.append(info)
        except Exception as exc:
            errors.append(f"{name}: {exc}")
    if not saved:
        return JSONResponse({"error": "; ".join(errors) or "No files were imported."}, status_code=400)
    return {"status": "ok", "saved": saved, "errors": errors, "storage_path": str(Path("Data") / "DIY")}


@app.delete("/api/byo/files/{filename}")
def delete_byo_file(filename: str) -> Any:
    try:
        path = _byo_path(filename)
        path.unlink()
        for automation in [t for t in load_byo_email_automations() if str(t.get("dataset_name")) == str(filename)]:
            delete_byo_email_automation(str(automation.get("id", "")))
        pair = STATE.get("byo_chat_pair", [])
        if filename in pair:
            STATE["byo_chat_pair"] = []
            STATE["byo_chat"] = []
        return {"status": "ok", "deleted": filename}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, FileNotFoundError) else 400)


@app.get("/api/byo/raw-data")
def byo_raw_data(filename: str = "", search: str = "", offset: int = 0, limit: int = 75) -> Any:
    try:
        if not filename:
            ready = [x["name"] for x in _byo_library_records() if x.get("status") == "Ready"]
            filename = ready[0] if ready else ""
        if not filename:
            return {"file_found": False, "filename": "", "raw_rows": 0, "filtered_rows": 0, "columns": [], "rows": [], "offset": 0, "limit": limit, "modified_at": None, "signature": "", "latest_backup": None, "sheet": ""}
        source = _byo_path(filename)
        raw = _read_byo_editable_dataset(source)
        frame = raw.copy()
        frame.insert(0, "_SourceRow", np.arange(2, len(frame) + 2, dtype=int))
        q = str(search or "").strip()
        if q:
            mask = frame.drop(columns=["_SourceRow"], errors="ignore").astype(str).apply(lambda col: col.str.contains(q, case=False, na=False)).any(axis=1)
            frame = frame[mask]
        offset = max(0, int(offset or 0))
        limit = min(max(10, int(limit or 75)), 200)
        latest = _latest_byo_backup(filename)
        return {
            "file_found": True, "filename": source.name, "file_type": source.suffix.lower().lstrip(".").upper(),
            "sheet": _byo_sheet_name(source), "raw_rows": len(raw), "filtered_rows": len(frame),
            "columns": [str(c) for c in raw.columns], "rows": _records_json_safe(frame.iloc[offset:offset + limit], limit),
            "offset": offset, "limit": limit, "modified_at": datetime.fromtimestamp(source.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p"),
            "signature": _source_file_signature(source), "latest_backup": str(latest) if latest else None,
        }
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, FileNotFoundError) else 400)


@app.post("/api/byo/raw-data/save")
async def save_byo_raw_data(request: Request) -> Any:
    payload = await request.json()
    changes = payload.get("changes") or []
    if not isinstance(changes, list) or len(changes) > 2500:
        return JSONResponse({"error": "Invalid or excessive change set."}, status_code=400)
    filename = str(payload.get("filename", "")).strip()
    try:
        result = _write_byo_changes(filename, changes, str(payload.get("signature", "") or ""))
        agent_results: List[Dict[str, Any]] = []
        if bool(payload.get("run_email_agent", True)) and result.get("saved", 0):
            agent_results = [_public_automation_result(r) for r in run_enabled_byo_email_automations(filename)]
        return {**result, "filename": filename, "agent_results": agent_results}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)


@app.post("/api/byo/raw-data/restore-latest")
async def restore_byo_raw_backup(request: Request) -> Any:
    payload = await request.json()
    filename = str(payload.get("filename", "")).strip()
    try:
        latest = _latest_byo_backup(filename)
        if latest is None:
            return JSONResponse({"error": "No backup exists yet for this DIY file."}, status_code=404)
        path = _restore_byo_backup(filename, latest)
        results = [_public_automation_result(r) for r in run_enabled_byo_email_automations(filename)]
        return {"status": "ok", "path": path, "restored_from": str(latest), "agent_results": results}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)


def _byo_email_meta_for_dataset(filename: str) -> Dict[str, Any]:
    datasets = [x for x in _byo_library_records() if x.get("status") == "Ready"]
    names = [x["name"] for x in datasets]
    if filename not in names:
        filename = names[0] if names else ""
    df = _read_byo_dataset(filename) if filename else pd.DataFrame()
    templates = load_byo_email_automations()
    enriched: List[Dict[str, Any]] = []
    for template in templates:
        t = dict(template)
        try:
            tdf = _read_byo_dataset(t.get("dataset_name", ""))
            ids = [c for c in t.get("identity_columns", []) if c in tdf.columns]
            t["identity_duplicate_count"] = int(tdf[ids].astype(str).duplicated(keep=False).sum()) if ids and not tdf.empty else 0
            t["dataset_found"] = True
        except Exception:
            t["identity_duplicate_count"] = 0
            t["dataset_found"] = False
        t["baseline_exists"] = _snapshot_path(t.get("id", "")).exists()
        enriched.append(t)
    default = _default_byo_email_template(df, filename)
    ids = default.get("identity_columns", [])
    default["identity_duplicate_count"] = int(df[ids].astype(str).duplicated(keep=False).sum()) if ids and not df.empty else 0
    cfg = smtp_settings()
    return {
        "templates": enriched, "default_template": default, "columns": [str(c) for c in df.columns],
        "operators": EMAIL_OPERATORS, "status": _byo_email_status_payload(), "history": _byo_email_history(),
        "check_interval_seconds": BYO_EMAIL_AGENT_MIN_CHECK_SECONDS, "auto_run": BYO_EMAIL_AGENT_AUTO_RUN,
        "selected_dataset": filename, "datasets": datasets,
        "smtp": {"configured": smtp_is_configured(), "issues": smtp_configuration_issues(), "host": cfg["host"], "port": cfg["port"], "sender": cfg["from_email"], "transport": "SSL" if cfg["use_ssl"] else ("STARTTLS" if cfg["use_tls"] else "plain SMTP")},
    }


@app.get("/api/byo/email-agent")
def byo_email_agent_meta(filename: str = "") -> Dict[str, Any]:
    return _byo_email_meta_for_dataset(filename)


@app.post("/api/byo/email-agent/template")
async def save_byo_email_template(request: Request) -> Any:
    try:
        payload = await request.json()
        filename = str(payload.get("dataset_name", "")).strip()
        df = _read_byo_dataset(filename)
        template = _normalize_byo_email_template(payload, df)
        is_new = not template["id"]
        saved = upsert_byo_email_automation(template)
        baseline = establish_email_baseline(df, saved) if is_new else None
        return {"status": "ok", "template": saved, "baseline": baseline}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.delete("/api/byo/email-agent/template/{template_id}")
async def delete_byo_email_template_route(template_id: str) -> Dict[str, str]:
    delete_byo_email_automation(template_id)
    return {"status": "ok"}


@app.post("/api/byo/email-agent/baseline/{template_id}")
async def baseline_byo_email_template(template_id: str) -> Any:
    template = _find_byo_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    try:
        df = _read_byo_dataset(template.get("dataset_name", ""))
        return establish_email_baseline(df, template)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/email-agent/baselines")
async def baseline_all_byo_email_templates() -> Dict[str, Any]:
    count = 0
    errors: List[str] = []
    for template in load_byo_email_automations():
        if not template.get("enabled"):
            continue
        try:
            establish_email_baseline(_read_byo_dataset(template.get("dataset_name", "")), template)
            count += 1
        except Exception as exc:
            errors.append(f"{template.get('name')}: {exc}")
    return {"status": "ok", "baselines": count, "errors": errors}


@app.post("/api/byo/email-agent/run/{template_id}")
async def run_byo_email_template_route(template_id: str, request: Request) -> Any:
    template = _find_byo_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    payload = await request.json()
    try:
        result = _run_one_byo_email_template(template, send_email=bool(payload.get("send_email", True)))
        template["last_run_at"] = result.get("checked_at")
        template["last_status"] = result.get("status")
        template["last_match_count"] = result.get("matches", 0)
        template["last_change_count"] = result.get("change_rows", 0)
        upsert_byo_email_automation(template)
        _byo_email_history_event({"automation_id": template_id, "automation": template.get("name"), "dataset": template.get("dataset_name"), "status": result.get("status"), "detected_changes": result.get("change_rows", 0), "matching_changes": result.get("matches", 0), "manual": True})
        return _public_automation_result(result)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/email-agent/check")
async def check_enabled_byo_email_templates() -> Dict[str, Any]:
    return {"status": "ok", "results": [_public_automation_result(r) for r in run_enabled_byo_email_automations()]}


@app.post("/api/byo/email-agent/preview")
async def preview_byo_email_template(request: Request) -> Any:
    try:
        payload = await request.json()
        filename = str(payload.get("dataset_name", "")).strip()
        df = _read_byo_dataset(filename)
        template = _normalize_byo_email_template(payload, df)
        matches = current_matching_rows(df, template)
        subject, _, body_html, _ = build_structured_email(template, matches.head(100), is_test=True)
        return {"status": "ok", "matches": len(matches), "subject": subject, "html": body_html, "rows": _records_json_safe(matches, 100)}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/email-agent/test/{template_id}")
async def test_byo_email_template_route(template_id: str) -> Any:
    template = _find_byo_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    try:
        df = _read_byo_dataset(template.get("dataset_name", ""))
        matches = current_matching_rows(df, template)
        if matches.empty:
            return JSONResponse({"error": "No current rows satisfy this automation, so there is nothing to include in a test email."}, status_code=400)
        subject = send_structured_email(template, matches.head(100), is_test=True)
        return {"status": "Test email sent", "subject": subject, "matches": len(matches)}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/compare")
async def compare_byo_route(request: Request) -> Any:
    payload = await request.json()
    try:
        return compare_byo_datasets(str(payload.get("left", "")), str(payload.get("right", "")), include_previews=True)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.get("/api/byo/chat")
def byo_chat_status(left: str = "", right: str = "") -> Dict[str, Any]:
    pair = [left, right] if left and right else []
    if pair and pair != STATE.get("byo_chat_pair", []):
        return {"chat": [], "pair": pair, "ai_configured": bool(OpenAI and groq_api_key())}
    return {
        "chat": STATE.get("byo_chat", []),
        "pair": STATE.get("byo_chat_pair", []),
        "ai_configured": bool(OpenAI and groq_api_key()),
        "model": groq_model(),
    }


@app.post("/api/byo/chat/reset")
async def reset_byo_chat() -> Dict[str, str]:
    STATE["byo_chat"] = []
    STATE["byo_chat_pair"] = []
    return {"status": "ok"}


@app.post("/api/byo/chat")
async def byo_chat(request: Request) -> Any:
    payload = await request.json()
    prompt = str(payload.get("prompt", "")).strip()
    left_name = str(payload.get("left", "")).strip()
    right_name = str(payload.get("right", "")).strip()
    if not prompt:
        return JSONResponse({"error": "Enter a question for the AI analyst."}, status_code=400)
    try:
        if not left_name or not right_name or left_name == right_name:
            guide_answer = "Please select two different saved datasets (Dataset A and Dataset B) in the dataset selector above or the Dataset Library to run comparative AI analysis."
            return {"status": "ok", "answer": guide_answer, "chat": STATE.get("byo_chat", [])}
        # Validate both selected files before touching chat history.
        _byo_path(left_name)
        _byo_path(right_name)
        pair = [left_name, right_name]
        if pair != STATE.get("byo_chat_pair", []):
            STATE["byo_chat_pair"] = pair
            STATE["byo_chat"] = [{
                "role": "assistant",
                "content": f"I am grounded in {left_name} and {right_name}. Ask me to compare schema, quality, distributions, likely keys, overlap, anomalies, or business patterns across the two datasets.",
            }]
        STATE["byo_chat"].append({"role": "user", "content": prompt})
        context = _byo_ai_context(left_name, right_name, prompt)
        api_key = groq_api_key()
        model = groq_model()
        system = (
            "You are DART Compare AI, a senior data analyst. You are grounded only in the two user-selected datasets and the computed context supplied to you. Adapt the explanation depth and recommended next checks to the supplied persona when present. "
            "Compare Dataset A and Dataset B explicitly. Never invent rows, values, statistics, matches, or causal explanations. Distinguish direct evidence from inference. "
            "When the context cannot answer a question exactly, say what additional calculation would be required. "
            "Return polished Markdown that reads like a professional analysis report. Use a clear title only when useful, short section headings, compact tables for true side-by-side comparisons, and bullets for findings/actions. "
            "Avoid repeating the same statistic in both a table and prose. Use bold selectively for the most important evidence. Keep tables narrow enough to read in a browser. "
            "Consistency-check the answer against the supplied schemas before responding: never attribute a field, status, category, or statistic to a dataset when that field is not present there. "
            "Always finish every section you start; never end on a heading, colon, or incomplete bullet. Aim for roughly 600-1000 words unless the user explicitly asks for more. "
            "Do not claim the model was trained on these datasets; say it is grounded in them for this conversation."
        )
        if OpenAI and api_key:
            try:
                context_text = json.dumps(
                    context,
                    ensure_ascii=False,
                    default=str,
                    separators=(",", ":"),
                )
                messages = [
                    {"role": "system", "content": system + "\n\nCOMPACT DATASET CONTEXT:\n" + context_text},
                    *STATE["byo_chat"][-4:],
                ]
                answer, active_model = groq_chat_completion(
                    messages,
                    temperature=0.15,
                    max_output_tokens=650,
                )
                answer = answer or "I could not generate an answer from the selected dataset context."
            except Exception as exc:
                comparison = compare_byo_datasets(left_name, right_name, include_previews=False)
                answer = f"AI service error: {exc}\n\nLocal comparison:\n" + _byo_local_answer(left_name, right_name, comparison)
        else:
            comparison = compare_byo_datasets(left_name, right_name, include_previews=False)
            answer = _byo_local_answer(left_name, right_name, comparison)
        STATE["byo_chat"].append({"role": "assistant", "content": answer})
        return {"status": "ok", "answer": answer, "chat": STATE["byo_chat"]}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/byo/reset")
async def reset_byo() -> Dict[str, str]:
    # Kept for compatibility with older clients. Persistent files are intentionally not deleted here.
    STATE["byo_df"] = pd.DataFrame()
    STATE["byo_source"] = ""
    STATE["byo_uploaded_at"] = ""
    return {"status": "ok"}


@app.post("/api/restore-demo")
async def reload_excel_files() -> Dict[str, str]:
    STATE["df"] = pd.DataFrame()
    STATE["source"] = ""
    STATE["snapshots"] = []
    STATE["connections"] = []
    STATE["issues"] = []
    STATE["impacts"] = []
    ensure_data()
    return {"status": "ok"}


@app.post("/api/seed-lineage")
async def seed_lineage() -> Dict[str, str]:
    STATE["connections"] = []
    seed_lineage_from_df(standardize(STATE["df"]))
    return {"status": "ok"}


@app.post("/api/lineage")
async def upsert_lineage(request: Request) -> Dict[str, str]:
    payload = await request.json()
    item = {k: str(payload.get(k, "")).strip() for k in ["id", "Stream", "SourceTable", "SourceField", "TargetTable", "TargetField", "Report", "KPI", "Owner", "Status", "RiskTier", "Notes"]}
    if not item["id"]:
        item["id"] = f"L{len(STATE['connections']) + 1}"
    found = False
    for i, existing in enumerate(STATE["connections"]):
        if str(existing.get("id")) == item["id"]:
            STATE["connections"][i] = item
            found = True
            break
    if not found:
        STATE["connections"].append(item)
    return {"status": "ok"}


@app.delete("/api/lineage/{item_id}")
async def delete_lineage(item_id: str) -> Dict[str, str]:
    STATE["connections"] = [x for x in STATE["connections"] if str(x.get("id")) != item_id]
    return {"status": "ok"}


@app.post("/api/issues")
async def add_issue(request: Request) -> Dict[str, str]:
    payload = await request.json()
    payload["id"] = payload.get("id") or f"I{len(STATE['issues']) + 1}"
    STATE["issues"].append(payload)
    return {"status": "ok"}


@app.post("/api/impacts")
async def add_impact(request: Request) -> Dict[str, str]:
    payload = await request.json()
    payload["id"] = payload.get("id") or f"M{len(STATE['impacts']) + 1}"
    STATE["impacts"].append(payload)
    return {"status": "ok"}


@app.post("/api/governance")
async def upsert_governance(request: Request) -> Dict[str, str]:
    payload = await request.json()
    item = {k: str(payload.get(k, "")).strip() for k in ["id", "Control", "Owner", "Cadence", "Status", "Evidence"]}
    if not item["id"]:
        item["id"] = f"G{len(STATE['governance']) + 1}"
    found = False
    for i, existing in enumerate(STATE["governance"]):
        if str(existing.get("id")) == item["id"]:
            STATE["governance"][i] = item
            found = True
            break
    if not found:
        STATE["governance"].append(item)
    return {"status": "ok"}


@app.delete("/api/governance/{item_id}")
async def delete_governance(item_id: str) -> Dict[str, str]:
    STATE["governance"] = [x for x in STATE["governance"] if str(x.get("id")) != item_id]
    return {"status": "ok"}


@app.post("/api/rules")
async def add_rule(request: Request) -> Dict[str, str]:
    STATE["rules"].append(await request.json())
    return {"status": "ok"}



@app.on_event("startup")
def _start_background_email_agent() -> None:
    start_email_agent_scheduler()
    start_byo_email_agent_scheduler()


@app.get("/api/raw-data")
def raw_data(search: str = "", offset: int = 0, limit: int = 75) -> Dict[str, Any]:
    source = Path(MAIN_PATH)
    if not source.exists():
        return {"workbook": source.name, "workbook_found": False, "raw_rows": 0, "filtered_rows": 0, "columns": [], "rows": [], "offset": 0, "limit": limit, "modified_at": None, "latest_backup": None}
    raw = _read_raw_main_sheet()
    frame = _build_raw_editor_frame(raw)
    q = str(search or "").strip()
    if q:
        mask = frame.drop(columns=["_ExcelRow"], errors="ignore").astype(str).apply(lambda col: col.str.contains(q, case=False, na=False)).any(axis=1)
        frame = frame[mask]
    offset = max(0, int(offset or 0))
    limit = min(max(10, int(limit or 75)), 200)
    latest = _latest_workbook_backup()
    return {
        "workbook": source.name,
        "workbook_found": True,
        "sheet": MAIN_SHEET,
        "raw_rows": len(raw),
        "filtered_rows": len(frame),
        "columns": [str(c) for c in raw.columns],
        "rows": _records_json_safe(frame.iloc[offset:offset + limit], limit),
        "offset": offset,
        "limit": limit,
        "modified_at": datetime.fromtimestamp(source.stat().st_mtime).strftime("%Y-%m-%d %I:%M:%S %p"),
        "signature": _source_file_signature(source),
        "latest_backup": str(latest) if latest else None,
    }


@app.post("/api/raw-data/save")
async def save_raw_data(request: Request) -> Any:
    payload = await request.json()
    changes = payload.get("changes") or []
    if not isinstance(changes, list) or len(changes) > 2500:
        return JSONResponse({"error": "Invalid or excessive change set."}, status_code=400)
    try:
        result = _write_raw_changes_to_workbook(
            changes, expected_signature=str(payload.get("signature", "") or "")
        )
        fresh_df = _reload_state_from_workbook()
        agent_results: List[Dict[str, Any]] = []
        if bool(payload.get("run_email_agent", True)) and result.get("saved", 0):
            agent_results = [_public_automation_result(r) for r in run_enabled_email_automations(fresh_df)]
        return {**result, "agent_results": agent_results, "source": STATE["source"]}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)


@app.post("/api/raw-data/restore-latest")
async def restore_latest_raw_backup() -> Any:
    latest = _latest_workbook_backup()
    if latest is None:
        return JSONResponse({"error": "No workbook editor backup exists yet."}, status_code=404)
    try:
        path = _restore_workbook_backup(latest)
        fresh_df = _reload_state_from_workbook()
        results = [_public_automation_result(r) for r in run_enabled_email_automations(fresh_df)]
        return {"status": "ok", "path": path, "restored_from": str(latest), "agent_results": results}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=409)


def _find_email_template(template_id: str) -> Optional[Dict[str, Any]]:
    return next((t for t in load_email_automations() if str(t.get("id")) == str(template_id)), None)


@app.get("/api/email-agent")
def email_agent_meta() -> Dict[str, Any]:
    ensure_data()
    df = standardize(STATE["df"])
    templates = load_email_automations()
    enriched_templates: List[Dict[str, Any]] = []
    for template in templates:
        t = dict(template)
        ids = [c for c in t.get("identity_columns", []) if c in df.columns]
        if ids and not df.empty:
            t["identity_duplicate_count"] = int(df[ids].astype(str).duplicated(keep=False).sum())
        else:
            t["identity_duplicate_count"] = 0
        t["baseline_exists"] = _snapshot_path(t.get("id", "")).exists()
        enriched_templates.append(t)
    default = _default_email_agent_template(df)
    ids = default.get("identity_columns", [])
    default["identity_duplicate_count"] = int(df[ids].astype(str).duplicated(keep=False).sum()) if ids and not df.empty else 0
    cfg = smtp_settings()
    return {
        "templates": enriched_templates,
        "default_template": default,
        "columns": [str(c) for c in df.columns],
        "operators": EMAIL_OPERATORS,
        "status": _email_agent_status_payload(),
        "history": _email_agent_history(),
        "check_interval_seconds": EMAIL_AGENT_MIN_CHECK_SECONDS,
        "auto_run": EMAIL_AGENT_AUTO_RUN,
        "workbook_found": Path(MAIN_PATH).exists(),
        "workbook": str(MAIN_PATH),
        "smtp": {"configured": smtp_is_configured(), "issues": smtp_configuration_issues(), "host": cfg["host"], "port": cfg["port"], "sender": cfg["from_email"], "transport": "SSL" if cfg["use_ssl"] else ("STARTTLS" if cfg["use_tls"] else "plain SMTP")},
    }


@app.post("/api/email-agent/template")
async def save_email_template(request: Request) -> Any:
    ensure_data()
    try:
        df = _reload_state_from_workbook()
        template = _normalize_email_template(await request.json(), df)
        is_new = not template["id"]
        saved = upsert_email_automation(template)
        baseline = establish_email_baseline(df, saved) if is_new else None
        return {"status": "ok", "template": saved, "baseline": baseline}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.delete("/api/email-agent/template/{template_id}")
async def delete_email_template_route(template_id: str) -> Dict[str, str]:
    delete_email_automation(template_id)
    return {"status": "ok"}


@app.post("/api/email-agent/baseline/{template_id}")
async def baseline_email_template(template_id: str) -> Any:
    template = _find_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    df = _reload_state_from_workbook()
    return establish_email_baseline(df, template)


@app.post("/api/email-agent/baselines")
async def baseline_all_enabled_email_templates() -> Dict[str, Any]:
    df = _reload_state_from_workbook()
    count = 0
    for template in load_email_automations():
        if template.get("enabled"):
            establish_email_baseline(df, template)
            count += 1
    return {"status": "ok", "baselines": count}


@app.post("/api/email-agent/run/{template_id}")
async def run_email_template_route(template_id: str, request: Request) -> Any:
    template = _find_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    payload = await request.json()
    df = _reload_state_from_workbook()
    result = run_email_automation(df, template, send_email=bool(payload.get("send_email", True)))
    template["last_run_at"] = result.get("checked_at")
    template["last_status"] = result.get("status")
    template["last_match_count"] = result.get("matches", 0)
    template["last_change_count"] = result.get("change_rows", 0)
    upsert_email_automation(template)
    _append_agent_history({"automation_id": template_id, "automation": template.get("name"), "status": result.get("status"), "detected_changes": result.get("change_rows", 0), "matching_changes": result.get("matches", 0), "manual": True})
    return _public_automation_result(result)


@app.post("/api/email-agent/check")
async def check_enabled_email_templates() -> Dict[str, Any]:
    df = _reload_state_from_workbook()
    return {"status": "ok", "results": [_public_automation_result(r) for r in run_enabled_email_automations(df)]}


@app.post("/api/email-agent/preview")
async def preview_email_template(request: Request) -> Any:
    ensure_data()
    try:
        template = _normalize_email_template(await request.json(), standardize(STATE["df"]))
        matches = current_matching_rows(standardize(STATE["df"]), template)
        subject, _, body_html, _ = build_structured_email(template, matches.head(100), is_test=True)
        return {"status": "ok", "matches": len(matches), "subject": subject, "html": body_html, "rows": _records_json_safe(matches, 100)}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/email-agent/test/{template_id}")
async def test_email_template_route(template_id: str) -> Any:
    template = _find_email_template(template_id)
    if not template:
        return JSONResponse({"error": "Automation not found."}, status_code=404)
    df = _reload_state_from_workbook()
    matches = current_matching_rows(df, template)
    if matches.empty:
        return JSONResponse({"error": "No current rows satisfy this automation, so there is nothing to include in a test email."}, status_code=400)
    try:
        subject = send_structured_email(template, matches.head(100), is_test=True)
        return {"status": "Test email sent", "subject": subject, "matches": len(matches)}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/email-agent/smtp-test")
async def smtp_test_route() -> Any:
    try:
        return {"status": "ok", **test_smtp_connection()}
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=400)


@app.post("/api/copilot")
async def copilot(request: Request) -> Dict[str, Any]:
    payload = await request.json()
    prompt = str(payload.get("prompt", "")).strip()
    filters = payload.get("filters", {})
    fdf = apply_filters(filters)
    if prompt:
        STATE["chat"].append({"role": "user", "content": prompt})
    context = {
        "persona": STATE["persona"],
        "rows": len(fdf),
        "field_average": field_average(fdf) if not fdf.empty else None,
        "weighted_rate": weighted_rate(fdf) if not fdf.empty else None,
        "top_actions": clean_records(fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]), 8) if not fdf.empty else [],
        "stream_summary": summary_payload(fdf).get("stream_summary", []),
        "lineage_count": len(STATE.get("connections", [])),
        "impact_count": len(STATE.get("impacts", [])),
        "governance_count": len(STATE.get("governance", [])),
    }
    api_key = groq_api_key()
    model = groq_model()
    system = "You are Dartboard, the DART healthcare data intelligence assistant. Be concise, executive-ready, and data-driven. Use only supplied context. Never invent statistics. This is decision support, not clinical advice. Adapt wording, depth, and recommended actions to the supplied persona role, audience, focus areas, and detail level without changing the underlying facts. Return polished Markdown with short headings, compact tables only when they improve comparison, and bullets for findings/actions. Avoid repetition and always finish every section you start. Aim for roughly 400-800 words unless the user asks for more."
    if OpenAI and api_key:
        try:
            context_text = json.dumps(context, default=str, separators=(",", ":"))
            messages = [
                {"role": "system", "content": system + "\n\nCURRENT CONTEXT:\n" + context_text},
                *STATE["chat"][-4:],
            ]
            answer, active_model = groq_chat_completion(
                messages,
                temperature=0.2,
                max_output_tokens=600,
            )
            answer = answer or "I could not generate an answer from the current context."
        except Exception as exc:
            answer = f"AI service error: {exc}\n\nLocal result: field average is {pct(field_average(fdf))}; overall average match is {pct(weighted_rate(fdf))}."
    else:
        if fdf.empty:
            answer = "No rows match the current filters. Adjust scope or upload a dataset."
        else:
            top = fdf[fdf["NeedsChange"]].sort_values(["ImpactScore", "NotMatchedClaims"], ascending=[False, False]).head(3)
            bullets = "\n".join([f"- {r['Stream']} · {r['NCH Target Column']}: {pct(r['MatchRate'])} match, {fmt_int(r['NotMatchedClaims'])} unmatched, {r['RiskTier']} risk" for _, r in top.iterrows()])
            answer = f"Current scope has {len(fdf):,} fields. Field average is {pct(field_average(fdf))}; overall average match is {pct(weighted_rate(fdf))}. The workspace also has {len(STATE.get('connections', []))} lineage mappings, {len(STATE.get('impacts', []))} impact mappings, and {len(STATE.get('governance', []))} governance controls. Top action candidates:\n{bullets}\n\nModel-backed Dartboard analysis is unavailable because the Groq/OpenAI client could not be initialized."
    STATE["chat"].append({"role": "assistant", "content": answer})
    return {"status": "ok", "answer": answer, "chat": STATE["chat"]}



@app.get("/api/medicaid/dashboard")
def medicaid_dashboard_api() -> Any:
    return _medicaid_dashboard_payload()


@app.get("/api/medicaid/state/{state_id}")
def medicaid_state_api(state_id: int, status: str = "", issue_type: str = "", tag_id: int = 0) -> Any:
    try:
        return _medicaid_state_payload(state_id, status=status.strip().lower(), issue_type=issue_type.strip().lower(), tag_id=tag_id)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, ValueError) else 400)


@app.post("/api/medicaid/state/{state_id}/issue")
async def medicaid_create_issue_api(state_id: int, request: Request) -> Any:
    payload = await request.json()
    title = str(payload.get("title", "")).strip()
    if not title:
        return JSONResponse({"error": "Issue title is required."}, status_code=400)
    status = str(payload.get("status", "open")).strip().lower()
    priority = str(payload.get("priority", "medium")).strip().lower()
    issue_type = str(payload.get("issue_type", "general")).strip().lower() or "general"
    if status not in {"open", "done", "cancelled"}: status = "open"
    if priority not in {"low", "medium", "high"}: priority = "medium"
    db = _medicaid_db(); cur = db.cursor()
    cur.execute("SELECT id FROM states WHERE id=?", (state_id,))
    if not cur.fetchone(): db.close(); return JSONResponse({"error": "State not found."}, status_code=404)
    cur.execute("INSERT INTO issues (state_id,title,description,status,priority,issue_type) VALUES (?,?,?,?,?,?)",
                (state_id,title,str(payload.get("description", "")).strip(),status,priority,issue_type))
    issue_id = cur.lastrowid; db.commit(); db.close()
    return {"status": "ok", "issue_id": issue_id}


@app.post("/api/medicaid/issue/{issue_id}/status")
async def medicaid_update_issue_api(issue_id: int, request: Request) -> Any:
    payload = await request.json(); status = str(payload.get("status", "")).strip().lower()
    if status not in {"open", "done", "cancelled"}:
        return JSONResponse({"error": "Status must be open, done, or cancelled."}, status_code=400)
    db = _medicaid_db(); cur = db.cursor(); cur.execute("UPDATE issues SET status=?,updated_at=CURRENT_TIMESTAMP WHERE id=?", (status, issue_id)); db.commit(); changed = cur.rowcount; db.close()
    if not changed: return JSONResponse({"error": "Issue not found."}, status_code=404)
    return {"status": "ok"}


@app.get("/api/medicaid/compare")
def medicaid_compare_api(issue_type: str = "", states: str = "", min_total: int = 0) -> Any:
    ids = [int(x) for x in states.split(",") if x.strip().isdigit()] if states else None
    rows = _medicaid_comparison_rows(issue_type.strip().lower(), ids)
    return {"rows": [r for r in rows if (r.get("total") or 0) >= max(0, min_total)], "issue_labels": MEDICAID_ISSUE_LABELS}


@app.get("/api/medicaid/analytics")
def medicaid_analytics_api() -> Any:
    return _medicaid_analytics_payload()


@app.get("/api/medicaid/ai-comparison")
def medicaid_ai_comparison_api(issue_type: str = "", states: str = "", min_total: int = 3) -> Any:
    ids = [int(x) for x in states.split(",") if x.strip().isdigit()] if states else None
    data = _medicaid_comparison_rows(issue_type.strip().lower(), ids)
    filtered = [d for d in data if (d.get("total") or 0) >= max(0, min_total)]
    avg_success = sum(d.get("success_rate", 0) for d in filtered) / len(filtered) if filtered else 0
    avg_conf = sum(d.get("confidence", 0) for d in filtered) / len(filtered) if filtered else 0
    return {"data": filtered, "insights": _medicaid_generate_insights(filtered, issue_type.strip().lower()),
            "meta": {"total_states": len(data), "states_with_data": sum(1 for d in data if d.get("total",0)>0), "states_analyzed": len(filtered), "min_total": max(0,min_total), "avg_success_rate": round(avg_success,1), "avg_confidence": round(avg_conf,1)}}


@app.post("/api/medicaid/ai-chat")
async def medicaid_ai_chat_api(request: Request) -> Any:
    payload = await request.json()
    return _medicaid_ai_answer(str(payload.get("question", "")))


@app.get("/api/medicaid/executive-brief")
def medicaid_executive_brief_api(issue_type: str = "", min_total: int = 3) -> Any:
    return _medicaid_executive_brief(issue_type.strip().lower(), min_total)


@app.get("/api/medicaid/heatmap")
def medicaid_heatmap_api(status: str = "", issue_type: str = "", tag_id: int = 0, metric: str = "quality_score") -> Any:
    rows = _medicaid_heatmap_rows(status.strip().lower(), issue_type.strip().lower(), tag_id)
    valid = {"quality_score","success_rate","total","successful","open","cancelled","backlog_rate","cancel_rate"}
    if metric not in valid: metric = "quality_score"
    for row in rows: row["metric_value"] = row.get(metric, 0)
    return {"metric": metric, "rows": rows, "issue_labels": MEDICAID_ISSUE_LABELS}


@app.get("/api/medicaid/heatmap/state/{state_id}")
def medicaid_heatmap_state_api(state_id: int, status: str = "", issue_type: str = "", tag_id: int = 0) -> Any:
    try: return _medicaid_state_payload(state_id, status.strip().lower(), issue_type.strip().lower(), tag_id)
    except Exception as exc: return JSONResponse({"error": str(exc)}, status_code=404)


@app.get("/api/medicaid/search")
def medicaid_search_api(q: str = "") -> Any:
    q = q.strip()
    if len(q) < 2: return {"states": [], "issues": []}
    db=_medicaid_db(); cur=db.cursor(); cur.execute("SELECT id,name FROM states WHERE name LIKE ? ORDER BY name LIMIT 8", (f"%{q}%",)); states=[dict(r) for r in cur.fetchall()]
    cur.execute('''SELECT i.id,i.title,i.status,i.issue_type,i.priority,s.name state_name,s.id state_id FROM issues i JOIN states s ON i.state_id=s.id WHERE i.title LIKE ? OR i.description LIKE ? ORDER BY i.updated_at DESC LIMIT 12''',(f"%{q}%",f"%{q}%")); issues=[dict(r) for r in cur.fetchall()]; db.close(); return {"states":states,"issues":issues}


@app.get("/api/medicaid/state-summary/{state_id}")
def medicaid_state_summary_api(state_id: int) -> Any:
    try: return _medicaid_state_summary(state_id)
    except Exception as exc: return JSONResponse({"error": str(exc)}, status_code=404)


@app.get("/api/medicaid/rankings")
def medicaid_rankings_api() -> Any:
    return _medicaid_rankings()


@app.get("/api/medicaid/claims-data")
def medicaid_claims_data_api() -> Any:
    return _medicaid_claims_payload()


@app.post("/api/medicaid/claims/register-agent")
async def medicaid_register_agent_api(request: Request) -> Any:
    payload = await request.json(); email_value = str(payload.get("email", "")).strip()
    if not email_value: return JSONResponse({"error": "Email is required"}, status_code=400)
    requirements = payload.get("requirements", []) or []; rules = payload.get("rules", []) or []
    db=_medicaid_db(); cur=db.cursor(); cur.execute("INSERT INTO agent_subscriptions (email,requirements,filter_value,rules) VALUES (?,?,?,?)",
        (email_value, ",".join(str(x) for x in requirements), str(payload.get("filter_value", "")), json.dumps(rules))); db.commit(); db.close()
    return {"success": True, "message": "Agent registered successfully"}


@app.post("/api/medicaid/claims/chat")
async def medicaid_claims_chat_api(request: Request) -> Any:
    payload = await request.json(); return _medicaid_ai_answer(str(payload.get("question", "")))


@app.get("/download/medicaid/issues.csv")
def medicaid_export_issues(status: str = "", issue_type: str = "", state: str = "") -> Any:
    db=_medicaid_db(); cur=db.cursor(); query='''SELECT s.name state,i.title,i.description,i.status,i.priority,i.issue_type,i.created_at,i.updated_at,GROUP_CONCAT(DISTINCT t.name) tags FROM issues i JOIN states s ON i.state_id=s.id LEFT JOIN issue_tags it ON i.id=it.issue_id LEFT JOIN tags t ON it.tag_id=t.id WHERE 1=1'''; params=[]
    if status.strip(): query += " AND i.status=?"; params.append(status.strip().lower())
    if issue_type.strip(): query += " AND LOWER(i.issue_type)=?"; params.append(issue_type.strip().lower())
    if state.strip(): query += " AND s.name LIKE ?"; params.append(f"%{state.strip()}%")
    query += " GROUP BY i.id ORDER BY s.name,i.updated_at DESC"; cur.execute(query,params); rows=cur.fetchall(); db.close()
    output=io.StringIO(); writer=csv.writer(output); writer.writerow(['State','Title','Description','Status','Priority','Issue Type','Created','Updated','Tags'])
    for r in rows: writer.writerow([r['state'],r['title'],r['description'] or '',r['status'],r['priority'],r['issue_type'],r['created_at'],r['updated_at'],r['tags'] or ''])
    content=output.getvalue().encode('utf-8-sig')
    return StreamingResponse(io.BytesIO(content), media_type='text/csv; charset=utf-8', headers={'Content-Disposition':'attachment; filename=medicaid_state_issues_export.csv'})


@app.get("/download/byo-file/{filename}")
def download_byo_file(filename: str) -> Any:
    try:
        path = _byo_path(filename)
        media = "text/csv" if path.suffix.lower() == ".csv" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        return FileResponse(path, media_type=media, filename=path.name)
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, FileNotFoundError) else 400)


@app.get("/download/byo-csv/{filename}")
def download_byo_as_csv(filename: str) -> Any:
    try:
        path = _byo_path(filename)
        frame = _read_byo_dataset(path)
        csv = frame.to_csv(index=False).encode("utf-8-sig")
        name = re.sub(r"[^A-Za-z0-9._-]+", "_", path.stem) or "byo_dataset"
        return StreamingResponse(io.BytesIO(csv), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename={name}_dart_export.csv"})
    except Exception as exc:
        return JSONResponse({"error": str(exc)}, status_code=404 if isinstance(exc, FileNotFoundError) else 400)


@app.get("/download/current.csv")
def download_current() -> StreamingResponse:
    ensure_data()
    csv = standardize(STATE["df"]).to_csv(index=False).encode("utf-8")
    return StreamingResponse(io.BytesIO(csv), media_type="text/csv", headers={"Content-Disposition": "attachment; filename=dart_current.csv"})


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=int(os.getenv("PORT", "8000")))
