import json
import os
import io
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from pathlib import Path
from typing import List, Optional

import markdown as md_lib
import numpy as np
import pandas as pd
import plotly.express as px
from fastapi import Depends, FastAPI, File, Form, Query, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from markupsafe import Markup
from starlette.middleware.sessions import SessionMiddleware
from sqlmodel import Session, select
from apscheduler.schedulers.background import BackgroundScheduler

from . import data_engine as de
from .database import engine, get_session, init_db
from .models import AutomationRule, Connection, Dataset, Impact, Issue

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


BASE_DIR = Path(__file__).resolve().parent

app = FastAPI(title="Medi Intelligence Hub")

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SECRET_KEY", "change-this-in-production"),
)

app.mount(
    "/static",
    StaticFiles(directory=str(BASE_DIR / "static")),
    name="static",
)

templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

templates.env.filters["pct"] = de.fmt_pct
templates.env.filters["num"] = de.fmt_num
templates.env.filters["markdown"] = lambda text: Markup(
    md_lib.markdown(text or "")
)


SYSTEM_PROMPT = """You are a healthcare data intelligence copilot. Be concise and data-driven.
Use only the supplied context; never invent statistics. Explain field-level average versus
event-weighted match rate when relevant. You are a decision-support assistant, not a clinical advisor."""


# -------------------------
# Filters
# -------------------------

class Filters:
    def __init__(
        self,
        stream: Optional[List[str]] = Query(None),
        classification: Optional[List[str]] = Query(None),
        min_rate: int = Query(0, ge=0, le=100),
        actions_only: bool = Query(False),
    ):
        self.stream = stream or []
        self.classification = classification or []
        self.min_rate = min_rate
        self.actions_only = actions_only


def apply_filters(df: pd.DataFrame, f: Filters) -> pd.DataFrame:
    out = df

    if f.stream:
        out = out[out["Stream"].astype(str).isin(f.stream)]

    if f.classification:
        out = out[
            out["Classification"].astype(str).isin(f.classification)
        ]

    out = out[
        out["MatchRate"].isna()
        | (out["MatchRate"] >= f.min_rate / 100)
    ]

    if f.actions_only:
        out = out[out["NeedsChange"]]

    return out


def get_active(session: Session = Depends(get_session)):
    ds = session.exec(
        select(Dataset).where(Dataset.is_active == True)
    ).first()

    if not ds:
        return de.demo_data(), "Built-in demo dataset"

    return de.load_df_from_path(ds.filepath), ds.name


def base_ctx(
    request: Request,
    df: pd.DataFrame,
    f: Filters,
    source: str,
    **extra,
) -> dict:
    ctx = {
        "request": request,
        "streams": sorted(
            df["Stream"].astype(str).unique()
        ),
        "classes": sorted(
            df["Classification"].astype(str).unique()
        ),
        "filters": f,
        "source": source,
    }

    ctx.update(extra)

    return ctx


# -------------------------
# Onboarding / Persona
# -------------------------

@app.get("/")
def root(request: Request):
    return RedirectResponse(
        "/home"
        if request.session.get("persona")
        else "/onboarding"
    )


@app.get("/onboarding", response_class=HTMLResponse)
def onboarding_form(request: Request):
    return templates.TemplateResponse(
        request,
        "onboarding.html",
        {"request": request},
    )


@app.post("/onboarding")
def onboarding_submit(
    request: Request,
    program: str = Form(...),
    role: str = Form(...),
    focus: List[str] = Form([]),
    geography: List[str] = Form([]),
    audience: str = Form(...),
    depth: str = Form("Balanced"),
    first_question: str = Form(""),
):
    request.session["persona"] = {
        "program": program,
        "role": role,
        "focus": focus,
        "geography": geography,
        "audience": audience,
        "depth": depth,
        "first_question": first_question,
    }

    return RedirectResponse(
        "/home",
        status_code=303,
    )


# -------------------------
# Home / Briefing
# -------------------------

@app.get("/home", response_class=HTMLResponse)
def home(
    request: Request,
    f: Filters = Depends(),
    data=Depends(get_active),
):
    persona = request.session.get("persona")

    if not persona:
        return RedirectResponse("/onboarding")

    df, source = data
    fdf = apply_filters(df, f)

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        persona=persona,
        field_avg=de.field_average(fdf),
        weighted=de.weighted_rate(fdf),
        needs_change=int(fdf["NeedsChange"].sum()),
        domains=fdf["Stream"].nunique(),
        rows=de.records(
            fdf.sort_values("MatchRate").head(75)
        ),
    )

    return templates.TemplateResponse(
        request,
        "home.html",
        ctx,
    )


@app.get("/briefing", response_class=HTMLResponse)
def briefing(
    request: Request,
    f: Filters = Depends(),
    data=Depends(get_active),
):
    persona = request.session.get("persona")

    if not persona:
        return RedirectResponse("/onboarding")

    df, source = data
    fdf = apply_filters(df, f)

    top = (
        fdf[fdf["NeedsChange"]]
        .sort_values(
            ["NotMatchedClaims", "MatchRate"],
            ascending=[False, True],
        )
        .head(10)
    )

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        persona=persona,
        field_avg=de.field_average(fdf),
        weighted=de.weighted_rate(fdf),
        needs_change=int(fdf["NeedsChange"].sum()),
        domains=fdf["Stream"].nunique(),
        top=de.records(top),
    )

    return templates.TemplateResponse(
        request,
        "briefing.html",
        ctx,
    )


# -------------------------
# System Explorer
# -------------------------

@app.get("/explorer", response_class=HTMLResponse)
def explorer(
    request: Request,
    q: str = Query(""),
    stream_filter: str = Query(""),
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, source = data

    v = apply_filters(df, f)

    if q:
        mask = (
            v.astype(str)
            .apply(
                lambda col: col.str.contains(
                    q,
                    case=False,
                    na=False,
                )
            )
            .any(axis=1)
        )

        v = v[mask]

    if stream_filter:
        v = v[
            v["Stream"].astype(str) == stream_filter
        ]

    v = v.sort_values("MatchRate")

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        q=q,
        stream_filter=stream_filter,
        count=len(v),
        rows=de.records_with_id(v),
    )

    return templates.TemplateResponse(
        request,
        "explorer.html",
        ctx,
    )


@app.get(
    "/explorer/field/{row_id}",
    response_class=HTMLResponse,
)
def explorer_field(
    request: Request,
    row_id: int,
    data=Depends(get_active),
):
    df, source = data

    if row_id not in df.index:
        return RedirectResponse("/explorer")

    return templates.TemplateResponse(
        request,
        "explorer_detail.html",
        {
            "request": request,
            "r": df.loc[row_id],
            "source": source,
        },
    )


# -------------------------
# Data Quality Lab
# -------------------------

@app.get("/quality", response_class=HTMLResponse)
def quality(
    request: Request,
    mode: str = Query("field"),
    threshold: int = Query(
        80,
        ge=0,
        le=100,
    ),
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, source = data
    fdf = apply_filters(df, f)

    table: list = []

    if mode == "weighted":
        g = (
            fdf.groupby("Stream", as_index=False)
            .agg(
                Matched=("MatchedClaims", "sum"),
                NotMatched=("NotMatchedClaims", "sum"),
            )
        )

        g["WeightedMatch"] = (
            g["Matched"]
            / (g["Matched"] + g["NotMatched"]).replace(
                0,
                np.nan,
            )
        )

        table = de.records(g)

        fig = px.bar(
            g,
            x="Stream",
            y="WeightedMatch",
            text="WeightedMatch",
            title="Event-Weighted Match Rate by Stream",
        )

        fig.update_yaxes(tickformat=".0%")

        fig.update_traces(
            texttemplate="%{text:.1%}",
            textposition="outside",
        )

    elif mode == "stream":
        g = (
            fdf[fdf["Comparable"]]
            .groupby("Stream", as_index=False)
            .agg(
                MatchRate=("MatchRate", "mean"),
                Fields=("MatchRate", "count"),
            )
        )

        fig = px.bar(
            g.sort_values("MatchRate"),
            x="Stream",
            y="MatchRate",
            text="MatchRate",
            title="Average Comparable Match Rate by Stream",
        )

        fig.update_yaxes(tickformat=".0%")

        fig.update_traces(
            texttemplate="%{text:.1%}",
            textposition="outside",
        )

    else:
        mode = "field"

        x = fdf[fdf["Comparable"]]

        fig = px.histogram(
            x,
            x="MatchRate",
            nbins=20,
            title="Comparable Field Match-Rate Distribution",
        )

        fig.update_xaxes(tickformat=".0%")

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    chart_html = fig.to_html(
        full_html=False,
        include_plotlyjs=False,
    )

    flagged = fdf[
        fdf["MatchRate"] < threshold / 100
    ]

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        mode=mode,
        chart=chart_html,
        table=table,
        threshold=threshold,
        field_avg=de.field_average(fdf),
        weighted=de.weighted_rate(fdf),
        comparable=int(fdf["Comparable"].sum()),
        needs_change=int(fdf["NeedsChange"].sum()),
        flagged=de.records(
            flagged.head(100)
        ),
        flagged_count=len(flagged),
    )

    return templates.TemplateResponse(
        request,
        "quality.html",
        ctx,
    )


# -------------------------
# Issue Studio
# -------------------------

@app.get("/issues", response_class=HTMLResponse)
def issues_page(
    request: Request,
    f: Filters = Depends(),
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, source = data
    fdf = apply_filters(df, f)

    candidates = fdf[fdf["NeedsChange"]]

    board = session.exec(
        select(Issue).order_by(
            Issue.created_at.desc()
        )
    ).all()

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        candidates=de.records(candidates),
        unmatched_volume=(
            candidates["NotMatchedClaims"]
            .fillna(0)
            .sum()
        ),
        affected_streams=candidates["Stream"].nunique(),
        board=board,
    )

    return templates.TemplateResponse(
        request,
        "issues.html",
        ctx,
    )


@app.post("/issues")
def create_issue(
    stream: str = Form(...),
    field: str = Form(...),
    match_rate: float = Form(0),
    not_matched: int = Form(0),
    owner: str = Form(""),
    status: str = Form("Open"),
    note: str = Form(""),
    session: Session = Depends(get_session),
):
    session.add(
        Issue(
            stream=stream,
            field=field,
            match_rate=match_rate,
            not_matched=not_matched,
            owner=owner,
            status=status,
            note=note,
        )
    )

    session.commit()

    return RedirectResponse(
        "/issues",
        status_code=303,
    )


@app.post("/issues/{issue_id}/status")
def update_issue_status(
    issue_id: int,
    status: str = Form(...),
    session: Session = Depends(get_session),
):
    issue = session.get(Issue, issue_id)

    if issue:
        issue.status = status
        session.add(issue)
        session.commit()

    return RedirectResponse(
        "/issues",
        status_code=303,
    )


# -------------------------
# Compare
# -------------------------

@app.get("/compare", response_class=HTMLResponse)
def compare(
    request: Request,
    group: str = Query("Stream"),
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, source = data
    fdf = apply_filters(df, f)

    if (
        group not in (
            "Stream",
            "Classification",
            "Disposition",
        )
        or group not in fdf.columns
    ):
        group = "Stream"

    g = fdf.copy()

    g["Group"] = g[group].astype(str)

    s = (
        g.groupby("Group")
        .agg(
            Fields=("MatchRate", "count"),
            AverageMatch=("MatchRate", "mean"),
            Matched=("MatchedClaims", "sum"),
            NotMatched=("NotMatchedClaims", "sum"),
        )
        .reset_index()
    )

    s["WeightedMatch"] = (
        s["Matched"]
        / (s["Matched"] + s["NotMatched"]).replace(
            0,
            np.nan,
        )
    )

    fig = px.scatter(
        s,
        x="AverageMatch",
        y="WeightedMatch",
        size="Fields",
        color="Group",
        hover_data=["Matched", "NotMatched"],
        title="Field Average vs Event-Weighted Performance",
    )

    fig.update_xaxes(tickformat=".0%")
    fig.update_yaxes(tickformat=".0%")

    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        group=group,
        table=de.records(s),
        chart=fig.to_html(
            full_html=False,
            include_plotlyjs=False,
        ),
    )

    return templates.TemplateResponse(
        request,
        "compare.html",
        ctx,
    )


# -------------------------
# Connections
# -------------------------

@app.get(
    "/connections",
    response_class=HTMLResponse,
)
def connections_page(
    request: Request,
    f: Filters = Depends(),
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, source = data
    fdf = apply_filters(df, f)

    saved = session.exec(
        select(Connection).order_by(
            Connection.id.desc()
        )
    ).all()

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        connections=de.records(fdf),
        saved=saved,
    )

    return templates.TemplateResponse(
        request,
        "connections.html",
        ctx,
    )


@app.post("/connections")
def create_connection(
    source: str = Form(...),
    target: str = Form(...),
    key: str = Form(""),
    report: str = Form(""),
    session: Session = Depends(get_session),
):
    session.add(
        Connection(
            source=source,
            target=target,
            key=key,
            report=report,
        )
    )

    session.commit()

    return RedirectResponse(
        "/connections",
        status_code=303,
    )


# -------------------------
# Report Impact
# -------------------------

@app.get(
    "/impact",
    response_class=HTMLResponse,
)
def impact_page(
    request: Request,
    f: Filters = Depends(),
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, source = data
    fdf = apply_filters(df, f)

    impacts = session.exec(
        select(Impact).order_by(
            Impact.id.desc()
        )
    ).all()

    ctx = base_ctx(
        request,
        df,
        f,
        source,
        fields=de.records_with_id(fdf),
        impacts=impacts,
    )

    return templates.TemplateResponse(
        request,
        "impact.html",
        ctx,
    )


@app.post("/impact")
def create_impact(
    row_id: int = Form(...),
    report: str = Form(""),
    kpi: str = Form(""),
    owner: str = Form(""),
    impact: str = Form("Low"),
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, _ = data

    stream = ""
    field = ""

    if row_id in df.index:
        row = df.loc[row_id]
        stream = row["Stream"]
        field = row["NCH Target Column"]

    session.add(
        Impact(
            stream=stream,
            field=field,
            report=report,
            kpi=kpi,
            owner=owner,
            impact=impact,
        )
    )

    session.commit()

    return RedirectResponse(
        "/impact",
        status_code=303,
    )


# -------------------------
# AI Copilot
# -------------------------

def call_ai(
    context: dict,
    history: list,
    api_key: str,
    model: str,
) -> str:

    if OpenAI and api_key:
        try:
            client = OpenAI(
                api_key=api_key,
                base_url="https://api.groq.com/openai/v1",
            )

            messages = [
                {
                    "role": "system",
                    "content": (
                        SYSTEM_PROMPT
                        + "\n\nCURRENT CONTEXT:\n"
                        + json.dumps(
                            context,
                            default=str,
                        )
                    ),
                }
            ]

            messages += history[-8:]

            res = client.chat.completions.create(
                model=model,
                temperature=0.2,
                messages=messages,
            )

            return res.choices[0].message.content

        except Exception as e:
            return (
                f"AI service error: `{e}`\n\n"
                f"Local result: the comparable-field average is "
                f"**{de.fmt_pct(context['field_average'])}** "
                f"and the event-weighted rate is "
                f"**{de.fmt_pct(context['weighted_rate'])}**."
            )

    g = context.get("stream_summary", [])

    worst = (
        max(
            g,
            key=lambda r: r.get("NotMatched") or 0,
        )["Stream"]
        if g
        else "N/A"
    )

    return (
        f"Based on current filters, the comparable-field average is "
        f"**{de.fmt_pct(context['field_average'])}**, "
        f"the event-weighted rate is "
        f"**{de.fmt_pct(context['weighted_rate'])}**, "
        f"and the largest unmatched volume is in "
        f"**{worst}**. Add a Groq API key above for conversational AI."
    )


@app.get(
    "/copilot",
    response_class=HTMLResponse,
)
def copilot_page(request: Request):
    persona = request.session.get("persona") or {}

    chat = request.session.get("chat") or [
        {
            "role": "assistant",
            "content": (
                f"Hi. I'm your Medi Intelligence Copilot, "
                f"tuned for **{persona.get('program', 'your')}** "
                f"and **{persona.get('role', 'your role')}**. "
                f"Ask about match rates, issues, streams, or priorities."
            ),
        }
    ]

    return templates.TemplateResponse(
        request,
        "copilot.html",
        {
            "request": request,
            "chat": chat,
        },
    )


@app.post(
    "/copilot",
    response_class=HTMLResponse,
)
def copilot_ask(
    request: Request,
    prompt: str = Form(...),
    api_key: str = Form(""),
    model: str = Form(
        "llama-3.3-70b-versatile"
    ),
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, source = data

    fdf = apply_filters(df, f)

    chat = request.session.get("chat") or []

    chat.append(
        {
            "role": "user",
            "content": prompt,
        }
    )

    stream_summary = (
        fdf.groupby("Stream")
        .agg(
            Fields=("MatchRate", "count"),
            AverageMatch=("MatchRate", "mean"),
            Matched=("MatchedClaims", "sum"),
            NotMatched=("NotMatchedClaims", "sum"),
        )
        .reset_index()
        .to_dict("records")
    )

    context = {
        "persona": request.session.get("persona"),
        "source": source,
        "rows": len(fdf),
        "field_average": de.field_average(fdf),
        "weighted_rate": de.weighted_rate(fdf),
        "stream_summary": stream_summary,
        "top_actions": de.records(
            fdf[fdf["NeedsChange"]][
                [
                    "Stream",
                    "NCH Target Column",
                    "MatchRate",
                    "NotMatchedClaims",
                    "Recommendation",
                ]
            ]
            .sort_values(
                "NotMatchedClaims",
                ascending=False,
            )
            .head(20)
        ),
    }

    answer = call_ai(
        context,
        chat,
        api_key,
        model,
    )

    chat.append(
        {
            "role": "assistant",
            "content": answer,
        }
    )

    request.session["chat"] = chat[-20:]

    return templates.TemplateResponse(
        request,
        "copilot.html",
        {
            "request": request,
            "chat": chat,
        },
    )


# -------------------------
# Automations
# -------------------------

@app.get(
    "/automations",
    response_class=HTMLResponse,
)
def automations_page(
    request: Request,
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, source = data

    rules = session.exec(
        select(AutomationRule)
    ).all()

    return templates.TemplateResponse(
        request,
        "automations.html",
        {
            "request": request,
            "rules": rules,
            "columns": list(df.columns),
            "source": source,
        },
    )


@app.post("/automations")
def create_rule(
    recipient: str = Form(""),
    field: str = Form(...),
    condition: str = Form(...),
    value: str = Form(""),
    subject: str = Form(
        "Medi Intelligence Alert"
    ),
    body: str = Form(""),
    session: Session = Depends(get_session),
):
    session.add(
        AutomationRule(
            recipient=recipient,
            field=field,
            condition=condition,
            value=value,
            subject=subject,
            body=body,
        )
    )

    session.commit()

    return RedirectResponse(
        "/automations",
        status_code=303,
    )


@app.post(
    "/automations/{rule_id}/test"
)
def test_rule(
    rule_id: int,
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, _ = data

    rule = session.get(
        AutomationRule,
        rule_id,
    )

    if rule:
        hits = de.evaluate_rule(
            df,
            rule,
        )

        rule.last_run = datetime.utcnow()
        rule.last_hits = len(hits)

        session.add(rule)
        session.commit()

    return RedirectResponse(
        "/automations",
        status_code=303,
    )


def send_alert_email(
    rule: AutomationRule,
    hits: pd.DataFrame,
) -> None:

    host = os.getenv("SMTP_HOST")

    if not host or not rule.recipient:
        return

    port = int(
        os.getenv(
            "SMTP_PORT",
            "587",
        )
    )

    user = os.getenv("SMTP_USER")
    pwd = os.getenv("SMTP_PASS")

    row = hits.iloc[0].to_dict()

    mapping = {
        "{{Stream}}": str(
            row.get("Stream", "")
        ),
        "{{Field}}": str(
            row.get(
                "NCH Target Column",
                "",
            )
        ),
        "{{MatchRate}}": de.fmt_pct(
            row.get("MatchRate")
        ),
        "{{MatchedClaims}}": str(
            row.get(
                "MatchedClaims",
                "",
            )
        ),
        "{{NotMatchedClaims}}": str(
            row.get(
                "NotMatchedClaims",
                "",
            )
        ),
        "{{Classification}}": str(
            row.get(
                "Classification",
                "",
            )
        ),
        "{{Recommendation}}": str(
            row.get(
                "Recommendation",
                "",
            )
        ),
    }

    body = rule.body

    for token, value in mapping.items():
        body = body.replace(
            token,
            value,
        )

    msg = MIMEText(
        f"{body}\n\n"
        f"({len(hits)} rows matched this rule.)"
    )

    msg["Subject"] = rule.subject
    msg["From"] = (
        user
        or "medi-intelligence@localhost"
    )
    msg["To"] = rule.recipient

    with smtplib.SMTP(
        host,
        port,
    ) as server:

        server.starttls()

        if user and pwd:
            server.login(
                user,
                pwd,
            )

        server.send_message(msg)


def run_automation_checks() -> None:
    with Session(engine) as session:

        ds = session.exec(
            select(Dataset).where(
                Dataset.is_active == True
            )
        ).first()

        df = (
            de.load_df_from_path(ds.filepath)
            if ds
            else de.demo_data()
        )

        rules = session.exec(
            select(AutomationRule).where(
                AutomationRule.active == True
            )
        ).all()

        for rule in rules:

            hits = de.evaluate_rule(
                df,
                rule,
            )

            rule.last_run = datetime.utcnow()
            rule.last_hits = len(hits)

            session.add(rule)

            if len(hits):
                try:
                    send_alert_email(
                        rule,
                        hits,
                    )
                except Exception:
                    pass

        session.commit()


# -------------------------
# My Data
# -------------------------

@app.get(
    "/mydata",
    response_class=HTMLResponse,
)
def mydata_page(
    request: Request,
    data=Depends(get_active),
    session: Session = Depends(get_session),
):
    df, source = data

    datasets = session.exec(
        select(Dataset).order_by(
            Dataset.uploaded_at.desc()
        )
    ).all()

    profile = [
        {
            "Column": c,
            "Type": str(df[c].dtype),
            "Non-null": int(
                df[c].notna().sum()
            ),
            "Unique": int(
                df[c].nunique(
                    dropna=True
                )
            ),
        }
        for c in df.columns
    ]

    return templates.TemplateResponse(
        request,
        "mydata.html",
        {
            "request": request,
            "source": source,
            "rows": len(df),
            "cols": len(df.columns),
            "datasets": datasets,
            "profile": profile,
        },
    )


@app.post("/mydata/upload")
async def upload_dataset(
    file: UploadFile = File(...),
    session: Session = Depends(get_session),
):
    de.UPLOAD_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    dest = (
        de.UPLOAD_DIR
        / f"{int(datetime.utcnow().timestamp())}_{file.filename}"
    )

    with open(dest, "wb") as out:
        out.write(
            await file.read()
        )

    for d in session.exec(
        select(Dataset)
    ).all():
        d.is_active = False
        session.add(d)

    session.add(
        Dataset(
            name=file.filename,
            filepath=str(dest),
            is_active=True,
        )
    )

    session.commit()

    return RedirectResponse(
        "/mydata",
        status_code=303,
    )


@app.post(
    "/mydata/activate/{dataset_id}"
)
def activate_dataset(
    dataset_id: int,
    session: Session = Depends(get_session),
):
    for d in session.exec(
        select(Dataset)
    ).all():

        d.is_active = (
            d.id == dataset_id
        )

        session.add(d)

    session.commit()

    return RedirectResponse(
        "/mydata",
        status_code=303,
    )


@app.get("/mydata/download.csv")
def download_csv(
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, _ = data
    fdf = apply_filters(df, f)

    buf = io.StringIO()

    fdf.to_csv(
        buf,
        index=False,
    )

    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="text/csv",
        headers={
            "Content-Disposition":
                "attachment; "
                "filename=medi_intelligence_filtered.csv"
        },
    )


@app.get("/mydata/download.xlsx")
def download_xlsx(
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, _ = data
    fdf = apply_filters(df, f)

    buf = io.BytesIO()

    with pd.ExcelWriter(
        buf,
        engine="openpyxl",
    ) as writer:

        fdf.to_excel(
            writer,
            index=False,
            sheet_name="Filtered",
        )

    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        headers={
            "Content-Disposition":
                "attachment; "
                "filename=medi_intelligence_filtered.xlsx"
        },
    )


# -------------------------
# JSON API
# -------------------------

@app.get("/api/summary")
def api_summary(
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, source = data
    fdf = apply_filters(df, f)

    return {
        "source": source,
        "rows": len(fdf),
        "field_average": de.field_average(fdf),
        "weighted_rate": de.weighted_rate(fdf),
        "needs_change": int(
            fdf["NeedsChange"].sum()
        ),
    }


@app.get("/api/fields")
def api_fields(
    f: Filters = Depends(),
    data=Depends(get_active),
):
    df, _ = data

    return de.records(
        apply_filters(df, f)
    )


# -------------------------
# Startup
# -------------------------

scheduler = BackgroundScheduler()


@app.on_event("startup")
def on_startup():
    init_db()

    with Session(engine) as session:

        if not session.exec(
            select(Dataset)
        ).first():

            de.UPLOAD_DIR.mkdir(
                parents=True,
                exist_ok=True,
            )

            demo_path = (
                de.UPLOAD_DIR
                / "demo.csv"
            )

            de.demo_data().to_csv(
                demo_path,
                index=False,
            )

            session.add(
                Dataset(
                    name="Built-in demo dataset",
                    filepath=str(
                        demo_path
                    ),
                    is_active=True,
                )
            )

            session.commit()

    scheduler.add_job(
        run_automation_checks,
        "interval",
        minutes=15,
        id="automation_checks",
        replace_existing=True,
    )

    scheduler.start()


@app.on_event("shutdown")
def on_shutdown():
    scheduler.shutdown()