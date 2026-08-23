
import os, json, re
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="Medi Intelligence Hub", page_icon="◆", layout="wide")

# ============================================================
# MEDI INTELLIGENCE HUB
# One-file interactive product combining:
#   • DART-style System A ↔ System B reconciliation
#   • State/issue-style tracking and comparison
#   • Persona-driven onboarding
#   • Bring-your-own-data analysis
#   • Interactive issue/action workflows
#   • Report-impact mapping
#   • AI Copilot
#   • Automation templates
# ============================================================

st.markdown("""
<style>
.stApp {background:linear-gradient(180deg,#07111f,#091827);color:#eaf2fb}
section[data-testid="stSidebar"]{background:#081421}
.block-container{max-width:1500px;padding-top:1.2rem}
.hero{padding:28px 32px;border:1px solid #23425f;border-radius:20px;
background:linear-gradient(135deg,#0d2035,#0a1726);margin-bottom:20px}
.hero h1{margin:0;font-size:2.25rem}.hero p{color:#9bb0c5;margin:.5rem 0 0}
.card{background:#0d1d30;border:1px solid #1f3853;border-radius:16px;padding:18px;min-height:110px}
.label{color:#9bb0c5;font-size:.8rem;text-transform:uppercase;letter-spacing:.06em}
.value{font-size:1.8rem;font-weight:750;margin-top:6px}.sub{color:#9bb0c5;font-size:.8rem}
.insight{border-left:4px solid #55c2ff;padding:14px 16px;background:#0b1b2c;border-radius:10px}
.small{color:#9bb0c5;font-size:.85rem}
</style>
""", unsafe_allow_html=True)


# ------------------------- Utilities -------------------------

def card(label, value, sub=""):
    return f'<div class="card"><div class="label">{label}</div><div class="value">{value}</div><div class="sub">{sub}</div></div>'

def pct(x):
    return "—" if pd.isna(x) else f"{x:.1%}"

def norm_cols(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()) for c in df.columns]
    return df

def find_col(df, choices):
    exact = {str(c).lower().strip(): c for c in df.columns}
    for x in choices:
        if x.lower().strip() in exact:
            return exact[x.lower().strip()]
    for c in df.columns:
        for x in choices:
            if x.lower() in str(c).lower():
                return c
    return None

def standardize(df):
    if df.empty:
        return df
    df = norm_cols(df)

    aliases = {
        "Stream":["Stream","Claim Stream","Domain"],
        "NCH Target Table":["NCH Target Table","NCH Table","System A Table","Source Table"],
        "NCH Target Column":["NCH Target Column","NCH Column","System A Column","Source Column"],
        "SS Table":["SS Table","Shared Systems Table","System B Table","Target Table"],
        "SS Column":["SS Column","Shared Systems Column","System B Column","Target Column"],
        "MatchedClaims":["MatchedClaims","Matched Claims","Matched Events","Matched"],
        "NotMatchedClaims":["NotMatchedClaims","Not Matched Claims","Not-Matched Claims","Unmatched"],
        "MatchRate":["MatchRate","Match Rate","Match %","Match Percentage"],
        "Classification":["Classification","ClassificationClean","Class"],
        "Sub-Classification":["Sub-Classification","Sub Classification","Subclass"],
        "Disposition":["Disposition","DispositionClean","Action"],
        "Recommendation":["Recommendation","Recommended Action"],
        "Lynette's recommendation":["Lynette's recommendation","Lynette Recommendation"],
        "Comments":["Comments","Comment","Notes"],
    }
    ren = {}
    for target, opts in aliases.items():
        c = find_col(df, opts)
        if c and c != target:
            ren[c] = target
    df = df.rename(columns=ren)

    for c in ["MatchedClaims","NotMatchedClaims","MatchRate"]:
        if c in df:
            df[c] = pd.to_numeric(
                df[c].astype(str).str.replace(",","",regex=False).str.replace("%","",regex=False),
                errors="coerce"
            )

    if "MatchRate" in df and df["MatchRate"].dropna().size and df["MatchRate"].dropna().max() > 1.5:
        df["MatchRate"] /= 100

    for c in ["MatchedClaims","NotMatchedClaims"]:
        if c not in df:
            df[c] = np.nan

    if "MatchRate" not in df:
        total = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
        df["MatchRate"] = np.where(total > 0, df["MatchedClaims"]/total, np.nan)

    if "Classification" not in df:
        df["Classification"] = np.where(
            df["MatchRate"] >= .9999, "Match",
            np.where(df["MatchRate"].notna(),"Different","Missing")
        )

    for c in ["Stream","NCH Target Table","NCH Target Column","SS Table","SS Column"]:
        if c not in df:
            df[c] = ""

    df["Comparable"] = df["Classification"].astype(str).str.strip().str.lower().isin(["match","different"])

    if "Disposition" in df:
        disp = df["Disposition"].astype(str).str.lower()
        df["NeedsChange"] = disp.str.contains("action|change")
    else:
        df["NeedsChange"] = False
    df["NeedsChange"] |= ~df["Comparable"]

    df["TotalClaims"] = df["MatchedClaims"].fillna(0) + df["NotMatchedClaims"].fillna(0)
    return df

def field_average(df):
    x = df[df["Comparable"] & df["MatchRate"].notna()]
    return x["MatchRate"].mean() if len(x) else np.nan

def weighted_rate(df):
    m = df["MatchedClaims"].fillna(0).sum()
    u = df["NotMatchedClaims"].fillna(0).sum()
    return m/(m+u) if m+u else np.nan

def read_file(upload):
    if upload.name.lower().endswith(".csv"):
        return pd.read_csv(upload)
    return pd.read_excel(upload)

def local_data():
    here = Path(__file__).resolve().parent
    files = []
    for root in [here, here/"data"]:
        if root.exists():
            files += list(root.glob("*.xlsx")) + list(root.glob("*.xls")) + list(root.glob("*.csv"))
    files = sorted(set(files), key=lambda p: (0 if any(k in p.name.lower() for k in ["recon","comparison","integrity","dart"]) else 1,p.name))
    for p in files:
        try:
            if p.suffix.lower() == ".csv":
                d = pd.read_csv(p)
                if len(d):
                    return standardize(d), p.name
            xl = pd.ExcelFile(p)
            for sheet in xl.sheet_names:
                d = pd.read_excel(p, sheet_name=sheet)
                if len(d) and len(d.columns) >= 3:
                    return standardize(d), f"{p.name} / {sheet}"
        except Exception:
            pass
    return pd.DataFrame(), ""

def demo_data():
    rng=np.random.default_rng(42)
    meta=[
        ("PROF","Professional","MCS"),
        ("INP/SNF","Inpatient / SNF","FISS"),
        ("OUT","Outpatient","FISS"),
        ("HH","Home Health","FISS"),
        ("HSPC","Hospice","FISS"),
        ("DME","DME","VMS")
    ]
    rows=[]
    for stream,domain,ss in meta:
        for i in range(20):
            rate=float(rng.uniform(.38,.99))
            matched=int(rng.integers(3000,200000))
            total=max(int(matched/max(rate,.05)),matched)
            rows.append({
                "Stream":stream,
                "NCH Target Table":f"{stream}_TABLE_{i%5+1}",
                "NCH Target Column":f"{stream}_FIELD_{i+1}",
                "SS Table":f"{ss}_TABLE_{i%5+1}",
                "SS Column":f"{ss}_FIELD_{i+1}",
                "MatchedClaims":matched,
                "NotMatchedClaims":total-matched,
                "MatchRate":matched/total,
                "Classification":"Match" if rate>=.9 else ("Different" if rate>=.55 else "Missing"),
                "Sub-Classification":"Example",
                "Disposition":"Good" if rate>=.9 else "Action needed",
                "Recommendation":"Review mapping" if rate<.9 else "No action",
                "Lynette's recommendation":"",
                "Comments":"Interactive demo record"
            })
    return standardize(pd.DataFrame(rows))


# ------------------------- Session state -------------------------

if "persona" not in st.session_state: st.session_state.persona=None
if "df" not in st.session_state: st.session_state.df=pd.DataFrame()
if "source" not in st.session_state: st.session_state.source=""
if "chat" not in st.session_state: st.session_state.chat=[]
if "rules" not in st.session_state: st.session_state.rules=[]
if "issues" not in st.session_state: st.session_state.issues=[]
if "connections" not in st.session_state: st.session_state.connections=[]
if "impacts" not in st.session_state: st.session_state.impacts=[]

# ------------------------- Personalized onboarding -------------------------

if st.session_state.persona is None:
    st.markdown("""
    <div class="hero">
      <h1>◆ Medi Intelligence Hub</h1>
      <p>Your interactive workspace for healthcare data, reconciliation, issues, reporting impact, and AI-powered insight.</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("## Let's build your workspace")
    st.caption("A few questions personalize what you see first. Nothing is permanent.")

    a,b=st.columns(2)
    with a:
        program=st.radio("Are you interested in Medicare or Medicaid?",["Medicare","Medicaid","Both","Other / General"],horizontal=True)
        role=st.selectbox("What's your primary role?",[
            "Executive / Leadership","Data / Analytics","Program / Policy",
            "Operations","Quality / Compliance","IT / Engineering","Research","Other"
        ])
        focus=st.multiselect("What do you care about most?",[
            "Data quality","Claims","System reconciliation","State comparisons",
            "Issues & remediation","Business rules","Reporting impact","Automation","AI insights"
        ],default=["Data quality","AI insights"])
    with b:
        geography=st.multiselect("What geography matters to you?",[
            "National","State","Multi-state","Specific states"
        ],default=["National"])
        audience=st.selectbox("Who is this workspace primarily for?",[
            "Myself","Leadership","Analysts","Program teams","Technical teams","External stakeholders"
        ])
        depth=st.select_slider("How much detail do you want?",["Executive","Balanced","Technical"],value="Balanced")
        first_question=st.text_input("What do you want to know first?",placeholder="Which areas need the most attention?")

    if st.button("Build My Workspace",type="primary",use_container_width=True):
        st.session_state.persona={
            "program":program,"role":role,"focus":focus,"geography":geography,
            "audience":audience,"depth":depth,"first_question":first_question
        }
        st.rerun()
    st.stop()

p=st.session_state.persona

# ------------------------- Load data -------------------------

if st.session_state.df.empty:
    d,name=local_data()
    if d.empty:
        st.session_state.df=demo_data()
        st.session_state.source="Built-in interactive demo data"
    else:
        st.session_state.df=d
        st.session_state.source=name

df=standardize(st.session_state.df)

# ------------------------- Sidebar -------------------------

st.sidebar.markdown("## ◆ Medi Intelligence")
st.sidebar.caption(f"{p['program']} · {p['role']}")

with st.sidebar.expander("My profile"):
    st.write("**Program:**",p["program"])
    st.write("**Focus:**",", ".join(p["focus"]) or "All")
    st.write("**Audience:**",p["audience"])
    if st.button("Change my answers"):
        st.session_state.persona=None
        st.rerun()

uploads=st.sidebar.file_uploader("Upload your own CSV/XLSX",type=["csv","xlsx","xls"],accept_multiple_files=True)
if uploads:
    chosen=uploads[0] if len(uploads)==1 else st.sidebar.selectbox("Active upload",uploads,format_func=lambda x:x.name)
    try:
        newdf=standardize(read_file(chosen))
        if st.sidebar.button("Use uploaded data",type="primary"):
            st.session_state.df=newdf
            st.session_state.source=chosen.name
            st.rerun()
    except Exception as e:
        st.sidebar.error(f"Upload error: {e}")

if st.sidebar.button("Restore demo data"):
    st.session_state.df=demo_data()
    st.session_state.source="Built-in interactive demo data"
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("### Filters")
streams=sorted(df["Stream"].astype(str).unique())
classes=sorted(df["Classification"].astype(str).unique())
sel_stream=st.sidebar.multiselect("Stream / domain",streams,default=streams)
sel_class=st.sidebar.multiselect("Classification",classes,default=classes)
min_rate=st.sidebar.slider("Minimum match rate",0,100,0)/100
actions_only=st.sidebar.checkbox("Only fields needing action")

fdf=df[df["Stream"].astype(str).isin(sel_stream)]
fdf=fdf[fdf["Classification"].astype(str).isin(sel_class)]
fdf=fdf[fdf["MatchRate"].isna() | (fdf["MatchRate"]>=min_rate)]
if actions_only:
    fdf=fdf[fdf["NeedsChange"]]

pages=[
    "Home","My Briefing","System Explorer","Data Quality Lab","Issue Studio",
    "Compare","Connections","Report Impact","AI Copilot","Automations","My Data"
]
page=st.sidebar.radio("Workspace",pages)

# ------------------------- HOME -------------------------

if page=="Home":
    st.markdown(f"""
    <div class="hero">
      <h1>Welcome back.</h1>
      <p>Your workspace is tuned for <b>{p['program']}</b>, with emphasis on {", ".join(p['focus'][:3]) or "the full data landscape"}.</p>
    </div>
    """,unsafe_allow_html=True)

    c=st.columns(4)
    c[0].markdown(card("Average comparable match",pct(field_average(fdf)),"Equal-weight field average"),unsafe_allow_html=True)
    c[1].markdown(card("Event-weighted match",pct(weighted_rate(fdf)),"Matched ÷ matched + not matched"),unsafe_allow_html=True)
    c[2].markdown(card("Fields needing attention",f"{int(fdf['NeedsChange'].sum()):,}","Current filters"),unsafe_allow_html=True)
    c[3].markdown(card("Active domains",f"{fdf['Stream'].nunique():,}","Streams in view"),unsafe_allow_html=True)

    st.markdown("### Your starting point")
    st.markdown(f'<div class="insight"><b>Your question:</b> {p["first_question"] or "What should I investigate first?"}<br><br>DART-style reconciliation, issue tracking, interactive exploration, reporting impact, and AI are combined into one workspace.</div>',unsafe_allow_html=True)

    st.markdown("### Jump into an experience")
    x,y,z=st.columns(3)
    x.subheader("Explore")
    x.write("Drill from stream → table → column → system relationship.")
    x.button("Use System Explorer",key="home_explore",on_click=lambda:None)
    y.subheader("Act")
    y.write("Turn mismatches into issues, owners, priorities, and recommendations.")
    y.button("Open Issue Studio",key="home_issue")
    z.subheader("Ask")
    z.write("Use natural language to investigate the current data.")
    z.button("Open AI Copilot",key="home_ai")

    st.markdown("### Live workspace")
    st.dataframe(fdf.head(75),use_container_width=True,hide_index=True,
                 column_config={"MatchRate":st.column_config.ProgressColumn("Match Rate",min_value=0,max_value=1,format="%.1f%%")})

# ------------------------- BRIEFING -------------------------

elif page=="My Briefing":
    st.markdown('<div class="hero"><h1>My Briefing</h1><p>A personalized decision view based on your role, program, interests, and current filters.</p></div>',unsafe_allow_html=True)
    c=st.columns(4)
    c[0].metric("Comparable field average",pct(field_average(fdf)))
    c[1].metric("Event-weighted rate",pct(weighted_rate(fdf)))
    c[2].metric("Action candidates",f"{int(fdf['NeedsChange'].sum()):,}")
    c[3].metric("Streams",f"{fdf['Stream'].nunique():,}")

    st.markdown(f'<div class="insight"><b>Personalized lens:</b> {p["program"]} · {p["role"]}. Your workspace emphasizes {", ".join(p["focus"]) or "all available capabilities"}.</div>',unsafe_allow_html=True)

    st.markdown("### What deserves attention?")
    top=fdf[fdf["NeedsChange"]].sort_values(["NotMatchedClaims","MatchRate"],ascending=[False,True]).head(10)
    if top.empty:
        st.success("Nothing currently requires attention under these filters.")
    else:
        for _,r in top.iterrows():
            st.markdown(
                f'<div class="insight"><b>{r["Stream"]} · {r["NCH Target Column"]}</b> — {pct(r["MatchRate"])} match · '
                f'{int(r["NotMatchedClaims"] or 0):,} not matched<br><span class="small">{r.get("Recommendation","")}</span></div><br>',
                unsafe_allow_html=True
            )

# ------------------------- SYSTEM EXPLORER -------------------------

elif page=="System Explorer":
    st.markdown('<div class="hero"><h1>System Explorer</h1><p>Interact with the relationship between System A and System B instead of only viewing a chart.</p></div>',unsafe_allow_html=True)
    q=st.text_input("Search any table, column, stream, or system")
    v=fdf.copy()
    if q:
        mask=v.astype(str).apply(lambda col:col.str.contains(q,case=False,na=False)).any(axis=1)
        v=v[mask]

    s=st.selectbox("Stream",["All"]+sorted(v["Stream"].astype(str).unique()))
    if s!="All": v=v[v["Stream"].astype(str)==s]

    st.write(f"**{len(v):,} fields** in this view.")
    st.dataframe(v[[
        "Stream","NCH Target Table","NCH Target Column","SS Table","SS Column",
        "MatchedClaims","NotMatchedClaims","MatchRate","Classification","Disposition","Recommendation"
    ]].sort_values("MatchRate"),use_container_width=True,hide_index=True,
    column_config={"MatchRate":st.column_config.ProgressColumn("Match Rate",min_value=0,max_value=1,format="%.1f%%")})

    if len(v):
        idx=st.selectbox("Inspect a field",list(v.index),format_func=lambda i:f"{v.loc[i,'Stream']} · {v.loc[i,'NCH Target Column']}")
        r=v.loc[idx]
        a,b,c,d=st.columns(4)
        a.metric("Match rate",pct(r["MatchRate"]))
        b.metric("Matched",f"{int(r['MatchedClaims'] or 0):,}")
        c.metric("Not matched",f"{int(r['NotMatchedClaims'] or 0):,}")
        d.metric("Class",str(r["Classification"]))
        st.markdown(f"""
        **System A:** `{r["NCH Target Table"]}` → `{r["NCH Target Column"]}`  
        **System B:** `{r["SS Table"]}` → `{r["SS Column"]}`  
        **Disposition:** `{r.get("Disposition","")}`  
        **Recommendation:** {r.get("Recommendation","")}
        """)

# ------------------------- QUALITY LAB -------------------------

elif page=="Data Quality Lab":
    st.markdown('<div class="hero"><h1>Data Quality Lab</h1><p>Test different definitions of quality and build your own thresholds.</p></div>',unsafe_allow_html=True)
    a,b,c,d=st.columns(4)
    a.metric("Field average",pct(field_average(fdf)))
    b.metric("Event weighted",pct(weighted_rate(fdf)))
    c.metric("Comparable fields",f"{int(fdf['Comparable'].sum()):,}")
    d.metric("Action fields",f"{int(fdf['NeedsChange'].sum()):,}")

    mode=st.radio("View quality as",["Field-level average","Event-volume weighted","By stream"],horizontal=True)
    if mode=="Field-level average":
        x=fdf[fdf["Comparable"]]
        fig=px.histogram(x,x="MatchRate",nbins=20,title="Comparable Field Match-Rate Distribution")
        fig.update_xaxes(tickformat=".0%")
        st.plotly_chart(fig,use_container_width=True)
        st.info(f"The current field-level average is {pct(field_average(fdf))}. Each comparable field contributes equally.")
    elif mode=="Event-volume weighted":
        g=fdf.groupby("Stream",as_index=False).agg(Matched=("MatchedClaims","sum"),NotMatched=("NotMatchedClaims","sum"))
        g["WeightedMatch"]=g["Matched"]/(g["Matched"]+g["NotMatched"]).replace(0,np.nan)
        st.dataframe(g.style.format({"Matched":"{:,.0f}","NotMatched":"{:,.0f}","WeightedMatch":"{:.2%}"}),use_container_width=True,hide_index=True)
        fig=px.bar(g,x="Stream",y="WeightedMatch",text="WeightedMatch",title="Event-Weighted Match Rate by Stream")
        fig.update_yaxes(tickformat=".0%");fig.update_traces(texttemplate="%{text:.1%}",textposition="outside")
        st.plotly_chart(fig,use_container_width=True)
    else:
        g=fdf[fdf["Comparable"]].groupby("Stream",as_index=False).agg(MatchRate=("MatchRate","mean"),Fields=("MatchRate","count"))
        fig=px.bar(g.sort_values("MatchRate"),x="Stream",y="MatchRate",text="MatchRate",title="Average Comparable Match Rate by Stream")
        fig.update_yaxes(tickformat=".0%");fig.update_traces(texttemplate="%{text:.1%}",textposition="outside")
        st.plotly_chart(fig,use_container_width=True)

    st.markdown("### Create your own quality threshold")
    threshold=st.slider("Flag fields below",0,100,80)
    flagged=fdf[fdf["MatchRate"]<threshold/100]
    st.write(f"**{len(flagged):,}** fields are below {threshold}%.")
    st.dataframe(flagged.head(100),use_container_width=True,hide_index=True)

# ------------------------- ISSUE STUDIO -------------------------

elif page=="Issue Studio":
    st.markdown('<div class="hero"><h1>Issue Studio</h1><p>Turn findings into a living work queue with owners, statuses, notes, and priorities.</p></div>',unsafe_allow_html=True)
    issues=fdf[fdf["NeedsChange"]].copy()
    a,b,c=st.columns(3)
    a.metric("Open candidates",f"{len(issues):,}")
    b.metric("Unmatched volume",f"{int(issues['NotMatchedClaims'].fillna(0).sum()):,}")
    c.metric("Affected streams",f"{issues['Stream'].nunique():,}")

    st.dataframe(issues[[
        "Stream","NCH Target Table","NCH Target Column","MatchRate",
        "NotMatchedClaims","Classification","Disposition","Recommendation"
    ]].sort_values("NotMatchedClaims",ascending=False),use_container_width=True,hide_index=True,
    column_config={"MatchRate":st.column_config.ProgressColumn("Match Rate",min_value=0,max_value=1,format="%.1f%%")})

    st.markdown("### Create an issue")
    if len(issues):
        idx=st.selectbox("Finding",list(issues.index),format_func=lambda i:f"{issues.loc[i,'Stream']} · {issues.loc[i,'NCH Target Column']}")
        r=issues.loc[idx]
        a,b=st.columns(2)
        owner=a.text_input("Owner / team")
        status=b.selectbox("Status",["Open","In progress","Blocked","Resolved"])
        note=st.text_area("Working note")
        if st.button("Add to issue board",type="primary"):
            st.session_state.issues.append({
                "Stream":r["Stream"],"Field":r["NCH Target Column"],
                "MatchRate":r["MatchRate"],"NotMatchedClaims":r["NotMatchedClaims"],
                "Owner":owner,"Status":status,"Note":note
            })
            st.success("Issue added.")
    if st.session_state.issues:
        st.markdown("### My issue board")
        st.dataframe(pd.DataFrame(st.session_state.issues),use_container_width=True,hide_index=True)

# ------------------------- COMPARE -------------------------

elif page=="Compare":
    st.markdown('<div class="hero"><h1>Compare</h1><p>Look at the same data through different lenses, or use your uploaded dataset as the new comparison.</p></div>',unsafe_allow_html=True)
    group=st.selectbox("Compare by",["Stream","Classification","Disposition"])
    if group not in fdf:
        st.warning("This field is not available in the current dataset.")
    else:
        g=fdf.copy()
        g["Group"]=g[group].astype(str)
        s=g.groupby("Group").agg(
            Fields=("MatchRate","count"),
            AverageMatch=("MatchRate","mean"),
            Matched=("MatchedClaims","sum"),
            NotMatched=("NotMatchedClaims","sum")
        ).reset_index()
        s["WeightedMatch"]=s["Matched"]/(s["Matched"]+s["NotMatched"]).replace(0,np.nan)
        st.dataframe(s.style.format({"AverageMatch":"{:.2%}","WeightedMatch":"{:.2%}","Matched":"{:,.0f}","NotMatched":"{:,.0f}"}),use_container_width=True,hide_index=True)
        fig=px.scatter(s,x="AverageMatch",y="WeightedMatch",size="Fields",color="Group",hover_data=["Matched","NotMatched"],title="Field Average vs Event-Weighted Performance")
        fig.update_xaxes(tickformat=".0%");fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig,use_container_width=True)

# ------------------------- CONNECTIONS -------------------------

elif page=="Connections":
    st.markdown('<div class="hero"><h1>Connections</h1><p>Make system relationships explicit and reusable.</p></div>',unsafe_allow_html=True)
    conn=fdf[["Stream","NCH Target Table","NCH Target Column","SS Table","SS Column","MatchRate","Classification"]].copy()
    st.metric("Current mapped comparisons",f"{len(conn):,}")
    st.dataframe(conn,use_container_width=True,hide_index=True)
    st.markdown("### Add a future connection")
    with st.form("connection"):
        source=st.text_input("System A / source",value="NCH / CVM")
        target=st.text_input("System B / target",value="Shared Systems")
        key=st.text_input("Comparison key",placeholder="Claim ID")
        report=st.text_input("Report or dashboard")
        if st.form_submit_button("Save connection"):
            st.session_state.connections.append({"Source":source,"Target":target,"Key":key,"Report":report})
            st.success("Connection saved for this session.")
    if st.session_state.connections:
        st.dataframe(pd.DataFrame(st.session_state.connections),use_container_width=True,hide_index=True)

# ------------------------- REPORT IMPACT -------------------------

elif page=="Report Impact":
    st.markdown('<div class="hero"><h1>Report Impact</h1><p>Connect data findings to the reports, KPIs, and owners that depend on them.</p></div>',unsafe_allow_html=True)
    st.markdown('<div class="insight"><b>Product direction:</b> move from “this field has a problem” to “this field may affect this report and KPI.” The mappings below are interactive and can later be populated from lineage metadata.</div>',unsafe_allow_html=True)
    if len(fdf):
        with st.form("impact"):
            idx=st.selectbox("Data field",list(fdf.index),format_func=lambda i:f"{fdf.loc[i,'Stream']} · {fdf.loc[i,'NCH Target Column']}")
            report=st.text_input("Report / dashboard")
            kpi=st.text_input("KPI")
            owner=st.text_input("Business owner")
            impact=st.selectbox("Potential impact",["Low","Moderate","High","Critical"])
            if st.form_submit_button("Map report impact"):
                r=fdf.loc[idx]
                st.session_state.impacts.append({"Stream":r["Stream"],"Field":r["NCH Target Column"],"Report":report,"KPI":kpi,"Owner":owner,"Impact":impact})
                st.success("Report impact saved.")
    if st.session_state.impacts:
        st.dataframe(pd.DataFrame(st.session_state.impacts),use_container_width=True,hide_index=True)

# ------------------------- AI COPILOT -------------------------

elif page=="AI Copilot":
    st.markdown('<div class="hero"><h1>AI Copilot</h1><p>A conversational analyst that understands your persona, filters, mappings, and current findings.</p></div>',unsafe_allow_html=True)
    api_key=st.text_input("Groq API key (optional)",value=os.getenv("GROQ_API_KEY",""),type="password")
    model=st.selectbox("Model",["llama-3.3-70b-versatile","llama-3.1-8b-instant"])

    if not st.session_state.chat:
        st.session_state.chat=[{"role":"assistant","content":f"Hi. I'm your Medi Intelligence Copilot. Your workspace is tuned for **{p['program']}** and **{p['role']}**. Ask me what is changing, where the biggest risks are, or what to investigate next."}]
    for m in st.session_state.chat:
        with st.chat_message(m["role"]): st.markdown(m["content"])

    prompt=st.chat_input("Ask about match rates, issues, streams, reports, or priorities…")
    if prompt:
        st.session_state.chat.append({"role":"user","content":prompt})
        with st.chat_message("user"): st.markdown(prompt)

        stream_summary=fdf.groupby("Stream").agg(
            Fields=("MatchRate","count"),
            AverageMatch=("MatchRate","mean"),
            Matched=("MatchedClaims","sum"),
            NotMatched=("NotMatchedClaims","sum")
        ).reset_index().to_dict("records")
        context={
            "persona":p,
            "rows":len(fdf),
            "field_average":field_average(fdf),
            "weighted_rate":weighted_rate(fdf),
            "stream_summary":stream_summary,
            "top_actions":fdf[fdf["NeedsChange"]][["Stream","NCH Target Column","MatchRate","NotMatchedClaims","Recommendation"]].sort_values("NotMatchedClaims",ascending=False).head(20).fillna("").to_dict("records")
        }

        system="""You are a healthcare data intelligence copilot. Be concise and data-driven.
Use only supplied context; never invent statistics. Explain field-level average versus event-weighted match rate.
You are a decision-support assistant, not a clinical advisor."""
        if OpenAI and api_key:
            try:
                client=OpenAI(api_key=api_key,base_url="https://api.groq.com/openai/v1")
                messages=[{"role":"system","content":system+"\n\nCURRENT CONTEXT:\n"+json.dumps(context,default=str)}]+st.session_state.chat[-8:]
                res=client.chat.completions.create(model=model,temperature=.2,messages=messages)
                answer=res.choices[0].message.content
            except Exception as e:
                answer=f"AI service error: `{e}`\n\nLocal result: the comparable-field average is **{pct(field_average(fdf))}** and the event-weighted rate is **{pct(weighted_rate(fdf))}**."
        else:
            g=fdf.groupby("Stream")["NotMatchedClaims"].sum().sort_values(ascending=False)
            worst=g.index[0] if len(g) else "N/A"
            answer=f"Based on the current filters, the comparable-field average is **{pct(field_average(fdf))}**, the event-weighted rate is **{pct(weighted_rate(fdf))}**, and the largest unmatched volume is in **{worst}**. Add a Groq API key for conversational AI."
        st.session_state.chat.append({"role":"assistant","content":answer})
        with st.chat_message("assistant"): st.markdown(answer)

# ------------------------- AUTOMATIONS -------------------------

elif page=="Automations":
    st.markdown('<div class="hero"><h1>Automations</h1><p>Create reusable rules that turn new data conditions into notifications or actions.</p></div>',unsafe_allow_html=True)
    with st.form("automation"):
        recipient=st.text_input("Notification email")
        idx=st.selectbox("Field",list(fdf.index),format_func=lambda i:f"{fdf.loc[i,'Stream']} · {fdf.loc[i,'NCH Target Column']}") if len(fdf) else None
        field=fdf.loc[idx,"NCH Target Column"] if idx is not None else ""
        condition=st.selectbox("Condition",[
            "is empty / null","is not empty","equals value","contains text",
            "above number","below number","match rate below threshold","needs action"
        ])
        value=st.text_input("Value / threshold")
        subject=st.text_input("Subject","Medi Intelligence Alert")
        body=st.text_area("Structured email template","""A monitored condition was triggered.

Stream: {{Stream}}
Field: {{Field}}
Match Rate: {{MatchRate}}
Matched Claims: {{MatchedClaims}}
Not Matched Claims: {{NotMatchedClaims}}
Classification: {{Classification}}
Recommendation: {{Recommendation}}

Please review the Medi Intelligence workspace.""")
        if st.form_submit_button("Create automation",type="primary"):
            st.session_state.rules.append({"Recipient":recipient,"Field":field,"Condition":condition,"Value":value,"Subject":subject,"Body":body})
            st.success("Automation template saved for this session.")

    if st.session_state.rules:
        st.markdown("### My automation templates")
        st.dataframe(pd.DataFrame(st.session_state.rules),use_container_width=True,hide_index=True)

    st.markdown("### Test a rule against current data")
    if st.session_state.rules:
        ri=st.selectbox("Template",range(len(st.session_state.rules)),format_func=lambda i:st.session_state.rules[i]["Subject"])
        r=st.session_state.rules[ri]
        if st.button("Run test"):
            hit=fdf.copy()
            col=r["Field"]
            try:
                if r["Condition"]=="is empty / null": hit=hit[hit[col].isna() | hit[col].astype(str).str.strip().eq("")]
                elif r["Condition"]=="is not empty": hit=hit[hit[col].notna() & hit[col].astype(str).str.strip().ne("")]
                elif r["Condition"]=="equals value": hit=hit[hit[col].astype(str).str.lower()==r["Value"].lower()]
                elif r["Condition"]=="contains text": hit=hit[hit[col].astype(str).str.contains(r["Value"],case=False,na=False)]
                elif r["Condition"]=="above number": hit=hit[pd.to_numeric(hit[col],errors="coerce")>float(r["Value"])]
                elif r["Condition"]=="below number": hit=hit[pd.to_numeric(hit[col],errors="coerce")<float(r["Value"])]
                elif r["Condition"]=="match rate below threshold": hit=hit[hit["MatchRate"]<float(r["Value"])/100]
                elif r["Condition"]=="needs action": hit=hit[hit["NeedsChange"]]
                st.metric("Triggered rows",f"{len(hit):,}")
                st.dataframe(hit.head(100),use_container_width=True,hide_index=True)
            except Exception as e:
                st.error(f"Could not test rule: {e}")

# ------------------------- MY DATA -------------------------

elif page=="My Data":
    st.markdown('<div class="hero"><h1>My Data</h1><p>Bring your own dataset into the product and immediately use the same exploration and analysis tools.</p></div>',unsafe_allow_html=True)
    st.write(f"**Source:** `{st.session_state.source}` · **Rows:** {len(df):,} · **Columns:** {len(df.columns):,}")
    st.markdown("### Detected structure")
    profile=pd.DataFrame({
        "Column":df.columns,
        "Type":[str(df[c].dtype) for c in df.columns],
        "Non-null":[int(df[c].notna().sum()) for c in df.columns],
        "Unique":[int(df[c].nunique(dropna=True)) for c in df.columns]
    })
    st.dataframe(profile,use_container_width=True,hide_index=True)

    st.markdown("### Interactive data table")
    rows=st.slider("Rows to display",10,min(500,max(10,len(df))),min(100,len(df)))
    st.dataframe(df.head(rows),use_container_width=True,hide_index=True)

    st.download_button(
        "Download current filtered data",
        fdf.to_csv(index=False).encode("utf-8"),
        "medi_intelligence_filtered.csv",
        "text/csv"
    )

# ------------------------- Footer -------------------------

st.sidebar.markdown("---")
st.sidebar.caption(f"Loaded: {st.session_state.source}")
st.sidebar.caption(f"{len(df):,} rows · one-file product edition")
