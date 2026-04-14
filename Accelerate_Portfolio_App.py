import streamlit as st
import pandas as pd
import json
import os
import io
import fitz  # PyMuPDF
from docx import Document as DocxDocument
from anthropic import Anthropic
from datetime import datetime

# Google Drive imports
try:
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaIoBaseDownload
    from google.oauth2 import service_account
    GDRIVE_AVAILABLE = True
except ImportError:
    GDRIVE_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Venture Intelligence | NEN Resources Network",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

client = Anthropic()

# ─── Session state init ───────────────────────────────────────────────────
if "chat_messages" not in st.session_state:
    st.session_state.chat_messages = []
if "analysis_cache" not in st.session_state:
    st.session_state.analysis_cache = {}

# ─────────────────────────────────────────────────────────────────────────────
# STYLES — light theme matching Expert Search app
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Light theme badge colours */
.badge-active  { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-alumni  { background:#ede9fe; color:#5b21b6; border:1px solid #c4b5fd; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-idea    { background:#fef3c7; color:#92400e; border:1px solid #fcd34d; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-mvp     { background:#dbeafe; color:#1e40af; border:1px solid #93c5fd; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }
.badge-growth  { background:#d1fae5; color:#065f46; border:1px solid #6ee7b7; padding:2px 10px; border-radius:20px; font-size:11px; font-weight:600; }

/* Cards */
.nen-card        { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; padding:20px 24px; margin-bottom:16px; }
.nen-card-accent { background:#eff6ff; border:1px solid #bfdbfe; border-radius:12px; padding:20px 24px; margin-bottom:16px; }

/* AI insight callout */
.ai-insight { background:#eff6ff; border-left:3px solid #2563eb; padding:14px 18px; border-radius:0 10px 10px 0; font-size:14px; line-height:1.7; color:#1e3a5f; margin:10px 0; }

/* Journey arrow */
.journey-arrow { color:#2563eb; font-size:18px; margin:0 8px; }

/* Star colours */
.star-5 { color:#d97706; }
.star-4 { color:#65a30d; }
.star-3 { color:#2563eb; }
.star-2 { color:#ea580c; }
.star-1 { color:#dc2626; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING & FILE HANDLING
# ─────────────────────────────────────────────────────────────────────────────

# Keyword maps for auto-classification
FILE_TYPE_KEYWORDS = {
    "deck":     ["deck", "pitch", "presentation"],
    "vp":       ["vp", "call", "meeting"],
    "expert":   ["expert", "session", "mentor"],
    "panelist": ["panelist", "panel", "investor"],
}
SUPPORTED_EXTS = {".pdf", ".docx", ".doc", ".txt", ".md"}


def classify_file(filename: str) -> str:
    """Classify a file into deck/vp/expert/panelist/other based on filename keywords."""
    name = filename.lower()
    for ftype, keywords in FILE_TYPE_KEYWORDS.items():
        if any(kw in name for kw in keywords):
            return ftype
    return "other"


def read_file_text(path: str, max_chars: int = 8000) -> str:
    """Read text from a local file path."""
    if not path or not os.path.exists(path):
        return ""
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".pdf":
            doc = fitz.open(path)
            return "\n".join(page.get_text() for page in doc)[:max_chars]
        elif ext in [".docx", ".doc"]:
            doc = DocxDocument(path)
            return "\n".join(p.text for p in doc.paragraphs)[:max_chars]
        elif ext in [".txt", ".md"]:
            return open(path, encoding="utf-8").read()[:max_chars]
    except Exception as e:
        return f"[Could not read file: {e}]"
    return ""


def read_file_bytes(file_bytes: bytes, filename: str, max_chars: int = 8000) -> str:
    """Read text from file bytes (used for Google Drive downloads)."""
    ext = os.path.splitext(filename)[1].lower()
    try:
        if ext == ".pdf":
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            return "\n".join(page.get_text() for page in doc)[:max_chars]
        elif ext in [".docx", ".doc"]:
            doc = DocxDocument(io.BytesIO(file_bytes))
            return "\n".join(p.text for p in doc.paragraphs)[:max_chars]
        elif ext in [".txt", ".md"]:
            return file_bytes.decode("utf-8", errors="ignore")[:max_chars]
    except Exception as e:
        return f"[Could not read file bytes: {e}]"
    return ""


@st.cache_resource
def get_gdrive_service():
    """Build and cache Google Drive service using Streamlit secrets."""
    if not GDRIVE_AVAILABLE:
        return None
    try:
        creds_dict = dict(st.secrets["gdrive_service_account"])
        # Fix escaped newlines in private key
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = service_account.Credentials.from_service_account_info(
            creds_dict,
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
        return build("drive", "v3", credentials=creds, cache_discovery=False)
    except Exception as e:
        return None


def gdrive_find_venture_folder(service, venture_name: str, parent_folder_name: str = "Venture_Docs") -> tuple:
    """
    Find a venture subfolder by name inside the parent Venture_Docs folder.
    Returns (folder_id or None, debug_message)
    """
    try:
        # Search for parent folder — try both Venture_Docs and Venture Docs
        parent_id = None
        for pname in [parent_folder_name, parent_folder_name.replace("_", " ")]:
            res = service.files().list(
                q=f"name='{pname}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                fields="files(id, name)",
                pageSize=5
            ).execute()
            parents = res.get("files", [])
            if parents:
                parent_id = parents[0]["id"]
                break

        if not parent_id:
            return None, f"Parent folder '{parent_folder_name}' not found in Drive. Make sure you shared it with the service account."

        # List ALL subfolders to help debug name mismatches
        res_all = service.files().list(
            q=f"mimeType='application/vnd.google-apps.folder' and '{parent_id}' in parents and trashed=false",
            fields="files(id, name)",
            pageSize=50
        ).execute()
        all_folders = res_all.get("files", [])
        all_names = [f["name"] for f in all_folders]

        # Case-insensitive match
        for f in all_folders:
            if f["name"].lower().strip() == venture_name.lower().strip():
                return f["id"], "ok"

        return None, f"Folder '{venture_name}' not found. Available folders: {', '.join(all_names) if all_names else 'none'}"

    except Exception as e:
        return None, f"Drive error: {str(e)}"


def gdrive_list_files(service, folder_id: str) -> list:
    """List all supported files in a Drive folder."""
    supported_mimes = {
        "application/pdf": ".pdf",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
        "text/plain": ".txt",
    }
    try:
        res = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="files(id, name, mimeType, size)",
            pageSize=50
        ).execute()
        files = []
        for f in res.get("files", []):
            ext = os.path.splitext(f["name"])[1].lower()
            mime = f.get("mimeType", "")
            if ext in SUPPORTED_EXTS or mime in supported_mimes:
                files.append(f)
        return files
    except Exception:
        return []


def gdrive_download_file(service, file_id: str) -> bytes:
    """Download a file from Drive and return raw bytes."""
    try:
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()
    except Exception:
        return b""


def scan_venture_folder(venture_name: str, folder_path: str = "") -> dict:
    """
    Scan venture files from Google Drive (primary) or local path (fallback).
    Matches venture folder by name inside Venture_Docs on Drive.
    Returns classified file dict.
    """
    result = {"deck": [], "vp": [], "expert": [], "panelist": [], "other": [],
              "all_files": [], "folder_exists": False, "total": 0,
              "source": "none"}

    # ── Try Google Drive first ─────────────────────────────────────────────
    service = get_gdrive_service()
    lookup_name = folder_path.strip() if folder_path.strip() else venture_name.strip()

    if service and lookup_name:
        folder_id, debug_msg = gdrive_find_venture_folder(service, lookup_name)
        result["gdrive_debug"] = debug_msg
        if folder_id:
            result["folder_exists"] = True
            result["source"] = "gdrive"
            drive_files = gdrive_list_files(service, folder_id)
            for f in drive_files:
                fname = f["name"]
                ftype = classify_file(fname)
                max_c = 6000 if ftype == "deck" else 4000
                raw = gdrive_download_file(service, f["id"])
                text = read_file_bytes(raw, fname, max_chars=max_c) if raw else ""
                entry = {"path": f"gdrive://{f['id']}", "name": fname,
                         "type": ftype, "text": text}
                result[ftype].append(entry)
                result["all_files"].append(entry)
            result["total"] = len(result["all_files"])
            return result

    # ── Fallback: local path ───────────────────────────────────────────────
    if folder_path and os.path.isdir(folder_path):
        result["folder_exists"] = True
        result["source"] = "local"
        for fname in sorted(os.listdir(folder_path)):
            ext = os.path.splitext(fname)[1].lower()
            if ext not in SUPPORTED_EXTS:
                continue
            fpath = os.path.join(folder_path, fname)
            ftype = classify_file(fname)
            max_c = 6000 if ftype == "deck" else 4000
            text = read_file_text(fpath, max_chars=max_c)
            entry = {"path": fpath, "name": fname, "type": ftype, "text": text}
            result[ftype].append(entry)
            result["all_files"].append(entry)
        result["total"] = len(result["all_files"])

    return result


@st.cache_data(show_spinner=False)
def load_excel(file_bytes: bytes) -> dict:
    """Load all sheets from uploaded Excel file."""
    xl = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None)
    data = {}
    for sheet, df in xl.items():
        df.columns = [str(c).strip() for c in df.columns]
        for col in df.columns:
            if "date" in col.lower() or "Date" in col:
                try:
                    df[col] = pd.to_datetime(df[col], errors="coerce")
                except Exception:
                    pass
        data[sheet.strip()] = df
    return data


def get_venture_meetings(venture_id: str, data: dict) -> dict:
    """Collect all meetings across three types for a venture."""
    result = {}
    for sheet_key, mtype in [("VP_Sessions","VP"), ("Expert_Sessions","Expert"), ("Panelist_Calls","Panelist")]:
        df = data.get(sheet_key, pd.DataFrame())
        if not df.empty and "Venture_ID" in df.columns:
            result[mtype] = df[df["Venture_ID"] == venture_id].copy()
        else:
            result[mtype] = pd.DataFrame()
    return result


# ─────────────────────────────────────────────────────────────────────────────
# AI ANALYSIS FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def ai_analyze_venture(venture_row: pd.Series, meetings: dict, scanned_files: dict) -> dict:
    """Use Claude to deeply analyze a single venture's journey and meeting effectiveness."""

    venture_info = venture_row.to_dict()

    # Build structured meeting text from Excel data
    meeting_text = ""
    for mtype, df in meetings.items():
        if df.empty:
            continue
        meeting_text += f"\n\n--- {mtype} SESSIONS (from Excel) ---\n"
        for _, row in df.iterrows():
            meeting_text += f"\nDate: {row.get('Date','N/A')}\n"
            if mtype == "VP":
                meeting_text += f"VP: {row.get('VP_Name','N/A')}\n"
                meeting_text += f"Action Items: {row.get('Action_Items','')}\n"
                meeting_text += f"Notes: {row.get('Notes','')}\n"
            elif mtype == "Expert":
                meeting_text += f"Expert: {row.get('Expert_Name','N/A')} ({row.get('Expertise_Area','')})\n"
                meeting_text += f"Problems: {row.get('Problems_Discussed','')}\n"
                meeting_text += f"Actions: {row.get('Action_Items','')}\n"
                meeting_text += f"Rating: {row.get('Effectiveness_Rating','N/A')}/5\n"
                meeting_text += f"Notes: {row.get('Notes','')}\n"
            elif mtype == "Panelist":
                meeting_text += f"Panelists: {row.get('Panelists','N/A')}\n"
                meeting_text += f"Focus: {row.get('Focus_Area','')}\n"
                meeting_text += f"Feedback: {row.get('Key_Feedback','')}\n"
                meeting_text += f"Actions: {row.get('Action_Items','')}\n"
                meeting_text += f"Rating: {row.get('Overall_Rating','N/A')}/5\n"

    # Append folder-scanned transcript files grouped by type
    for ftype_label, sheet_key in [("VP","vp"), ("Expert","expert"), ("Panelist","panelist")]:
        files = scanned_files.get(sheet_key, [])
        if files:
            meeting_text += f"\n\n--- {ftype_label} TRANSCRIPTS (from folder) ---\n"
            for f in files:
                meeting_text += f"\n[File: {f['name']}]\n{f['text'][:2500]}\n"

    # Deck content — concatenate all deck files
    deck_parts = [f["text"] for f in scanned_files.get("deck", []) if f["text"]]
    deck_text = "\n\n---\n\n".join(deck_parts)
    deck_section = f"\nPITCH DECK CONTENT:\n{deck_text}\n" if deck_text else ""
    has_deck_str = "true" if deck_text else "false"

    # Other unclassified files — include as additional context
    other_files = scanned_files.get("other", [])
    other_section = ""
    if other_files:
        other_section = "\n\nADDITIONAL DOCUMENTS (from folder):\n"
        for f in other_files:
            other_section += f"\n[File: {f['name']}]\n{f['text'][:1500]}\n"

    prompt = f"""You are an analyst for a startup accelerator program. Analyze this venture using all available data including their pitch deck and meeting transcripts.

VENTURE:
Name: {venture_info.get('Venture_Name')}
Founder: {venture_info.get('Founder_Name')}
Sector: {venture_info.get('Sector')}
Starting Stage: {venture_info.get('Stage_Start')}
Current Stage: {venture_info.get('Stage_Current')}
Description: {venture_info.get('Description')}
Status: {venture_info.get('Status')}
{deck_section}
MEETINGS & SESSIONS:
{meeting_text}{other_section}

Provide a comprehensive analysis in this EXACT JSON format (no markdown, no extra text):
{{
  "deck_insights": {{
    "has_deck": {has_deck_str},
    "problem_statement": "problem as stated in deck, or Not available if no deck",
    "solution_summary": "solution as described in deck, or Not available",
    "deck_strengths": ["strength 1", "strength 2"],
    "deck_gaps": ["gap 1", "gap 2"],
    "deck_vs_reality": "one sentence on how deck narrative compares to what emerged in sessions"
  }},
  "key_problems": [
    {{"problem": "Problem description", "source": "deck/VP session/Expert session/Panelist call", "status": "resolved/ongoing/unaddressed"}}
  ],
  "journey_narrative": "2-3 paragraph narrative of where they started, what happened in the program, where they are now. Reference deck claims vs session realities where relevant. Write it like a story.",
  "stage_progression": {{
    "from": "starting stage",
    "to": "current stage",
    "momentum": "strong/moderate/slow/stalled",
    "momentum_reason": "one sentence"
  }},
  "meeting_effectiveness": [
    {{
      "meeting_ref": "session id or date + type",
      "type": "VP/Expert/Panelist",
      "effectiveness_score": 1-5,
      "impact": "what specifically changed or was unlocked",
      "is_most_impactful": true/false
    }}
  ],
  "most_impactful_meeting": {{
    "ref": "session reference",
    "type": "type",
    "reason": "why this was the turning point"
  }},
  "biggest_unresolved_problem": "one sentence",
  "overall_health_score": 1-10,
  "health_rationale": "one sentence",
  "next_priority": "top recommended next action for this venture"
}}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=2500,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip().rstrip("```").strip()
    return json.loads(raw)


def ai_portfolio_analysis(ventures_df: pd.DataFrame, data: dict) -> dict:
    """Portfolio-level AI analysis."""

    summary_rows = []
    for _, row in ventures_df.iterrows():
        vid = row.get("Venture_ID","")
        vp_count = len(data.get("VP_Sessions", pd.DataFrame()).pipe(lambda df: df[df["Venture_ID"]==vid] if "Venture_ID" in df.columns else pd.DataFrame()))
        ex_count = len(data.get("Expert_Sessions", pd.DataFrame()).pipe(lambda df: df[df["Venture_ID"]==vid] if "Venture_ID" in df.columns else pd.DataFrame()))
        pa_count = len(data.get("Panelist_Calls", pd.DataFrame()).pipe(lambda df: df[df["Venture_ID"]==vid] if "Venture_ID" in df.columns else pd.DataFrame()))

        # Avg expert rating
        ex_df = data.get("Expert_Sessions", pd.DataFrame())
        avg_ex = ""
        if not ex_df.empty and "Venture_ID" in ex_df.columns and "Effectiveness_Rating" in ex_df.columns:
            subset = ex_df[ex_df["Venture_ID"]==vid]["Effectiveness_Rating"].dropna()
            avg_ex = f"{subset.mean():.1f}" if len(subset) > 0 else "N/A"

        summary_rows.append({
            "id": vid,
            "name": row.get("Venture_Name",""),
            "sector": row.get("Sector",""),
            "stage_start": row.get("Stage_Start",""),
            "stage_current": row.get("Stage_Current",""),
            "status": row.get("Status",""),
            "vp_sessions": vp_count,
            "expert_sessions": ex_count,
            "panelist_calls": pa_count,
            "avg_expert_rating": avg_ex
        })

    prompt = f"""You are a portfolio analyst for a startup accelerator. Analyze this portfolio.

PORTFOLIO DATA:
{json.dumps(summary_rows, indent=2)}

Provide portfolio-level insights in this EXACT JSON format (no markdown):
{{
  "portfolio_health": {{
    "score": 1-10,
    "summary": "2-3 sentence portfolio health summary"
  }},
  "stage_distribution": {{
    "summary": "one sentence on where most ventures are"
  }},
  "most_engaged_venture": {{
    "name": "venture name",
    "reason": "why"
  }},
  "needs_attention": [
    {{"name": "venture name", "reason": "specific concern"}}
  ],
  "best_performing": [
    {{"name": "venture name", "reason": "specific achievement"}}
  ],
  "meeting_type_effectiveness": {{
    "most_effective_type": "VP/Expert/Panelist",
    "reason": "why this type seems most impactful across portfolio",
    "least_effective_type": "type",
    "least_effective_reason": "why"
  }},
  "portfolio_recommendations": [
    "recommendation 1",
    "recommendation 2",
    "recommendation 3"
  ],
  "cohort_narrative": "A 3-4 sentence story of this cohort overall - their collective journey, what worked, what didn't."
}}"""

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1500,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = response.content[0].text.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
    raw = raw.strip().rstrip("```").strip()
    return json.loads(raw)


# ─────────────────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def stage_badge(stage: str) -> str:
    s = str(stage).lower()
    if "idea" in s:
        return f'<span class="badge-idea">💡 Idea</span>'
    elif "mvp" in s:
        return f'<span class="badge-mvp">🔧 MVP</span>'
    elif "growth" in s:
        return f'<span class="badge-growth">📈 Growth</span>'
    return f'<span class="badge-mvp">{stage}</span>'

def status_badge(status: str) -> str:
    s = str(status).lower()
    if "active" in s:
        return f'<span class="badge-active">● Active</span>'
    elif "alumni" in s:
        return f'<span class="badge-alumni">★ Alumni</span>'
    return f'<span class="badge-idea">{status}</span>'

def rating_stars(r) -> str:
    try:
        n = int(float(r))
    except:
        return "—"
    stars = "★" * n + "☆" * (5 - n)
    cls = f"star-{n}"
    return f'<span class="{cls}">{stars}</span>'

def health_color(score: int) -> str:
    if score >= 8: return "#16a34a"
    if score >= 6: return "#60a5fa"
    if score >= 4: return "#fbbf24"
    return "#dc2626"


# ─────────────────────────────────────────────────────────────────────────────
# PAGES
# ─────────────────────────────────────────────────────────────────────────────

def render_venture_detail(venture_row: pd.Series, meetings: dict, data: dict):
    vid = venture_row.get("Venture_ID","")
    vname = venture_row.get("Venture_Name","Unknown")

    st.markdown(f"## 🚀 {vname}")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"**Sector** &nbsp; {venture_row.get('Sector','—')}", unsafe_allow_html=True)
    with col2:
        st.markdown(f"**Founder** &nbsp; {venture_row.get('Founder_Name','—')}", unsafe_allow_html=True)
    with col3:
        st.markdown(f"**Cohort** &nbsp; {venture_row.get('Cohort','—')}", unsafe_allow_html=True)
    with col4:
        st.markdown(status_badge(str(venture_row.get("Status",""))), unsafe_allow_html=True)

    st.markdown(f'<div class="nen-card"><em style="color:#374151;">"{venture_row.get("Description","")}"</em></div>', unsafe_allow_html=True)

    # Journey stage progression
    st.markdown("### Stage Journey")
    stage_start = venture_row.get("Stage_Start","?")
    stage_now = venture_row.get("Stage_Current","?")
    st.markdown(
        f'<div class="nen-card" style="display:flex;align-items:center;gap:8px;">'
        f'<span style="color:#475569;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Started as</span>&nbsp;'
        f'{stage_badge(stage_start)}'
        f'<span class="journey-arrow">→</span>'
        f'<span style="color:#475569;font-size:12px;text-transform:uppercase;letter-spacing:1px;">Now</span>&nbsp;'
        f'{stage_badge(stage_now)}'
        f'</div>',
        unsafe_allow_html=True
    )

    # ── Folder scanning ────────────────────────────────────────────────────
    vp_df = meetings.get("VP", pd.DataFrame())
    ex_df = meetings.get("Expert", pd.DataFrame())
    pa_df = meetings.get("Panelist", pd.DataFrame())

    folder_path = str(venture_row.get("Folder_Path", "")).strip()
    scanned = scan_venture_folder(venture_name=vname, folder_path=folder_path)

    # ── Source badge ───────────────────────────────────────────────────────
    source = scanned.get("source", "none")
    source_badge = {
        "gdrive": '<span style="background:#f0fdf4;color:#16a34a;border:1px solid #6ee7b7;padding:2px 10px;border-radius:20px;font-size:11px;">☁️ Google Drive</span>',
        "local":  '<span style="background:#eff6ff;color:#2563eb;border:1px solid #93c5fd;padding:2px 10px;border-radius:20px;font-size:11px;">💾 Local</span>',
        "none":   '<span style="background:#fefce8;color:#d97706;border:1px solid #fcd34d;padding:2px 10px;border-radius:20px;font-size:11px;">⚠️ No files</span>',
    }.get(source, "")

    # ── Metrics row ────────────────────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("VP Sessions", len(vp_df))
    c2.metric("Expert Sessions", len(ex_df))
    c3.metric("Panelist Calls", len(pa_df))
    c4.metric("Files Found", scanned["total"] if scanned["folder_exists"] else "—")
    c5.metric("Pitch Deck", f"{len(scanned['deck'])} file(s)" if scanned["deck"] else ("No deck" if scanned["folder_exists"] else "—"))

    if source_badge:
        st.markdown(f'<div style="margin-bottom:8px;">Files source: {source_badge}</div>', unsafe_allow_html=True)

    # ── Folder file browser ────────────────────────────────────────────────
    if not scanned["folder_exists"]:
        if source == "none":
            gdrive_ok = get_gdrive_service() is not None
            debug_msg = scanned.get("gdrive_debug", "")
            if gdrive_ok:
                if "Available folders:" in debug_msg:
                    st.warning(f"⚠️ {debug_msg}")
                elif "not found in Drive" in debug_msg:
                    st.error(f"🔑 {debug_msg} — go to Google Drive, right-click **Venture_Docs** and share with the service account email.")
                else:
                    st.warning(f"⚠️ {debug_msg or f'No folder found for {vname}'}")
            else:
                st.info("💡 Google Drive not configured yet. Add `gdrive_service_account` to Streamlit secrets.")
    else:
        type_colors = {"deck":"#fbbf24","vp":"#c4b5fd","expert":"#6ee7b7","panelist":"#fcd34d","other":"#94a3b8"}
        type_labels = {"deck":"📋 Deck","vp":"🎙 VP","expert":"💡 Expert","panelist":"🏛 Panelist","other":"📄 Other"}

        with st.expander(f"📁 Venture Folder  ·  {scanned['total']} file(s) found  ·  Click to browse", expanded=False):
            if scanned["all_files"]:
                # Summary chips
                chip_html = ""
                for ftype in ["deck","vp","expert","panelist","other"]:
                    count = len(scanned[ftype])
                    if count:
                        chip_html += (
                            f'<span style="background:#f8fafc;border:1px solid {type_colors[ftype]};'
                            f'color:{type_colors[ftype]};padding:3px 10px;border-radius:20px;'
                            f'font-size:11px;margin-right:6px;">{type_labels[ftype]} × {count}</span>'
                        )
                st.markdown(f'<div style="margin-bottom:14px;">{chip_html}</div>', unsafe_allow_html=True)

                # File list
                for entry in scanned["all_files"]:
                    tc = type_colors.get(entry["type"], "#94a3b8")
                    tl = type_labels.get(entry["type"], "📄 Other")
                    preview = entry["text"][:400].replace("\n"," ").strip() if entry["text"] else "[No text extracted]"
                    with st.expander(f"{entry['name']}  ·  {tl}", expanded=False):
                        st.markdown(
                            f'<div style="background:#f1f5f9;border:1px solid #e2e8f0;border-radius:8px;'
                            f'padding:12px;font-size:12px;color:#475569;line-height:1.7;white-space:pre-wrap;">'
                            f'{preview}{"…" if len(entry["text"])>400 else ""}</div>',
                            unsafe_allow_html=True
                        )
            else:
                st.warning("No supported files found (.pdf, .docx, .txt, .md)")

    # ── AI Analysis ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 🤖 AI Intelligence Report")

    if "analysis_cache" not in st.session_state:
        st.session_state.analysis_cache = {}

    can_analyze = scanned["folder_exists"] or (not vp_df.empty or not ex_df.empty or not pa_df.empty)

    if vid not in st.session_state.analysis_cache:
        file_summary = f"{scanned['total']} folder file(s)" if scanned["folder_exists"] else "Excel data only"
        if st.button(f"Generate Full Analysis for {vname}  ({file_summary})", key=f"analyze_{vid}"):
            with st.spinner(f"Claude is reading {scanned['total']} files and all session data…"):
                try:
                    analysis = ai_analyze_venture(venture_row, meetings, scanned)
                    st.session_state.analysis_cache[vid] = analysis
                    st.rerun()
                except Exception as e:
                    st.error(f"Analysis failed: {e}")
    else:
        analysis = st.session_state.analysis_cache[vid]
        _render_analysis(analysis, meetings)


def _render_analysis(analysis: dict, meetings: dict):
    """Render the AI analysis results."""

    # ── Deck Insights ──────────────────────────────────────────────────────
    di = analysis.get("deck_insights", {})
    if di.get("has_deck"):
        st.markdown("#### 📋 Pitch Deck Analysis")
        col_d1, col_d2 = st.columns(2)
        with col_d1:
            st.markdown(
                f'<div class="nen-card-accent">'                f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Problem Statement (from deck)</div>'                f'<div style="font-size:14px;color:#1e293b;line-height:1.6;">{di.get("problem_statement","Not available")}</div>'                f'<div style="height:12px"></div>'                f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Solution (from deck)</div>'                f'<div style="font-size:14px;color:#1e293b;line-height:1.6;">{di.get("solution_summary","Not available")}</div>'                f'</div>',
                unsafe_allow_html=True
            )
        with col_d2:
            strengths = di.get("deck_strengths", [])
            gaps = di.get("deck_gaps", [])
            s_html = "".join(f'<div style="color:#16a34a;font-size:13px;margin-bottom:4px;">✓ {s}</div>' for s in strengths)
            g_html = "".join(f'<div style="color:#dc2626;font-size:13px;margin-bottom:4px;">✗ {g}</div>' for g in gaps)
            st.markdown(
                f'<div class="nen-card">'                f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Deck Strengths</div>'                f'{s_html}'                f'<div style="height:10px"></div>'                f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Deck Gaps</div>'                f'{g_html}'                f'</div>',
                unsafe_allow_html=True
            )
        dvr = di.get("deck_vs_reality", "")
        if dvr:
            st.markdown(
                f'<div class="ai-insight"><span style="color:#d97706;font-weight:600;">💡 Deck vs Reality: </span>{dvr}</div>',
                unsafe_allow_html=True
            )
        st.markdown("---")

    # Health score
    hs = analysis.get("overall_health_score", 5)
    col1, col2 = st.columns([1, 3])
    with col1:
        st.markdown(
            f'<div class="nen-card" style="text-align:center;">'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Health Score</div>'
            f'<div style="font-size:52px;font-family:inherit;font-weight:800;color:{health_color(hs)}">{hs}</div>'
            f'<div style="font-size:11px;color:#475569;">/ 10</div>'
            f'<div style="font-size:12px;color:#475569;margin-top:8px;">{analysis.get("health_rationale","")}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    with col2:
        sp = analysis.get("stage_progression", {})
        momentum_colors = {"strong":"#16a34a","moderate":"#60a5fa","slow":"#fbbf24","stalled":"#dc2626"}
        m = sp.get("momentum","moderate")
        mc = momentum_colors.get(m, "#60a5fa")
        st.markdown(
            f'<div class="nen-card-accent">'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;">Journey Narrative</div>'
            f'<div style="font-size:14px;line-height:1.8;color:#1e3a5f;">{analysis.get("journey_narrative","")}</div>'
            f'<div style="margin-top:14px;display:flex;align-items:center;gap:10px;">'
            f'<span style="font-size:11px;color:#475569;">Momentum:</span>'
            f'<span style="color:{mc};font-weight:600;text-transform:capitalize;">▲ {m}</span>'
            f'<span style="font-size:12px;color:#475569;">— {sp.get("momentum_reason","")}</span>'
            f'</div></div>',
            unsafe_allow_html=True
        )

    # Problems
    st.markdown("#### 🔍 Key Problems Identified")
    problems = analysis.get("key_problems", [])
    if problems:
        status_map = {
            "resolved":    ("✅", "#f0fdf4", "#16a34a", "#86efac"),
            "ongoing":     ("⚠️", "#fefce8", "#d97706", "#fcd34d"),
            "unaddressed": ("🔴", "#fef2f2", "#dc2626", "#fca5a5")
        }
        for p in problems:
            stat = p.get("status","ongoing").lower()
            icon, bg, tc, bc = status_map.get(stat, status_map["ongoing"])
            st.markdown(
                f'<div style="background:{bg};border:1px solid {bc};border-radius:10px;padding:12px 16px;margin-bottom:8px;">'
                f'<div style="display:flex;align-items:center;gap:10px;">'
                f'<span style="font-size:16px;">{icon}</span>'
                f'<span style="color:{tc};font-size:12px;text-transform:uppercase;letter-spacing:0.5px;font-weight:600;">{stat}</span>'
                f'<span style="color:#475569;font-size:11px;">via {p.get("source","")}</span>'
                f'</div>'
                f'<div style="color:#1e293b;margin-top:6px;font-size:14px;">{p.get("problem","")}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
    else:
        st.info("No problems data extracted.")

    # Meeting effectiveness
    st.markdown("#### 📊 Meeting Effectiveness")
    me_list = analysis.get("meeting_effectiveness", [])
    if me_list:
        type_colors = {"VP":"#7c3aed","Expert":"#059669","Panelist":"#d97706"}
        for me in sorted(me_list, key=lambda x: x.get("effectiveness_score",0), reverse=True):
            score = me.get("effectiveness_score", 0)
            tc = type_colors.get(me.get("type","Expert"), "#7eb8f7")
            is_top = me.get("is_most_impactful", False)
            border = "2px solid #d97706" if is_top else "1px solid #e2e8f0"
            st.markdown(
                f'<div style="background:#f8fafc;border:{border};border-radius:10px;padding:14px 18px;margin-bottom:8px;">'
                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                f'<div style="display:flex;align-items:center;gap:10px;">'
                f'{"🏆 " if is_top else ""}'
                f'<span style="color:{tc};font-size:11px;text-transform:uppercase;letter-spacing:1px;font-weight:700;">{me.get("type","")}</span>'
                f'<span style="color:#475569;font-size:12px;">{me.get("meeting_ref","")}</span>'
                f'</div>'
                f'<div style="color:#d97706;font-size:16px;">{"★"*score}{"☆"*(5-score)}</div>'
                f'</div>'
                f'<div style="color:#374151;font-size:13px;margin-top:8px;">{me.get("impact","")}</div>'
                f'</div>',
                unsafe_allow_html=True
            )

    # Most impactful meeting
    mim = analysis.get("most_impactful_meeting", {})
    if mim:
        st.markdown(
            f'<div class="nen-card-accent">'
            f'<div style="font-size:11px;color:#d97706;text-transform:uppercase;letter-spacing:1px;">🏆 Most Impactful Session</div>'
            f'<div style="font-size:16px;color:#111827;font-family:inherit;font-weight:700;margin:8px 0;">{mim.get("ref","")} · {mim.get("type","")}</div>'
            f'<div style="color:#374151;font-size:14px;">{mim.get("reason","")}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    # Priority
    st.markdown(
        f'<div class="ai-insight">'
        f'<span style="color:#2563eb;font-weight:600;">⚡ Top Priority: </span>{analysis.get("next_priority","")}'
        f'</div>',
        unsafe_allow_html=True
    )


def render_portfolio_view(data: dict):
    ventures_df = data.get("Ventures", pd.DataFrame())
    if ventures_df.empty:
        st.warning("No ventures data found. Check your Excel sheet name is 'Ventures'.")
        return

    st.markdown("## 📊 Portfolio Overview")

    # Top-line metrics
    total = len(ventures_df)
    active = len(ventures_df[ventures_df.get("Status","") == "Active"]) if "Status" in ventures_df.columns else 0
    alumni = len(ventures_df[ventures_df.get("Status","") == "Alumni"]) if "Status" in ventures_df.columns else 0

    vp_df = data.get("VP_Sessions", pd.DataFrame())
    ex_df = data.get("Expert_Sessions", pd.DataFrame())
    pa_df = data.get("Panelist_Calls", pd.DataFrame())
    total_meetings = len(vp_df) + len(ex_df) + len(pa_df)

    avg_expert_rating = None
    if not ex_df.empty and "Effectiveness_Rating" in ex_df.columns:
        avg_expert_rating = ex_df["Effectiveness_Rating"].dropna().mean()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Ventures", total)
    c2.metric("Active", active)
    c3.metric("Alumni", alumni)
    c4.metric("Total Sessions", total_meetings)
    c5.metric("Avg Expert Rating", f"{avg_expert_rating:.1f}/5" if avg_expert_rating else "—")

    st.markdown("---")

    # Venture table
    st.markdown("### All Ventures")
    for _, row in ventures_df.iterrows():
        vid = row.get("Venture_ID","")
        vp_c = len(vp_df[vp_df["Venture_ID"]==vid]) if not vp_df.empty and "Venture_ID" in vp_df.columns else 0
        ex_c = len(ex_df[ex_df["Venture_ID"]==vid]) if not ex_df.empty and "Venture_ID" in ex_df.columns else 0
        pa_c = len(pa_df[pa_df["Venture_ID"]==vid]) if not pa_df.empty and "Venture_ID" in pa_df.columns else 0

        avg_ex = "—"
        if not ex_df.empty and "Venture_ID" in ex_df.columns and "Effectiveness_Rating" in ex_df.columns:
            subset = ex_df[ex_df["Venture_ID"]==vid]["Effectiveness_Rating"].dropna()
            if len(subset) > 0:
                avg_ex = f"{subset.mean():.1f}"

        st.markdown(
            f'<div class="nen-card">'
            f'<div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;">'
            f'<div>'
            f'<div style="font-family:inherit;font-weight:700;font-size:16px;color:#fff;">{row.get("Venture_Name","")}</div>'
            f'<div style="color:#475569;font-size:12px;margin-top:2px;">{row.get("Founder_Name","")} · {row.get("Sector","")}</div>'
            f'</div>'
            f'<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
            f'{stage_badge(str(row.get("Stage_Start","")))} → {stage_badge(str(row.get("Stage_Current","")))}'
            f'&nbsp;&nbsp;{status_badge(str(row.get("Status","")))}'
            f'</div>'
            f'</div>'
            f'<div style="display:flex;gap:20px;margin-top:12px;">'
            f'<span style="color:#7c3aed;font-size:12px;">VP: {vp_c}</span>'
            f'<span style="color:#059669;font-size:12px;">Expert: {ex_c}</span>'
            f'<span style="color:#d97706;font-size:12px;">Panelist: {pa_c}</span>'
            f'<span style="color:#2563eb;font-size:12px;">Avg Expert Rating: {avg_ex}/5</span>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### 🤖 Portfolio AI Analysis")

    if "portfolio_analysis" not in st.session_state:
        if st.button("Generate Portfolio Intelligence Report"):
            with st.spinner("Analyzing entire portfolio…"):
                try:
                    pa = ai_portfolio_analysis(ventures_df, data)
                    st.session_state.portfolio_analysis = pa
                    st.rerun()
                except Exception as e:
                    st.error(f"Portfolio analysis failed: {e}")
    else:
        pa = st.session_state.portfolio_analysis
        _render_portfolio_analysis(pa)


def _render_portfolio_analysis(pa: dict):
    ph = pa.get("portfolio_health", {})
    hs = ph.get("score", 5)

    col1, col2 = st.columns([1,3])
    with col1:
        st.markdown(
            f'<div class="nen-card" style="text-align:center;">'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">Portfolio Health</div>'
            f'<div style="font-size:52px;font-family:inherit;font-weight:800;color:{health_color(hs)}">{hs}</div>'
            f'<div style="font-size:11px;color:#475569;">/ 10</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            f'<div class="nen-card-accent">'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;">Cohort Story</div>'
            f'<div style="font-size:14px;line-height:1.8;color:#1e3a5f;">{pa.get("cohort_narrative","")}</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### ⚠️ Needs Attention")
        for item in pa.get("needs_attention", []):
            st.markdown(
                f'<div style="background:#fef2f2;border:1px solid #fca5a5;border-radius:8px;padding:10px 14px;margin-bottom:8px;">'
                f'<div style="color:#dc2626;font-weight:600;">{item.get("name","")}</div>'
                f'<div style="color:#374151;font-size:12px;margin-top:4px;">{item.get("reason","")}</div>'
                f'</div>',
                unsafe_allow_html=True
            )

    with col2:
        st.markdown("#### 🌟 Best Performing")
        for item in pa.get("best_performing", []):
            st.markdown(
                f'<div style="background:#f0fdf4;border:1px solid #86efac;border-radius:8px;padding:10px 14px;margin-bottom:8px;">'
                f'<div style="color:#16a34a;font-weight:600;">{item.get("name","")}</div>'
                f'<div style="color:#374151;font-size:12px;margin-top:4px;">{item.get("reason","")}</div>'
                f'</div>',
                unsafe_allow_html=True
            )

    # Meeting type effectiveness
    mte = pa.get("meeting_type_effectiveness", {})
    if mte:
        st.markdown("#### 🏆 Most Impactful Meeting Type (Portfolio-Wide)")
        st.markdown(
            f'<div class="nen-card-accent">'
            f'<div style="display:flex;gap:40px;flex-wrap:wrap;">'
            f'<div>'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;">Most Effective</div>'
            f'<div style="font-size:20px;font-family:inherit;font-weight:700;color:#16a34a;margin:6px 0;">{mte.get("most_effective_type","")}</div>'
            f'<div style="font-size:13px;color:#374151;">{mte.get("reason","")}</div>'
            f'</div>'
            f'<div>'
            f'<div style="font-size:11px;color:#475569;font-weight:600;text-transform:uppercase;letter-spacing:1px;">Least Effective</div>'
            f'<div style="font-size:20px;font-family:inherit;font-weight:700;color:#dc2626;margin:6px 0;">{mte.get("least_effective_type","")}</div>'
            f'<div style="font-size:13px;color:#374151;">{mte.get("least_effective_reason","")}</div>'
            f'</div>'
            f'</div>'
            f'</div>',
            unsafe_allow_html=True
        )

    st.markdown("#### 💡 Portfolio Recommendations")
    for i, rec in enumerate(pa.get("portfolio_recommendations", []), 1):
        st.markdown(
            f'<div class="ai-insight">'
            f'<span style="color:#2563eb;font-weight:700;">{i}.</span> {rec}'
            f'</div>',
            unsafe_allow_html=True
        )


# ─────────────────────────────────────────────────────────────────────────────
# HEADER — matches Expert Search app logo layout
# ─────────────────────────────────────────────────────────────────────────────

hcol1, hcol2, hcol3 = st.columns([1, 2, 1])
with hcol1:
    if os.path.exists("DP_BG1.png"):
        st.image("DP_BG1.png", width=150)
    else:
        st.markdown(
            '<div style="font-family:inherit;font-weight:800;font-size:22px;'
            'color:#111827;padding:10px 0;">NEN</div>',
            unsafe_allow_html=True
        )
with hcol2:
    st.markdown(
        "<h2 style='text-align:center;font-family:inherit;font-weight:800;"
        "color:#111827;margin:0;padding:8px 0;'>"
        "🚀 Resources Network — Venture Intelligence"
        "</h2>",
        unsafe_allow_html=True
    )
st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    if os.path.exists("DP_BG1.png"):
        st.image("DP_BG1.png", width=130)
    st.markdown(
        '<div style="font-size:11px;color:#475569;text-transform:uppercase;'
        'letter-spacing:1.5px;margin:8px 0 16px;">Portfolio Analyzer</div>',
        unsafe_allow_html=True
    )
    st.markdown("---")

    uploaded = st.file_uploader(
        "📊 Upload Master Excel File",
        type=["xlsx","xls"],
        help="Upload your NEN master data file with Ventures, VP_Sessions, Expert_Sessions, Panelist_Calls sheets"
    )

    if uploaded:
        data = load_excel(uploaded.read())
        available_sheets = list(data.keys())
        st.success(f"✓ Loaded: {', '.join(available_sheets)}")

        st.markdown("---")
        view_mode = st.selectbox("View Mode", ["Portfolio Overview", "Venture Deep-Dive"])

        if view_mode == "Venture Deep-Dive":
            ventures_df = data.get("Ventures", pd.DataFrame())
            if not ventures_df.empty and "Venture_Name" in ventures_df.columns:
                venture_names = ventures_df["Venture_Name"].tolist()
                selected_name = st.selectbox("Select Venture", venture_names)
            else:
                selected_name = None
                st.warning("No 'Ventures' sheet found.")
        else:
            selected_name = None

        st.markdown("---")
        if st.sidebar.button("🗑️ Clear Chat & Cache"):
            st.session_state.chat_messages = []
            st.session_state.analysis_cache = {}
            st.session_state.pop("portfolio_analysis", None)
            st.rerun()
    else:
        data = None
        view_mode = None
        selected_name = None

    st.markdown("---")
    st.markdown(
        '<div style="font-size:10px;color:#475569;text-align:center;">'
        'Powered by Claude AI · NEN Resources Network</div>',
        unsafe_allow_html=True
    )


# ─────────────────────────────────────────────────────────────────────────────
# CHAT INTELLIGENCE
# ─────────────────────────────────────────────────────────────────────────────

def build_chat_context(data: dict, selected_name: str = None, view_mode: str = None) -> str:
    """Build rich context string for the chat AI from loaded data + any cached analyses."""
    ctx = ""

    ventures_df = data.get("Ventures", pd.DataFrame()) if data else pd.DataFrame()
    vp_df = data.get("VP_Sessions", pd.DataFrame()) if data else pd.DataFrame()
    ex_df = data.get("Expert_Sessions", pd.DataFrame()) if data else pd.DataFrame()
    pa_df = data.get("Panelist_Calls", pd.DataFrame()) if data else pd.DataFrame()

    if not ventures_df.empty:
        ctx += "PORTFOLIO VENTURES:\n"
        for _, r in ventures_df.iterrows():
            ctx += (f"- {r.get('Venture_Name','')} ({r.get('Venture_ID','')}) | "
                    f"Sector: {r.get('Sector','')} | "
                    f"Stage: {r.get('Stage_Start','')} → {r.get('Stage_Current','')} | "
                    f"Status: {r.get('Status','')}\n")
        ctx += "\n"

    # Include cached analyses
    for vid, analysis in st.session_state.get("analysis_cache", {}).items():
        vrow = ventures_df[ventures_df["Venture_ID"] == vid] if not ventures_df.empty else pd.DataFrame()
        vname = vrow.iloc[0].get("Venture_Name", vid) if not vrow.empty else vid
        ctx += f"\nAI ANALYSIS — {vname}:\n"
        ctx += f"Health Score: {analysis.get('overall_health_score','N/A')}/10\n"
        ctx += f"Stage: {analysis.get('stage_progression',{}).get('from','')} → {analysis.get('stage_progression',{}).get('to','')} | Momentum: {analysis.get('stage_progression',{}).get('momentum','')}\n"
        ctx += f"Journey: {analysis.get('journey_narrative','')[:600]}\n"
        ctx += f"Most Impactful Session: {analysis.get('most_impactful_meeting',{}).get('ref','')} ({analysis.get('most_impactful_meeting',{}).get('type','')}) — {analysis.get('most_impactful_meeting',{}).get('reason','')}\n"
        ctx += f"Biggest Unresolved Problem: {analysis.get('biggest_unresolved_problem','')}\n"
        ctx += f"Top Priority: {analysis.get('next_priority','')}\n"
        problems = analysis.get("key_problems", [])
        if problems:
            ctx += "Problems: " + "; ".join(f"{p.get('problem','')} [{p.get('status','')}]" for p in problems) + "\n"

    if "portfolio_analysis" in st.session_state:
        pa = st.session_state.portfolio_analysis
        ctx += f"\nPORTFOLIO ANALYSIS:\n"
        ctx += f"Health: {pa.get('portfolio_health',{}).get('score','N/A')}/10\n"
        ctx += f"Cohort Story: {pa.get('cohort_narrative','')}\n"
        ctx += f"Most Effective Meeting Type: {pa.get('meeting_type_effectiveness',{}).get('most_effective_type','')}\n"
        needs = pa.get("needs_attention", [])
        best = pa.get("best_performing", [])
        if needs:
            ctx += "Needs Attention: " + ", ".join(n.get("name","") for n in needs) + "\n"
        if best:
            ctx += "Best Performing: " + ", ".join(b.get("name","") for b in best) + "\n"

    # Session counts per venture
    if not ventures_df.empty:
        ctx += "\nSESSION COUNTS PER VENTURE:\n"
        for _, r in ventures_df.iterrows():
            vid = r.get("Venture_ID","")
            vp_c = len(vp_df[vp_df["Venture_ID"]==vid]) if not vp_df.empty and "Venture_ID" in vp_df.columns else 0
            ex_c = len(ex_df[ex_df["Venture_ID"]==vid]) if not ex_df.empty and "Venture_ID" in ex_df.columns else 0
            pa_c = len(pa_df[pa_df["Venture_ID"]==vid]) if not pa_df.empty and "Venture_ID" in pa_df.columns else 0
            avg_ex = ""
            if not ex_df.empty and "Venture_ID" in ex_df.columns and "Effectiveness_Rating" in ex_df.columns:
                subset = ex_df[ex_df["Venture_ID"]==vid]["Effectiveness_Rating"].dropna()
                avg_ex = f" | Avg Expert Rating: {subset.mean():.1f}" if len(subset) > 0 else ""
            ctx += f"- {r.get('Venture_Name','')}: VP={vp_c}, Expert={ex_c}, Panelist={pa_c}{avg_ex}\n"

    return ctx


def chat_respond(user_msg: str, data: dict, selected_name: str, view_mode: str) -> str:
    """Send user message + full context to Claude and return response."""
    context = build_chat_context(data, selected_name, view_mode)

    # Build conversation history (last 8 messages)
    history = []
    for msg in st.session_state.chat_messages[-8:]:
        history.append({"role": msg["role"], "content": msg["content"]})

    system = """You are an expert analyst assistant for NEN Resources Network, helping program managers understand their startup portfolio.

You have full context about ventures, their meeting histories, AI analyses, and portfolio health.

Guidelines:
- Answer questions about specific ventures by name — pull from the context provided
- For portfolio questions, synthesize across all ventures
- Be concise and direct — use bullet points for lists, prose for narratives
- When asked about meeting effectiveness, reference specific session types and outcomes
- If analysis hasn't been run yet for a venture, say so and suggest they generate it
- Always be helpful, factual, and grounded in the data provided"""

    messages = history + [{"role": "user", "content": f"DATA CONTEXT:\n{context}\n\nQUESTION: {user_msg}"}]

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=system,
        messages=messages
    )
    return response.content[0].text


def render_chat(data: dict, selected_name: str = None, view_mode: str = None):
    """Render the persistent chat interface below any view."""
    st.markdown("---")
    st.markdown("### 💬 Ask Anything")

    # Render chat history
    for msg in st.session_state.chat_messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Placeholder prompts
    if not st.session_state.chat_messages:
        hint = ""
        if view_mode == "Venture Deep-Dive" and selected_name:
            hint = (f"*Try: What are {selected_name}'s biggest unresolved problems? "
                    f"\u00b7 Which session helped {selected_name} the most? "
                    f"\u00b7 Summarise {selected_name}'s journey*")
        elif view_mode == "Portfolio Overview":
            hint = ("*Try: Which ventures need the most attention? "
                    "\u00b7 Which meeting type was most effective overall? "
                    "\u00b7 Compare all ventures by stage progression*")
        if hint:
            st.markdown(
                f'<div style="color:#475569;font-size:13px;padding:8px 0;">{hint}</div>',
                unsafe_allow_html=True
            )

    # Chat input — matches Expert Search app pattern exactly
    user_input = st.chat_input("Ask about a venture, the portfolio, or any session...")

    if user_input:
        with st.chat_message("user"):
            st.markdown(user_input)
        st.session_state.chat_messages.append({"role": "user", "content": user_input})

        with st.chat_message("assistant"):
            with st.spinner("Thinking..."):
                try:
                    reply = chat_respond(user_input, data, selected_name, view_mode)
                    st.markdown(reply)
                    st.session_state.chat_messages.append({"role": "assistant", "content": reply})
                except Exception as e:
                    err = f"Chat error: {e}"
                    st.error(err)
                    st.session_state.chat_messages.append({"role": "assistant", "content": err})


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

if not data:
    # Landing state
    st.markdown(
        '<div style="text-align:center;padding:40px 20px 20px;">'
        '<div style="font-family:inherit;font-weight:800;font-size:38px;color:#111827;line-height:1.1;margin-bottom:16px;">'
        'Venture Intelligence<br><span style="color:#3b82f6;">Portfolio Analyzer</span>'
        '</div>'
        '<div style="font-size:16px;color:#475569;max-width:500px;margin:0 auto 32px;">'
        'Upload your NEN master Excel file to analyze venture journeys, meeting effectiveness, and portfolio health — powered by Claude AI.'
        '</div>'
        '</div>',
        unsafe_allow_html=True
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            '<div class="nen-card" style="text-align:center;padding:28px 20px;">'
            '<div style="font-size:28px;margin-bottom:10px;">🔍</div>'
            '<div style="font-family:inherit;font-weight:700;color:#fff;margin-bottom:6px;">Venture Deep-Dive</div>'
            '<div style="font-size:13px;color:#475569;">Problems, journey narrative, meeting effectiveness per venture</div>'
            '</div>',
            unsafe_allow_html=True
        )
    with col2:
        st.markdown(
            '<div class="nen-card" style="text-align:center;padding:28px 20px;">'
            '<div style="font-size:28px;margin-bottom:10px;">📊</div>'
            '<div style="font-family:inherit;font-weight:700;color:#fff;margin-bottom:6px;">Portfolio Overview</div>'
            '<div style="font-size:13px;color:#475569;">Health scores, best performers, attention flags across all ventures</div>'
            '</div>',
            unsafe_allow_html=True
        )
    with col3:
        st.markdown(
            '<div class="nen-card" style="text-align:center;padding:28px 20px;">'
            '<div style="font-size:28px;margin-bottom:10px;">🏆</div>'
            '<div style="font-family:inherit;font-weight:700;color:#fff;margin-bottom:6px;">Meeting Intelligence</div>'
            '<div style="font-size:13px;color:#475569;">Which sessions drove the most impact — on venture & portfolio level</div>'
            '</div>',
            unsafe_allow_html=True
        )

    st.markdown("---")
    st.markdown("### Expected Excel Structure")
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.markdown(
            '<div class="nen-card">'
            '<div style="color:#2563eb;font-weight:700;margin-bottom:8px;">Sheet: Ventures</div>'
            '<div style="font-size:12px;color:#475569;line-height:1.8;">'
            'Venture_ID<br>Venture_Name<br>Founder_Name<br>Founder_Email<br>Sector<br>Stage_Start<br>Stage_Current<br>Cohort<br>Status<br>Description<br><span style="color:#d97706">Folder_Path</span>'
            '</div></div>',
            unsafe_allow_html=True
        )
    with col_b:
        st.markdown(
            '<div class="nen-card">'
            '<div style="color:#7c3aed;font-weight:700;margin-bottom:8px;">Sheet: VP_Sessions</div>'
            '<div style="font-size:12px;color:#475569;line-height:1.8;">'
            'Session_ID<br>Venture_ID<br>Date<br>VP_Name<br>Duration_Mins<br>Transcript_Path<br>Action_Items<br>Notes'
            '</div></div>',
            unsafe_allow_html=True
        )
    with col_c:
        st.markdown(
            '<div class="nen-card">'
            '<div style="color:#059669;font-weight:700;margin-bottom:8px;">Sheet: Expert_Sessions</div>'
            '<div style="font-size:12px;color:#475569;line-height:1.8;">'
            'Session_ID<br>Venture_ID<br>Date<br>Expert_Name<br>Expertise_Area<br>Duration_Mins<br>Problems_Discussed<br>Action_Items<br>Effectiveness_Rating<br>Notes'
            '</div></div>',
            unsafe_allow_html=True
        )
    with col_d:
        st.markdown(
            '<div class="nen-card">'
            '<div style="color:#d97706;font-weight:700;margin-bottom:8px;">Sheet: Panelist_Calls</div>'
            '<div style="font-size:12px;color:#475569;line-height:1.8;">'
            'Call_ID<br>Venture_ID<br>Date<br>Panelists<br>Duration_Mins<br>Focus_Area<br>Key_Feedback<br>Action_Items<br>Overall_Rating<br>Notes'
            '</div></div>',
            unsafe_allow_html=True
        )

else:
    if view_mode == "Portfolio Overview":
        render_portfolio_view(data)
        render_chat(data, None, "Portfolio Overview")

    elif view_mode == "Venture Deep-Dive" and selected_name:
        ventures_df = data.get("Ventures", pd.DataFrame())
        venture_row = ventures_df[ventures_df["Venture_Name"] == selected_name].iloc[0]
        vid = venture_row.get("Venture_ID","")
        meetings = get_venture_meetings(vid, data)
        render_venture_detail(venture_row, meetings, data)
        render_chat(data, selected_name, "Venture Deep-Dive")
