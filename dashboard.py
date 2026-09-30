#!/usr/bin/env python3
"""Course attendance & seating control panel.

One dashboard over the whole pipeline:
  1. Allocate seats     - roster in, seat allocation + seating workbook + hall-plan HTML out.
  2. Signature sheets   - allocation in, printable signature-sheet PDFs + template.json out.
  3. Seat allotment PDF - allocation(s) in, the plain student-facing block-map + roll-list PDF out.
  4. Find your block    - allocation(s) + optional TA duty file in, a shareable self-lookup
                           page out (`finder.py`) - roll number in, block or TA duty out.
  5. Move a student     - moves one student into a vacant core seat, same room (block-to-block)
                           or a different one (batch transfer), then regenerates every
                           downstream artifact (workbook, hall plan, signature sheets) for
                           whichever room(s) changed, and offers to mail the student.
  6. Email students     - mails seating to <enrolment>@<EMAIL_DOMAIN> (`mailer.py`): a block
                           transition (block only, never a seat number) or an exam seating
                           notice. Dry run by default; real sends are logged.

Separately, "Exam seating" allocates a roster into an *exam* venue parsed
straight from `LHC Seating Plan.xlsx` (`exam_rooms.py`) - a different model
from the classroom rooms above: one exam can span several physical rooms at
once, and seats are conventionally spaced out rather than filled solid.

Run with:
    streamlit run dashboard.py
"""
from __future__ import annotations
import contextlib
import io
import json
import os
import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import altair as alt          # ships with streamlit; the attendance trend chart
import streamlit as st

import answer_showing as ansshow
import attendance_pdf as attpdf
import attendance_report as attrep
import bundle
import checks
import course_settings
import course_setup
import config as C
import exam_plan
import exam_rooms
import finder
import handout as handoutmod
import history
import lockfile
import lookup
import mailer
import posters
import reports
import sheets as sheetmod
import seating
import ta_duty
from seating import (allocate, allocate_grouped, block_summary, load_roster,
                     move_student, seat_map, vacant_seats)

try:
    from pypdf import PdfWriter
    HAVE_PYPDF = True
except ImportError:
    HAVE_PYPDF = False

st.set_page_config(page_title=f"{C.COURSE_CODE} Control Panel", page_icon="🎓", layout="wide",
                   initial_sidebar_state="expanded")


@st.cache_resource(show_spinner="Installing the sample data (first start only)...")
def _bootstrap_sample_data() -> bool:
    """A fresh clone or a hosted demo has no workbooks yet: install the
    invented set from sample_data/. Never touches a workbook already there."""
    import make_sample_data
    return make_sample_data.ensure_installed()


# Every page reads the workbooks and out/ relative to the project folder. The
# launchers cd there first; a host that starts Streamlit from elsewhere does not.
os.chdir(Path(__file__).resolve().parent)
try:
    _bootstrap_sample_data()
except Exception as e:                              # noqa: BLE001 - shown verbatim
    st.error(f"Could not install the sample data: {type(e).__name__}: {e}")

# Colours and fonts are design tokens in .streamlit/config.toml, including a
# full dark palette. What's left here is layout and the few components
# Streamlit has no token for - written against the *active* theme, so dark
# mode doesn't fall back to light-grey cards on a black page.
try:
    _DARK = getattr(st.context.theme, "type", "light") == "dark"
except Exception:                                   # noqa: BLE001 - older runtimes
    _DARK = False

_T = {
    "card": "#141920" if _DARK else "#FFFFFF",
    "card_alt": "#181D24" if _DARK else "#F8F9FB",
    "line": "#232A33" if _DARK else "#E4E7EB",
    "text": "#E6E9ED" if _DARK else "#14181D",
    "muted": "#8B95A1" if _DARK else "#6B7280",
    "accent": "#2DD4BF" if _DARK else "#0F766E",
    "accent_soft": "#0E2E2C" if _DARK else "#E6F2F0",
    "warn_bg": "#2A2113" if _DARK else "#FEF6E7",
    "warn_line": "#4A3A1B" if _DARK else "#F0DFBB",
    "warn_text": "#E0A94A" if _DARK else "#8A5A11",
    "shadow": "0 1px 2px rgba(0,0,0,.35), 0 8px 24px rgba(0,0,0,.35)" if _DARK
              else "0 1px 2px rgba(16,24,40,.04), 0 8px 24px rgba(16,24,40,.06)",
}

st.markdown(f"""
<style>
  /* Material *Symbols* (not the older Material Icons): Streamlit's own
     `:material/…:` widgets draw from this set, so a name that works on a nav
     button also works in a page header. The older set silently renders an
     empty box for anything added since 2021 - "stethoscope", for one. */
  @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,400,0,0');
  .phl-sym {{
      font-family:'Material Symbols Rounded'; font-weight:normal; font-style:normal;
      line-height:1; letter-spacing:normal; text-transform:none; display:inline-block;
      white-space:nowrap; direction:ltr; font-feature-settings:'liga';
      -webkit-font-smoothing:antialiased;
  }}
  .block-container {{padding-top: 2.4rem; padding-bottom: 4rem; max-width: 1180px;}}

  /* ── sidebar ───────────────────────────────────────────────────────── */
  .phl-brand {{display:flex; align-items:center; gap:.65rem; padding:.1rem 0 1rem;}}
  .phl-brand .logo {{
      width:34px; height:34px; border-radius:10px; flex:none;
      background:linear-gradient(140deg,#2DD4BF,#0F766E);
      color:#04211F; font-weight:700; font-size:.82rem; letter-spacing:.02em;
      display:flex; align-items:center; justify-content:center;
  }}
  .phl-brand .t {{font-weight:650; font-size:1rem; line-height:1.15; color:#F3F5F7;}}
  .phl-brand .s {{font-size:.75rem; color:#8B95A1;}}
  section[data-testid="stSidebar"] .stCaption {{margin:.9rem 0 .1rem .2rem;}}
  section[data-testid="stSidebar"] .stCaption p {{
      text-transform:uppercase; letter-spacing:.08em; font-size:.64rem;
      color:#6E7885; font-weight:600;
  }}
  section[data-testid="stSidebar"] .stButton button {{
      justify-content:flex-start !important; text-align:left; font-weight:500;
      padding:.4rem .6rem; border:1px solid transparent; min-height:0;
      color:#C3CAD3;
  }}
  section[data-testid="stSidebar"] .stButton button > div {{
      justify-content:flex-start !important; width:100%; gap:.6rem;
  }}
  section[data-testid="stSidebar"] .stButton button p {{font-size:.87rem;}}
  section[data-testid="stSidebar"] .stButton button:hover {{
      background:#1B2027; color:#FFFFFF;
  }}
  section[data-testid="stSidebar"] .stButton button[kind="primary"] {{
      background:#152A2A !important; color:#5EEAD4 !important; font-weight:600;
      box-shadow: inset 2px 0 0 0 #2DD4BF;
  }}
  section[data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {{gap:.1rem;}}
  .phl-side-foot {{
      margin-top:1.6rem; padding-top:.8rem; border-top:1px solid #232A33;
      color:#6E7885; font-size:.72rem; line-height:1.4;
  }}
  .phl-side-foot code {{color:#8B95A1; background:transparent; font-size:.7rem;}}

  /* ── status chips ──────────────────────────────────────────────────── */
  .phl-chips {{display:flex; gap:.35rem; flex-wrap:wrap; margin:.1rem 0 .3rem;}}
  .phl-chips.vert {{flex-direction:column; gap:.3rem; margin-bottom:.9rem;}}
  .phl-chip {{
      background:#161B21; border:1px solid #232A33; border-radius:9px;
      padding:.35rem .6rem; font-size:.76rem; color:#A9B2BD; white-space:nowrap;
  }}
  .phl-chip b {{color:#5EEAD4; font-weight:600;}}
  .phl-chip.warn {{background:#2A2113; border-color:#4A3A1B;}}
  .phl-chip.warn b {{color:#E0A94A;}}

  /* ── page identity ─────────────────────────────────────────────────── */
  .phl-page {{display:flex; gap:.85rem; align-items:flex-start; margin:0 0 1.4rem;}}
  .phl-page .ic {{
      width:40px; height:40px; border-radius:11px; flex:none; font-size:21px;
      background:{_T['accent_soft']}; color:{_T['accent']};
      display:flex; align-items:center; justify-content:center; margin-top:.15rem;
  }}
  .phl-page h2 {{
      margin:0; font-size:1.5rem; font-weight:680; letter-spacing:-.015em;
      line-height:1.2; color:inherit; border:0; padding:0;
  }}
  .phl-page p {{margin:.25rem 0 0; color:inherit; opacity:.62;
                font-size:.88rem; max-width:70ch;}}

  /* ── content ───────────────────────────────────────────────────────── */
  h3 {{font-size:1.05rem; font-weight:620; letter-spacing:-.01em; margin-top:1.2rem;}}
  h4 {{font-size:.95rem; font-weight:600; color:inherit; opacity:.72;
       text-transform:none; margin-top:1rem;}}
  div[data-testid="stVerticalBlockBorderWrapper"] {{
      background:{_T['card']}; border-radius:14px; box-shadow:{_T['shadow']};
  }}
  details[data-testid="stExpander"] {{
      border-radius:12px; overflow:hidden; background:{_T['card']};
      box-shadow:{_T['shadow']};
  }}
  div[data-testid="stMetric"] {{
      background:{_T['card_alt']}; border:1px solid {_T['line']};
      border-radius:12px; padding:.85rem 1rem;
  }}
  div[data-testid="stMetricLabel"] p {{
      font-size:.72rem; text-transform:uppercase; letter-spacing:.06em;
      color:inherit; opacity:.62; font-weight:600;
  }}
  div[data-testid="stDataFrame"] {{border-radius:12px; overflow:hidden;}}
  button[kind="primary"] {{font-weight:600; box-shadow:none;}}
  hr {{margin:2rem 0 1.4rem; border-color:{_T['line']};}}
</style>
""", unsafe_allow_html=True)

STATUS_HELP = {
    "Y": "present",
    "N": "absent",
}


def embed_html(path, height: int) -> None:
    """Preview a generated HTML file (hall plan, finder page) inside the app.

    `st.components.v1.html` is deprecated and goes away after 2026-06-01, so
    prefer `st.iframe`. The three machines this runs on (macOS/Linux/Windows)
    each have their own venv built at a different time, so the older call has
    to stay as a fallback until they're all on Streamlit >= 1.62.

    `st.iframe` reads a Path itself and sizes/scrolls a fixed-height iframe on
    its own -- there is no `scrolling` argument any more.
    """
    path = Path(path)
    if hasattr(st, "iframe"):
        st.iframe(path, height=height)
    else:
        st.components.v1.html(path.read_text(encoding="utf-8"),
                              height=height, scrolling=True)


def discover_cohorts(root: str = "out") -> dict[str, dict]:
    """Find every out/<cohort>/sheets/template.json produced by step 2."""
    found = {}
    for tpl_path in sorted(Path(root).glob("*/sheets/template.json")):
        cohort_dir = tpl_path.parent.parent
        info = json.loads(tpl_path.read_text())
        label = f"{cohort_dir.name.title()} Batch - {info.get('room', '?')}"
        found[label] = dict(
            template=str(tpl_path),
            attendance=str(cohort_dir / "attendance.xlsx"),
            room=info.get("room", "?"),
            blocks=sorted(info.get("blocks", {})),
        )
    return found


def discover_allocations(root: str = "out") -> dict[str, str]:
    """Find every out/<cohort>/allocation.csv produced by step 1."""
    return {
        f"{p.parent.name.title()} Batch  ·  {p}": str(p)
        for p in sorted(Path(root).glob("*/allocation.csv"))
    }


def list_project_rosters() -> list[str]:
    return sorted(str(p) for p in Path(".").glob("*.xls*") if "seating" not in p.stem.lower()
                  and "allotment" not in p.stem.lower())


def _window_id(group: list) -> str:
    """Identify one compose window across reruns. Streamlit rebuilds the list
    every rerun, so a positional index would drift the moment the selection
    changes underneath it - key on who the window is actually for."""
    return f"{group[0].roll}|{group[0].subject}|{len(group)}"


def gmail_handoff_ui(mails: list, key: str, group: bool = True) -> None:
    """Hand the composed messages to Gmail's compose window and let the human
    press Send.

    This is the path that works on an Institute Workspace account whose admin
    has disabled app passwords: nothing here authenticates or sends, so there
    is nothing for SMTP to be blocked on. The browser is already signed in as
    that account; all this does is pre-fill the window.
    """
    gc1, gc2 = st.columns(2)
    account = gc1.text_input(
        "Send from (the Google account composing it)",
        st.session_state.get("gmail_account", ""),
        placeholder=f"you@{C.EMAIL_DOMAIN} - or 0/1/2 for the browser's account slot",
        key=f"{key}_account",
        help="Which signed-in account Gmail should compose as. An address works; "
             "so does the number in Gmail's own /mail/u/N/ URL, which is more "
             "reliable when several accounts are signed in at once.")
    st.session_state["gmail_account"] = account
    visible_to = gc2.text_input(
        "Visible To: address", account,
        key=f"{key}_visible_to",
        help="Students go in Bcc so none of them sees another's address - but a "
             "mail with an empty To looks like spam to some filters, so put "
             "yourself here.")

    # A stale empty To: box (it defaults from the account, but Streamlit keeps
    # whatever was typed first) would produce a Bcc-only mail - fall back.
    links = mailer.compose_links(mails, account=account, group=group,
                                 to=(visible_to.strip() or account))
    personalised = len(links) > 1 and all(len(g) == 1 for _, _, g in links)

    if not group:
        per = ("student - each carries that student's own figures"
               if any(m.kind.startswith("attendance") for m in mails)
               else "block - each goes to a different set of recipients")
        st.caption(f"{len(links)} Gmail window(s), one per {per}, so these are never "
                   "merged into one shared window. Any Cc goes on every one of them.")
    elif personalised:
        st.warning(
            f"These {len(mails)} messages differ from each other (each names its own "
            "student), so Gmail needs **one window per student**. To send one window "
            "to a whole block instead, tick *Address the batch generically* above - "
            "then everyone gets an identical message and can go in a single Bcc.")
    else:
        st.caption(f"{len(links)} Gmail window(s) for {len(mails)} recipient(s) - "
                   "students are in **Bcc**. Each opens pre-filled; you press Send.")

    # Clicking through a couple of dozen windows, one student at a time, is
    # exactly where a person loses their place and either skips someone or
    # mails them twice. Each row carries its own tick box; ticked rows drop
    # out of the list, so what's left on screen is what's left to do.
    done_key = f"{key}_done"
    done: set = st.session_state.setdefault(done_key, set())
    shown = links[:60]
    if len(links) > 60:
        st.warning(f"Showing the first 60 of {len(links)} windows - narrow the "
                   "selection rather than assuming the rest were covered.")

    pending = [(i, l, u, g) for i, (l, u, g) in enumerate(shown)
               if _window_id(g) not in done]
    if done:
        st.progress(len(done) / max(len(shown), 1),
                    text=f"{len(done)} of {len(shown)} window(s) marked sent")

    for n, (i, label, url, group) in enumerate(pending, start=len(done) + 1):
        who = (group[0].name or group[0].roll) if len(group) == 1 else label
        c1, c2 = st.columns([6, 1])
        c1.link_button(f"✉️ {n}/{len(shown)} · Open in Gmail - {who}", url,
                       key=f"{key}_link_{i}")
        if c2.checkbox("sent", value=False, key=f"{key}_chk_{i}"):
            done.add(_window_id(group))
            st.rerun()

    if not pending:
        st.success(f"All {len(shown)} window(s) marked sent.")
    if done:
        with st.expander(f"Marked sent ({len(done)}) - untick to bring one back"):
            for i, (label, url, group) in enumerate(shown):
                if _window_id(group) in done:
                    d1, d2 = st.columns([6, 1])
                    d1.caption(f"✓ {label}")
                    if not d2.checkbox("sent", value=True, key=f"{key}_undo_{i}"):
                        done.discard(_window_id(group))
                        st.rerun()

    with st.expander("Doesn't open pre-filled? Copy the text instead"):
        st.caption("Some browsers truncate very long links. This is the exact same "
                   "message - paste it into a compose window by hand.")
        st.text_input("To / Bcc", ", ".join(m.to for m in mails), key=f"{key}_copy_to")
        st.text_input("Subject", mails[0].subject, key=f"{key}_copy_subj")
        st.text_area("Body", mails[0].body, height=220, key=f"{key}_copy_body")
        st.link_button("Or open your default mail app instead",
                       mailer.mailto_url(to=(mails[0].to if len(mails) == 1 else ""),
                                         subject=mails[0].subject, body=mails[0].body,
                                         bcc=(",".join(m.to for m in mails)
                                              if len(mails) > 1 else "")),
                       key=f"{key}_mailto")

    st.caption("Gmail's own Sent folder is the record of what actually went out - "
               "this tool can only note which windows you ticked off here.")
    ticked = [m for label, url, g in links if _window_id(g) in done for m in g]
    to_log = ticked or mails
    what = (f"the {len(ticked)} ticked above" if ticked
            else f"all {len(mails)}, ticked or not")
    if st.button(f"Note in the log that {what} went out", key=f"{key}_log_handoff"):
        mailer.log_handoff(to_log)
        st.success(f"Logged {len(to_log)} as `handed-off` in `{mailer.MAIL_LOG}` "
                   "(not `sent` - only you know whether you pressed Send).")


def smtp_send_ui(mails: list, key: str) -> None:
    """Direct send from the toolkit. Needs an SMTP password, so it does not
    work on an account whose admin has disabled app passwords - use the Gmail
    hand-off above there."""
    cfg = mailer.load_smtp_config()
    gaps = cfg.missing()
    if gaps:
        st.warning("SMTP is not configured - missing " + ", ".join(f"`{g}`" for g in gaps) +
                   ". Run `python3 mailer.py --configure`, or see README (*Gmail SMTP*). "
                   "If this is an Institute account with app passwords disabled, that "
                   "is expected - use **Open in Gmail** instead. Dry run below still "
                   "works and writes nothing.")
    else:
        st.caption(f"Sending as **{cfg.from_header()}** via `{cfg.host}:{cfg.port}`"
                   + (f", Bcc `{cfg.bcc}`" if cfg.bcc else "") + ".")
        if st.button("Test SMTP connection", key=f"{key}_smtp_test"):
            ok, detail = mailer.smtp_check(cfg)
            (st.success if ok else st.error)(detail)

    dry = st.checkbox("Dry run - compose and list only, send nothing",
                      value=True, key=f"{key}_dry")
    label = f"Rehearse {len(mails)} mail(s)" if dry else f"Send {len(mails)} mail(s) for real"
    if st.button(label, type="primary", disabled=bool(gaps) and not dry, key=f"{key}_send"):
        try:
            results = mailer.send_mails(mails, cfg, dry_run=dry)
        except Exception as e:                      # noqa: BLE001 - shown verbatim
            st.error(f"{type(e).__name__}: {e}")
            return
        st.dataframe(pd.DataFrame(results)[["roll", "to", "status", "detail"]],
                     width="stretch", hide_index=True)
        if dry:
            st.info("Dry run - nothing was sent and nothing was logged.")
        else:
            ok = sum(r["status"] == "sent" for r in results)
            (st.success if ok == len(results) else st.warning)(
                f"{ok}/{len(results)} sent. Logged to `{mailer.MAIL_LOG}`.")


def mail_send_ui(mails: list, key: str, *, intro: str = "", group: bool = True) -> None:
    """Preview, then either hand off to Gmail or send over SMTP.

    Shared by the move tab and the notify tab. Nothing reaches a student
    without a human action after this widget renders - a press of Send in
    Gmail, or unticking dry-run here.
    """
    if not mails:
        return
    if intro:
        st.caption(intro)

    grouped = any(getattr(m, "bcc", "") or "," in m.to for m in mails)
    rows = []
    for m in mails:
        row = {"Roll": m.roll}
        if grouped:
            # a block mail's To is forty addresses; the count is what's readable
            row["To"] = (f"{m.to.count('@')} students" if "," in m.to
                         else (m.to or "-"))
            row["Bcc"] = (f"{m.bcc.count('@')} TA(s)" if m.bcc else "-")
        else:
            row["To"] = m.to
        row["Subject"] = m.subject
        rows.append(row)
    if any(m.cc for m in mails):
        for row, m in zip(rows, mails):
            row["Cc"] = m.cc
    st.dataframe(pd.DataFrame(rows), width="stretch", hide_index=True)
    with st.expander(f"Preview the message that goes to {mails[0].roll}"):
        st.code(mails[0].preview(), language="text")

    mode = st.radio(
        "How to send",
        ["Open in Gmail - I press Send", "Send directly from here (SMTP)"],
        horizontal=True, key=f"{key}_mode",
        help="Gmail hand-off needs no setup and works on an Institute account "
             "with app passwords disabled. Direct send needs an SMTP password "
             "but can do a whole batch unattended.")
    if mode.startswith("Open in Gmail"):
        gmail_handoff_ui(mails, key, group=group)
    else:
        smtp_send_ui(mails, key)


@st.cache_data(show_spinner=False)
def master_attendance_sessions(path: str, mtime: float) -> dict[str, dict]:
    """{sheet: {held, dated, asof}} from the master Class Attendance workbook.

    `mtime` is only there to key the cache - the workbook is edited by hand
    outside this app, so a cached read has to expire when the file changes.

    `held` counts sessions that actually have marks. The workbook pre-dates
    the whole term (39 dated columns, a fraction of them marked), so counting
    dated columns would report classes that haven't happened.
    """
    out = {}
    for sh in attrep.sheet_names(path):
        try:
            _, sessions = attrep.load(path, sh)
        except (ValueError, KeyError):
            continue
        if sessions.empty:
            out[sh] = dict(held=0, dated=0, asof="")
            continue
        marked = sessions[(sessions.Present + sessions.Absent) > 0]
        asof = ""
        if len(marked):
            last = marked.Date.max()
            asof = last.strftime("%d %b %Y") if hasattr(last, "strftime") else str(last)
        out[sh] = dict(held=len(marked), dated=len(sessions), asof=asof)
    return out


def cohort_dirs(root: str = "out") -> list[Path]:
    """Real cohort folders - `out/.history` and friends are the toolkit's own
    state, not a batch of students."""
    base = Path(root)
    # `out/exams` holds saved exams, not a batch: listing it here gave the
    # Overview an "Exams" cohort with no seats, no sheets and no classes held.
    exams = Path(exam_plan.EXAM_ROOT).resolve()
    return sorted(q for q in base.glob("*")
                  if q.is_dir() and not q.name.startswith(".")
                  and q.resolve() != exams) if base.exists() else []


def _status_chips(vertical: bool = False) -> str:
    """Where the course stands, shown in the sidebar on every page: who's
    seated, how many classes are recorded, whether anyone's waiting to be told
    they moved. Every number is read fresh from the files the rest of the app
    writes, so a stale chip means a stale file, not a display bug."""
    chips = []
    for cdir in cohort_dirs():
        csv_path = cdir / "allocation.csv"
        if csv_path.exists():
            try:
                n = len(pd.read_csv(csv_path, usecols=["Roll"]))
            except (ValueError, OSError):
                continue
            chips.append(f'<span class="phl-chip">{cdir.name.title()} <b>{n}</b> seated</span>')

    att = [q for q in Path(".").glob("*Class Attendance*.xls*") if "backup" not in q.stem.lower()]
    if att:
        try:
            by_sheet = master_attendance_sessions(str(att[0]), att[0].stat().st_mtime)
            held = max((v["held"] for v in by_sheet.values()), default=0)
            asof = next((v["asof"] for v in by_sheet.values() if v["asof"]), "")
            if held:
                chips.append(f'<span class="phl-chip"><b>{held}</b> classes held'
                             + (f' · to {asof}' if asof else "") + "</span>")
        except Exception:                       # noqa: BLE001 - a chip must never break the page
            pass

    try:
        pending = len(mailer.unmailed(seating.block_changed()))
        if pending:
            chips.append(f'<span class="phl-chip warn"><b>{pending}</b> moved, not yet '
                         f'mailed</span>')
    except Exception:                           # noqa: BLE001 - same
        pass
    if not chips:
        return ""
    cls = "phl-chips vert" if vertical else "phl-chips"
    return f'<div class="{cls}">{"".join(chips)}</div>' 



# ───────────────────────── navigation ─────────────────────────
# Material Symbols rather than emoji: emoji render differently on every OS
# (and the three machines this runs on are all different), while these come
# from the same icon set Streamlit's own widgets use, so the sidebar looks
# identical everywhere.
PAGES = {
    "Seating": [
        ("dashboard", "Overview", "What's already been generated"),
        ("chair_alt", "Allocate seats", "Roster in, seats + workbook + hall plan out"),
        ("print", "Signature sheets", "Printable per-block sheets to sign"),
        ("picture_as_pdf", "Seat allotment PDF", "The student-facing block map"),
        ("travel_explore", "Find your block", "Shareable self-lookup page"),
        ("swap_horiz", "Move a student", "Transfers, block to block or batch to batch"),
        ("wallpaper", "Posters & packs", "Hall-door posters and exam invigilator packs"),
    ],
    "Exams": [
        ("school", "Exam seating", "Spaced seating across exam venues"),
        ("fact_check", "Answer script showing", "Quiz marks-entry sheet, per TA/room/slot"),
    ],
    "Attendance & mail": [
        ("person_search", "Student lookup", "Everything about one student"),
        ("insights", "Attendance summary", "Percentages and defaulters"),
        ("mail", "Email students", "Seating, exam and attendance notices"),
    ],
    "Maintenance": [
        ("stethoscope", "Health check", "Verify every invariant, undo a bad build"),
        ("tune", "Course setup", "Course details, people, rules, halls, files"),
    ],
}
PAGE_META = {name: (icon, tip) for group in PAGES.values() for icon, name, tip in group}
PAGE_ORDER = list(PAGE_META)

with st.sidebar:
    st.markdown(
        f'<div class="phl-brand"><div class="logo">{C.COURSE_CODE[:2].upper()}</div>'
        f'<div><div class="t">{C.COURSE_CODE}</div><div class="s">Control Panel</div></div></div>',
        unsafe_allow_html=True)
    st.markdown(_status_chips(vertical=True), unsafe_allow_html=True)
    if not course_settings.exists():
        st.caption("Running on the **demo course**. Set up yours under "
                   "Maintenance › Course setup.")

    if st.session_state.get("phl_page") not in PAGE_META:
        # first run: no course set up yet, so open on the setup form
        st.session_state["phl_page"] = (PAGE_ORDER[0] if course_settings.exists()
                                        else "Course setup")
    for group, items in PAGES.items():
        st.caption(group.upper())
        for icon, name, tip in items:
            selected = st.session_state["phl_page"] == name
            if st.button(name, key=f"nav_{name}", icon=f":material/{icon}:",
                         width="stretch",
                         type="primary" if selected else "tertiary"):
                st.session_state["phl_page"] = name
                st.rerun()

    st.markdown('<div class="phl-side-foot">Reads and writes <code>out/&lt;cohort&gt;/</code> '
                'and the Institute\'s Class Attendance workbook in this folder.</div>',
                unsafe_allow_html=True)

PAGE = st.session_state["phl_page"]


def page_header(title: str, subtitle: str = "") -> None:
    """One consistent page identity: the nav's own icon, the title, and a line
    saying what the page is for."""
    icon, tip = PAGE_META.get(PAGE, ("circle", ""))
    st.markdown(
        f'<div class="phl-page"><span class="ic phl-sym">{icon}</span>'
        f'<div><h2>{title}</h2><p>{subtitle or tip}</p></div></div>',
        unsafe_allow_html=True)


# Advisory only - see lockfile.py on why this can't be a real lock across a
# Dropbox-synced folder.
_others = lockfile.heartbeat()
if _others:
    st.warning("⚠️ " + lockfile.describe(_others))


def snapshot_before(paths, label: str, note: str = "") -> None:
    """Save whatever is about to be overwritten. Never raises: losing a
    snapshot is bad, but failing the build the user actually asked for because
    the snapshot failed is worse."""
    try:
        history.snapshot(paths, label, note=note)
    except Exception as e:                      # noqa: BLE001 - surfaced, not fatal
        st.caption(f"(Could not snapshot before this write: {type(e).__name__}: {e})")

# ───────────────────────── Overview ─────────────────────────
if PAGE == "Overview":
    page_header("Overview", "Everything generated so far, per cohort - and what each file is.")
    out_root = Path("out")
    cohorts_found = cohort_dirs()
    if not cohorts_found:
        st.info("Nothing in `out/` yet - start with **🪑 1 · Allocate seats** in the sidebar.")
    master_att = sorted(Path(".").glob("*Class Attendance*.xls*"))
    master_att = [q for q in master_att if "backup" not in q.stem.lower()]
    sessions_by_sheet: dict[str, dict] = {}
    if master_att:
        try:
            sessions_by_sheet = master_attendance_sessions(
                str(master_att[0]), master_att[0].stat().st_mtime)
        except Exception as e:                      # noqa: BLE001 - shown below
            st.caption(f"Couldn't read `{master_att[0].name}`: {e}")

    for cdir in cohorts_found:
        alloc_csv = cdir / "allocation.csv"
        html_plans = list(cdir.glob("*_seating_plan.html"))
        xlsx_plans = list(cdir.glob("*_seating.xlsx"))
        tpl = cdir / "sheets" / "template.json"
        att = cdir / "attendance.xlsx"
        # Sessions come from the Institute's master workbook, matched to this
        # cohort by sheet name (out/english -> the "English" sheet). The old
        # count read `out/<cohort>/attendance.xlsx` - the retired scan
        # pipeline's file, which only one cohort ever had, so English read 0
        # while nine classes had been held.
        sheet_for_cohort = next((sh for sh in sessions_by_sheet
                                 if sh.lower() == cdir.name.lower()), None)
        sess = sessions_by_sheet.get(sheet_for_cohort) if sheet_for_cohort else None
        with st.expander(f"**{cdir.name.title()}** - `{cdir}`", expanded=False):
            c1, c2, c3 = st.columns(3)
            c1.metric("Seats allocated", "yes" if alloc_csv.exists() else "no")
            c2.metric("Sheets printed", "yes" if tpl.exists() else "no")
            c3.metric("Classes held", sess["held"] if sess else "-",
                      help=(f"Sessions with marks on the **{sheet_for_cohort}** sheet of "
                            f"`{master_att[0].name}`, out of {sess['dated']} dated columns. "
                            f"Latest: {sess['asof']}." if sess else
                            "No sheet in the master Class Attendance workbook matches this "
                            "cohort's folder name."))
            if sess and sess["asof"]:
                st.caption(f"Attendance recorded up to **{sess['asof']}** "
                           f"({sess['held']} of {sess['dated']} scheduled classes).")
            if alloc_csv.exists():
                st.caption(f"Allocation: `{alloc_csv}` · {len(pd.read_csv(alloc_csv))} students")
            for hp in html_plans:
                with st.popover(f"Preview {hp.name}"):
                    embed_html(hp, height=500)
                st.download_button(f"Download {hp.name}", hp.read_bytes(), file_name=hp.name,
                                    key=f"dl_{hp}")
            for xp in xlsx_plans:
                st.download_button(f"Download {xp.name}", xp.read_bytes(), file_name=xp.name,
                                    key=f"dl_{xp}")
            if tpl.exists():
                info = json.loads(tpl.read_text())
                st.caption(f"Signature sheets: room **{info.get('room')}**, "
                           f"blocks {', '.join(sorted(info.get('blocks', {})))}")
            if att.exists():
                st.download_button("Download attendance.xlsx", att.read_bytes(),
                                    file_name=f"{cdir.name}_attendance.xlsx", key=f"dl_{att}")

    # ── one ZIP per cohort or exam, for the print run ─────────────────
    st.markdown("#### Print bundle")
    st.caption("Every printable built so far for one batch or one exam, in one ZIP, with a "
               "MANIFEST.txt of what each file is and when it was built. Nothing is rebuilt "
               "here - build on the other pages first, then bundle.")
    bundle_opts = {f"{c.name.title()} batch": ("cohort", c) for c in cohorts_found}
    bundle_opts.update({f"Exam · {m['title']}": ("exam", m) for m in exam_plan.list_exams()})
    if not bundle_opts:
        st.info("Nothing built yet.")
    else:
        pb1, pb2 = st.columns([2, 1], vertical_alignment="bottom")
        pick_b = pb1.selectbox("Bundle", list(bundle_opts), key="bundle_pick")
        kind_b, what_b = bundle_opts[pick_b]
        if kind_b == "cohort":
            files_b = bundle.cohort_files(what_b)
            zname_b = f"{C.COURSE_CODE}_{what_b.name}_printables.zip"
            title_b = f"{what_b.name.title()} batch printables"
        else:
            files_b = bundle.exam_files(what_b["dir"])
            zname_b = f"{C.COURSE_CODE}_{what_b['slug']}_printables.zip"
            title_b = f"{what_b['title']} printables"
        alloc_b = (what_b if kind_b == "cohort" else Path(what_b["dir"])) / "allocation.csv"
        if not files_b:
            pb2.caption("No printables built for this yet.")
        else:
            pb2.download_button(f"Download ZIP ({len(files_b)} files)",
                                bundle.make_zip(files_b, title_b, alloc_b), file_name=zname_b,
                                mime="application/zip", key="bundle_dl", type="primary",
                                width="stretch")
            old_b = bundle.stale(files_b, alloc_b)
            if old_b:
                st.warning(f"{len(old_b)} file(s) were built before the seats last changed - "
                           "rebuild them before printing: " + ", ".join(old_b[:5])
                           + (" …" if len(old_b) > 5 else ""))
            with st.expander("What's in it"):
                st.code(bundle.manifest(files_b, title_b, alloc_b), language=None)

# ───────────────────────── 1. Allocate seats ─────────────────────────
if PAGE == "Allocate seats":
    page_header("Allocate seats", "Roster in; seat allocation, seating workbook and interactive hall plan out.")
    have_master = Path(seating.ATTENDANCE_WORKBOOK).exists()
    modes = (["Class Attendance workbook (recommended)"] if have_master else []) + \
            ["Upload a file", "Use a file already in the project folder"]
    src_mode = st.radio("Roster source", modes, horizontal=True, key="alloc_src_mode")
    roster_bytes = None
    roster_name = None
    master_df = None
    if src_mode.startswith("Class Attendance"):
        # The department keeps this workbook current, one sheet per medium, so
        # it is right about who is in which batch even when a Section & Group
        # export is not. Its row order is also the order signature sheets and
        # the attendance register share, so seats follow it directly.
        st.caption(f"Reading `{seating.ATTENDANCE_WORKBOOK}` - one sheet per medium, "
                   "in the register's own row order.")
        try:
            master_all = seating.load_attendance_roster()
        except Exception as e:
            st.error(f"Could not read the Class Attendance workbook: {e}")
            st.stop()
        media = sorted(master_all.Medium.unique())
        pick_med = st.selectbox("Batch (sheet)", media, key="alloc_master_med")
        master_df = (master_all[master_all.Medium == pick_med]
                     .sort_values("Order").reset_index(drop=True))
        st.caption(f"{len(master_df)} students on sheet **{pick_med}**.")
    elif src_mode == "Upload a file":
        up = st.file_uploader("Roster workbook (.xlsx/.xls)", type=["xlsx", "xls"], key="alloc_upload")
        if up:
            roster_bytes, roster_name = up.getbuffer(), up.name
    else:
        choices = list_project_rosters()
        pick = st.selectbox("Roster file", ["(choose)"] + choices, key="alloc_pick")
        if pick != "(choose)":
            roster_name = pick
            roster_bytes = Path(pick).read_bytes()

    if master_df is None:
        c1, c2, c3 = st.columns(3)
        sheet_name = c1.text_input("Sheet name", "Section & Group", key="alloc_sheet")
        header_row = c2.number_input("Header row (0-based)", min_value=0, value=1, key="alloc_header")
        medium = c3.text_input("Filter by medium (optional)", "", placeholder="e.g. Hindi",
                               key="alloc_medium")
    else:
        # Sheet, header and medium are all settled by the batch picked above.
        sheet_name, header_row, medium = None, None, st.session_state["alloc_master_med"]

    c4, c5 = st.columns(2)
    room = c4.selectbox("Room", list(C.ROOMS), key="alloc_room")
    cohort = c5.text_input("Cohort label", "Hindi Batch" if medium.lower() == "hindi" else "English Batch",
                           key="alloc_cohort")
    c6, c7 = st.columns(2)
    course = c6.text_input("Course", C.COURSE, key="alloc_course")
    session = c7.text_input("Session", C.SESSION, key="alloc_session")

    default_out = f"out/{medium.strip().lower() or 'plan'}"
    out_dir = st.text_input("Output folder", default_out, key="alloc_out")

    group_available = room in C.BRANCH_GROUPS
    group_seating = st.checkbox(
        "Group students by branch (per config.BRANCH_GROUPS)",
        value=False, disabled=not group_available, key="alloc_group",
        help="Seats each block's assigned branches together instead of plain roll order. "
             "Backlog students and any branch not assigned to a block fill whatever's left, "
             "front-to-rear, same fallback plain allocation always uses."
             if group_available else "No branch grouping is defined for this room in config.BRANCH_GROUPS.",
    )

    run_alloc = st.button("Allocate seats", type="primary",
                          disabled=roster_bytes is None and master_df is None, key="alloc_run")
    if run_alloc and (roster_bytes is not None or master_df is not None):
        with tempfile.TemporaryDirectory() as tmp:
            roster_path = Path(tmp) / (roster_name or "roster.xlsx")
            if roster_bytes is not None:
                roster_path.write_bytes(roster_bytes)
            try:
                if master_df is not None:
                    df = master_df[["Roll", "Name", "Medium"]].copy()
                else:
                    df = load_roster(str(roster_path), sheet=sheet_name, header=int(header_row))
                if master_df is None and medium.strip():
                    if "Medium" not in df:
                        st.error("Roster has no Medium/Language column to filter by.")
                        st.stop()
                    df = df[df.Medium.str.lower() == medium.strip().lower()].reset_index(drop=True)
                    if df.empty:
                        st.error(f"No students with medium {medium!r}.")
                        st.stop()
                if group_seating:
                    alloc, seats = allocate_grouped(df, room, C.BRANCH_GROUPS[room])
                else:
                    alloc, seats = allocate(df, room)
            except (ValueError, KeyError) as e:
                st.error(str(e))
                st.stop()

            out = Path(out_dir)
            out.mkdir(parents=True, exist_ok=True)
            alloc.to_csv(out / "allocation.csv", index=False)
            snapshot_before([out], f"allocate-{Path(out).name}",
                            note=f"before re-allocating {room} from a roster")
            xlsx_path = reports.build_workbook(alloc, seats, str(out / f"{room}_seating.xlsx"),
                                               room=room, course=course, cohort=cohort,
                                               session=session, source=roster_name or "")
            html_path = reports.build_html(alloc, seats, str(out / f"{room}_seating_plan.html"),
                                           room=room, course=course, cohort=cohort, session=session)

            st.session_state["last_alloc"] = dict(
                alloc=alloc, room=room, course=course, cohort=cohort,
                session=session, out=str(out),
            )

        st.success(f"{len(alloc)} students seated in {room}.")
        st.dataframe(block_summary(alloc, seats, room), width="stretch", hide_index=True)

        d1, d2, d3 = st.columns(3)
        d1.download_button("Download allocation.csv", (out / "allocation.csv").read_bytes(),
                           file_name="allocation.csv")
        d2.download_button("Download seating workbook", Path(xlsx_path).read_bytes(),
                           file_name=Path(xlsx_path).name)
        d3.download_button("Download seating plan (HTML)", Path(html_path).read_bytes(),
                           file_name=Path(html_path).name)
        with st.expander("Preview seating plan"):
            embed_html(html_path, height=600)

# ───────────────────────── 2. Signature sheets ─────────────────────────
if PAGE == "Signature sheets":
    page_header("Signature sheets", "One printable PDF per block, in the attendance register's own order.")
    have_session_alloc = "last_alloc" in st.session_state
    src = st.radio(
        "Allocation source",
        (["Just-generated allocation (page 1)"] if have_session_alloc else [])
        + ["Load an existing allocation.csv"],
        horizontal=True, key="sheets_src",
    )

    alloc_df = room2 = course2 = cohort2 = session2 = out2 = None
    if src == "Just-generated allocation (page 1)":
        la = st.session_state["last_alloc"]
        alloc_df, room2 = la["alloc"], la["room"]
        course2, cohort2, session2 = la["course"], la["cohort"], la["session"]
        out2 = f"{la['out']}/sheets"
        st.caption(f"Using the allocation for **{cohort2}** in **{room2}** from `{la['out']}`.")
    else:
        found = discover_allocations()
        pick = st.selectbox("allocation.csv", ["(choose)"] + list(found), key="sheets_pick")
        if pick != "(choose)":
            alloc_path = Path(found[pick])
            alloc_df = pd.read_csv(alloc_path, dtype={"Roll": str})
            # Room and output folder are keyed *per allocation*. With one fixed
            # key each, Streamlit keeps whatever was set for the previously
            # picked cohort - which is how a Hindi allocation once got built
            # into out/english/sheets under room LHC110, wiping that cohort's
            # E and F sheets. Switching the allocation now re-derives both.
            slug = alloc_path.parent.name
            room_guess = next((q.stem[: -len("_seating")]
                               for q in alloc_path.parent.glob("*_seating.xlsx")
                               if q.stem.endswith("_seating") and q.stem[: -len("_seating")] in C.ROOMS),
                              list(C.ROOMS)[0])
            c1, c2 = st.columns(2)
            room2 = c1.selectbox("Room", list(C.ROOMS),
                                 index=list(C.ROOMS).index(room_guess),
                                 key=f"sheets_room_{slug}")
            cohort2 = c2.text_input("Cohort label", slug.title() + " Batch",
                                    key=f"sheets_cohort_{slug}")
            c3, c4 = st.columns(2)
            course2 = c3.text_input("Course", C.COURSE, key="sheets_course")
            session2 = c4.text_input("Session", C.SESSION, key="sheets_session")
            out2 = st.text_input("Output folder", str(alloc_path.parent / "sheets"),
                                 key=f"sheets_out_{slug}")

    date_label = st.text_input("Date printed on the sheet (optional)", "",
                               placeholder=date.today().strftime("%d %b %Y"), key="sheets_date")

    ta_name = st.text_input("Attendance TA name (optional, printed on the sheet)", "",
                            key="sheets_ta_manual")
    ta_candidates = sorted(str(p) for p in Path(".").glob("*TA*.xls*") if "template" not in p.stem.lower())
    with st.expander("Auto-fill from the TA duty schedule instead"):
        auto_ta = st.checkbox("Look up the on-duty TA from a schedule file", key="sheets_ta_auto")
        if auto_ta:
            slot = st.radio("Slot", ["Morning", "Afternoon"], horizontal=True, key="sheets_ta_slot")
            ta_src = st.radio("TA file source", ["Upload a file", "Use a file already in the project folder"],
                              horizontal=True, key="sheets_ta_src")
            ta_bytes = ta_fname = None
            if ta_src == "Upload a file":
                ta_up = st.file_uploader("TA duty workbook (.xlsx)", type=["xlsx", "xls"], key="sheets_ta_upload")
                if ta_up:
                    ta_bytes, ta_fname = ta_up.getbuffer(), ta_up.name
            else:
                ta_pick = st.selectbox("TA file", ["(choose)"] + ta_candidates, key="sheets_ta_pick")
                if ta_pick != "(choose)":
                    ta_fname, ta_bytes = ta_pick, Path(ta_pick).read_bytes()

            parsed_date = pd.to_datetime(date_label, errors="coerce") if date_label else None
            if not date_label:
                st.info("Enter a date above first - the schedule is per weekday.")
            elif pd.isna(parsed_date):
                st.warning(f"Couldn't parse {date_label!r} as a date; auto-fill needs a recognizable date.")
            elif ta_bytes is not None:
                with tempfile.TemporaryDirectory() as tmp:
                    ta_path = Path(tmp) / (ta_fname or "tas.xlsx")
                    ta_path.write_bytes(ta_bytes)
                    try:
                        sched = finder.ta_schedule(finder.load_tas(str(ta_path)))
                    except Exception as e:
                        st.error(f"Couldn't read the TA file: {e}")
                        sched = None
                if sched is not None:
                    looked_up = finder.ta_for_date(sched, parsed_date, slot)
                    if looked_up:
                        ta_name = looked_up
                        st.success(f"{parsed_date.strftime('%A')} {slot}: **{ta_name}**")
                    else:
                        st.warning(f"No {slot} TA scheduled for {parsed_date.strftime('%A')} in this file - "
                                  f"printing with no TA name unless you fill it in manually above.")

    # ── optional attendance column (the "next class" sheet) ────────────────
    att_map, att_asof = None, ""
    show_att = st.checkbox(
        "Show each student's attendance % up to the last class held",
        value=False, key="sheets_show_att",
        help="Adds an Attendance column to every row - the figure as of the last "
             "class that has marks in the workbook, not counting the class this "
             "sheet is for. Print these for the next class so each student sees "
             "where they stand as they sign.")
    if show_att:
        att_files_s = sorted(str(q) for q in Path(".").glob("*Class Attendance*.xls*")
                             if "backup" not in q.stem.lower())
        if not att_files_s:
            st.error("No `*Class Attendance*.xlsx` in the project folder - can't read "
                     "attendance without it.")
        else:
            sc1, sc2 = st.columns(2)
            att_file_s = sc1.selectbox("Attendance workbook", att_files_s,
                                       key="sheets_att_file")
            try:
                # one cohort's sheets are being printed, so the combined sheet
                # (both batches on one list) is not a choice here
                sheet_opts = attrep.cohort_sheets(att_file_s)
            except Exception as e:                     # noqa: BLE001 - shown as-is
                st.error(f"Could not read that workbook: {e}")
                sheet_opts = []
            # The batch sheet has to match the cohort being printed; guess from the
            # output folder's name (out/english → "English") and let it be changed.
            guess = next((x for x in sheet_opts
                          if x.lower() in str(out2).lower()), sheet_opts[0] if sheet_opts else None)
            att_sheet_s = sc2.selectbox("Batch sheet", sheet_opts,
                                        index=sheet_opts.index(guess) if guess in sheet_opts else 0,
                                        key="sheets_att_sheet") if sheet_opts else None
            if att_sheet_s:
                try:
                    att_map, att_asof = attrep.percent_map(att_file_s, att_sheet_s)
                except (ValueError, KeyError) as e:
                    st.error(str(e))
                    att_map = None
            if att_map is not None and alloc_df is not None:
                missing = [r for r in alloc_df.Roll.astype(str) if r not in att_map]
                st.caption(f"Attendance as of **{att_asof or 'the last class held'}** · "
                           f"{len(att_map)} students in sheet **{att_sheet_s}**"
                           + (f" · ⚠️ {len(missing)} seated student(s) not in it, "
                              f"they print as “-”: {', '.join(missing[:6])}"
                              f"{'…' if len(missing) > 6 else ''}" if missing else ""))
                st.caption("Note: a signature sheet is passed down the row, so every "
                           "student in a block can read their neighbours' percentages.")

    merge_all = st.checkbox("Also merge all blocks into one combined PDF", value=True,
                            disabled=not HAVE_PYPDF, key="sheets_merge")
    if not HAVE_PYPDF:
        st.caption("Install `pypdf` to enable the combined-PDF merge.")

    # Last line of defence: building into another cohort's folder deletes that
    # cohort's sheets for blocks this allocation doesn't use (`build_all`
    # clears stale block PDFs), so say so before it happens rather than after.
    mismatch = ""
    if alloc_df is not None and src == "Load an existing allocation.csv":
        try:
            if Path(out2).resolve().parent != alloc_path.parent.resolve():
                mismatch = (f"Output folder `{out2}` is not "
                            f"`{alloc_path.parent / 'sheets'}` - this allocation "
                            f"({len(alloc_df)} students, blocks "
                            f"{', '.join(sorted(alloc_df.Block.unique()))}) would overwrite "
                            f"what's in there, and delete its sheets for any block not "
                            f"in this list.")
        except (OSError, NameError):
            pass
    if mismatch:
        st.warning(mismatch)
        confirm_out = st.checkbox("I mean to write there", key="sheets_out_confirm")
    else:
        confirm_out = True

    run_sheets = st.button("Build signature sheets", type="primary",
                           disabled=alloc_df is None or not confirm_out, key="sheets_run")
    if run_sheets and alloc_df is not None:
        try:
            snapshot_before([out2], f"sheets-{Path(out2).parent.name}",
                            note=f"before rebuilding {room2} signature sheets")
            tpl_path = sheetmod.build_all(alloc_df, out2, course=course2, room=room2,
                                          session=session2, date_label=date_label,
                                          ta_name=ta_name, attendance=att_map,
                                          att_asof=att_asof)
        except ValueError as e:
            st.error(str(e))
            st.stop()

        out_dir2 = Path(out2)
        pdfs = sorted(out_dir2.glob("signature_block_*.pdf"))
        st.success(f"{len(alloc_df)} boxes across {len(pdfs)} blocks → `{out_dir2}`")

        merged_path = None
        if merge_all and HAVE_PYPDF and pdfs:
            w = PdfWriter()
            for p in pdfs:
                w.append(str(p))
            merged_name = f"ALL_signature_sheets_{cohort2.split()[0]}.pdf" if cohort2 else "ALL_signature_sheets.pdf"
            merged_path = out_dir2 / merged_name
            with open(merged_path, "wb") as f:
                w.write(f)

        cols = st.columns(len(pdfs) + (1 if merged_path else 0) + 1)
        for col, p in zip(cols, pdfs):
            col.download_button(p.name, p.read_bytes(), file_name=p.name, key=f"dl_{p}")
        if merged_path:
            cols[len(pdfs)].download_button(merged_path.name, merged_path.read_bytes(),
                                            file_name=merged_path.name, key=f"dl_{merged_path}")
        cols[-1].download_button("template.json", Path(tpl_path).read_bytes(),
                                 file_name="template.json", key=f"dl_{tpl_path}")
        st.caption("Print at 100% scale (no 'fit to page') so the seat/roll columns "
                   "line up with the hall plan.")

# ───────────────────────── 3. Seat allotment PDF ─────────────────────────
if PAGE == "Seat allotment PDF":
    page_header("Seat allotment PDF", "The plain student-facing handout: block map plus the roll list per block.")
    st.caption(
        "A plain PDF for students, not TAs: one page showing which block is where in the "
        "hall, then the full roll-number list grouped by block. One or more cohorts can be "
        "combined into a single file - the original combined both English and Hindi."
    )

    found_allocs = discover_allocations()
    if not found_allocs:
        st.info("No `allocation.csv` found yet - allocate seats on page **🪑 1 · Allocate seats** first.")
    else:
        picked = st.multiselect("Cohorts to include", list(found_allocs),
                                default=list(found_allocs), key="handout_cohorts")

        cohort_cfgs = []
        for label in picked:
            alloc_path = Path(found_allocs[label])
            cdir = alloc_path.parent
            existing_xlsx = list(cdir.glob("*_seating.xlsx"))
            guessed_room = (existing_xlsx[0].stem[:-len("_seating")]
                            if existing_xlsx and existing_xlsx[0].stem.endswith("_seating")
                            else (existing_xlsx[0].stem if existing_xlsx else list(C.ROOMS)[0]))
            guessed_slot = "Morning" if "english" in cdir.name.lower() else (
                "Afternoon" if "hindi" in cdir.name.lower() else "None")
            with st.expander(f"{label}", expanded=True):
                c1, c2 = st.columns(2)
                room_h = c1.selectbox("Room", list(C.ROOMS),
                                      index=list(C.ROOMS).index(guessed_room) if guessed_room in C.ROOMS else 0,
                                      key=f"handout_room_{label}")
                cohort_h = c2.text_input("Cohort label", f"{cdir.name.title()} batch",
                                         key=f"handout_label_{label}")
                slot_h = st.radio("Attendance TA slot for this cohort (for the weekday roster footer)",
                                  ["None", "Morning", "Afternoon"], horizontal=True,
                                  index=["None", "Morning", "Afternoon"].index(guessed_slot),
                                  key=f"handout_slot_{label}")
                cohort_cfgs.append(dict(path=alloc_path, room=room_h, cohort=cohort_h,
                                        slot=None if slot_h == "None" else slot_h))

        st.markdown("#### Attendance TA weekday roster (optional)")
        st.caption(
            "Prints each cohort's Mon/Wed/Fri Attendance-TA roster in the footer, in place of "
            "the generic \"ask a TA\" note - this PDF has no single date, so it shows the whole "
            "weekly rotation for whichever slot that cohort is set to above."
        )
        ta_candidates_h = sorted(str(p) for p in Path(".").glob("*TA*.xls*") if "template" not in p.stem.lower())
        include_tas_h = st.checkbox("Include the TA weekday roster", value=bool(ta_candidates_h),
                                    key="handout_include_tas")
        ta_bytes_h = ta_fname_h = None
        if include_tas_h:
            ta_src_h = st.radio("TA file source", ["Upload a file", "Use a file already in the project folder"],
                                horizontal=True, key="handout_ta_src")
            if ta_src_h == "Upload a file":
                ta_up_h = st.file_uploader("TA duty workbook (.xlsx)", type=["xlsx", "xls"], key="handout_ta_upload")
                if ta_up_h:
                    ta_bytes_h, ta_fname_h = ta_up_h.getbuffer(), ta_up_h.name
            else:
                ta_pick_h = st.selectbox("TA file", ["(choose)"] + ta_candidates_h, key="handout_ta_pick")
                if ta_pick_h != "(choose)":
                    ta_fname_h, ta_bytes_h = ta_pick_h, Path(ta_pick_h).read_bytes()

        c3, c4 = st.columns(2)
        course_h = c3.text_input("Course", C.COURSE, key="handout_course")
        session_h = c4.text_input("Session", C.SESSION, key="handout_session")
        roster_h = st.checkbox("Include the full roll-number roster pages", value=True,
                               key="handout_roster")
        out_name = st.text_input("Output file", f"{C.COURSE_CODE} Seating.pdf", key="handout_out")

        run_handout = st.button("Build seat allotment PDF", type="primary",
                                disabled=not cohort_cfgs, key="handout_run")
        if run_handout and cohort_cfgs:
            cohorts_data = [
                dict(alloc=pd.read_csv(cfg["path"], dtype={"Roll": str}),
                     room=cfg["room"], cohort=cfg["cohort"], slot=cfg["slot"])
                for cfg in cohort_cfgs
            ]
            tas_df_h = None
            if include_tas_h and ta_bytes_h is not None:
                with tempfile.TemporaryDirectory() as tmp:
                    ta_path_h = Path(tmp) / (ta_fname_h or "tas.xlsx")
                    ta_path_h.write_bytes(ta_bytes_h)
                    try:
                        tas_df_h = finder.load_tas(str(ta_path_h))
                    except Exception as e:
                        st.error(f"Couldn't read the TA file: {e}")
                        st.stop()

            try:
                snapshot_before([out_name], "seat-allotment-pdf")
                out_pdf = handoutmod.build(cohorts_data, out_name, course=course_h,
                                           session=session_h, roster=roster_h, tas=tas_df_h)
            except Exception as e:
                st.error(str(e))
                st.stop()

            st.success(f"Built `{out_pdf}` - {', '.join(c['cohort'] for c in cohorts_data)}.")
            st.download_button("Download seat allotment PDF", Path(out_pdf).read_bytes(),
                               file_name=Path(out_pdf).name, key="handout_dl")

# ───────────────────────── 4. Find your block ─────────────────────────
if PAGE == "Find your block":
    page_header("Find your block", "A shareable self-lookup page - roll number in, block out. Safe to publish.")
    st.caption(
        "A single static page: a student types their roll number and gets their block, "
        "highlighted on the hall plan; a TA types theirs and gets their duty assignment "
        "instead. Roll numbers are stored as hashes, never in plain text, and no names are "
        "embedded - this can be shared publicly."
    )

    found_allocs_f = discover_allocations()
    if not found_allocs_f:
        st.info("No `allocation.csv` found yet - allocate seats on page **🪑 1 · Allocate seats** first.")
    else:
        picked_f = st.multiselect("Cohorts to include", list(found_allocs_f),
                                  default=list(found_allocs_f), key="finder_cohorts")

        cohort_cfgs_f = []
        for label in picked_f:
            alloc_path = Path(found_allocs_f[label])
            cdir = alloc_path.parent
            existing_xlsx = list(cdir.glob("*_seating.xlsx"))
            guessed_room = (existing_xlsx[0].stem[:-len("_seating")]
                            if existing_xlsx and existing_xlsx[0].stem.endswith("_seating")
                            else (existing_xlsx[0].stem if existing_xlsx else list(C.ROOMS)[0]))
            with st.expander(f"{label}", expanded=True):
                fc1, fc2 = st.columns(2)
                room_f = fc1.selectbox("Room", list(C.ROOMS),
                                       index=list(C.ROOMS).index(guessed_room) if guessed_room in C.ROOMS else 0,
                                       key=f"finder_room_{label}")
                cohort_f = fc2.text_input("Cohort label", f"{cdir.name.title()} Batch",
                                          key=f"finder_label_{label}")
                cohort_cfgs_f.append(dict(path=alloc_path, room=room_f, cohort=cohort_f))

        st.markdown("#### TA duty lookups (optional)")
        ta_candidates = sorted(str(p) for p in Path(".").glob("*TA*.xls*") if "template" not in p.stem.lower())
        include_tas = st.checkbox("Include TA duty lookups too", value=bool(ta_candidates),
                                  key="finder_include_tas")
        ta_bytes = ta_name_f = None
        if include_tas:
            ta_src = st.radio("TA file source", ["Upload a file", "Use a file already in the project folder"],
                              horizontal=True, key="finder_ta_src")
            if ta_src == "Upload a file":
                ta_up = st.file_uploader("TA duty workbook (.xlsx)", type=["xlsx", "xls"], key="finder_ta_upload")
                if ta_up:
                    ta_bytes, ta_name_f = ta_up.getbuffer(), ta_up.name
            else:
                ta_pick = st.selectbox("TA file", ["(choose)"] + ta_candidates, key="finder_ta_pick")
                if ta_pick != "(choose)":
                    ta_name_f = ta_pick
                    ta_bytes = Path(ta_pick).read_bytes()

        fc3, fc4 = st.columns(2)
        course_f = fc3.text_input("Course", C.COURSE, key="finder_course")
        session_f = fc4.text_input("Session", C.SESSION, key="finder_session")
        # Attendance on the lookup page: same one-way hash as the block lookup,
        # so a student reads their own figure only by knowing their own roll
        # number, and the published file still carries no names.
        show_att_f = st.checkbox(
            "Also show each student their attendance % (up to the last class held)",
            value=False, key="finder_show_att",
            help="Cuts down 'what's my attendance?' mails. The page stays safe to publish - "
                 "roll numbers are hashed and no names are embedded.")
        att_map_f, att_asof_f = None, ""
        if show_att_f:
            att_files_f = sorted(str(q) for q in Path(".").glob("*Class Attendance*.xls*")
                                 if "backup" not in q.stem.lower())
            if not att_files_f:
                st.error("No `*Class Attendance*.xlsx` in the project folder.")
            else:
                att_file_f = st.selectbox("Attendance workbook", att_files_f, key="finder_att_file")
                merged, asofs = {}, []
                for cfg in cohort_cfgs_f:
                    sheet_f = checks.sheet_for(Path(cfg["path"]).parent.name, att_file_f)
                    if not sheet_f:
                        continue
                    try:
                        m, asof = attrep.percent_map(att_file_f, sheet_f)
                    except (ValueError, KeyError) as e:
                        st.error(f"{sheet_f}: {e}")
                        continue
                    merged.update(m)
                    if asof:
                        asofs.append(asof)
                att_map_f = merged or None
                att_asof_f = asofs[0] if asofs else ""
                if att_map_f:
                    st.caption(f"{len(att_map_f)} students' figures, as of "
                               f"**{att_asof_f or 'the last class held'}**.")

        out_name_f = st.text_input("Output file", "find-your-block.html", key="finder_out")

        run_finder = st.button("Build find-your-block page", type="primary",
                               disabled=not cohort_cfgs_f, key="finder_run")
        if run_finder and cohort_cfgs_f:
            cohorts_data_f = [
                dict(alloc=pd.read_csv(cfg["path"], dtype={"Roll": str}),
                     room=cfg["room"], cohort=cfg["cohort"], label=cfg["cohort"])
                for cfg in cohort_cfgs_f
            ]
            tas_df = None
            if include_tas and ta_bytes is not None:
                with tempfile.TemporaryDirectory() as tmp:
                    ta_path = Path(tmp) / (ta_name_f or "tas.xlsx")
                    ta_path.write_bytes(ta_bytes)
                    try:
                        tas_df = finder.load_tas(str(ta_path))
                    except Exception as e:
                        st.error(f"Couldn't read the TA file: {e}")
                        st.stop()

            try:
                snapshot_before([out_name_f], "find-your-block")
                out_path_f = finder.build(cohorts_data_f, out_name_f, course=course_f,
                                          session=session_f, tas=tas_df,
                                          attendance=att_map_f, att_asof=att_asof_f)
            except Exception as e:
                st.error(str(e))
                st.stop()

            n_students = sum(len(c["alloc"]) for c in cohorts_data_f)
            n_tas = len(tas_df) if tas_df is not None else 0
            st.success(f"Built `{out_path_f}` - {n_students} students"
                       + (f", {n_tas} TAs" if n_tas else "") + ".")
            st.download_button("Download find-your-block.html", Path(out_path_f).read_bytes(),
                               file_name=Path(out_path_f).name, key="finder_dl")
            with st.expander("Preview"):
                embed_html(out_path_f, height=700)

# ───────────────────────── 5. Move a student ─────────────────────────
if PAGE == "Move a student":
    page_header("Move a student", "Block to block, or a transfer between batches - every downstream file follows.")
    st.caption(
        "Same room, different block - or a different room entirely, a batch transfer. Just "
        "pick the batch and block; it opens straight into the first vacant seat there - "
        "including transition/wedge seats by default, the reserve pool front blocks with a "
        "taper hold back at initial allocation, since that's exactly the capacity a transfer "
        "like this needs - or tells you clearly if that block is already full. Regenerates "
        "the seating workbook, "
        "hall-plan HTML, and signature sheets for every room that changes. "
        "**Attendance already recorded under the student's old seat is not moved "
        "retroactively** - only sessions after the move will reflect the new seat."
    )

    found_allocs_m = discover_allocations()
    if not found_allocs_m:
        st.info("No `allocation.csv` found yet - allocate seats on page **🪑 1 · Allocate seats** first.")
    else:
        def _guess_room(cdir: Path) -> str:
            existing = list(cdir.glob("*_seating.xlsx"))
            guess = (existing[0].stem[:-len("_seating")]
                     if existing and existing[0].stem.endswith("_seating")
                     else (existing[0].stem if existing else list(C.ROOMS)[0]))
            return guess if guess in C.ROOMS else list(C.ROOMS)[0]

        st.markdown("#### Find the student")
        from_label = st.selectbox("From cohort", list(found_allocs_m), key="move_from_label")
        from_path = Path(found_allocs_m[from_label])
        from_alloc = pd.read_csv(from_path, dtype={"Roll": str})
        from_dir = from_path.parent
        from_room = _guess_room(from_dir)
        st.caption(f"Room: **{from_room}** (from `{from_dir}`'s own `*_seating.xlsx`).")

        has_name = "Name" in from_alloc
        options = [f"{r.Roll} - {r.Name}" if has_name else r.Roll for r in from_alloc.itertuples()]
        pick = st.selectbox("Student", ["(choose)"] + options, key="move_pick")

        if pick != "(choose)":
            roll = pick.split(" - ")[0]
            cur = from_alloc.loc[from_alloc.Roll == roll].iloc[0]
            st.caption(f"Currently: **{from_room}**, Block **{cur.Block}**, Seat **{cur.Seat}**.")

            st.markdown("#### Move to")
            mc1, mc2 = st.columns(2)
            to_label = mc1.selectbox("To cohort", list(found_allocs_m),
                                     index=list(found_allocs_m).index(from_label), key="move_to_label")
            to_path = Path(found_allocs_m[to_label])
            same_room_move = to_path == from_path
            to_alloc = from_alloc if same_room_move else pd.read_csv(to_path, dtype={"Roll": str})
            to_dir = to_path.parent
            to_room = from_room if same_room_move else _guess_room(to_dir)
            mc2.text_input("To room", to_room, disabled=True, key="move_to_room_display")

            use_wedge = st.checkbox(
                "Include transition/wedge seats (the reserve pool held back at initial "
                "allocation - front blocks with a taper often have several)",
                value=True, key="move_use_wedge")

            def _block_label(b):
                core_n = len(vacant_seats(to_alloc, to_room, b.key))
                if not use_wedge or b.extra == 0:
                    return f"{b.key} - {core_n} vacant"
                total_n = len(vacant_seats(to_alloc, to_room, b.key, include_wedge=True))
                return f"{b.key} - {core_n} vacant + {total_n - core_n} transition"

            block_opts = [_block_label(b) for b in C.ROOMS[to_room]]
            to_block_pick = st.selectbox("To block", block_opts, key="move_to_block")
            to_block = to_block_pick.split(" - ")[0]
            seat_dst = None  # always the first vacant seat in the block, in room order

            st.markdown("#### Regeneration details")
            mc3, mc4 = st.columns(2)
            course_m = mc3.text_input("Course", C.COURSE, key="move_course")
            session_m = mc4.text_input("Session", C.SESSION, key="move_session")
            mc5, mc6 = st.columns(2)
            from_cohort_m = mc5.text_input("From cohort label", from_dir.name.title() + " Batch",
                                           key="move_from_cohort_label")
            to_cohort_m = mc6.text_input("To cohort label", to_dir.name.title() + " Batch",
                                         key="move_to_cohort_label")
            date_m = st.text_input("Date printed on the sheet (optional)", "", key="move_date")
            merge_m = st.checkbox("Also merge each affected room's blocks into one combined PDF",
                                  value=True, disabled=not HAVE_PYPDF, key="move_merge")

            reorder_m = st.checkbox(
                "Keep seats in roll order within affected block(s) (contiguous from 1, no block changes)",
                value=True, key="move_reorder_m",
                help="Re-sorts the destination block so the new student slots into roll order within their branch, "
                     "and compacts the source block so no seat gaps remain. Nobody changes block.")

            run_move = st.button("Move student", type="primary", key="move_run")
            if run_move:
                try:
                    new_src, new_dst, seat_used = move_student(
                        from_alloc, from_room, roll, to_room, to_block,
                        alloc_dst=None if same_room_move else to_alloc, seat_dst=seat_dst,
                        include_wedge=use_wedge, reorder=reorder_m,
                    )
                except ValueError as e:
                    st.error(str(e))
                    st.stop()

                affected = {from_dir: (new_src, from_room, from_cohort_m)}
                if not same_room_move:
                    affected[to_dir] = (new_dst, to_room, to_cohort_m)

                snapshot_before(list(affected), f"move-{roll}",
                                note=f"{from_room} {cur.Block} → {to_room} {to_block}")

                results = []
                for cdir, (alloc_df, room, cohort_label) in affected.items():
                    cdir.mkdir(parents=True, exist_ok=True)
                    alloc_df.to_csv(cdir / "allocation.csv", index=False)
                    # Re-baseline so this move isn't later re-reported as a
                    # hand edit by seating.detect_untracked_moves.
                    history.save_state(cdir.name, cdir / "allocation.csv")
                    seats_df = seat_map(room)
                    seats_df["Occupied"] = seats_df.Seat.isin(set(alloc_df.Seat))
                    xlsx_path = reports.build_workbook(
                        alloc_df, seats_df, str(cdir / f"{room}_seating.xlsx"), room=room,
                        course=course_m, cohort=cohort_label, session=session_m, source="move_student")
                    html_path = reports.build_html(
                        alloc_df, seats_df, str(cdir / f"{room}_seating_plan.html"), room=room,
                        course=course_m, cohort=cohort_label, session=session_m)
                    sheets_dir = cdir / "sheets"
                    tpl_path = sheetmod.build_all(alloc_df, str(sheets_dir), course=course_m, room=room,
                                                  session=session_m, date_label=date_m)
                    merged_path = None
                    if merge_m and HAVE_PYPDF:
                        pdfs = sorted(Path(sheets_dir).glob("signature_block_*.pdf"))
                        if pdfs:
                            w = PdfWriter()
                            for p in pdfs:
                                w.append(str(p))
                            merged_path = Path(sheets_dir) / f"ALL_signature_sheets_{cohort_label.split()[0]}.pdf"
                            with open(merged_path, "wb") as f:
                                w.write(f)
                    results.append(dict(dir=cdir, room=room, alloc=xlsx_path, html=html_path,
                                        tpl=tpl_path, merged=merged_path, n=len(alloc_df)))

                seating.log_transition(
                    roll, name=(cur.Name if has_name else ""),
                    from_room=from_room, from_block=str(cur.Block), from_seat=str(cur.Seat),
                    to_room=to_room, to_block=to_block, to_seat=seat_used,
                    from_cohort=from_dir.name, to_cohort=to_dir.name)

                st.session_state["move_last"] = dict(
                    roll=roll, name=(cur.Name if has_name else ""),
                    old_block=str(cur.Block), new_block=to_block,
                    old_room=from_room, room=to_room,
                    # the cohort moved into, not the row's own Medium flag, which
                    # tracks language of instruction and can disagree with the room
                    medium=to_dir.name.title(),
                    course=course_m, session=session_m, cohort=to_cohort_m)

                st.success(f"Moved **{roll}** - {from_room} Block {cur.Block} Seat {cur.Seat} "
                          f"→ {to_room} Block {to_block} Seat {seat_used}.")
                for r in results:
                    st.caption(f"Regenerated `{r['dir']}` - {r['n']} students, room {r['room']}.")
                    dcols = st.columns(3 if r["merged"] else 2)
                    dcols[0].download_button(f"{Path(r['alloc']).name}", Path(r['alloc']).read_bytes(),
                                             file_name=Path(r['alloc']).name, key=f"move_dl_alloc_{r['dir']}")
                    dcols[1].download_button(f"{Path(r['html']).name}", Path(r['html']).read_bytes(),
                                             file_name=Path(r['html']).name, key=f"move_dl_html_{r['dir']}")
                    if r["merged"]:
                        dcols[2].download_button(f"{r['merged'].name}", r["merged"].read_bytes(),
                                                 file_name=r["merged"].name, key=f"move_dl_merged_{r['dir']}")
                st.caption("Signature sheets regenerated too - see the room's `sheets/` folder or page "
                          "**🖨️ 2 · Signature sheets**. Re-run pages **3** and **4** if you also want the seat "
                          "allotment PDF and find-your-block page refreshed with this change.")

        # ── tell the student ───────────────────────────────────────────────
        last = st.session_state.get("move_last")
        if last:
            st.divider()
            st.markdown("#### 📧 Tell the student")
            st.caption(
                f"Mail goes to **{mailer.student_email(last['roll'])}** - the address is "
                "derived from the enrolment number, not read from any list. **The message "
                "names the block only, never a seat number**: classroom seating is "
                "block-wise, so a seat id would be telling them something untrue."
            )
            ec1, ec2 = st.columns(2)
            eff = ec1.text_input("Effective from (optional)", "the next class",
                                 key="move_mail_effective")
            contact = ec2.text_input("Contact line in the signature (optional)", mailer.DEFAULT_CONTACT,
                                     key="move_mail_contact")
            try:
                mails = [mailer.build_transition_mail(
                    last["roll"], last["new_block"], name=last["name"],
                    old_block=last["old_block"], room=last["room"], old_room=last["old_room"],
                    course=last["course"], session=last["session"], medium=last["medium"],
                    effective=eff, contact=contact)]
            except ValueError as e:
                st.error(str(e))
                mails = []
            mail_send_ui(mails, key="move_mail")
            if st.button("Dismiss - don't mail this move", key="move_mail_dismiss"):
                del st.session_state["move_last"]
                st.rerun()

        # ── reorder / compact blocks ───────────────────────────────────────
        st.divider()
        st.markdown("### 🔀 Reorder / compact seats within blocks")
        st.caption(
            "Re-sorts students in a block into roll order without changing blocks, and compacts "
            "seat numbers so they run contiguous from 1 with vacant seats at the end. "
            "Downstream files (seating workbook, hall-plan HTML, signature sheets) are regenerated automatically."
        )

        rcol1, rcol2, rcol3 = st.columns(3)
        rc_label = rcol1.selectbox("Cohort", list(found_allocs_m), key="reorder_cohort_label")
        rc_path = Path(found_allocs_m[rc_label])
        rc_alloc = pd.read_csv(rc_path, dtype={"Roll": str})
        rc_dir = rc_path.parent
        rc_room = _guess_room(rc_dir)

        rc_block_opts = ["All blocks"] + [b.key for b in C.ROOMS[rc_room]]
        rc_block_pick = rcol2.selectbox("Block", rc_block_opts, key="reorder_block_pick")

        # Register order first: it is what the signature sheets are meant to
        # follow, and what the health check verifies. The UI used to offer only
        # the two roll-number orders, so the default reorder put a block out of
        # register order and the health check flagged the dashboard's own work.
        rc_mode_opts = [
            "Attendance register order (matches the signature sheets)",
            "Roll order within branches (preserves branch grouping)",
            "Pure roll number order (A-Z)",
        ]
        rc_mode_pick = rcol3.selectbox("Ordering mode", rc_mode_opts, key="reorder_mode_pick")
        rc_mode = ("register" if "register" in rc_mode_pick
                   else "branch_roll" if "branches" in rc_mode_pick else "roll")
        rc_reg_order = None
        if rc_mode == "register":
            try:
                rc_sheet = checks.sheet_for(rc_dir.name)
                rc_reg_order = list(seating.load_attendance_roster(
                    seating.ATTENDANCE_WORKBOOK, rc_sheet).Roll) if rc_sheet else None
            except Exception as e:                  # noqa: BLE001 - fall back, say so
                st.caption(f"Couldn't read the register ({e}); using every batch sheet.")

        run_reorder = st.button("Reorder seats", type="secondary", key="reorder_run")
        if run_reorder:
            target_blocks = [b.key for b in C.ROOMS[rc_room]] if rc_block_pick == "All blocks" else [rc_block_pick]
            new_alloc = rc_alloc.copy()
            for tb in target_blocks:
                new_alloc = seating.reorder_block_in_roll_order(new_alloc, rc_room, tb, mode=rc_mode,
                                                                reg_order=rc_reg_order)

            # Compare seats
            seat_diff = []
            m_old = rc_alloc.set_index("Roll")
            for r in new_alloc.itertuples():
                old_s = str(m_old.at[r.Roll, "Seat"]) if r.Roll in m_old.index else ""
                if old_s != str(r.Seat):
                    seat_diff.append({
                        "Roll": r.Roll,
                        "Name": getattr(r, "Name", ""),
                        "Block": r.Block,
                        "Old Seat": old_s,
                        "New Seat": r.Seat,
                    })

            snapshot_before([rc_dir], f"reorder-{rc_block_pick}")
            new_alloc.to_csv(rc_path, index=False)
            history.save_state(rc_dir.name, rc_path)

            seats_df = seat_map(rc_room)
            seats_df["Occupied"] = seats_df.Seat.isin(set(new_alloc.Seat))
            cohort_label = rc_dir.name.title() + " Batch"
            xlsx_path = reports.build_workbook(
                new_alloc, seats_df, str(rc_dir / f"{rc_room}_seating.xlsx"), room=rc_room,
                course=C.COURSE, cohort=cohort_label, session=C.SESSION, source="reorder")
            html_path = reports.build_html(
                new_alloc, seats_df, str(rc_dir / f"{rc_room}_seating_plan.html"), room=rc_room,
                course=C.COURSE, cohort=cohort_label, session=C.SESSION)
            sheets_dir = rc_dir / "sheets"
            sheetmod.build_all(new_alloc, str(sheets_dir), course=C.COURSE, room=rc_room,
                               session=C.SESSION)
            if HAVE_PYPDF:
                pdfs = sorted(Path(sheets_dir).glob("signature_block_*.pdf"))
                if pdfs:
                    w = PdfWriter()
                    for p in pdfs:
                        w.append(str(p))
                    with open(Path(sheets_dir) / f"ALL_signature_sheets_{cohort_label.split()[0]}.pdf", "wb") as f:
                        w.write(f)

            st.success(f"Reordered **{rc_block_pick}** in **{rc_room}** ({rc_label}). {len(seat_diff)} student(s) updated seats.")
            if seat_diff:
                st.dataframe(pd.DataFrame(seat_diff), width="stretch")
            else:
                st.info("All seats were already in optimal roll order contiguous from 1.")

# ───────────────────────── Exam seating (multi-room) ─────────────────────────
if PAGE == "Exam seating":
    page_header("Exam seating", "A separate model: one exam can span several rooms, seats spaced out.")
    st.caption(
        "Separate from the classroom rooms above - an exam venue is parsed straight from "
        f"`{exam_rooms.SOURCE_XLSX}` and can combine blocks from more than one physical room "
        "into a single seating plan, the way the Institute's own multi-room sheet does."
    )

    try:
        exam_room_map = exam_rooms.load_exam_rooms()
    except FileNotFoundError:
        exam_room_map = {}
        st.error(f"`{exam_rooms.SOURCE_XLSX}` not found in the project folder.")

    sc1, sc2 = st.columns(2)
    spacing_labels = {
        "Gap beside and behind": "alternate",
        "Gap beside only (may sit behind)": "side",
        "Fill every seat": "every",
    }
    spacing = sc1.radio(
        "Seat spacing", list(spacing_labels), key="exam_spacing",
        captions=["Alternate seats in alternate rows - no neighbour in any direction.",
                  "Alternate seats, every row used - stops the sideways glance at "
                  "double the capacity.",
                  "Every chair. For a hall used only to hold people."])
    spacing_mode = spacing_labels[spacing]
    spaced = spacing_mode != "every"
    order_mode = sc2.radio(
        "Seating order", ["Mix branches (recommended)", "Roll-number order"],
        key="exam_order",
        help="A roster in roll order seats CS beside CS for forty seats - the people most "
             "likely to have prepared together end up side by side. Mixing deals students "
             "round-robin from per-branch piles, so neighbours are rarely from the same "
             "branch. It is deterministic: the same roster always produces the same plan.")
    mix_branches = order_mode.startswith("Mix")

    for label, room in exam_room_map.items():
        if not room.supported:
            with st.expander(f"⚠️ {label} - not usable"):
                st.caption(room.note)

    supported_labels = [l for l, r in exam_room_map.items() if r.supported]
    st.markdown("#### Rooms, in fill order")
    st.caption("First room picked fills first; students overflow into the next room only once "
               "one fills up. Reorder by clearing and re-picking in the order you want.")
    room_order = st.multiselect("Rooms to use, in the order they should fill", supported_labels,
                                key="exam_room_order")

    selected: list[tuple[str, str, str]] = []          # (room_label, block_key, seat)
    overflow: list[tuple[str, str, str]] = []          # transition seats, last resort only
    venue_parts: list[str] = []
    for label in room_order:
        room = exam_room_map[label]
        row_col = exam_rooms.is_row_col_room(label)
        unit, units = ("row", "rows") if row_col else ("block", "blocks")
        keys_sorted = sorted(room.blocks)
        pretty = {k: exam_rooms.group_label(label, k) for k in keys_sorted}
        with st.expander(
            f"{label} - {room.capacity(spacing_mode)} seats available "
            f"({'spaced' if spaced else 'dense'}) across {units} "
            f"{', '.join(pretty[k].split()[-1] for k in keys_sorted)}",
            expanded=True,
        ):
            cap_line = ", ".join(f"{pretty[k]}: {len(room.blocks[k].seats(spacing_mode))}"
                                 for k in keys_sorted)
            st.caption(f"Per {unit}: {cap_line}")
            picked = st.multiselect(f"{units.capitalize()} to use, in fill order", keys_sorted,
                                    default=keys_sorted, key=f"exam_blocks_{label}",
                                    format_func=lambda k, _p=pretty: _p[k])
            if picked:
                venue_parts.append(f"{label} ({exam_rooms.venue_note(label, picked)})")
                for bk in picked:
                    seats = room.blocks[bk].seats(spacing_mode)
                    selected.extend((label, bk, s) for s in seats)
                    overflow.extend((label, bk, s)
                                    for s in room.blocks[bk].overflow_seats(spacing_mode))

    total_capacity = len(selected)
    mc1, mc2 = st.columns(2)
    mc1.metric("Total seats selected", total_capacity)
    if overflow:
        mc2.metric("Transition seats in reserve", len(overflow),
                   help=f"Held back, and used only if at most {exam_rooms.OVERFLOW_LIMIT} "
                        "students would otherwise be left over. Opening another room - and "
                        "another invigilator - for one or two people is the worse trade.")

    st.markdown("#### Roster")
    esrc = st.radio("Roster source",
                    ["Roll list, in its own order (+ anyone missing, first)",
                     "Class Attendance workbook (both batches)", "Upload a file",
                     "Use a file already in the project folder"],
                    horizontal=True, key="exam_src_mode",
                    help="The roll list is the order the department reads students in, so "
                         "seating that follows it can be checked against their own "
                         "paperwork - but it is a snapshot and currently misses a late "
                         "admission. That student is put at the *front* rather than "
                         "dropped or buried in the last row.")
    eroster_bytes = eroster_name = None
    eroster_df = None                      # set directly when reading the workbook
    if esrc.startswith("Roll list"):
        rl_choices = [c for c in list_project_rosters()
                      if "rolllist" in c.lower().replace("_", "").replace("-", "")] \
                     or list_project_rosters()
        rl_pick = st.selectbox("Roll list", rl_choices, key="exam_rolllist_pick")
        try:
            eroster_df = exam_plan.exam_roster(rl_pick)
            eroster_name = rl_pick
            extra = len(eroster_df) - len(seating.load_roster(rl_pick))
            st.caption(f"{len(eroster_df)} students in roll-list order"
                       + (f" · **{extra} not in the export**, placed first: "
                          + ", ".join(eroster_df.head(extra).Roll) if extra else "")
                       + ".")
        except Exception as e:                        # noqa: BLE001
            st.error(f"Couldn't build the roster: {e}")
    elif esrc.startswith("Class Attendance"):
        att_for_exam = sorted(str(q) for q in Path(".").glob("*Class Attendance*.xls*")
                              if "backup" not in q.stem.lower())
        if not att_for_exam:
            st.error("No `*Class Attendance*.xlsx` in the project folder.")
        else:
            att_pick_e = st.selectbox("Workbook", att_for_exam, key="exam_att_pick")
            try:
                eroster_df = seating.load_attendance_roster(att_pick_e)[["Roll", "Name", "Medium"]]
                eroster_name = att_pick_e
                counts = eroster_df.Medium.value_counts().to_dict()
                st.caption(f"{len(eroster_df)} students - "
                           + ", ".join(f"{k} {v}" for k, v in counts.items())
                           + ". Both batches sit every scheduled exam together.")
            except Exception as e:                    # noqa: BLE001
                st.error(f"Couldn't read the workbook: {e}")
    elif esrc == "Upload a file":
        eup = st.file_uploader("Roster workbook (.xlsx/.xls)", type=["xlsx", "xls"], key="exam_upload")
        if eup:
            eroster_bytes, eroster_name = eup.getbuffer(), eup.name
    else:
        echoices = list_project_rosters()
        # Default to the plain Institute roll-list export if it's there: that is
        # the file an exam is almost always allocated from, and hunting for it
        # in a list of a dozen workbooks is the first friction of every exam.
        eopts = ["(choose)"] + echoices
        edefault = next((i for i, c in enumerate(eopts)
                         if "rolllist" in c.lower().replace("_", "").replace("-", "")), 0)
        epick = st.selectbox("Roster file", eopts, index=edefault, key="exam_pick")
        if epick != "(choose)":
            eroster_name = epick
            eroster_bytes = Path(epick).read_bytes()

    ec1, ec2, ec3 = st.columns(3)
    esheet = ec1.text_input("Sheet name", "Section & Group", key="exam_sheet")
    eheader = ec2.number_input("Header row (0-based)", min_value=0, value=1, key="exam_header")
    emedium = ec3.text_input("Filter by medium (optional)", "", key="exam_medium")

    st.markdown("#### Exam details")
    # The term's two scheduled exams are prefilled rather than retyped: the
    # date and time end up on the roll list, the seat map, the door poster and
    # every student mail, and one mistyped time propagates to all of them.
    preset_labels = {f"{e['title']} - {exam_plan.when_label(e)}": e for e in exam_plan.SCHEDULE}
    preset_pick = st.selectbox("Exam", ["Something else…"] + list(preset_labels),
                               key="exam_preset",
                               help="Both scheduled exams are sat by both batches together.")
    preset = preset_labels.get(preset_pick)
    if preset:
        st.caption(f"**{preset['title']}** · {preset['batches']} · all "
                   f"{sum(len(pd.read_csv(c / 'allocation.csv')) for c in cohort_dirs() if (c / 'allocation.csv').exists())} "
                   "students seated together - pick enough rooms below for that many.")

    ed1, ed2 = st.columns(2)
    exam_title = ed1.text_input("Exam title",
                                preset["title"] if preset else "Quiz-1, Semester-1, Academic Year 2025-26",
                                key=f"exam_title_{preset['key'] if preset else 'custom'}")
    course_e = ed2.text_input("Course", C.COURSE, key="exam_course")
    ed3, ed4 = st.columns(2)
    datetime_label = ed3.text_input(
        "Date & time",
        value=exam_plan.when_label(preset) if preset else "",
        placeholder="November 10, 2025 (6:30 PM to 7:30 PM)",
        key=f"exam_datetime_{preset['key'] if preset else 'custom'}")
    ta_name = ed4.text_input("TA", "", key="exam_ta")
    out_xlsx_name = st.text_input("Output file", "Exam Seating.xlsx", key="exam_out")

    have_roster_e = eroster_bytes is not None or eroster_df is not None
    run_exam = st.button("Allocate exam seats", type="primary",
                         disabled=not (have_roster_e and selected), key="exam_run")
    if not selected:
        st.info("Pick at least one block above to enable allocation.")

    if run_exam and have_roster_e:
        with tempfile.TemporaryDirectory() as tmp:
            if eroster_df is not None:
                edf = eroster_df.copy()
            else:
                eroster_path = Path(tmp) / (eroster_name or "roster.xlsx")
                eroster_path.write_bytes(eroster_bytes)
            try:
                if eroster_df is None:
                    edf = load_roster(str(eroster_path), sheet=esheet, header=int(eheader))
                if emedium.strip():
                    if "Medium" not in edf:
                        st.error("Roster has no Medium/Language column to filter by.")
                        st.stop()
                    edf = edf[edf.Medium.str.lower() == emedium.strip().lower()].reset_index(drop=True)
                    if edf.empty:
                        st.error(f"No students with medium {emedium!r}.")
                        st.stop()
            except (ValueError, KeyError) as e:
                st.error(str(e))
                st.stop()

            short = len(edf) - total_capacity
            used_overflow = []
            if 0 < short <= exam_rooms.OVERFLOW_LIMIT and len(overflow) >= short:
                used_overflow = overflow[:short]
                selected = selected + used_overflow
                total_capacity = len(selected)

            if len(edf) > total_capacity:
                st.error(f"Only {total_capacity} seats selected for {len(edf)} students - "
                         f"pick more blocks or switch to 'Fill every seat'.")
                st.stop()

            if mix_branches:
                ranked = {r: i for i, r in enumerate(
                    exam_plan.interleave_by_branch(list(edf.Roll.astype(str))))}
                edf = (edf.assign(_k=[ranked.get(str(r), 10 ** 6) for r in edf.Roll])
                          .sort_values("_k").drop(columns="_k").reset_index(drop=True))

            take = selected[: len(edf)]
            alloc_e = edf.copy()
            alloc_e["Room"] = [r for r, _, _ in take]
            alloc_e["Block"] = [b for _, b, _ in take]
            alloc_e["Seat"] = [s for _, _, s in take]

            venue_label = "; ".join(venue_parts)
            out_path = exam_rooms.build_exam_roll_list(
                alloc_e, out_xlsx_name, exam_title=exam_title, course=course_e,
                datetime_label=datetime_label, venue_label=venue_label, ta=ta_name,
                room_map=exam_room_map,
            )

        # kept for the 📧 Email students page, so an exam mail-out doesn't have to
        # re-upload and re-allocate the same roster just to name the same seats.
        st.session_state["exam_last"] = dict(
            alloc=alloc_e.copy(), exam=exam_title, course=course_e, when=datetime_label,
            venue=venue_label, spacing=spacing_mode)

        st.success(f"{len(alloc_e)} students seated across {len(venue_parts)} venue selection(s), "
                   f"{total_capacity - len(alloc_e)} seats left unused.")
        if used_overflow:
            where = ", ".join(f"{r} block {b} seat {s}" for r, b, s in used_overflow)
            st.info(f"**{len(used_overflow)} transition seat(s) used** rather than opening "
                    f"another room for so few: {where}. They are marked on the seat map.")
        st.dataframe(alloc_e[["Roll", "Name", "Room", "Block", "Seat"]], width="stretch", hide_index=True)
        st.caption(f"Exam Venue : {venue_label}")
        st.download_button("Download exam roll list (roll list + physical layout)",
                           Path(out_path).read_bytes(), file_name=Path(out_path).name, key="exam_dl")

        # An exam is worth keeping: saved here it survives the session, and the
        # seat map, the student lookup, the invigilator pack and the mail-out
        # all read it back instead of asking for the roster again.
        exam_dir = exam_plan.EXAM_ROOT / exam_plan.slugify(exam_title)
        snapshot_before([exam_dir], f"exam-{exam_plan.slugify(exam_title)}")
        built = exam_plan.build_all(
            alloc_e, exam_room_map,
            meta=dict(title=exam_title, when=datetime_label, venue=venue_label,
                      course=course_e, spaced=spacing_mode))
        st.session_state["exam_last"]["slug"] = exam_plan.slugify(exam_title)
        st.success(f"Saved to `{built['dir']}` - allocation, seat map and student lookup page.")

        pc1, pc2 = st.columns(2)
        pc1.download_button("Seat map (HTML - invigilators / notice board)",
                            Path(built["seat_map"]).read_bytes(),
                            file_name=f"{exam_plan.slugify(exam_title)}_seat_map.html",
                            key="exam_dl_map")
        pc2.download_button("Find-my-seat page (HTML - safe to publish)",
                            Path(built["finder"]).read_bytes(),
                            file_name=f"{exam_plan.slugify(exam_title)}_find-my-seat.html",
                            key="exam_dl_find")
        st.caption("The seat map lists roll numbers in position and is for staff. The "
                   "find-my-seat page holds only a one-way hash of each roll number and no "
                   "names - that is the one to put behind a QR code or a link.")
        with st.expander("Preview the seat map"):
            embed_html(built["seat_map"], height=620)
        with st.expander("Preview the student lookup page"):
            embed_html(built["finder"], height=620)

    # ── invigilators and the printable pack ────────────────────────────
    st.divider()
    st.markdown("#### Invigilators & printable pack")
    last_x = st.session_state.get("exam_last")
    if not last_x:
        st.caption("Allocate an exam above (or load a saved one below) to assign TAs "
                   "and build the PDF.")
    else:
        alloc_t = last_x["alloc"]
        tc1, tc2, tc3 = st.columns(3)
        n_front = tc1.number_input("TAs per front block", 1, 5, ta_duty.SECTION_TAS["front"],
                                   key="exam_ta_front")
        n_rear = tc2.number_input("TAs per long back block", 1, 5, ta_duty.SECTION_TAS["rear"],
                                  key="exam_ta_rear",
                                  help="E/F/G in LHC 110 run fourteen rows deep. Two is the "
                                       "default; three is reasonable for a full hall if the "
                                       "duty sheet has the people.")
        n_room = tc3.number_input("TAs per smaller room", 1, 5, ta_duty.PER_ROOM_TAS,
                                  key="exam_ta_room")
        reqs_t = ta_duty.requirements(alloc_t,
                                      section_tas={"front": int(n_front), "rear": int(n_rear)},
                                      per_room=int(n_room))
        ta_files = sorted(str(q) for q in Path(".").glob("*TA_Duty*.xls*")
                          if "template" not in q.stem.lower())
        if not ta_files:
            st.warning("No TA duty file in the project folder - the pack will print with "
                       "blank invigilator lines.")
            tas_t = pd.DataFrame(columns=["Roll", "Name", "Duty", "Day", "Notes"])
        else:
            ta_file_t = st.selectbox("TA duty file", ta_files, key="exam_ta_file")
            try:
                tas_t = finder.load_tas(ta_file_t)
            except Exception as e:                      # noqa: BLE001
                st.error(f"Couldn't read the TA file: {e}")
                tas_t = pd.DataFrame(columns=["Roll", "Name", "Duty", "Day", "Notes"])

        st.caption(ta_duty.summary(reqs_t, tas_t,
                                   {"Question paper printing": int(
                                       st.session_state.get("exam_ta_print",
                                       ta_duty.SUPPORT_ROLES["Question paper printing"]))}))

        mode_t = st.radio("Assignment", ["Automatic", "Manual"], horizontal=True,
                          key="exam_ta_mode",
                          help="Automatic fills posts in room order from the duty sheet, "
                               "Evaluation TAs first (exam duty is theirs) and heads last. "
                               "Manual starts from that and lets you change any post.")
        # Printing is chosen first, so whoever takes it can be kept out of the
        # block pool - a TA at the printer cannot also be standing in a block.
        st.markdown("###### Question paper printing")
        pc1, pc2 = st.columns([1, 3])
        n_print = pc1.number_input("How many", 0, 5,
                                   ta_duty.SUPPORT_ROLES["Question paper printing"],
                                   key="exam_ta_print")
        all_names = ([f"{r.Name}" for r in tas_t.itertuples()] if len(tas_t) else [])
        head_names_all = set(tas_t[tas_t.Duty.isin(ta_duty.HEAD_DUTIES)].Name) if len(tas_t) else set()
        auto_names = [n for n in ta_duty.assign_support(tas_t, None,
                                                        {"Question paper printing": int(n_print)}).TA
                      if n] if (len(tas_t) and n_print) else []
        saved_print = st.session_state.get("exam_print_names")
        printing_names = pc2.multiselect(
            "Who prints the papers", all_names,
            default=[n for n in (saved_print or auto_names) if n in all_names],
            key="exam_ta_print_who",
            help="Heads are not picked automatically - they supervise - but naming one "
                 "here is an instruction and is honoured. Whoever is named is kept out "
                 "of the block posting, as long as the rest still covers every post.")
        st.session_state["exam_print_names"] = printing_names
        named_heads = sorted(set(printing_names) & head_names_all)
        if named_heads:
            st.caption("· " + ", ".join(named_heads)
                       + (" is" if len(named_heads) == 1 else " are")
                       + " a head - posted here because you named them, not by the "
                         "automatic pick.")

        auto_t = ta_duty.auto_assign(reqs_t, tas_t, exclude=printing_names) if len(tas_t) \
            else pd.DataFrame(columns=["Room", "Block", "Section", "Students", "TA",
                                       "Roll", "Duty"])

        if mode_t == "Automatic":
            assign_t = auto_t
            st.dataframe(assign_t[["Room", "Block", "Students", "TA", "Duty"]],
                         width="stretch", hide_index=True)
        else:
            names_t = list(tas_t.Name) if len(tas_t) else []
            picked_rows = []
            for r in reqs_t.itertuples():
                default = [n for n in
                           auto_t[(auto_t.Room == r.Room) & (auto_t.Block == r.Block)].TA
                           if n]
                label = (f"{r.Room} · Block {r.Block}" if r.Block != ta_duty.WHOLE_ROOM
                         else f"{r.Room} · whole room")
                chosen = st.multiselect(f"{label} - {r.Students} students, {r.TAs} needed",
                                        names_t, default=default,
                                        key=f"exam_ta_{r.Room}_{r.Block}")
                if len(chosen) != r.TAs:
                    st.caption(f"⚠️ {len(chosen)} of {r.TAs} posts filled here.")
                for n in chosen:
                    row = tas_t[tas_t.Name == n].iloc[0]
                    picked_rows.append({"Room": r.Room, "Block": r.Block,
                                        "Section": r.Section, "Students": r.Students,
                                        "TA": n, "Roll": row.Roll, "Duty": row.Duty})
            assign_t = pd.DataFrame(picked_rows, columns=auto_t.columns)
            twice = [n for n, c in assign_t.TA.value_counts().items() if c > 1]
            if twice:
                st.warning("Posted to more than one place at once: " + ", ".join(twice)
                           + " - a TA in two rooms is a gap that looks filled.")

        # ── off-block duties ───────────────────────────────────────────
        roles_t = {"Question paper printing": int(n_print)}
        support_t = (ta_duty.assign_support(tas_t, assign_t, roles_t,
                                            names=printing_names or None)
                     if len(tas_t) and n_print else
                     pd.DataFrame(columns=["Role", "TA", "Roll", "Duty", "AlsoInvigilating"]))
        if len(support_t):
            st.dataframe(support_t[["Role", "TA", "Duty", "AlsoInvigilating"]],
                         width="stretch", hide_index=True)
            doubled = [r.TA for r in support_t.itertuples() if r.AlsoInvigilating]
            if doubled:
                st.caption("⚠️ " + ", ".join(doubled) + " also holds a block post - fine if "
                           "printing is finished before the hall opens, otherwise swap them out.")

        # ── heads ──────────────────────────────────────────────────────
        heads_t = ta_duty.head_duties(tas_t, last_x.get("exam", ""))
        if len(heads_t):
            head_names = list(dict.fromkeys(heads_t.TA))
            lead_now = heads_t[heads_t.InCharge].TA.iloc[0] if heads_t.InCharge.any() else head_names[0]
            lead_pick = st.selectbox(
                "Overall in charge", head_names, index=head_names.index(lead_now),
                key="exam_head_lead",
                help="Taken from the duty sheet's own wording - a quiz is the Attendance "
                     "Head's (\u201c… and Quiz Management\u201d), a Minor or Major the "
                     "Evaluation Head's. Change it here for one exam.")
            heads_t = heads_t.assign(
                InCharge=[r.TA == lead_pick and i == list(heads_t.TA).index(r.TA)
                          for i, r in enumerate(heads_t.itertuples())])
            st.dataframe(heads_t[["TA", "Duty", "Responsibility"]],
                         width="stretch", hide_index=True)
        else:
            st.caption("No Evaluation/Attendance Head on the duty sheet - the pack will "
                       "print no management section.")

        gap = ta_duty.shortfall(assign_t) if len(assign_t) else int(reqs_t.TAs.sum())
        if gap:
            st.warning(f"{gap} post(s) still unfilled - the pack prints those as "
                       "“- not assigned -”.")

        pack_paper_opt = st.radio("Paper size", ["Legal size (14×8.5 in)", "A4"],
                                  horizontal=True, key="exam_pack_paper_choice",
                                  help="Legal size (14×8.5 in) landscape gives wide cells and spacious signature sheets.")
        if st.button("Build printable exam pack (PDF)", type="primary", key="exam_pack_run"):
            slug_t = last_x.get("slug") or exam_plan.slugify(last_x.get("exam", "exam"))
            outdir_t = exam_plan.EXAM_ROOT / slug_t
            outdir_t.mkdir(parents=True, exist_ok=True)
            pdf_t = outdir_t / f"{slug_t}_exam_pack.pdf"
            snapshot_before([pdf_t, outdir_t / "ta_assignment.csv"], f"pack-{slug_t}")
            if len(assign_t):
                assign_t.to_csv(outdir_t / "ta_assignment.csv", index=False)
            if len(support_t):
                support_t.to_csv(outdir_t / "ta_other_duties.csv", index=False)
            if len(heads_t):
                heads_t.to_csv(outdir_t / "ta_head_duties.csv", index=False)
            try:
                posters.exam_pack(
                    alloc_t, str(pdf_t), room_map=exam_room_map,
                    exam=last_x.get("exam", "Exam"),
                    course=last_x.get("course", C.COURSE),
                    when=last_x.get("when", ""), venue=last_x.get("venue", ""),
                    spacing=exam_rooms.SPACING_MODES.get(
                        str(last_x.get("spacing", "")), ""),
                    duties=ta_duty.by_block(assign_t) if len(assign_t) else {},
                    support=support_t if len(support_t) else None,
                    heads=heads_t if len(heads_t) else None,
                    landscape=True,
                    paper="legal" if pack_paper_opt.startswith("Legal") else "a4")
            except Exception as e:                      # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            st.success(f"Built `{pdf_t}` - cover with the TA posting, a seat grid per block, "
                       "then a signature list per room.")
            st.download_button("Download exam pack (PDF)", Path(pdf_t).read_bytes(),
                               file_name=pdf_t.name, key="exam_pack_dl")

    # ── exams already on disk ──────────────────────────────────────────
    st.divider()
    st.markdown("#### Saved exams")
    saved_exams = exam_plan.list_exams()
    if not saved_exams:
        st.caption("Nothing saved yet - allocating above writes the exam to "
                   f"`{exam_plan.EXAM_ROOT}/<exam>/` so it survives this session.")
    else:
        labels_x = {f"{m['title']} · {m.get('students', '?')} students"
                    f"{' · ' + m['when'] if m.get('when') else ''}": m for m in saved_exams}
        pick_x = st.selectbox("Exam", list(labels_x), key="exam_saved_pick")
        meta_x = labels_x[pick_x]
        d_x = Path(meta_x["dir"])
        alloc_x = pd.read_csv(d_x / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        st.caption(f"`{d_x}` · saved {meta_x.get('saved', '?')} · "
                   f"{', '.join(meta_x.get('rooms', [])) or 'room unknown'} · "
                   + ("alternate seating" if meta_x.get("spaced") else "every seat filled"))

        fc = st.columns(4)
        for i, (label, fname) in enumerate([("Allocation (.csv)", exam_plan.ALLOC_FILE),
                                            ("Seat map", "seat_map.html"),
                                            ("Find-my-seat", "find-my-seat.html")]):
            f = d_x / fname
            if f.exists():
                fc[i].download_button(label, f.read_bytes(),
                                      file_name=f"{meta_x['slug']}_{fname}",
                                      key=f"exam_saved_dl_{i}")
        if fc[3].button("Load for mail & packs", key="exam_saved_load"):
            st.session_state["exam_last"] = dict(
                alloc=alloc_x.copy(), exam=meta_x.get("title", ""),
                course=meta_x.get("course", C.COURSE), when=meta_x.get("when", ""),
                venue=meta_x.get("venue", ""), slug=meta_x.get("slug", ""))
            st.success(f"**{meta_x.get('title')}** loaded - the **🖼️ Posters & packs** and "
                       "**📧 Email students** pages will now offer it.")

        st.dataframe(alloc_x[["Roll", "Name", "Room", "Block", "Seat"]].head(200),
                     width="stretch", hide_index=True)
        if len(alloc_x) > 200:
            st.caption(f"Showing the first 200 of {len(alloc_x)} rows - the CSV has all of them.")

# ───────────────────────── Answer script showing ─────────────────────────
if PAGE == "Answer script showing":
    page_header("Answer script showing",
                "One printable marks-entry sheet per TA, room and time slot, from the duty roster.")
    st.caption(
        "Reads the **Quiz 1 Sheets Showing** tab of the TA duty workbook for the schedule - "
        "who shows which group's scripts, in which room and slot, with which Evaluation TAs "
        "on standby - and the Class Attendance workbook for the roster. Builds a cover page "
        "(the whole schedule) plus one sheet per duty row: Sr, Roll, Name, and blank Marks / "
        "Tick / Remarks boxes for the TA to fill by pen while showing the scripts."
    )

    duty_candidates = sorted(str(q) for q in Path(".").glob("*TA*.xls*")
                             if "template" not in q.stem.lower())
    roster_candidates = sorted(str(q) for q in Path(".").glob("*Class Attendance*.xls*")
                               if "backup" not in q.stem.lower())

    if not duty_candidates:
        st.info(f"No `*TA*.xlsx` found in the project folder - need `{ansshow.DUTY_BOOK}` "
                f"with a **{ansshow.DUTY_SHEET}** tab.")
    elif not roster_candidates:
        st.info(f"No `*Class Attendance*.xlsx` found in the project folder - "
                f"need `{ansshow.ROSTER}`, the roster this schedule is checked against.")
    else:
        c1, c2 = st.columns(2)
        duty_pick = c1.selectbox(
            "TA duty workbook", duty_candidates,
            index=duty_candidates.index(ansshow.DUTY_BOOK) if ansshow.DUTY_BOOK in duty_candidates else 0,
            key="ansshow_duty")
        roster_pick = c2.selectbox(
            "Roster workbook", roster_candidates,
            index=roster_candidates.index(ansshow.ROSTER) if ansshow.ROSTER in roster_candidates else 0,
            key="ansshow_roster")

        try:
            duties, sched_date = ansshow.read_schedule(duty_pick)
        except Exception as e:                        # noqa: BLE001
            st.error(f"Couldn't read the **{ansshow.DUTY_SHEET}** tab of `{duty_pick}`: "
                     f"{type(e).__name__}: {e}")
            st.stop()

        try:
            groups = ansshow.load_groups(roster_pick)
        except Exception as e:                        # noqa: BLE001
            st.error(f"Couldn't read `{roster_pick}`: {type(e).__name__}: {e}")
            st.stop()

        total = sum(len(g) for g in groups.values())
        st.caption(f"{len(duties)} duty row(s) · {total} students on the roster · "
                   f"date on the sheet: **{sched_date or '(none found)'}**")
        st.dataframe(
            pd.DataFrame([{"TA": d.name, "Roll": d.roll, "Group": d.label,
                           "Room": d.room, "Slot": d.slot,
                           "Evaluation TAs": ", ".join(d.evaluators)} for d in duties]),
            width="stretch", hide_index=True)

        try:
            ansshow.check(duties, groups)
        except SystemExit as e:
            st.error(str(e))
            st.caption(f"Fix the **{ansshow.DUTY_SHEET}** tab (or the roster) and reload this page.")
            st.stop()
        st.success("Schedule matches the roster: every student shown exactly once, counts agree.")

        if st.button("Build answer-showing pack", type="primary", key="ansshow_run"):
            snapshot_before([ansshow.OUT], "answer-showing",
                            note="before rebuilding the answer-showing pack")
            with st.spinner("Building…"):
                built = ansshow.build(duties, groups, total, sched_date)
            st.session_state["ansshow_built"] = {
                "pack": str(built["pack"]), "xlsx": str(built["xlsx"]),
                "per_ta": [str(q) for q in built["per_ta"]],
            }
            st.success(f"Built {len(duties)} sheet(s) for {total} students → `{ansshow.OUT}`")

        built_s = st.session_state.get("ansshow_built")
        if built_s and Path(built_s["pack"]).exists():
            dl1, dl2 = st.columns(2)
            dl1.download_button("Download the full pack (cover + every TA's sheet)",
                                Path(built_s["pack"]).read_bytes(),
                                file_name=Path(built_s["pack"]).name, key="ansshow_dl_pack")
            dl2.download_button("Download the tracking workbook (.xlsx)",
                                Path(built_s["xlsx"]).read_bytes(),
                                file_name=Path(built_s["xlsx"]).name, key="ansshow_dl_xlsx")
            with st.expander(f"Individual TA sheets ({len(built_s['per_ta'])})"):
                for pt in built_s["per_ta"]:
                    ptp = Path(pt)
                    if ptp.exists():
                        st.download_button(ptp.name, ptp.read_bytes(), file_name=ptp.name,
                                           key=f"ansshow_dl_{ptp.name}")
            st.caption("Print the pack at 100% scale. Marks / Tick / Remarks are filled by pen "
                       "during the showing, then the same sheets come back here to enter into "
                       f"the course grade sheet.")

# ───────────────────────── Attendance summary ─────────────────────────
if PAGE == "Attendance summary":
    page_header("Attendance summary", "Read from the Institute's master workbook, percentages recomputed here.")
    st.caption(
        "Reads the Institute's own master **Class Attendance** workbook - one sheet per "
        "medium, Y/N per student per session held so far - not any per-cohort "
        "`attendance.xlsx` left over from the old scan pipeline. Percent is recomputed here "
        "(present / sessions held, ignoring blanks and holiday columns) rather than trusted "
        "from the workbook's cached formula cells."
    )

    att_candidates = sorted(str(p) for p in Path(".").glob("*Class Attendance*.xls*")
                            if "backup" not in p.stem.lower())
    att_src = st.radio("Workbook source", ["Use a file already in the project folder", "Upload a file"],
                       horizontal=True, key="attrep_src")
    att_bytes = att_fname = None
    if att_src == "Use a file already in the project folder":
        att_pick = st.selectbox("Attendance workbook", ["(choose)"] + att_candidates,
                                index=1 if len(att_candidates) == 1 else 0, key="attrep_pick")
        if att_pick != "(choose)":
            att_fname, att_bytes = att_pick, Path(att_pick).read_bytes()
    else:
        att_up = st.file_uploader("Class Attendance workbook (.xlsx)", type=["xlsx", "xls"],
                                  key="attrep_upload")
        if att_up:
            att_bytes, att_fname = att_up.getbuffer(), att_up.name

    if att_bytes is None:
        st.info("Pick the Class Attendance workbook to see the summary.")
    else:
        with tempfile.TemporaryDirectory() as tmp:
            att_path = Path(tmp) / (att_fname or "attendance.xlsx")
            att_path.write_bytes(att_bytes)
            try:
                # registers only: the workbook also carries a Dashboard tab,
                # which holds no marks and cannot be summed over
                sheets_avail = attrep.register_sheets(str(att_path))
            except Exception as e:
                st.error(f"Couldn't read the workbook: {e}")
                st.stop()

            cohort_only = attrep.cohort_sheets(str(att_path))
            combined_avail = [s for s in sheets_avail if s not in cohort_only]
            # the combined sheet is the whole course on one list, which is what a
            # summary and a defaulter list are: one table, batch in a column.
            # The batch sheets are still there for a per-batch read.
            cohorts_pick = st.multiselect("Sheets to include", sheets_avail,
                                          default=combined_avail or cohort_only,
                                          key="attrep_cohorts")
            if combined_avail:
                st.caption(
                    f"**{', '.join(combined_avail)}** is both batches on one list, and the "
                    "batch each student is in comes from its own Batch column. Picking it "
                    "*alongside* the batch sheets counts every student twice."
                )
            if [x for x in cohorts_pick if x in combined_avail] and \
                    [x for x in cohorts_pick if x in cohort_only]:
                st.warning("A combined sheet is selected together with a batch sheet, so "
                           "every student is counted twice in the figures below.")
            # the course keeps its own live view in the workbook's Dashboard sheet;
            # its settings are the course's, so they are read rather than a
            # second copy being kept here
            wb_dash = attrep.dashboard_blocks(str(att_path))
            wb_set = wb_dash.get("settings", {})
            thr_default = int(float(wb_set.get("threshold (%)") or C.ATTENDANCE_THRESHOLD))
            bench_default = float(wb_set.get("class benchmark (%)") or 0) or None
            total_course = int(float(wb_set.get("total classes (course)") or 0) or C.TOTAL_CLASSES)
            threshold = st.slider("Defaulter threshold - below this % is flagged", 0, 100,
                                  thr_default, key="attrep_threshold")
            if wb_set:
                st.caption(
                    f"Defaults read from the workbook's **{attrep.DASHBOARD_SHEET}** sheet: "
                    f"threshold {mailer._pct(thr_default)}%"
                    + (f", class benchmark {mailer._pct(bench_default)}%" if bench_default else "")
                    + (f", {total_course} classes in the course" if total_course else "")
                    + ". Change the slider to ask a different question here; the sheet is "
                      "not written back to."
                )
            exc_mode = st.radio(
                "Excused absences (**E** - medical receipt or another accepted reason)",
                ["Don't count that class at all", "Count it as present"],
                index=0 if C.EXCUSED_MODE == "exclude" else 1,
                horizontal=True, key="attrep_excused",
                help="An E is never counted as an absence either way, so it can't trigger a "
                     "shortfall notice. This only decides what it does to the percentage: "
                     "drop the session from that student's denominator, or credit it.")
            exc_arg = "exclude" if exc_mode.startswith("Don't") else "present"

            if not cohorts_pick:
                st.info("Pick at least one sheet.")
                st.stop()

            students_parts, sessions_parts = [], []
            try:
                for sh in cohorts_pick:
                    stu, sess = attrep.load(str(att_path), sh, excused=exc_arg)
                    # on the combined sheet the batch is a column, not the sheet
                    # name, so a defaulter list read from it still says who is
                    # in which batch
                    stu["Cohort"] = stu.Batch.where(stu.Batch.astype(bool), sh) \
                        if "Batch" in stu else sh
                    sess["Cohort"] = sh
                    students_parts.append(stu)
                    sessions_parts.append(sess)
            except Exception as e:
                st.error(f"Couldn't parse sheet {sh!r}: {e}")
                st.stop()

            # per-student marks for the trend chart: the combined sheet if it is
            # picked (one read, batch in a column), otherwise each batch sheet
            trend_sheets = ([x for x in cohorts_pick if x in combined_avail][:1]
                            or [x for x in cohorts_pick if x not in combined_avail])
            try:
                marks_df = pd.concat([attrep.load_marks(str(att_path), sh)
                                      for sh in trend_sheets], ignore_index=True)
            except Exception as e:                  # noqa: BLE001 - chart is optional
                marks_df = pd.DataFrame()
                st.caption(f"Couldn't read per-class marks for the trend chart: {e}")

        students_df = pd.concat(students_parts, ignore_index=True)
        sessions_df = pd.concat(sessions_parts, ignore_index=True)
        held_df = students_df[students_df.Held > 0]

        held_sessions = sessions_df[sessions_df.Present + sessions_df.Absent > 0].groupby("Cohort").size()

        held_now = int(held_sessions.max()) if len(held_sessions) else 0
        remaining_now = max(0, (total_course or len(sessions_df) // max(len(cohorts_pick), 1))
                            - held_now)

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Students", len(students_df))
        c2.metric("Classes held", held_now,
                  help=f"{remaining_now} still to be held" if remaining_now else None)
        c3.metric("Class average", f"{held_df.Percent.mean():.1f}%" if len(held_df) else "-")
        c4.metric(f"Below {threshold}%", int((held_df.Percent < threshold).sum()))

        # Per batch, plus the two counts the course's own sheet carries: who is
        # critically low, and who can no longer reach the threshold at all -
        # that second one is the line between the two mail templates.
        def _cannot(df):
            return sum(1 for r in df.itertuples()
                       if not mailer.can_still_reach(int(r.Present), int(r.Absent),
                                                     remaining_now, float(threshold)))

        rows_sum = []
        for label, part in [("All", held_df)] + [(c, held_df[held_df.Cohort == c])
                                                 for c in dict.fromkeys(held_df.Cohort)]:
            if not len(part):
                continue
            rows_sum.append({
                "": label, "Students": len(part),
                f"Below {threshold}%": int((part.Percent < threshold).sum()),
                "Average": round(part.Percent.mean(), 1),
                "Critical (<50%)": int((part.Percent < 50).sum()),
                "Cannot reach": _cannot(part[part.Percent < threshold]),
            })
        if bench_default:
            for row, part in zip(rows_sum, [held_df] + [held_df[held_df.Cohort == c]
                                                        for c in dict.fromkeys(held_df.Cohort)]):
                row[f"At or above {mailer._pct(bench_default)}%"] = \
                    int((part.Percent >= bench_default).sum())
        st.dataframe(pd.DataFrame(rows_sum).style.format({"Average": "{:.1f}%"}, na_rep="-"),
                     width="stretch", hide_index=True)

        bands = [("< 50%", 0, 50), ("50-75%", 50, 75), ("75-90%", 75, 90), ("90-100%", 90, 101)]
        band_df = pd.DataFrame([
            {"Attendance band": name,
             "Students": int(((held_df.Percent >= lo) & (held_df.Percent < hi)).sum())}
            for name, lo, hi in bands])
        bc1, bc2 = st.columns([1, 2])
        bc1.dataframe(band_df, width="stretch", hide_index=True)
        # an explicit sort: st.bar_chart orders the bands as text, which put
        # "< 50%" after "90-100%"
        bc2.altair_chart(
            alt.Chart(band_df).mark_bar().encode(
                x=alt.X("Attendance band:N", sort=[b[0] for b in bands], title=None,
                        axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Students:Q"), tooltip=["Attendance band", "Students"]),
            width="stretch")

        if wb_dash.get("summary"):
            # the sheet's own figures, against the same figures recomputed here
            wb_sum = wb_dash["summary"]
            here = {r[""]: r for r in rows_sum}
            drift = []
            for label, key in (("Students", "students"),
                               ("Below threshold", "below threshold"),
                               ("Critical (below 50%)", "critical (below 50%)"),
                               ("Can't reach threshold by end", "can't reach threshold by end")):
                sheet_row = wb_sum.get(key, {})
                for col, val in sheet_row.items():
                    mine = here.get(col, {})
                    ours = {"Students": mine.get("Students"),
                            "Below threshold": mine.get(f"Below {threshold}%"),
                            "Critical (below 50%)": mine.get("Critical (<50%)"),
                            "Can't reach threshold by end": mine.get("Cannot reach")}[label]
                    if ours is not None and val is not None and int(float(val)) != int(ours):
                        drift.append(f"{label} · {col}: sheet {int(float(val))}, register {ours}")
            if drift:
                st.warning("The workbook's Dashboard sheet disagrees with the register it is "
                           "built from - " + "; ".join(drift[:4])
                           + ". Its formulas may not have recalculated since the last edit.")
            else:
                st.caption(f"Matches the workbook's {attrep.DASHBOARD_SHEET} sheet exactly.")

        trend = attrep.batch_trend(marks_df, excused=exc_arg)
        if not trend.empty:
            st.markdown("#### Attendance over the term, by batch")
            tc1, tc2 = st.columns([2, 1])
            trend_kind = tc1.radio(
                "Show", ["Turnout at each class", "Running attendance to date"],
                horizontal=True, key="attrep_trend_kind",
                help="Turnout is who came to that class. Running attendance is the batch's "
                     "percentage over every class so far - the same measure a student's own "
                     "percentage uses, so it is the one to hold against the threshold.")
            show_all = tc2.checkbox("Whole course line", value=trend.Batch.nunique() > 1,
                                    key="attrep_trend_all",
                                    disabled=trend.Batch.nunique() < 2)
            ycol = "Percent" if trend_kind.startswith("Turnout") else "Cumulative"
            plot = trend.copy()
            if show_all and trend.Batch.nunique() > 1:
                allc = (trend.groupby(["Date"], as_index=False)[["Present", "Absent", "Excused"]]
                        .sum().sort_values("Date"))
                cred = allc.Present + (allc.Excused if exc_arg == "present" else 0)
                cnt = cred + allc.Absent
                allc["Percent"] = cred / cnt.where(cnt > 0) * 100
                allc["Cumulative"] = cred.cumsum() / cnt.cumsum().where(cnt.cumsum() > 0) * 100
                allc["Batch"], allc["Session"] = "Whole course", ""
                plot = pd.concat([plot, allc[plot.columns]], ignore_index=True)
            plot["Date"] = pd.to_datetime(plot["Date"])
            lo = max(0, min(float(plot[ycol].min()), float(threshold)) - 5)
            base = alt.Chart(plot).encode(
                x=alt.X("Date:T", title=None),
                y=alt.Y(f"{ycol}:Q", title="Attendance (%)",
                        scale=alt.Scale(domain=[lo, 100])),
                color=alt.Color("Batch:N", title=None, legend=alt.Legend(orient="top")),
                tooltip=[alt.Tooltip("Batch:N"), alt.Tooltip("Session:N"),
                         alt.Tooltip("Date:T", format="%d %b"),
                         alt.Tooltip(f"{ycol}:Q", format=".1f", title="%"),
                         alt.Tooltip("Present:Q"), alt.Tooltip("Absent:Q")])
            rules = [alt.Chart(pd.DataFrame({"y": [float(threshold)],
                                             "label": [f"Threshold {mailer._pct(threshold)}%"]}))
                     .mark_rule(strokeDash=[6, 4], color="#c0392b").encode(y="y:Q",
                                                                           tooltip=["label:N"])]
            if bench_default:
                rules.append(alt.Chart(pd.DataFrame({"y": [float(bench_default)],
                                                     "label": [f"Benchmark {mailer._pct(bench_default)}%"]}))
                             .mark_rule(strokeDash=[2, 3], color="#7f8c8d").encode(
                                 y="y:Q", tooltip=["label:N"]))
            chart = alt.layer(base.mark_line(point=True), *rules).properties(height=320)
            st.altair_chart(chart, width="stretch")
            below_cls = trend[trend.Percent < float(bench_default or threshold)]
            if len(below_cls):
                st.caption(f"{len(below_cls)} class(es) had turnout under "
                           f"{mailer._pct(bench_default or threshold)}%: "
                           + ", ".join(f"{r.Batch} {r.Session} ({r.Date:%d %b}, "
                                       f"{mailer._pct(r.Percent)}%)"
                                       for r in below_cls.head(6).itertuples())
                           + (" …" if len(below_cls) > 6 else ""))

        st.markdown(f"#### Defaulters - below {threshold}%")
        defaulters = (held_df[held_df.Percent < threshold]
                     .sort_values("Percent")[["Cohort", "Roll", "Name", "Present", "Absent",
                                               "Excused", "Held", "Percent"]])
        if defaulters.empty:
            st.success("No one below the threshold.")
        else:
            st.dataframe(defaulters.style.format({"Percent": "{:.1f}%"}, na_rep="-"),
                        width="stretch", hide_index=True)
            st.download_button("Download defaulter list (.csv)",
                               defaulters.to_csv(index=False).encode("utf-8"),
                               file_name="defaulters.csv", key="attrep_dl_def")

        # ── at-risk forecast ─────────────────────────────────────────────
        st.markdown(f"#### At risk - on or above {threshold}% today, but not for long")
        fr1, fr2 = st.columns([1, 2])
        horizon_max = max(1, remaining_now) if remaining_now else 10
        horizon = fr1.number_input(
            "If they miss the next … classes", min_value=1, max_value=int(horizon_max),
            value=min(3, int(horizon_max)), step=1, key="attrep_risk_horizon",
            help="Capped at the classes still to be held: nobody is at risk from a class "
                 "that won't happen.")
        fr2.caption(
            "Everyone listed is fine *today*, so none of them gets a shortfall notice yet. "
            "**Can miss** is how many classes in a row they can still miss and stay on the "
            "threshold - the same arithmetic as the *attend the next N* figure in a notice. "
            "A word now is cheaper than a recovery plan later.")
        risk_df = attrep.at_risk(held_df, float(threshold), int(horizon),
                                 remaining=remaining_now or None)
        if risk_df.empty:
            st.success(f"Nobody at or above {threshold}% would drop below it by missing the "
                       f"next {int(horizon)} class(es).")
        else:
            risk_show = risk_df[["Cohort", "Roll", "Name", "Present", "Absent", "Held",
                                 "Percent", "CanMiss", "IfMissed"]].rename(
                columns={"CanMiss": "Can miss", "IfMissed": f"After missing {int(horizon)}"})
            rk1, rk2 = st.columns(2)
            rk1.metric("At risk", len(risk_df))
            rk2.metric("One absence from the line", int((risk_df.CanMiss == 0).sum()),
                       help="Can miss 0: the very next absence takes them below the threshold.")
            st.dataframe(risk_show.style.format({"Percent": "{:.1f}%",
                                                 f"After missing {int(horizon)}": "{:.1f}%"},
                                                na_rep="-"),
                         width="stretch", hide_index=True)
            st.download_button("Download at-risk list (.csv)",
                               risk_show.to_csv(index=False).encode("utf-8"),
                               file_name="at_risk.csv", key="attrep_dl_risk")

        with st.expander(f"Full summary - all {len(students_df)} students"):
            full = students_df.sort_values(["Cohort", "Percent"],
                                           na_position="first")[
                ["Cohort", "Roll", "Name", "Present", "Absent", "Excused", "Held", "Percent"]]
            st.dataframe(full.style.format({"Percent": "{:.1f}%"}, na_rep="-"),
                        width="stretch", hide_index=True)
            st.download_button("Download full summary (.csv)",
                               full.to_csv(index=False).encode("utf-8"),
                               file_name="attendance_summary.csv", key="attrep_dl_full")

        # ── the printable report ─────────────────────────────────────────
        st.divider()
        st.markdown("#### Attendance report (PDF)")
        st.caption(
            "The same figures as a report to hand over or attach: where the class stands "
            "and a batch comparison, class by class with anything under the benchmark "
            "marked, everyone below the requirement with how many consecutive classes "
            "would bring them back and the best they can still finish on, and a method "
            "page. Built from the register, so it cannot drift from the mail a student "
            "gets."
        )
        rp1, rp2, rp3 = st.columns(3)
        rep_sheet = rp1.selectbox(
            "Read from", cohorts_pick, key="attrep_pdf_sheet",
            help="A combined sheet covers both batches in one report. A batch sheet "
                 "reports that batch alone.")
        rep_total = rp2.number_input("Classes in the course", min_value=1, step=1,
                                     value=int(total_course or len(sessions_df)
                                               // max(len(cohorts_pick), 1)),
                                     key="attrep_pdf_total")
        rep_bench = rp3.number_input("Class benchmark (%)", min_value=0.0, max_value=100.0,
                                     step=1.0,
                                     value=float(bench_default or attpdf.DEFAULT_BENCHMARK),
                                     key="attrep_pdf_bench",
                                     help="A class whose turnout is under this is marked on "
                                          "the class-wise page.")
        rp4, rp5 = st.columns(2)
        rep_classwise = rp4.checkbox("Include the class-wise page", value=True,
                                     key="attrep_pdf_classwise")
        rep_method = rp5.checkbox("Include the method page", value=True,
                                  key="attrep_pdf_method",
                                  help="What the percentage means, what an excused mark "
                                       "does, and which students are special cases. A "
                                       "report that states its method can be checked.")
        rep_name = st.text_input(
            "Output file", f"{C.COURSE_CODE}_Attendance_Report_{date.today():%d%b%Y}.pdf",
            key="attrep_pdf_name")

        if st.button("Build the attendance report", type="primary", key="attrep_pdf_run"):
            out_rep = Path(rep_name.strip() or "attendance_report.pdf")
            snapshot_before([out_rep], "attendance-report")
            try:
                # `att_path` lived in the TemporaryDirectory above, which is
                # gone by now - this button always failed with "No such file".
                # Write the same bytes to a fresh one for the build.
                with tempfile.TemporaryDirectory() as tmp_rep:
                    rep_src = Path(tmp_rep) / (att_fname and Path(att_fname).name
                                               or "attendance.xlsx")
                    rep_src.write_bytes(att_bytes)
                    attpdf.build_report(
                        str(rep_src), str(out_rep), sheet=rep_sheet,
                        threshold=float(threshold), total_classes=int(rep_total),
                        benchmark=float(rep_bench), class_wise=rep_classwise,
                        method=rep_method, excused=exc_arg)
            except Exception as e:                      # noqa: BLE001 - surfaced as-is
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            st.success(f"Built `{out_rep}`.")
            st.download_button("Download the report (PDF)", out_rep.read_bytes(),
                               file_name=out_rep.name, key="attrep_pdf_dl")

# ───────────────────────── Student lookup ─────────────────────────
if PAGE == "Student lookup":
    page_header("Student lookup",
                "One student's seat, exam seats, attendance, moves and mail - for answering a query.")
    st.caption("Read-only: gathered from the allocations, saved exams, the Class Attendance "
               "workbook, `out/transitions.csv` and the mail log. Nothing here writes a file.")
    people_l = lookup.directory()
    q_l = st.text_input("Roll number or name", "", key="lookup_query",
                        placeholder="B26CY1695, or part of a name")
    if not q_l.strip():
        st.info(f"{len(people_l)} students on file. Type a roll number or a name.")
    else:
        hits_l = lookup.search(q_l, people_l)
        if hits_l.empty:
            st.warning(f"No student matches **{q_l.strip()}**.")
        else:
            opts_l = {f"{r.Roll} - {r.Name} ({r.Batch})": r.Roll for r in hits_l.itertuples()}
            pick_l = (next(iter(opts_l)) if len(opts_l) == 1 else
                      st.selectbox(f"{len(opts_l)} matches", list(opts_l), key="lookup_pick"))
            rec = lookup.student_record(opts_l[pick_l])
            att_l = rec.attendance or {}

            st.markdown(f"### {rec.name or rec.roll}")
            st.caption(" · ".join(x for x in [rec.roll, rec.batch and f"{rec.batch} batch",
                                               rec.email] if x))
            m1, m2, m3, m4 = st.columns(4)
            seat_l = rec.classroom[0] if rec.classroom else None
            m1.metric("Classroom seat", f"Block {seat_l['Block']}" if seat_l else "-",
                      help=(f"{seat_l['Room']}, seat {seat_l['Seat']}. Classroom notices name "
                            "the block only." if seat_l else "Not in any allocation."))
            if att_l.get("percent") is not None:
                m2.metric("Attendance", f"{mailer._pct(att_l['percent'])}%",
                          help=f"{att_l['present']} present, {att_l['absent']} absent, "
                               f"{att_l['excused']} excused, as of {att_l['asof']}.")
                if att_l["below"]:
                    m3.metric("To get back", f"{att_l['to_recover']} in a row"
                              if att_l["can_reach"] else "out of reach",
                              help=(f"Best still possible: "
                                    f"{mailer._pct(att_l['best_possible'])}% with "
                                    f"{att_l['remaining']} classes left."))
                else:
                    m3.metric("Can still miss", att_l["can_miss"],
                              help=f"Classes in a row before dropping below "
                                   f"{mailer._pct(att_l['threshold'])}%.")
            else:
                m2.metric("Attendance", "-")
                m3.metric("Can still miss", "-")
            m4.metric("Mails logged", len(rec.mail))

            if att_l.get("below"):
                st.error(f"Below the {mailer._pct(att_l['threshold'])}% requirement."
                         + ("" if att_l["can_reach"] else
                            f" Cannot reach it: at best {mailer._pct(att_l['best_possible'])}%."))
            elif att_l and att_l.get("can_miss", 99) <= 1:
                st.warning("On the line: "
                           + ("the next absence" if att_l["can_miss"] == 0
                              else "one more absence is fine, a second")
                           + f" takes them below {mailer._pct(att_l['threshold'])}%.")

            t_att, t_seat, t_hist = st.tabs(["Attendance", "Seats", "Moves & mail"])
            with t_att:
                if rec.marks.empty:
                    st.info("Not found on any batch sheet of the Class Attendance workbook.")
                else:
                    held_l = rec.marks[rec.marks.Mark != ""].copy()
                    label_l = {"Y": "Present", "N": "Absent", "E": "Excused", "NA": "N/A"}
                    held_l["Status"] = held_l.Mark.map(label_l).fillna(held_l.Mark)
                    held_l["Date"] = pd.to_datetime(held_l.Date)
                    strip = alt.Chart(held_l).mark_square(size=180).encode(
                        x=alt.X("Date:T", title=None),
                        color=alt.Color("Status:N", title=None, legend=alt.Legend(orient="top"),
                                        scale=alt.Scale(domain=["Present", "Absent", "Excused", "N/A"],
                                                        range=["#2e8b57", "#c0392b", "#e0a526", "#9aa0a6"])),
                        tooltip=["Session:N", alt.Tooltip("Date:T", format="%a %d %b"), "Status:N"],
                    ).properties(height=70)
                    st.altair_chart(strip, width="stretch")
                    missed_l = held_l[held_l.Mark.isin(["N", "E"])]
                    if missed_l.empty:
                        st.success("No absences.")
                    else:
                        st.markdown("**Absent or excused**")
                        st.dataframe(missed_l.assign(Date=missed_l.Date.dt.strftime("%a %d %b %Y"))
                                     [["Session", "Date", "Status"]],
                                     width="stretch", hide_index=True)
                    st.caption(f"{att_l.get('classes_held', 0)} classes held, "
                               f"{att_l.get('remaining', 0)} still to come. Read from the "
                               f"**{att_l.get('sheet', '')}** sheet.")
            with t_seat:
                if rec.classroom:
                    st.markdown("**Classroom**")
                    st.dataframe(pd.DataFrame(rec.classroom), width="stretch", hide_index=True)
                else:
                    st.info("Not seated in any classroom allocation.")
                if rec.exams:
                    st.markdown("**Exams**")
                    st.dataframe(pd.DataFrame(rec.exams), width="stretch", hide_index=True)
                else:
                    st.caption("No saved exam seats this student.")
            with t_hist:
                if len(rec.moves):
                    st.markdown("**Seat moves**")
                    st.dataframe(rec.moves, width="stretch", hide_index=True)
                else:
                    st.caption("Never moved through the toolkit.")
                if len(rec.mail):
                    st.markdown("**Mail log**")
                    st.dataframe(rec.mail[["timestamp", "kind", "subject", "status"]],
                                 width="stretch", hide_index=True)
                else:
                    st.caption("Nothing in the mail log for this student.")

            # a plain answer to paste into a reply
            lines_l = [f"{rec.name} ({rec.roll})" if rec.name else rec.roll]
            if seat_l:
                lines_l.append(f"Classroom: {seat_l['Room']}, Block {seat_l['Block']}")
            for e in rec.exams[:3]:
                lines_l.append(f"{e['Exam']}: {e['Room']}, Block {e['Block']}, Seat {e['Seat']}"
                               + (f" - {e['When']}" if e["When"] else ""))
            if att_l.get("percent") is not None:
                absent_on = ", ".join(pd.to_datetime(rec.marks[rec.marks.Mark == "N"].Date)
                                      .dt.strftime("%d %b"))
                lines_l.append(f"Attendance: {mailer._pct(att_l['percent'])}% "
                               f"({att_l['present']} present, {att_l['absent']} absent"
                               + (f", {att_l['excused']} excused" if att_l["excused"] else "")
                               + f") as of {att_l['asof']}")
                if absent_on:
                    lines_l.append(f"Absent on: {absent_on}")
            with st.expander("Summary to paste into a reply"):
                st.code("\n".join(lines_l), language=None)

# ───────────────────────── Course setup ─────────────────────────
if PAGE == "Course setup":
    page_header("Course setup",
                "The course's details, entered once - every page, sheet, poster and mail uses them.")
    flash = st.session_state.pop("setup_flash", None)
    if flash:
        st.success(flash)
    have_settings = course_settings.exists()
    cur = course_settings.load()
    if not have_settings:
        st.info("**No course set up yet** - the dashboard is running on the demo course "
                f"(**{C.COURSE}**). Fill in the tabs below and press **Save course settings** "
                "at the bottom. You can come back and change anything later from "
                "**Maintenance › Course setup**.")
        if st.button("Just explore the demo for now", key="setup_skip"):
            st.session_state["phl_page"] = "Overview"
            st.rerun()
    else:
        leftovers = course_settings.demo_leftovers(cur)
        st.caption(f"Saved in `{course_settings.SETTINGS_FILE}`"
                   + (f" on {cur['saved'].replace('T', ' ')}" if cur.get("saved") else "")
                   + ". It syncs with the folder, so every machine uses the same settings.")
        if leftovers:
            st.warning("Still the demo's value: " + ", ".join(f"`{x}`" for x in leftovers))

    c_, i_, t_, a_ = cur["course"], cur["institute"], cur["ta"], cur["attendance"]
    with st.form("course_setup_form", border=False):
        tab_c, tab_p, tab_a, tab_h, tab_f = st.tabs(
            ["Course", "Institute & people", "Attendance rules", "Halls", "Files"])

        with tab_c:
            s1, s2 = st.columns([1, 2])
            f_code = s1.text_input("Course code", c_["code"], key="setup_code",
                                   help="Letters and digits, e.g. PHL1010. Used in file names.")
            f_title = s2.text_input("Course title", c_["title"], key="setup_title")
            s3, s4, s5, s6 = st.columns(4)
            f_session = s3.text_input("Session / semester", c_["session"], key="setup_session",
                                      placeholder="AY 2026-27 Sem 1")

            def _d(v):
                try:
                    return date.fromisoformat(v) if v else None
                except ValueError:
                    return None
            f_start = s4.date_input("First class", _d(c_.get("start_date")), key="setup_start",
                                    format="DD/MM/YYYY")
            f_end = s5.date_input("Last class", _d(c_.get("end_date")), key="setup_end",
                                  format="DD/MM/YYYY")
            f_total = s6.number_input("Classes planned", min_value=1, step=1,
                                      value=max(1, int(c_.get("total_classes") or 1)),
                                      key="setup_total",
                                      help="The whole course. Used for 'classes still to come' "
                                           "when the workbook's Dashboard sheet doesn't say.")

        with tab_p:
            p1, p2 = st.columns(2)
            f_inst = p1.text_input("Institute", i_["name"], key="setup_inst")
            f_short = p2.text_input("Short name", i_["short"], key="setup_short",
                                    help="Used where space is tight, e.g. IITJ.")
            p3, p4 = st.columns(2)
            f_dept = p3.text_input("Department", i_["department"], key="setup_dept")
            f_domain = p4.text_input("Student email domain", i_["email_domain"], key="setup_domain",
                                     help="Mail goes to <roll number>@this domain.")
            f_url = st.text_input("Public lookup page URL (optional)", i_.get("lookup_url", ""),
                                  key="setup_url",
                                  help="Where the find-your-block page is published; printed as "
                                       "a QR code on the hall posters.")
            st.markdown("**Course instructors** - printed on sheets and posters, Cc'd on every mail")
            f_instr = st.text_area("One per line: Name, email",
                                   course_settings.format_people(cur["instructors"]),
                                   key="setup_instructors", height=100)
            st.markdown("**You, as TA** - the signature on every mail")
            q1, q2, q3 = st.columns(3)
            f_ta_name = q1.text_input("Name", t_["name"], key="setup_ta_name")
            f_ta_local = q2.text_input("Name in second language (optional)", t_.get("name_local", ""),
                                       key="setup_ta_local",
                                       help="Printed after your name on bilingual mail, "
                                            "e.g. Hindi. Leave blank to omit.")
            f_ta_roll = q3.text_input("Roll number", t_.get("roll", ""), key="setup_ta_roll")
            q4, q5, q6 = st.columns(3)
            f_ta_role = q4.text_input("Role", t_.get("role", "Teaching Assistant"), key="setup_ta_role")
            f_ta_email = q5.text_input("Email (replies come here)", t_.get("email", ""),
                                       key="setup_ta_email")
            f_sender = q6.text_input("Send mail from (optional)", cur["mail"].get("sender", ""),
                                     key="setup_sender",
                                     help="The From address for direct SMTP sending. The password "
                                          "is never stored here - it stays in your environment.")
            f_contact = st.text_input("Contact line under each mail (optional)",
                                      t_.get("contact", ""), key="setup_contact",
                                      placeholder="Office hours: Mon 4-5 PM, Room 204")
            with st.expander("Preview the signatures"):
                _nm = (f"{f_ta_name} / {f_ta_local}" if f_ta_local and f_ta_name
                       else f_ta_name or f_ta_local)
                st.code("\n".join(x for x in [_nm, f_ta_roll, f_ta_role, f_dept, f_inst] if x)
                        + "\n\n--- attendance notices ---\n"
                        + "\n".join(x for x in [f_ta_name, f"{f_ta_role}, {f_code} {f_title}",
                                                ", ".join(y for y in (f_dept, f_short) if y)] if x),
                        language=None)
                st.caption("Updates when you save.")

        with tab_a:
            r1, r2, r3 = st.columns(3)
            f_thr = r1.number_input("Attendance requirement (%)", 1.0, 100.0,
                                    float(a_["threshold"]), 1.0, key="setup_thr")
            f_bench = r2.number_input("Class benchmark (%)", 0.0, 100.0, float(a_["benchmark"]),
                                      1.0, key="setup_bench",
                                      help="A class whose turnout is under this is marked in "
                                           "the report and the trend chart.")
            f_levels = r3.text_input("Notice after this many absences",
                                     ", ".join(str(x) for x in a_["levels"]), key="setup_levels",
                                     help="One number per notice, increasing: 5, 9, 13 sends a "
                                          "first notice at 5 absences, a second at 9, a final at 13.")
            f_exc = st.radio("An excused absence (E)",
                             ["Doesn't count at all (left out of the percentage)",
                              "Counts as present"],
                             index=0 if a_["excused"] == "exclude" else 1, key="setup_excused",
                             horizontal=True)
            st.caption("The workbook's own Dashboard sheet, when it has one, still sets the "
                       "defaults on the Attendance summary page; these apply everywhere else "
                       "and whenever it doesn't.")

        with tab_h:
            st.markdown("**Classroom halls** - one row per block. Add a row for a new block "
                        "or hall; select rows and press Delete to remove them.")
            f_rooms_df = st.data_editor(
                course_setup.rooms_table(cur), num_rows="dynamic", width="stretch",
                hide_index=True, key="setup_rooms",
                column_config={
                    "Hall": st.column_config.TextColumn(required=True, help="e.g. LHC110"),
                    "Block": st.column_config.TextColumn(required=True, max_chars=1,
                                                         help="One capital letter"),
                    "Side": st.column_config.TextColumn(help="Where it is, e.g. Front left"),
                    "Rows": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "Seats per row": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                    "Wedge": st.column_config.TextColumn(
                        help="Tapering rows behind the last full row, seats in each: 3, 2, 2, 1. "
                             "Wedge seats are held back for late admissions and transfers."),
                    "Section": st.column_config.SelectboxColumn(options=["front", "rear"],
                                                                required=True),
                })
            st.markdown("**Branch grouping (optional)** - which branches sit together in "
                        "which block. Blocks not listed take everyone else.")
            f_groups_df = st.data_editor(
                course_setup.groups_table(cur), num_rows="dynamic", width="stretch",
                hide_index=True, key="setup_groups",
                column_config={"Branches": st.column_config.TextColumn(
                    help="Two-letter branch codes from the roll number, e.g. CS, EE, ME")})
            st.caption("Exam venues are read from the exam-hall seating plan workbook "
                       "(Files tab), not from this table.")

        with tab_f:
            st.caption("Each file is checked the way the page that uses it reads it, then saved "
                       "in the project folder under the name the dashboard looks for. A file "
                       "already there is snapshotted first (Health check › history).")
            now = course_setup.installed_files(C.COURSE_CODE)
            uploads = {}
            for kind, (label, namer) in course_setup.FILE_KINDS.items():
                have = now.get(kind)
                uploads[kind] = st.file_uploader(
                    f"{label} → `{namer(f_code or C.COURSE_CODE)}`"
                    + (" (replaces the one there)" if have else ""),
                    type=["xlsx", "xls"] if kind == "roll_list" else ["xlsx"],
                    key=f"setup_file_{kind}")

        saved = st.form_submit_button("Save course settings", type="primary")

    # current state of the halls, below the form
    with st.expander("Seats per hall (as saved)"):
        st.dataframe(course_setup.capacity(cur["rooms"]), width="stretch", hide_index=True)

    if saved:
        rooms_new, room_errs = course_setup.rooms_from_table(f_rooms_df)
        try:
            levels = course_settings.parse_int_list(f_levels)
        except ValueError:
            levels, room_errs = [], room_errs + ["Notice levels must be numbers, e.g. 5, 9, 13."]
        new = {
            "course": {"code": f_code.strip(), "title": f_title.strip(),
                       "session": f_session.strip(),
                       "start_date": f_start.isoformat() if f_start else "",
                       "end_date": f_end.isoformat() if f_end else "",
                       "total_classes": int(f_total)},
            "institute": {"name": f_inst.strip(), "short": f_short.strip(),
                          "department": f_dept.strip(),
                          "email_domain": f_domain.strip().lstrip("@").lower(),
                          "lookup_url": f_url.strip()},
            "instructors": course_settings.parse_people(f_instr),
            "ta": {"name": f_ta_name.strip(), "name_local": f_ta_local.strip(),
                   "roll": f_ta_roll.strip(), "email": f_ta_email.strip(),
                   "role": f_ta_role.strip() or "Teaching Assistant",
                   "contact": f_contact.strip()},
            "attendance": {"threshold": float(f_thr), "benchmark": float(f_bench),
                           "levels": levels,
                           "excused": "exclude" if f_exc.startswith("Doesn't") else "present"},
            "mail": {"sender": f_sender.strip()},
            "rooms": rooms_new,
            "branch_groups": course_setup.groups_from_table(f_groups_df),
        }
        errs = room_errs + course_settings.problems(course_settings._merge(course_settings.DEFAULTS, new))
        file_errs = {}
        for kind, up in uploads.items():
            if up is not None:
                why = course_setup.check_upload(kind, up.getvalue(), up.name)
                if why:
                    file_errs[kind] = why
        if errs or file_errs:
            st.error("Not saved - fix these first:\n\n" + "\n".join(
                [f"- {e}" for e in errs]
                + [f"- {course_setup.FILE_KINDS[k][0]}: {w}" for k, w in file_errs.items()]))
        else:
            gone = course_setup.halls_in_use() - set(rooms_new)
            renames = course_setup.renames_for_code(C.COURSE_CODE, new["course"]["code"])
            to_replace = [course_setup.FILE_KINDS[k][1](new["course"]["code"])
                          for k, up in uploads.items() if up is not None]
            snapshot_before([course_settings.SETTINGS_FILE] + [p for p, _ in renames] + to_replace,
                            "course-setup")
            moved = course_setup.apply_renames(renames)
            course_settings.save(new)
            put = [course_setup.install_file(k, up.getvalue(), new["course"]["code"])
                   for k, up in uploads.items() if up is not None]
            course_setup.reload_toolkit()
            msg = f"Saved. The dashboard now runs **{new['course']['code']} {new['course']['title']}**."
            if moved:
                msg += " Renamed to follow the new course code: " + ", ".join(
                    f"`{a.name}` → `{b.name}`" for a, b in moved) + "."
            if put:
                msg += " Installed: " + ", ".join(f"`{p.name}`" for p in put) + "."
            if gone:
                msg += (" Note: existing seating was built for " + ", ".join(sorted(gone))
                        + ", which is no longer a hall here - re-allocate those batches.")
            st.session_state["setup_flash"] = msg
            st.rerun()

# ───────────────────────── Email students ─────────────────────────
if PAGE == "Email students":
    page_header("Email students", f"Seating changes, exam venues and attendance notices - to <enrolment>@{C.EMAIL_DOMAIN}.")
    st.caption(
        "Addresses are derived from the enrolment number - `B26BB1901` → "
        f"`b26bb1901@{mailer.INSTITUTE_DOMAIN}` - so there is no address list to keep "
        "current. A **classroom** notice names the block only (seating is block-wise); an "
        "**exam** notice may name the actual seat, because there it is a real chair. "
        f"Every real send is appended to `{mailer.MAIL_LOG}`."
    )

    mail_mode = st.radio(
        "Who are you telling?",
        ["Students whose block changed", "A whole block", "Exam seating",
         "Exam block/room (one mail per block)", "Attendance shortfall"],
        horizontal=True, key="mail_mode",
        help="The first reads `out/transitions.csv` - every move made through "
             "**5 · Move a student** - and offers only the students who actually "
             "changed block and haven't been told yet. The second mails everyone "
             "in a block, for a first announcement.")

    mails: list = []
    group_windows = True   # attendance mail is per-student; seating mail batches

    if mail_mode == "Students whose block changed":
        all_trans = seating.load_transitions()
        if all_trans.empty:
            st.info(f"No moves recorded yet - `{seating.TRANSITION_LOG}` is written by "
                    "the **🔀 5 · Move a student** page. A move made by hand in Excel or "
                    "in `allocation.csv` leaves no trace here; mail those with **A whole "
                    "block**, or move them through page 5 so they're logged.")
        else:
            changed = seating.block_changed(all_trans)
            pending = mailer.unmailed(changed)
            within = len(all_trans) - len(changed)

            tc1, tc2 = st.columns(2)
            only_pending = tc1.checkbox("Only those not mailed about this move yet",
                                        value=True, key="mail_tr_pending")
            days = tc2.number_input("Only moves from the last N days (0 = all)",
                                    min_value=0, value=0, step=1, key="mail_tr_days")

            show = pending if only_pending else changed
            if days:
                cutoff = (pd.Timestamp.now() - pd.Timedelta(days=int(days))).isoformat()
                show = show[show.timestamp > cutoff]

            st.caption(
                f"{len(all_trans)} move(s) logged · {len(changed)} changed block "
                f"({within} stayed inside one block - a seat renumber, nothing to tell "
                f"anyone) · {len(pending)} not yet mailed.")

            if show.empty:
                st.success("Nobody is waiting to be told." if only_pending else
                           "No moves match that filter.")
            else:
                st.dataframe(show[["timestamp", "roll", "name", "from_room", "from_block",
                                   "to_room", "to_block"]],
                             width="stretch", hide_index=True)
                labels_t = {f"{r.roll} - {r.name or '?'}  ({r.from_block} → {r.to_block})": i
                            for i, r in enumerate(show.itertuples())}
                chosen_t = st.multiselect(
                    "Students to mail - leave empty to mail all of the above",
                    list(labels_t), default=[], key="mail_tr_students")
                if chosen_t:
                    show = show.iloc[[labels_t[c] for c in chosen_t]]

                tm1, tm2 = st.columns(2)
                course_t = tm1.text_input("Course", C.COURSE, key="mail_tr_course")
                session_t = tm2.text_input("Session", C.SESSION, key="mail_tr_session")
                tm3, tm4 = st.columns(2)
                eff_t = tm3.text_input("Effective from (optional)", "the next class",
                                       key="mail_tr_effective")
                contact_t = tm4.text_input("Contact line (optional)", mailer.DEFAULT_CONTACT, key="mail_tr_contact")
                st.caption(
                    "Each of these names that student's own old and new block, so they "
                    "are all different messages - Gmail gets one compose window per "
                    "student. That is the price of telling people something true about "
                    "themselves.")

                bad = []
                for r in show.itertuples():
                    try:
                        mails.append(mailer.build_transition_mail(
                            r.roll, r.to_block, name=r.name,
                            old_block=r.from_block, room=r.to_room, old_room=r.from_room,
                            course=course_t, session=session_t,
                            medium=(r.to_cohort.title() if r.to_cohort else ""),
                            effective=eff_t, contact=contact_t))
                    except ValueError as e:
                        bad.append(f"{r.roll}: {e}")
                if bad:
                    st.error("Skipped - " + " · ".join(bad))

    elif mail_mode == "A whole block":
        found_allocs_e = discover_allocations()
        if not found_allocs_e:
            st.info("No `allocation.csv` found yet - allocate seats on page **🪑 1 · Allocate seats** first.")
        else:
            src_label = st.selectbox("Cohort", list(found_allocs_e), key="mail_cohort")
            src_path = Path(found_allocs_e[src_label])
            alloc_mail = pd.read_csv(src_path, dtype={"Roll": str})
            room_guess = next((p.stem[: -len("_seating")]
                               for p in src_path.parent.glob("*_seating.xlsx")
                               if p.stem.endswith("_seating")), list(C.ROOMS)[0])
            room_mail = st.selectbox("Room named in the mail", list(C.ROOMS),
                                     index=list(C.ROOMS).index(room_guess)
                                     if room_guess in C.ROOMS else 0, key="mail_room")

            blocks_all = sorted(alloc_mail.Block.astype(str).unique())
            pick_blocks = st.multiselect("Blocks to include", blocks_all, default=[],
                                         key="mail_blocks")
            subset = alloc_mail[alloc_mail.Block.astype(str).isin(pick_blocks)] if pick_blocks \
                else alloc_mail.iloc[0:0]
            names = "Name" in alloc_mail.columns
            labels = {f"{r.Roll} - {r.Name}" if names else r.Roll: r.Roll
                      for r in subset.itertuples()}
            chosen = st.multiselect(
                f"Students ({len(labels)} in the selected block(s)) - leave empty to mail all of them",
                list(labels), default=[], key="mail_students")
            if chosen:
                subset = subset[subset.Roll.isin([labels[c] for c in chosen])]

            mc1, mc2 = st.columns(2)
            course_mail = mc1.text_input("Course", C.COURSE, key="mail_course")
            session_mail = mc2.text_input("Session", C.SESSION, key="mail_session")
            mc3, mc4 = st.columns(2)
            eff_mail = mc3.text_input("Effective from (optional)", "the next class",
                                      key="mail_effective")
            contact_mail = mc4.text_input("Contact line (optional)", mailer.DEFAULT_CONTACT, key="mail_contact")
            medium_mail = st.text_input(
                "Batch/medium named in the mail (blank to leave it out)",
                src_path.parent.name.title(), key="mail_medium",
                help="Taken from the cohort, not from each row's `Medium` value - "
                     "that column reflects the roster's language-of-instruction flag "
                     "and can disagree with the room a student is actually seated in "
                     "(out/english currently holds several rows marked Hindi).")
            generic = st.checkbox(
                'Address the batch generically ("Dear student") - one identical '
                "message per block, so a whole block fits in a single Gmail window "
                "with everyone in Bcc",
                value=True, key="mail_generic",
                help="Untick to greet each student by name. That personalises every "
                     "message, so the Gmail hand-off then needs one window per "
                     "student - fine for a handful, impractical for a block.")

            if subset.empty:
                st.info("Pick at least one block to compose the mails.")
            else:
                bad = []
                for r in subset.itertuples():
                    try:
                        mails.append(mailer.build_transition_mail(
                            r.Roll, str(r.Block),
                            name=("student" if generic else (r.Name if names else "")),
                            room=room_mail, course=course_mail, session=session_mail,
                            medium=medium_mail.strip(),
                            effective=eff_mail, contact=contact_mail))
                    except ValueError as e:
                        bad.append(f"{r.Roll}: {e}")
                if bad:
                    st.error("Skipped - " + " · ".join(bad))

    elif mail_mode == "Exam seating":
        last_exam = st.session_state.get("exam_last")
        # Saved exams are on disk, so they are offered whether or not this
        # browser session allocated or loaded one - after a restart the only
        # choice used to be uploading a roll list the toolkit had just written.
        saved_x = exam_plan.list_exams()
        exam_src = st.radio(
            "Seating source",
            (["The allocation from the Exam seating page"] if last_exam else [])
            + (["A saved exam"] if saved_x else [])
            + ["Upload a roll list (.xlsx/.csv with Roll, Room, Block, Seat)"],
            key="mail_exam_src")

        edf_mail = None
        exam_defaults = last_exam or {}
        if last_exam and exam_src.startswith("The allocation"):
            edf_mail = last_exam["alloc"]
            st.caption(f"{len(edf_mail)} students · {last_exam['venue']}")
        elif exam_src == "A saved exam":
            labels_sx = {f"{m['title']} · {m.get('when') or m.get('saved', '')}": m
                         for m in saved_x}
            meta_sx = labels_sx[st.selectbox("Exam", list(labels_sx), key="mail_exam_saved_pick")]
            edf_mail = pd.read_csv(Path(meta_sx["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
            exam_defaults = {"exam": meta_sx.get("title", ""),
                             "course": meta_sx.get("course", C.COURSE),
                             "when": meta_sx.get("when", ""),
                             "slug": meta_sx.get("slug", "")}
            st.caption(f"{len(edf_mail)} students · {meta_sx.get('venue', '')}")
        else:
            up = st.file_uploader("Exam roll list", type=["xlsx", "xls", "csv"],
                                  key="mail_exam_upload")
            if up:
                edf_mail = (pd.read_csv(up, dtype={"Roll": str}) if up.name.endswith(".csv")
                            else pd.read_excel(up, dtype={"Roll": str}))
                missing_cols = [c for c in ("Roll", "Block") if c not in edf_mail.columns]
                if missing_cols:
                    st.error("Missing column(s): " + ", ".join(missing_cols))
                    edf_mail = None

        if edf_mail is None:
            st.info("Allocate an exam on the **🧑‍🎓 Exam seating** page, or upload a roll list here.")
        else:
            xc1, xc2 = st.columns(2)
            exam_name = xc1.text_input("Exam title", exam_defaults.get("exam", "Quiz 1"),
                                       key=f"mail_exam_title_{exam_defaults.get('slug', '')}")
            course_x = xc2.text_input("Course", exam_defaults.get("course", C.COURSE),
                                      key=f"mail_exam_course_{exam_defaults.get('slug', '')}")
            xc3, xc4 = st.columns(2)
            when_x = xc3.text_input("Date & time", exam_defaults.get("when", ""),
                                    key=f"mail_exam_when_{exam_defaults.get('slug', '')}")
            session_x = xc4.text_input("Session", C.SESSION, key="mail_exam_session")
            include_seat = st.checkbox("Name the exact seat (uncheck for a block-only exam venue)",
                                       value=True, key="mail_exam_seat")
            note_x = st.text_area("Extra note (optional)", "", key="mail_exam_note")
            generic_x = st.checkbox(
                'Address the batch generically ("Dear student")', value=False,
                key="mail_exam_generic",
                help="Only groups into one Gmail window if the seating text is also "
                     "identical - which it is for a block-only venue, but not once "
                     "each student has their own seat number.")
            contact_x = st.text_input("Contact line (optional)", mailer.DEFAULT_CONTACT, key="mail_exam_contact")

            names_x = "Name" in edf_mail.columns
            labels_x = {f"{r.Roll} - {r.Name}" if names_x else r.Roll: r.Roll
                        for r in edf_mail.itertuples()}
            chosen_x = st.multiselect(
                f"Students ({len(labels_x)} allocated) - leave empty to mail all of them",
                list(labels_x), default=[], key="mail_exam_students")
            subset_x = (edf_mail[edf_mail.Roll.isin([labels_x[c] for c in chosen_x])]
                        if chosen_x else edf_mail)

            bad = []
            for r in subset_x.itertuples():
                try:
                    mails.append(mailer.build_exam_mail(
                        r.Roll,
                        name=("student" if generic_x else (r.Name if names_x else "")),
                        room=(str(r.Room) if "Room" in subset_x.columns else ""),
                        block=str(r.Block),
                        seat=(str(r.Seat) if "Seat" in subset_x.columns else ""),
                        exam=exam_name, when=when_x, course=course_x, session=session_x,
                        include_seat=include_seat, note=note_x, contact=contact_x))
                except ValueError as e:
                    bad.append(f"{r.Roll}: {e}")
            if bad:
                st.error("Skipped - " + " · ".join(bad))

    elif mail_mode == "Exam block/room (one mail per block)":
        group_windows = False
        st.caption(
            "One mail per block, in every room. **Students in To, instructors in Cc, "
            "that block's TAs in Bcc.** The mail names the room and block, which is what "
            "a student needs to walk in and sit down; the seat itself is on the block "
            "sheet at the door. A hall numbered by row and column (LHC 105) has no blocks, "
            "so it gets one mail for the room."
        )
        exams_b = exam_plan.list_exams()
        if not exams_b:
            st.info("No saved exam yet - allocate one on the **🧑‍🎓 Exam seating** page.")
        else:
            labels_b = {f"{m['title']} · {m.get('when') or m.get('saved', '')}": m
                        for m in exams_b}
            pick_b = st.selectbox("Exam", list(labels_b), key="mail_exam_block_pick")
            meta_b = labels_b[pick_b]
            dir_b = Path(meta_b["dir"])
            alloc_b = pd.read_csv(dir_b / exam_plan.ALLOC_FILE, dtype={"Roll": str})

            # TAs per block come from the saved posting, if the pack was built
            duty_path = dir_b / "ta_assignment.csv"
            duty_b = pd.read_csv(duty_path) if duty_path.exists() else pd.DataFrame()

            # One mail per block wherever the allocation has blocks, in every
            # room. The mail's job is to tell a student their room and block,
            # and a room-wide mail cannot do that for a hall of ten blocks.
            # Printing stays room-wise for the small halls; that is about
            # paper, not about what a student needs to be told.
            groups_b = []
            for room in dict.fromkeys(alloc_b.Room):
                sub_room = alloc_b[alloc_b.Room == room]
                blocks_here = [b for b in sorted(set(sub_room.Block.astype(str))) if b.strip()]
                # a hall numbered by row and column has no blocks to mail: one
                # mail names the room, and the row is on the plan at the door
                if exam_rooms.is_row_col_room(room):
                    groups_b.append((room, "", sub_room))
                elif blocks_here:
                    for block in blocks_here:
                        groups_b.append((room, block,
                                         sub_room[sub_room.Block.astype(str) == block]))
                else:
                    groups_b.append((room, "", sub_room))

            def tas_for(room, block):
                if duty_b.empty:
                    return []
                rows = duty_b[(duty_b.Room == room)
                              & ((duty_b.Block == block) if block
                                 else (duty_b.Block == ta_duty.WHOLE_ROOM))]
                if rows.empty:
                    rows = duty_b[duty_b.Room == room]
                return [r.Roll for r in rows.itertuples() if isinstance(r.Roll, str) and r.Roll]

            bc1, bc2 = st.columns(2)
            cc_b = bc1.text_input("Cc (instructors)", mailer.DEFAULT_ATTENDANCE_CC,
                                  key="mail_block_cc")
            url_b = bc2.text_input("Find-my-seat URL (optional)", C.LOOKUP_URL, key="mail_block_url",
                                   placeholder="https://example.edu/find-my-seat.html",
                                   help="Where a student looks up their own seat number.")
            seats_b = st.checkbox(
                "Also list each student's seat number", value=True,
                key="mail_block_seats",
                help="On by default: block-wise rooms (LHC 110, LHC 308) give every "
                     "student a real numbered seat, so the mail lists each roll number's "
                     "own seat rather than sending them to the block sheet at the door "
                     "for it. Turn off to name just the room and block.")
            students_in_b = st.radio(
                "Students go in", ["To (they can see it's their block)", "Bcc (addresses hidden)"],
                horizontal=True, key="mail_block_students_in",
                help="To is what was asked for. It does mean every student in the block can "
                     "see the others' addresses and can reply to all of them.")
            labels_pick = [f"{r} · {exam_rooms.group_label(r, b)}" if b
                           else (f"{r} · whole room (rows and columns)"
                                 if exam_rooms.is_row_col_room(r) else f"{r} · whole room")
                           for r, b, _ in groups_b]
            chosen_b = st.multiselect("Blocks to mail - leave empty for all",
                                      labels_pick, default=[], key="mail_block_which")
            wanted = set(chosen_b) if chosen_b else set(labels_pick)

            bad = []
            for (room, block, sub_g), label in zip(groups_b, labels_pick):
                if label not in wanted:
                    continue
                ta_rolls = tas_for(room, block)
                # the seat label carries the block where one mail covers several,
                # since seat 1 exists in every block of a room
                many_blocks = sub_g.Block.nunique() > 1
                if exam_rooms.is_row_col_room(room):
                    seat_map_b = ({str(r.Roll): str(r.Seat)
                                   for r in sub_g.itertuples()} if seats_b else None)
                else:
                    seat_map_b = ({str(r.Roll): (f"{r.Block}-{r.Seat}" if many_blocks
                                                 else str(r.Seat))
                                   for r in sub_g.itertuples()} if seats_b else None)
                try:
                    mails.append(mailer.build_exam_block_mail(
                        list(sub_g.Roll), exam=meta_b.get("title", "Exam"),
                        when=meta_b.get("when", ""), room=room, block=block,
                        course=meta_b.get("course", C.COURSE),
                        cc=cc_b.strip(), lookup_url=url_b.strip(),
                        bcc=",".join(mailer.ta_email(r) for r in ta_rolls),
                        students_in="to" if students_in_b.startswith("To") else "bcc",
                        seats=seat_map_b,
                        seat_note="Your seat number is on the block sheet posted at the "
                                  "room door and on the seating plan inside."))
                except ValueError as e:
                    bad.append(f"{label}: {e}")
            if bad:
                st.error("Skipped - " + " · ".join(bad))
            if mails:
                st.caption(f"{len(mails)} mail(s) · "
                           + ", ".join(f"{m.meta['label']} ({m.meta['n']})" for m in mails[:6])
                           + ("…" if len(mails) > 6 else "")
                           + (" · TAs in Bcc from the saved posting"
                              if not duty_b.empty else
                              " · no `ta_assignment.csv` for this exam, so no TAs in Bcc"))

    else:   # Attendance shortfall
        group_windows = False
        st.caption(
            "Reads the Institute's master **Class Attendance** workbook - the same file "
            "the 📊 summary uses - and writes in the course's agreed wording. **Two "
            "templates:** a student who can still reach the requirement is told how many "
            "classes in a row get them back; one who cannot, however many they attend, is "
            "told the highest figure still open to them and asked to meet the instructors. "
            "Excused absences count towards neither. No block or seat is mentioned."
        )

        att_files = sorted(str(q) for q in Path(".").glob("*Class Attendance*.xls*")
                           if "backup" not in q.stem.lower())
        if not att_files:
            st.error("No `*Class Attendance*.xlsx` found in the project folder.")
        else:
            ac1, ac2 = st.columns(2)
            att_pick_m = ac1.selectbox("Attendance workbook", att_files, key="mail_att_file")
            try:
                sheets_m = attrep.register_sheets(att_pick_m)
            except Exception as e:                    # noqa: BLE001 - surfaced as-is
                st.error(f"Could not read that workbook: {e}")
                sheets_m = []
            cohort_m = attrep.cohort_sheets(att_pick_m) if sheets_m else []
            sheet_m = ac2.multiselect("Batch sheet(s)", sheets_m,
                                      default=cohort_m or sheets_m[:1],
                                      key="mail_att_sheets",
                                      help="The two batch sheets by default. A combined "
                                           "sheet holds the same students again, so a "
                                           "student on both would otherwise be written to "
                                           "twice; duplicates are dropped below.")

            escalate = st.checkbox(
                "Escalation mode - thresholds "
                + ", ".join(f"{lim}+ ({name})" for lim, name in mailer.ATTENDANCE_LEVELS)
                + ", each student written to once per level",
                value=False, key="mail_att_escalate",
                help="Off: everyone over the single limit below, whether or not they've been "
                     "written to before. On: only students who have reached a level they "
                     "haven't been mailed at yet - someone who slips from one level to the "
                     "next reappears, someone sitting still does not. The level is recorded "
                     "in the mail log, and the second and third notices add one line saying "
                     "which notice it is.")

            ac3, ac4 = st.columns(2)
            pick_by = ac3.radio("Who gets a notice", ["Below the requirement", "Missed more than"],
                                horizontal=True, key="mail_att_pickby", disabled=escalate,
                                help="The notices are written against the attendance "
                                     "requirement, so selecting by percentage is the one "
                                     "that matches what they say. The absence count is "
                                     "kept for a plain \"you have missed N classes\" sweep.")
            if pick_by.startswith("Below"):
                threshold_a = ac3.number_input("Requirement (%)", min_value=1.0,
                                               max_value=100.0, step=1.0,
                                               value=float(mailer.ATTENDANCE_THRESHOLD),
                                               key="mail_att_threshold", disabled=escalate)
                limit = 0
            else:
                threshold_a = float(mailer.ATTENDANCE_THRESHOLD)
                limit = ac3.number_input("Write to students who missed MORE than", min_value=1,
                                         value=4, step=1, key="mail_att_limit",
                                         disabled=escalate,
                                         help="Counted as `N` marks in the workbook.")
            instructor_cc = ac4.text_input("Cc (course instructors)",
                                           mailer.DEFAULT_ATTENDANCE_CC,
                                           key="mail_att_cc",
                                           help="Copied on every one of these mails. "
                                                "Comma-separate for more than one.")

            ac5, ac6 = st.columns(2)
            course_a = ac5.text_input("Course", C.COURSE, key="mail_att_course",
                                      help="Named in the body and its first word is the "
                                           "subject line's prefix.")
            ac6.caption("Signed off by:")
            ac6.code(mailer.ATTENDANCE_SIGNATURE, language="text")

            st.caption(
                "**Two templates, picked by arithmetic.** A student who can still reach "
                "the requirement is told how many classes in a row get them back; one who "
                "cannot is told the highest figure still open to them and asked to meet "
                "the instructors. That split needs to know how many classes are still to "
                "be held."
            )
            if not sheet_m:
                st.info("Pick at least one batch sheet.")
            else:
                frames, sessions_m = [], None
                for sh in sheet_m:
                    try:
                        stu_m, sess_m = attrep.load(att_pick_m, sh)
                    except (ValueError, KeyError) as e:
                        st.error(f"{sh}: {e}")
                        continue
                    stu_m["Batch"] = sh
                    frames.append(stu_m)
                    if sessions_m is None:
                        sessions_m = sess_m

                # the dated columns the register already carries are the course's
                # own calendar, so "how many classes remain" is read, not typed
                scheduled_m = len(sessions_m) if sessions_m is not None else 0
                marked_m = (sessions_m[(sessions_m.Present + sessions_m.Absent
                                        + sessions_m.Excused) > 0]
                            if sessions_m is not None else None)
                held_reg = len(marked_m) if marked_m is not None else 0
                asof_default = ""
                if marked_m is not None and len(marked_m):
                    last_d = marked_m.Date.max()
                    asof_default = (last_d.strftime("%d/%m/%Y")
                                    if hasattr(last_d, "strftime") else str(last_d))

                ac7, ac8, ac9 = st.columns(3)
                asof_a = ac7.text_input("Figures as of", asof_default, key="mail_att_asof",
                                        help="The last class with marks against it, as the "
                                             "mail should print it.")
                remaining_a = ac8.number_input(
                    "Classes still to be held", min_value=0, step=1,
                    value=max(0, scheduled_m - held_reg), key="mail_att_remaining",
                    help=f"The register has {scheduled_m} dated columns and {held_reg} of "
                         "them are marked. This decides which template each student gets, "
                         "so correct it if the calendar has changed.")
                meet_by_a = ac9.text_input(
                    "Meet the instructors by", "", key="mail_att_meetby",
                    placeholder="e.g. 3 October 2026",
                    help="Printed in the second template only. Left blank it reads "
                         "\"as soon as possible\".")
                consequence_a = st.text_input(
                    "Consequence line (optional, second template only)", "",
                    key="mail_att_consequence",
                    placeholder="per institute or course policy",
                    help="One sentence, printed only in the notice to students who can no "
                         "longer reach the requirement. Left blank, nothing is said - "
                         "better silent than inventing a policy.")
                note_a = st.text_area(
                    "Extra paragraph (optional)", "", key="mail_att_note",
                    help="Inserted before 'Regards'. The rest of the wording is the "
                         "course's agreed text and is not edited here - see "
                         "`mailer.build_attendance_mail`.")
                contact_a = st.text_input("Contact line (optional)", mailer.DEFAULT_CONTACT,
                                          key="mail_att_contact")
                if frames:
                    allstu = pd.concat(frames, ignore_index=True)
                    # a combined sheet repeats the batch sheets' students; one
                    # mail per student, whichever sheets were picked
                    dropped = int(allstu.Roll.duplicated().sum())
                    if dropped:
                        allstu = allstu.drop_duplicates("Roll").reset_index(drop=True)
                        st.caption(f"{dropped} duplicate row(s) dropped - the sheets picked "
                                   "overlap, and nobody is written to twice.")
                    held_max = int(allstu.Held.max()) if len(allstu) else 0
                    if escalate:
                        over = mailer.unmailed_attendance(allstu).sort_values(
                            "Absent", ascending=False).reset_index(drop=True)
                        by_level = over.Level.value_counts().to_dict() if len(over) else {}
                        st.caption(f"{len(allstu)} students · up to {held_max} classes held · "
                                   f"**{len(over)}** due a notice they haven't had yet"
                                   + (f" ({', '.join(f'{k}: {v}' for k, v in sorted(by_level.items()))})"
                                      if by_level else ""))
                    elif pick_by.startswith("Below"):
                        seen_a = allstu[allstu.Held > 0]
                        over = seen_a[seen_a.Percent < float(threshold_a)].sort_values(
                            "Percent").reset_index(drop=True)
                        over["Level"] = ""
                        st.caption(f"{len(allstu)} students · up to {held_max} classes held "
                                   f"so far · **{len(over)}** below "
                                   f"{mailer._pct(threshold_a)}%.")
                    else:
                        over = allstu[allstu.Absent > int(limit)].sort_values(
                            "Absent", ascending=False).reset_index(drop=True)
                        over["Level"] = ""
                        st.caption(f"{len(allstu)} students · up to {held_max} classes held "
                                   f"so far · **{len(over)}** have missed more than "
                                   f"{int(limit)}.")
                    if over.empty:
                        st.success("Nobody is over the limit.")
                    else:
                        # which template each student gets, and the number that
                        # template turns on - shown before anything is composed
                        over["Needs"] = [
                            mailer.classes_to_recover(int(r.Present), int(r.Absent),
                                                      float(threshold_a))
                            for r in over.itertuples()]
                        over["Best possible"] = [
                            mailer.best_possible_percent(int(r.Present), int(r.Absent),
                                                         int(remaining_a))
                            for r in over.itertuples()]
                        over["Notice"] = [
                            "can recover" if mailer.can_still_reach(
                                int(r.Present), int(r.Absent), int(remaining_a),
                                float(threshold_a)) else "cannot reach"
                            for r in over.itertuples()]
                        n_lost = int((over.Notice == "cannot reach").sum())
                        st.caption(
                            f"**{len(over) - n_lost}** can still reach "
                            f"{mailer._pct(threshold_a)}% with {int(remaining_a)} classes "
                            f"left and get the first template; **{n_lost}** cannot and get "
                            "the second, which asks them to meet the instructors.")
                        cols_over = (["Batch", "Roll", "Name", "Level"] if escalate
                                     else ["Batch", "Roll", "Name"]) + \
                            ["Present", "Absent", "Excused", "Held", "Percent",
                             "Needs", "Best possible", "Notice"]
                        st.dataframe(over[cols_over].style.format(
                                               {"Percent": "{:.1f}%",
                                                "Best possible": "{:.1f}%"}, na_rep="-"),
                                     width="stretch", hide_index=True)
                        labels_a = {f"{r.Roll} - {r.Name} ({mailer._pct(r.Percent)}%, "
                                    f"{r.Absent} missed)": i
                                    for i, r in enumerate(over.itertuples())}
                        chosen_a = st.multiselect(
                            "Students to mail - leave empty to mail all of them",
                            list(labels_a), default=[], key="mail_att_students")
                        if chosen_a:
                            over = over.iloc[[labels_a[c] for c in chosen_a]]

                        bad = []
                        for r in over.itertuples():
                            try:
                                mails.append(mailer.build_attendance_mail(
                                    r.Roll, name=r.Name, absent=int(r.Absent),
                                    held=int(r.Held), present=int(r.Present),
                                    percent=(None if pd.isna(r.Percent) else float(r.Percent)),
                                    remaining=int(remaining_a), asof=asof_a.strip(),
                                    threshold=float(threshold_a),
                                    meet_by=meet_by_a.strip(),
                                    consequence=consequence_a.strip(),
                                    missed_limit=int(limit),
                                    level=(str(r.Level) if escalate else ""),
                                    cc=instructor_cc.strip(),
                                    course=course_a, medium=r.Batch,
                                    note=note_a, contact=contact_a))
                            except ValueError as e:
                                bad.append(f"{r.Roll}: {e}")
                        if bad:
                            st.error("Skipped - " + " · ".join(bad))

    if mails:
        st.divider()
        mail_send_ui(mails, key="bulk_mail", group=group_windows,
                     intro=f"{len(mails)} message(s) composed. Nothing reaches a student "
                           "until you press Send in Gmail - or untick **Dry run** under "
                           "the SMTP option.")

    if Path(mailer.MAIL_LOG).exists():
        with st.expander("Send log - what students have already been told"):
            log_df = pd.read_csv(mailer.MAIL_LOG)
            st.dataframe(log_df.tail(200).iloc[::-1], width="stretch", hide_index=True)
            st.download_button("Download mail log (.csv)",
                               Path(mailer.MAIL_LOG).read_bytes(),
                               file_name="mail_log.csv", key="mail_log_dl")

# ───────────────────────── Posters & packs ─────────────────────────
if PAGE == "Posters & packs":
    page_header("Posters & packs", "A poster for the hall door, an invigilator pack for the exam room.")
    st.caption(
        "A poster is what a student reads while standing in the doorway: the lookup "
        "URL, a QR to the same, and the block map. An invigilator pack is what the "
        "classroom side has always had and the exam side hasn't - the seat grid with "
        "roll numbers in position, plus a signature list per room."
    )

    st.markdown("#### Hall-door poster")
    found_p = discover_allocations()
    if not found_p:
        st.info("No `allocation.csv` yet - allocate seats on page **🪑 1 · Allocate seats** first.")
    else:
        pc1, pc2 = st.columns(2)
        cohort_p = pc1.selectbox("Cohort", list(found_p), key="poster_cohort")
        alloc_path_p = Path(found_p[cohort_p])
        alloc_p = pd.read_csv(alloc_path_p, dtype={"Roll": str})
        room_guess_p = next((q.stem[: -len("_seating")]
                             for q in alloc_path_p.parent.glob("*_seating.xlsx")
                             if q.stem.endswith("_seating") and q.stem[: -len("_seating")] in C.ROOMS),
                            list(C.ROOMS)[0])
        room_p = pc2.selectbox("Room", list(C.ROOMS), index=list(C.ROOMS).index(room_guess_p),
                               key="poster_room")
        url_p = st.text_input(
            "URL of the find-your-block page", C.LOOKUP_URL or "https://example.edu/find-your-block.html",
            key="poster_url",
            help="Whatever students will actually type or scan. The QR encodes exactly this, "
                 "so a placeholder here prints a placeholder on the wall.")
        pc3, pc4 = st.columns(2)
        course_p = pc3.text_input("Course", C.COURSE, key="poster_course")
        session_p = pc4.text_input("Session", C.SESSION, key="poster_session")
        note_p = st.text_input("Footer note (optional)", "", key="poster_note")

        if st.button("Build hall poster", type="primary", key="poster_run"):
            out_p = f"poster_{room_p.replace(' ', '_')}.pdf"
            snapshot_before([out_p], "hall-poster")
            try:
                posters.hall_poster(out_p, room=room_p, url=url_p,
                                    cohort=alloc_path_p.parent.name.title() + " Batch",
                                    course=course_p, session=session_p,
                                    counts=alloc_p.Block.value_counts().to_dict(), note=note_p)
            except Exception as e:                    # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            if posters._qr_image("x") is None:
                st.warning("`qrcode` isn't installed, so the poster printed without a QR code - "
                           "the URL is still there. `pip install qrcode` to get it.")
            st.success(f"Built `{out_p}` - {len(alloc_p)} students, {room_p}.")
            st.download_button("Download poster", Path(out_p).read_bytes(),
                               file_name=out_p, key="poster_dl")

    st.divider()
    st.markdown("#### Where to report (exam) - one page")
    st.caption(
        "The single sheet a student needs: roll-number ranges against room and block, "
        "no seats and no names. For the noticeboard, the department group and the mail. "
        "A group that is not one unbroken stretch of roll numbers prints each stretch, "
        "and the seniors' block is described in words rather than as a range that spans "
        "four batches."
    )
    exams_s = exam_plan.list_exams()
    if not exams_s:
        st.info("No saved exam yet. Allocate one on the **Exam seating** page.")
    else:
        labels_s = {f"{m['title']} · {m.get('when') or m.get('saved', '')}": m for m in exams_s}
        pick_s = st.selectbox("Exam", list(labels_s), key="poster_summary_exam")
        meta_s = labels_s[pick_s]
        alloc_s = pd.read_csv(Path(meta_s["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        sc1, sc2 = st.columns(2)
        by_s = sc1.text_input("Reporting line", "Reach your room at least ten minutes "
                                                "before the start time.", key="poster_by")
        url_s = sc2.text_input("Find-my-seat URL (optional)", "", key="poster_summary_url",
                               placeholder="https://…/find-my-seat.html")
        late_s = st.text_input(
            "Late or lost line",
            "If you are late or cannot find your room, report to the front of LHC 110 - "
            "seats are kept free there and a TA will place you.", key="poster_late")
        st.dataframe(pd.DataFrame([
            {"Room": r["room"], "Block": r["block"] or "whole room",
             "Roll numbers": ", ".join(posters._range_text(x) for x in r["runs"]),
             "Students": r["n"]}
            for r in posters.roll_ranges(alloc_s)]),
            width="stretch", hide_index=True)
        if st.button("Build the one-page summary", type="primary", key="poster_summary_run"):
            out_s = Path(meta_s["dir"]) / f"{meta_s['slug']}_where_to_report.pdf"
            snapshot_before([out_s], f"summary-{meta_s['slug']}")
            try:
                posters.report_summary(
                    alloc_s, str(out_s), exam=meta_s.get("title", "Exam"),
                    course=meta_s.get("course", C.COURSE),
                    when=meta_s.get("when", ""), report_by=by_s.strip(),
                    late_note=late_s.strip(), lookup_url=url_s.strip())
            except Exception as e:                    # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            st.success(f"Built `{out_s}` - one A4 page, {len(alloc_s)} students.")
            st.download_button("Download the summary", Path(out_s).read_bytes(),
                               file_name=out_s.name, key="poster_summary_dl")

    st.divider()
    st.markdown("#### Notice-board seat map (exam)")
    st.caption("One sheet per room, the whole room drawn as it is laid out, for the board "
               "outside the hall: **A3 for the big halls (LHC 110, 308, 105), A4 for the "
               "smaller rooms**, since a forty-student room on A3 is harder to read, not "
               "easier. Pick **A4 only** for a printer with no A3: a big hall is then split "
               "across numbered A4 sheets at the same size, to be taped up side by side. "
               "Different object from the grid in the pack: that one is held, this one is "
               "read from arm's length with people standing in front of it.")
    exams_n = exam_plan.list_exams()
    if not exams_n:
        st.info("No saved exam yet. Allocate one on the **Exam seating** page.")
    else:
        labels_n = {f"{m['title']} · {m.get('when') or m.get('saved', '')}": m for m in exams_n}
        pick_n = st.selectbox("Exam", list(labels_n), key="poster_notice_exam")
        meta_n = labels_n[pick_n]
        alloc_n = pd.read_csv(Path(meta_n["dir"]) / exam_plan.ALLOC_FILE, dtype={"Roll": str})
        rooms_n = sorted(set(alloc_n.Room)) if "Room" in alloc_n else []
        st.caption(f"{len(alloc_n)} students · {', '.join(rooms_n)}")
        paper_n = st.radio(
            "Paper", ["Legal size (14×8.5 in)", "A3 where it helps", "A4 only"], horizontal=True,
            key="poster_notice_paper",
            help="Legal size (14×8.5 in) or A3 for big halls, A4 only splits across sheets.")
        if st.button("Build notice-board map", type="primary", key="poster_notice_run"):
            out_n = Path(meta_n["dir"]) / f"{meta_n['slug']}_notice_board.pdf"
            snapshot_before([out_n], f"notice-{meta_n['slug']}")
            paper_choice = "legal" if paper_n.startswith("Legal") else ("a4" if paper_n.startswith("A4") else "auto")
            try:
                posters.notice_board_map(
                    alloc_n, str(out_n), room_map=exam_rooms.load_exam_rooms(),
                    exam=meta_n.get("title", "Exam"),
                    course=meta_n.get("course", C.COURSE),
                    when=meta_n.get("when", ""), venue=meta_n.get("venue", ""),
                    paper=paper_choice)
            except Exception as e:                    # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            a3 = [r for r in rooms_n if r in posters.A3_ROOMS]
            if paper_n.startswith("A4"):
                try:
                    from pypdf import PdfReader
                    sheets = len(PdfReader(str(out_n)).pages)
                except Exception:                     # noqa: BLE001 - count is a nicety
                    sheets = 0
                st.success(f"Built `{out_n}` - A4 throughout"
                           + (f", {sheets} sheet(s)" if sheets else "")
                           + f" for {len(rooms_n)} room(s); a big hall is split across "
                             "sheets marked 'sheet n of m', to be taped up side by side. "
                             "Print at 100%: some sheets are landscape.")
            else:
                st.success(f"Built `{out_n}` - one page per room ({len(rooms_n)} room(s)"
                           + (f", A3 for {', '.join(a3)}" if a3 else "") + "). Print at 100%: "
                           "the page size is set per room, so 'fit to page' undoes it.")
            st.download_button("Download notice-board map", Path(out_n).read_bytes(),
                               file_name=out_n.name, key="poster_notice_dl")

    st.divider()
    st.markdown("#### Invigilator pack (exam)")
    last_exam_p = st.session_state.get("exam_last")
    src_p = st.radio("Seating source",
                     (["The allocation from the Exam seating page"] if last_exam_p else [])
                     + ["Upload a roll list (.xlsx/.csv with Roll, Room, Block, Seat)"],
                     key="pack_src")
    edf_p = None
    if last_exam_p and src_p.startswith("The allocation"):
        edf_p = last_exam_p["alloc"]
        st.caption(f"{len(edf_p)} students · {last_exam_p['venue']}")
    else:
        up_p = st.file_uploader("Exam roll list", type=["xlsx", "xls", "csv"], key="pack_upload")
        if up_p:
            edf_p = (pd.read_csv(up_p, dtype={"Roll": str}) if up_p.name.endswith(".csv")
                     else pd.read_excel(up_p, dtype={"Roll": str}))
            missing_p = [c for c in ("Roll", "Block", "Seat") if c not in edf_p.columns]
            if missing_p:
                st.error("Missing column(s): " + ", ".join(missing_p))
                edf_p = None

    if edf_p is None:
        st.info("Allocate an exam on the **🧑‍🎓 Exam seating** page, or upload a roll list here.")
    else:
        kc1, kc2 = st.columns(2)
        exam_p = kc1.text_input("Exam title", (last_exam_p or {}).get("exam", "Quiz 1"),
                                key="pack_exam")
        when_p = kc2.text_input("Date & time", (last_exam_p or {}).get("when", ""), key="pack_when")
        out_pack = st.text_input("Output file", "invigilator_pack.pdf", key="pack_out")
        if st.button("Build invigilator pack", type="primary", key="pack_run"):
            snapshot_before([out_pack], "invigilator-pack")
            try:
                posters.invigilator_pack(edf_p, out_pack, exam=exam_p,
                                         course=(last_exam_p or {}).get("course", C.COURSE),
                                         when=when_p, venue=(last_exam_p or {}).get("venue", ""))
            except Exception as e:                    # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
                st.stop()
            rooms_n = edf_p.Room.nunique() if "Room" in edf_p else 1
            st.success(f"Built `{out_pack}` - {len(edf_p)} students across {rooms_n} room(s): "
                       "seat grid then signature list, per room.")
            st.download_button("Download pack", Path(out_pack).read_bytes(),
                               file_name=out_pack, key="pack_dl")

# ───────────────────────── Health check ─────────────────────────
if PAGE == "Health check":
    page_header("Health check", "Every invariant this toolkit depends on, checked in one pass - plus undo.")
    st.caption(
        "Runs every invariant this toolkit depends on - roll sets against the register, "
        "seat numbers contiguous, printed sheets describing the allocation they claim to, "
        "published files newer than what they were built from. Two real bugs in this "
        "project's history would have shown up here on their own: a counter reading zero "
        "classes while nine had been held, and one cohort's sheets rebuilt from the other "
        "cohort's allocation."
    )

    if st.button("Run checks", type="primary", key="hc_run") or st.session_state.get("hc_df") is None:
        with st.spinner("Checking…"):
            try:
                st.session_state["hc_df"] = checks.run_all()
            except Exception as e:                    # noqa: BLE001
                st.error(f"The checker itself failed: {type(e).__name__}: {e}")
                st.session_state["hc_df"] = pd.DataFrame()

    hc = st.session_state.get("hc_df")
    if hc is not None and len(hc):
        n_ok = int((hc.status == "ok").sum())
        n_warn = int((hc.status == "warn").sum())
        n_fail = int((hc.status == "fail").sum())
        m1, m2, m3 = st.columns(3)
        m1.metric("Passing", n_ok)
        m2.metric("Warnings", n_warn)
        m3.metric("Failures", n_fail)
        if n_fail:
            st.error(f"{n_fail} check(s) failed - someone could act on something wrong.")
        elif n_warn:
            st.warning(f"{n_warn} warning(s) - true but stale, or unverifiable from here.")
        else:
            st.success("Everything consistent.")

        icon = {"ok": "✅", "warn": "⚠️", "fail": "❌"}
        show = hc.copy()
        show.insert(0, "", show.status.map(icon))
        only_bad = st.checkbox("Only show what isn't passing", value=bool(n_warn or n_fail),
                               key="hc_only_bad")
        if only_bad:
            show = show[show.status != "ok"]
        st.dataframe(
            show[["", "group", "check", "detail", "fix"]],
            width="stretch", hide_index=True,
            column_config={
                "": st.column_config.TextColumn("", width="small"),
                "group": st.column_config.TextColumn("Where", width="small"),
                "check": st.column_config.TextColumn("Check", width="medium"),
                "detail": st.column_config.TextColumn("What was found", width="large"),
                "fix": st.column_config.TextColumn("What to do", width="medium"),
            })

    st.divider()
    st.markdown("#### Block changes made outside the toolkit")
    st.caption("An edit made by hand in Excel or straight in `allocation.csv` leaves no row in "
               "`transitions.csv`, so those students never appear in “who still needs telling”. "
               "This compares each allocation against the copy kept from the last time the "
               "toolkit wrote it.")
    for cdir in cohort_dirs():
        if not (cdir / "allocation.csv").exists():
            continue
        drift = seating.detect_untracked_moves(cdir.name)
        base = history.state_path(cdir.name)
        if not base.exists():
            st.info(f"**{cdir.name.title()}** - no baseline yet. "
                    f"Press the button below to record the current file as the reference point.")
        elif drift.empty:
            st.success(f"**{cdir.name.title()}** - no unrecorded block changes.")
        else:
            st.warning(f"**{cdir.name.title()}** - {len(drift)} block change(s) not in the log:")
            st.dataframe(drift[["roll", "name", "from_block", "to_block"]],
                         width="stretch", hide_index=True)
        if st.button(f"Adopt & re-baseline - {cdir.name.title()}", key=f"hc_adopt_{cdir.name}"):
            n = seating.adopt_untracked_moves(cdir.name)
            st.success(f"Recorded {n} move(s) in `{seating.TRANSITION_LOG}` and re-baselined. "
                       + ("They can now be mailed from **📧 Email students**." if n else ""))
            st.rerun()

    st.divider()
    st.markdown("#### Snapshots - undo a bad build")
    st.caption(f"Every build and move saves what it is about to overwrite into "
               f"`{history.HISTORY_DIR}` first. The newest {history.KEEP} are kept "
               f"({history.total_bytes() / 1e6:.1f} MB now); older ones are pruned "
               "automatically, and the folder is git-ignored so it never clutters the project.")
    snaps = history.list_snapshots()
    if not snaps:
        st.info("Nothing saved yet - snapshots start appearing the first time you build "
                "or move someone from here.")
    else:
        labels = {f"{sn['created'] or sn['name']} · {sn['label']} · {len(sn['files'])} file(s)": sn
                  for sn in snaps}
        pick_s = st.selectbox("Snapshot", list(labels), key="hc_snap")
        sn = labels[pick_s]
        if not sn["ok"]:
            st.error("This snapshot can't be read - leave it alone and use an older one.")
        else:
            if sn["note"]:
                st.caption(sn["note"])
            st.code("\n".join(sn["files"]), language="text")
            st.warning("Restoring overwrites those files with the saved copies. What's there now "
                       "is snapshotted first, so this is itself undoable.")
            if st.checkbox("I want to restore these files", key="hc_restore_confirm"):
                if st.button("Restore snapshot", type="primary", key="hc_restore"):
                    done = history.restore(sn["path"])
                    st.success(f"Restored {len(done)} file(s). Re-run the checks above.")
                    st.session_state["hc_df"] = None
