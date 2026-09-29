# AGENTS.md - course attendance & seating toolkit

Context for any AI coding CLI (Claude Code, Cursor, Aider, Codex CLI, Gemini
CLI, etc.) picking this project up - including on a fresh machine after a
transfer. Read this before making changes; it captures conventions and
gotchas that aren't obvious from the code alone, several of them found the
hard way over the course of this project's development.

> **Any course.** Everything course-specific (code, title, session,
> institute, mail domain, workbook names, instructors) is one block at the top
> of `config.py`; `<CODE>` below means `config.COURSE_CODE`. This file was
> written while running the toolkit for a real course and is kept because it
> explains *why* the code is the way it is. Every roll number, name and count
> in it refers to the invented dataset built by `make_sample_data.py`.

## The workbook's combined sheet

The Class Attendance workbook has three sheets: `English`, `Hindi` and
`Combined`. **Combined is a view, not a third cohort**: every student appears
on it *and* on their own batch sheet, with a `Batch` column saying which. Adding
it to the other two counts everyone twice, which is exactly what
`seating.load_attendance_roster()` used to do the moment the sheet appeared (it
raised on duplicate roll numbers).

So: `attendance_report.cohort_sheets(path)` is what to sum over,
`combined_sheets(path)` is the other one, and `is_combined(path, sheet)` decides
by looking for a `Batch` header rather than by sheet name.
`load_attendance_roster()` with no sheet now reads the cohort sheets only; asked
for the combined sheet by name it reads it, taking each student's `Medium` from
the `Batch` column instead of from the sheet name. The dashboard defaults every
sheet picker to the cohort sheets, warns if both views are picked at once, and
drops duplicate rolls before composing attendance mail.

The **attendance summary and defaulter list are read from the combined sheet**:
one table for the whole course, with the batch in a column rather than in the
sheet name (`attendance_report.load` returns `Batch`, blank on a per-medium
sheet). The dashboard's Attendance summary page defaults to it where the
workbook has one, and the rebuilt files live in `out/attendance/`
(`attendance_summary.csv`, `defaulters.csv`, and an `.xlsx` with Summary,
Below 75% and Per session sheets).

A sheet counts as a register only if its label row carries `Roll Number`
(`attendance_report.is_register`); `register_sheets()` is what a sheet picker
should offer. The workbook has since grown a `Dashboard` sheet of live figures,
and without that test it was offered as a third batch to read and mail from.

**The `Dashboard` sheet is the course's own live view** of the same marks:
threshold, total classes, class benchmark, a per-batch summary, attendance
bands, and the list below the threshold. Nothing is *read from* it for a mail or
a seat, but `attendance_report.dashboard_blocks()` parses its Settings, Summary
and Bands blocks so the Attendance summary page takes the course's threshold,
benchmark and course length from there instead of keeping a second copy, and
shows the same per-batch summary and bands recomputed from the register beside
them. The **report PDF** (`attendance_pdf`, Attendance summary page) prints the same
figures for handing over: four headline numbers, a batch comparison, attendance
bands, observations generated from the data rather than written, a class-by-class
page marking anything under the class benchmark, the below-threshold list
colour-coded amber (can still recover) or purple (cannot), and a method page.
Nothing in it is calculated twice: `classes_to_recover`, `best_possible_percent`
and `can_still_reach` are `mailer`'s, the same functions that decide which
template a student's mail uses, so the report and the mail can never disagree.
Sections can be switched off; a dated column with no marks is reported as not
held rather than quietly skipped.

`checks.check_dashboard_sheet` warns (never fails) when the two disagree:
Excel caches formula results, so a sheet saved without a recalculation shows
yesterday's counts, and that sheet is what the course reads. The register is the
authority either way.

`checks.check_combined_sheet` is the reason to keep the sheet honest: it fails
if the roll sets differ, if a `Batch` label disagrees with the sheet a student
is actually on, or if any student's Y/N/E counts differ between the two copies.
The batch sheets are the register; the combined sheet is the copy.


## What this is

**`Class Attendance <CODE>.xlsx` is the primary roster.** The
department keeps it current - one sheet per medium (English/Hindi), `Roll
Number` + `Name` under a header on row 3 - so it is right about who is in
which batch even when a `Section & Group` export has gone stale after batch
switches. `seating.load_attendance_roster()` reads it **in the workbook's own
row order**, and seats follow that order within each block, so reading a
signature sheet top-to-bottom reads the attendance register top-to-bottom.
Tab 1 of the dashboard defaults to it. `seating.reorder_within_blocks()`
re-sorts an existing allocation into that order without moving anyone between
blocks or changing which seats are occupied.

A pipeline + Streamlit dashboard for one course (set in `config.py`; the demo
is `DEMO101`) across two
mediums (English in room LHC110, Hindi in room LHC2 101): allocate students
to seating blocks → print signature sheets for them to sign. Attendance is
recorded in the Institute's master Class Attendance workbook and summarised
here. Plus: exam seating for a separate multi-room
exam-venue model, a public "find your block" self-lookup page, an
attendance summary/defaulter view, and a seat-allotment PDF for students.

No CLAUDE.md-specific instructions exist beyond this file; this file *is*
the project-level guidance, tool-agnostic on purpose.

## Where this lives

This folder is synced via **Dropbox** between a Windows machine and a Linux
machine, so both sides see the same files without a manual transfer step -
edit on either side and it propagates once Dropbox catches up.

**Don't run the dashboard on both machines at the same time.** Dropbox
syncs on save, not continuously mid-write, and this pipeline rewrites a lot
of files in quick succession during any regeneration (`allocation.csv`,
the Class Attendance workbook, every signature sheet, workbook, hall plan -
often a dozen-plus files per change). Two devices writing around the same
moment produces a `... (Conflicted copy from ...).xlsx` file from Dropbox,
which then needs manual untangling since it's not obvious which copy is
current. Fine to switch which device you're working from; not safe to use
both at once.

If you keep a second copy of the project outside the synced folder, treat
it as stale - the synced copy is the one to edit.

## Getting it running (fresh machine)

macOS / Linux:

```bash
./setup.sh              # creates .venv-$(uname -s), installs requirements.txt,
                         # then on Linux generates a .desktop launcher and on
                         # macOS an .app bundle (see make_mac_app.py)
./launch_dashboard.sh   # starts Streamlit headless, opens it in a browser
```

Windows: double-click `setup.bat` once, then `launch_dashboard_chrome.bat`.

macOS gets a real app: `setup.sh` runs `make_mac_app.py`, which builds
`~/Applications/<CODE> Dashboard.app` (generated `.icns` included) so the
dashboard is in Spotlight and can be pinned to the Dock. The bundle holds no
logic - its executable is one line handing `launch_dashboard.sh` to
Terminal.app. **That handoff is not cosmetic:** the project lives under
`~/Library/CloudStorage` (Dropbox), which macOS protects with TCC, and an app
launched from Finder gets `deny(1) file-read-data` on the launcher script -
silently, with no consent prompt, and an ad-hoc `codesign` doesn't change it.
Terminal already holds that access. To lose the Terminal window, grant the
`.app` Full Disk Access in System Settings › Privacy & Security, then have it
exec `launch_dashboard.sh` directly. Re-run `make_mac_app.py` after moving the
project - the path is baked into the bundle's executable.

**The venv is per-OS: `.venv-Darwin`, `.venv-Linux`, `.venv-Windows`.** This
folder is Dropbox-synced across all three machines and a venv is *not*
portable - its interpreter symlinks and console-script shebangs are absolute
paths from the machine that built it, so one shared `.venv` only works on
whichever machine ran setup last. Each machine builds and uses its own dir;
`setup.sh`/`setup.bat` rebuild it automatically if it turns out to be
another machine's. `setup.sh` also marks the dir `com.dropbox.ignored` where
the client supports it, so hundreds of MB of machine-specific binaries don't
sync. A legacy shared `.venv/` is still honoured by the launchers as a
fallback if it actually works on the current machine.

Or manually: `python3 -m venv .venv-$(uname -s)` then
`.venv-$(uname -s)/bin/pip install -r requirements.txt`, and
`.venv-$(uname -s)/bin/streamlit run dashboard.py` or `… /bin/python run.py --help`.

`launch_dashboard.sh` / `setup.sh` (bash, macOS + Linux) are the counterparts
of `launch_dashboard_chrome.bat` / `setup.bat` / the Desktop shortcut on
Windows; `run_dashboard.bat`, `run_static_server.bat` and `_python.bat` (the
shared interpreter resolver the `.bat` files `call`) are Windows-only,
harmless to keep around but won't run here. The `.bat` files no longer
hardcode a Python path - `_python.bat` resolves `.venv-Windows` → legacy
`.venv` → the old `Python312` install path → `py -3` → `python`.

**Streamlit does not hot-reload imported local modules** - only the main
script (`dashboard.py`) it's actually watching. After editing `seating.py`,
`sheets.py`, `handout.py`, `reports.py`, `finder.py`,
`exam_rooms.py`, or `attendance_report.py`, the running server must be
fully stopped and restarted, not just left to auto-reload.

## File map

| File | Role |
|---|---|
| `config.py` | Course instructors (`COURSE_INSTRUCTORS`: names + addresses, the single source for both the printed names and the mail Cc), hall geometry (`ROOMS`, `Block` dataclass), branch-grouping (`BRANCH_GROUPS`), per-block colours, signature-sheet page geometry. Edit this first for a new room. |
| `seating.py` | Roster loading (auto-detects sheet/header shape), `load_attendance_roster` (the master Class Attendance workbook, in register order), seat map, `allocate`/`allocate_grouped`, `reorder_within_blocks`, `move_student`, `vacant_seats`, `block_summary`. |
| `reports.py` | Seating workbook (xlsx) + interactive hall-plan HTML (`plan_template.html`, real interactive plan - falls back to a one-line stub if that template file is ever missing). |
| `sheets.py` | Signature-sheet PDFs (one per block) + `template.json` recording what was printed - the dashboard discovers cohorts by that file. Optional `attendance=`/`att_asof=` add a per-student attendance column (the "next class" sheet). |
| `handout.py` | Plain student-facing seat-allotment PDF (block map + roll roster), optional weekly Attendance-TA roster footer. |
| `finder.py` | "Find your block" self-lookup page (`finder template.html`) - hashed roll numbers, safe to publish; TA duty lookup too. |
| `ta_duty.py` | How many TAs a block needs and who fills each post (auto or manual). |
| `exam_plan.py` | Saved exams under `out/exams/<slug>/`, seat map + student lookup pages, branch interleaving, the term's exam `SCHEDULE`. |
| `exam_rooms.py` | Separate exam-seating model (`LHC Seating Plan.xlsx`), multi-room, alternate-seat spacing. |
| `attendance_pdf.py` | The attendance report as a PDF (`report_data` + `build_report`): where the class stands, class by class, everyone below the requirement with what it would take to recover, and a method page. Drawn over `attendance_report`/`mailer` figures so the report cannot drift from the mail a student gets. |
| `attendance_report.py` | Reads the master Class Attendance workbook for the dashboard's summary/defaulter view - **column positions are detected by header text, not hardcoded** (see gotcha below). |
| `out/transitions.csv` | Append-only history of every move made through tab 5 - the source for "who needs telling". Written by `seating.log_transition`, never rewritten by a regeneration. |
| `checks.py` | Every invariant in one pass (`run_all`), behind the 🩺 Health check page; also runnable as `python checks.py` (exit 1 on any failure). |
| `history.py` | Rolling snapshots in `out/.history/*.zip` taken before each write, plus the `state/` baseline drift detection compares against. |
| `lockfile.py` | Advisory "another machine has this open" marker for the Dropbox folder. An entry from a process that is gone on *this* machine (a Ctrl-C'd dashboard) is dropped, so a restart doesn't warn about itself; each machine is named once. |
| `lookup.py` | Read-only "everything about one student" (`directory`, `search`, `student_record`): classroom seat, exam seats, attendance with class-by-class marks, moves, mail log. Behind the Student lookup page. |
| `bundle.py` | Every printable for one cohort or exam as one ZIP (`cohort_bundle`, `exam_bundle`) with a MANIFEST.txt; files built before the folder's `allocation.csv` are flagged as stale. Overview page. |
| `posters.py` | Hall-door poster (QR + block map) and `exam_pack` - the printable exam PDF: TA posting cover, per-block seat grids, per-room signature lists. |
| `tests/` | pytest suite over seating, mail, history, sheets and exams - plus an AppTest smoke test that renders every dashboard page. |
| `mailer.py` | Student mail-out: `<enrolment>@EMAIL_DOMAIN` addresses, transition/exam message builders, SMTP send + `out/mail_log.csv`. No secrets in this folder - the password is env-only. |
| `run.py` | CLI entry point (`plan` / `sheets` subcommands). |
| `dashboard.py` | Streamlit control panel wrapping everything above - the primary way this gets used day to day. |
| `.streamlit/config.toml` | Dashboard theme (colours, radius, toolbar mode). Project-level on purpose, so all three synced machines look the same. |

## Data flow - what's source of truth

```
roster (.xlsx/.xls)  --[plan]-->  out/<cohort>/allocation.csv  (source of truth for seating)
                                        |
                                        +--> out/<cohort>/<room>_seating.xlsx
                                        +--> out/<cohort>/<room>_seating_plan.html
                                        +--> out/<cohort>/sheets/*.pdf + template.json
                                        +--> <CODE> Seating.pdf (both cohorts combined)
                                        +--> find-your-block.html (both cohorts + TA duty)

out/<cohort>/allocation.csv  <-- kept in sync with -->  Class Attendance <CODE>.xlsx
  (Roll/Name/Program/Group/Medium/Block/Seat/SeatRow/SeatCol)   (the Institute's own per-medium attendance record)
```

`out/<cohort>/allocation.csv` is what every generator (`reports.py`,
`sheets.py`, `handout.py`, `finder.py`) actually reads from. **The Class
Attendance workbook and `allocation.csv` are two independently-editable
files that must be kept in sync by hand** (or by whoever's driving this
tool) - there's no automatic reconciliation. In practice: the human edits
rows in the Excel workbook (medium transfers, reordering within a block)
and/or uses the dashboard's "Move a student" tab (which edits
`allocation.csv` directly); either one can get ahead of the other.

## Established workflow for "a student transitioned" / "rows got reordered"

This has come up repeatedly. The reliable sequence:

1. **Diff first, don't assume.** Compare roll sets between
   `out/<cohort>/allocation.csv` and the workbook's two sheets. Sometimes
   the membership move already happened in `allocation.csv` (e.g. via the
   dashboard's Move-a-student tab) and only the *order* is stale; sometimes
   neither has moved yet. Never assume which file is ahead.
2. **Per-block order + set check**, both cohorts: for each block, compare
   the roll sequence in the workbook (row order, filtered to that block)
   against `allocation.csv` sorted by `(Block, seat-number)`. A block
   showing "same set, different order" needs its `Seat`/`SeatRow`/`SeatCol`
   re-synced to the workbook's order - reassign the block's existing
   seat-map slots (in seat-map row/col order) to students in the *desired*
   order, preserving `Block`. A block showing a numeric gap (a departed
   student's old seat number now unused mid-sequence) needs recompacting:
   same technique, but preserving the block's *current* relative order
   rather than importing a new one.
3. **Seat numbers must stay contiguous from 1, vacant capacity at the
   end, block never changed** as a side effect of any reorder/compaction -
   this has been an explicit standing requirement.
4. **Regenerate downstream in this order**: signature sheets → seating
   workbook (xlsx) → hall-plan (html) → (if asked) seat-allotment PDF →
   (if asked) find-your-block page. Build to a scratch path first, verify
   (roll-by-roll diff against the pre-change version, order-match check
   against the workbook), *then* overwrite the live file.
5. **No persistent backup files.** Earlier in this project's life every
   change left a `.pre-something.xlsx` safety copy lying around; the human
   explicitly asked for those to stop accumulating. Verify in a scratch
   copy, apply once confident, delete the scratch copy - don't leave
   `.bak`/`.pre-*` files behind unless asked to.
6. **Never re-derive allocation from the roster** (`run.py sheets --group
   ...` without pointing at the existing `allocation.csv`) once manual
   moves exist - that command re-runs `allocate()`/`allocate_grouped()`
   fresh from the roster file and silently discards every hand-edit made
   since. Load `allocation.csv` via pandas and pass that DataFrame to the
   builder functions directly instead.

## Mailing students (`mailer.py`, dashboard tab **📧 Email students**)

Addresses are **derived, never stored**: `student_email("B26BB1901")` →
`b26bb1901@example.edu`, lower-cased. That is deliberate - the master Class
Attendance workbook has gained and lost an `Email` column more than once
(see the gotcha below), so no address list in this project can be trusted to
be current, while the enrolment number keys everything already.

**A classroom transition mail names the block, never the seat.** Seating here
is block-level by design (`Seat`/`SeatRow`/`SeatCol` are a print and
attendance order, not a chair), so a seat id in a student mail states
something untrue. `build_transition_mail` takes no seat argument at all, and
`_no_seat_tokens` re-scans the finished subject and body for anything shaped
like `A-004` and refuses to hand back the message if it finds one - so a seat
can't sneak in through a free-text field like "Effective from". Exam mail
(`build_exam_mail`) is the opposite case: an exam seat *is* a real chair, so
room + block + seat may all be named (`include_seat=False` for a block-only
exam venue re-applies the classroom rule).

Configuration is environment-first, because this folder is Dropbox-synced
across three machines and anything written here syncs with it:
`COURSE_SMTP_HOST` / `_PORT` / `_USER` / `_PASSWORD` / `_FROM` /
`_FROM_NAME` / `_REPLY_TO` / `_SSL`, plus `COURSE_MAIL_BCC` for a course
office archive copy. Precedence is **environment → `~/.course_dashboard_smtp.env` →
`mail_config.json`**. `mail_config.json` (see `mail_config.example.json`) may
hold everything **except the password**, which `load_smtp_config` ignores in
that file on purpose. `~/.course_dashboard_smtp.env` (plain `KEY=value`, read by
`load_env_file`) is where the app password actually lives: it is in the home
directory rather than this folder so it doesn't sync, and unlike a shell
`export` it is still found when the dashboard is launched from Spotlight, the
macOS `.app`, the Windows shortcut, or a `.desktop` entry - none of which
inherit a login shell's environment. Port 465 flips `use_ssl` on by itself,
and a Gmail app password pasted in Google's own `abcd efgh ijkl mnop` display
form is de-spaced (`_clean_password`, that exact 4×4 shape only - a real
passphrase containing spaces is left alone). Unconfigured, the dashboard
still composes and shows every message - only sending is blocked. README's
"Gmail SMTP (mailing students)" section is the step-by-step for the human.

**Two delivery routes, and the Gmail one is the default in practice.**
Institute Workspace accounts commonly have app passwords disabled by the
admin, which leaves SMTP unable to authenticate at all - so `mailer` can also
just build a Gmail compose URL (`gmail_compose_url`, `compose_links`,
`mailto_url`) and let the human press Send from an account the browser is
already signed into. The dashboard renders one **Open in Gmail** button per window with a `sent`
tick box beside it; ticking removes that row from the list and moves it into a
"Marked sent" expander, so what's still on screen is what's still to do - the
failure mode with two dozen one-per-student windows is losing your place and
either skipping someone or mailing them twice. Windows are identified by
`_window_id` (roll + subject + size), never by list position, which drifts as
soon as the selection changes underneath a rerun. The log button then records
the ticked ones if any are ticked, otherwise the whole batch.

`compose_links` groups messages with an *identical*
subject and body into one window with every recipient in **Bcc** (never To -
students must not see each other's addresses), measuring the real URL as it
packs so a browser can't truncate the list; personalised messages fall back to
one window per student, which is why the dashboard offers a "Dear student"
generic-greeting option for block-wide notices. Those get logged by
`log_handoff` as **`handed-off`**, deliberately not `sent`: the toolkit opened
a window, it cannot know what the human did next.

Sending over SMTP is a two-step in the UI and in the API: `send_mails(..., dry_run=True)`
is the default and touches nothing, and the dashboard's send button starts on
"Dry run" ticked every rerun. A real send appends one row per recipient to
`out/mail_log.csv` (timestamp, kind, roll, to, subject, status, detail) -
that log is the only record a student was told anything, so check it before
re-sending to a block. A refused recipient is logged and the batch continues;
only a connection/auth failure aborts.

`python mailer.py --configure` writes `~/.course_dashboard_smtp.env` interactively
(hidden password prompt, mode 600, merges rather than clobbers keys it didn't
ask about). `python mailer.py --to <address> --send` is the self-test: one sample
message to one address, connection check first, `--send` omitted stops at
composing. Use it before any batch, and after any SMTP change.

The `Medium` column is **not** used to fill the "Batch/medium" line in a
mail. `out/english/allocation.csv` currently holds 21 rows marked
`Medium=Hindi` (students seated in LHC110 whose roster language flag says
otherwise), so a per-row read would tell those students they're in the Hindi
batch. Both the bulk tab and the move tab name the cohort/room instead.

**Block mail** (`build_exam_block_mail` → `build_group_mail`). **One mail per
block, in every room**, naming room and block and not the seat: that is what a
student needs to walk in and sit down, and the seat is on the block sheet at
the door. A hall numbered by row and column (`exam_rooms.ROW_COL_ROOMS`, i.e.
LHC 105) has no blocks to name, so it gets **one mail for the room**, saying
the seat is found by row and column; every lettered room still mails block by
block, since a room-wide mail cannot tell a student which of seven blocks is
theirs. `seats=` can list roll and seat in
the body (two columns, no names) but is off by default. One message per
block - per room where the room is posted whole - with **students in To,
instructors in Cc, that block's TAs in Bcc**, which is what the course asked
for. `students_in="bcc"` is offered next to it: To lets the block see it is
their notice and reply, at the cost of publishing forty classmates' addresses
to each other and inviting reply-all. The mail names room and block and
**never a seat** - one message reaches the whole block, and a seat is per
student, so it points at the find-my-seat page or the door sheet instead. TA
addresses come from the saved `ta_assignment.csv`, so the Bcc list follows
whatever posting the pack was built with; `ta_email` is `student_email` under
another name, since TA rolls follow the same pattern.

**Escalation.** `ATTENDANCE_LEVELS` (5/9/13 absences → L1/L2/L3) plus
`unmailed_attendance()` answer "who has reached a level they haven't been
written to at yet". The level goes into the log's `kind`
(`attendance_shortfall_L2`), which is what stops a student getting the same
notice twice while still catching them when they get worse. L2 and L3 add one
sentence naming which notice it is; the agreed wording is otherwise untouched.

**Attendance shortfall mail** (`build_attendance_mail`) is the one kind that
isn't about seating: it reads the master Class Attendance workbook through
`attendance_report.load` - so it inherits that module's header-text column
detection, and must never grow its own hardcoded column positions - counts
`N` marks per student, and writes to everyone above the limit with both course
instructors in **Cc** (`DEFAULT_ATTENDANCE_CC` = `instructor.one@`, `instructor.two@`;
`Mail.cc` carries into both the SMTP message and the Gmail compose URL).

**Two templates, and which one a student gets is arithmetic, not judgement.**
`can_still_reach(present, absent, remaining, threshold)` decides. A student who
can still make the requirement is told exactly how - `classes_to_recover()`, the
number of classes in a row that gets them back, which at 75% is `3N - Y`,
because every class missed costs three to make up. A student who cannot reach it
however many they attend is *not* told to "attend regularly" as though it were
still in their hands: they get `best_possible_percent()`, the highest figure
still open to them, a request to meet the instructors by a date, and a line
saying documents for genuine absences change the arithmetic. Telling the second
group the first message is a promise they would find out was untrue.

`remaining` (how many classes are still to be held) is what splits the two, so
the caller has to supply it - with none given, everyone reads as unrecoverable.
The dashboard reads it from the register's own dated columns (39 scheduled, 21
marked → 18 remain) rather than asking anyone to type it, and shows which notice
each student would get before a single message is composed. `meet_by` left blank
degrades to "as soon as possible"; the consequence line is blank by default and
is never invented, since it states institute policy.

These notices sign off with `ATTENDANCE_SIGNATURE` (Your Name / Teaching
Assistant / `config.DEPARTMENT`, `config.INSTITUTE_SHORT`) over `Regards,` - shorter than
the seating mail's `SIGNATURE`, because a hundred of them go out at once and
they read as a note from the student's TA, not a letter from an office.

**The body is otherwise the course's own agreed wording - don't improve it.** It
names the figures the policy turns on; `held`, `present` and `percent` are
computed and kept in `Mail.meta` (and shown in the dashboard's check table)
but deliberately stay out of the message. It closes `Best Regards,` over
`mailer.SIGNATURE` - a named person (Devanagari included; both the UTF-8 SMTP
body and the percent-encoded Gmail URL carry it intact), not "the course
office", because the mail asks the student to come and talk to someone. There
is no session/term line: the agreed wording has none, which is why
`build_attendance_mail` takes no `session` argument while the seating builders
do. The dashboard's "extra paragraph" box inserts before `Best Regards,`;
everything else is fixed in code.

These are composed with `compose_links(..., group=False)` - one window per
student, never a shared Bcc window, since each body carries that student's own
record. `mail_send_ui(..., group=False)` is what the dashboard passes here.

The Overview tab's per-cohort **Classes held** metric reads the master Class
Attendance workbook (sheet matched to the cohort folder name), counting only
sessions that have marks - not `out/<cohort>/attendance.xlsx`, the retired
scan pipeline's file, which only the Hindi cohort ever had and which made
English read 0 while nine classes had been held. The read is cached on the
workbook's mtime, since it's edited by hand outside this app.

**Who gets mailed for a *seating* change is driven by `out/transitions.csv`**, not by the current
allocation. `allocation.csv` is state - it says where everyone sits now and is
rewritten wholesale by every regeneration - so after a move there is nothing
left saying *who changed*. `seating.log_transition` appends one append-only row
per move made through tab 5 (roll, name, from/to room+block+seat, cohorts,
timestamp); `seating.block_changed` drops moves that stayed inside one block
(a seat renumber or a compaction - nothing a student needs told, and mail
never names a seat anyway); `mailer.unmailed` drops the ones already covered.
That last one compares **timestamps, not roll numbers**: a student can move
twice in a term, and a mail about the first move does not cover the second.
`handed-off` counts as covered, `refused`/`failed` do not.

A move made by hand in Excel or straight in `allocation.csv` leaves no row
here and so will never appear in that list - that is the price of the log
being written at the point of the move. Mail those through **A whole block**,
or make the move through tab 5 so it's recorded.

Two entry points: **📧 Email students** (bulk - pick blocks in a cohort, or an
exam allocation carried over from the Exam seating tab / an uploaded roll
list), and the **🔀 5 · Move a student** tab, which after a move offers to
mail exactly that one student. Adding a new kind of notice means one more
`build_*_mail` function returning a `Mail`; the preview/confirm/send widget
(`mail_send_ui` in `dashboard.py`) takes any list of them.

## The "next class" signature sheet (attendance column)

`sheets.build_all(..., attendance=…, att_asof=…)` prints each student's
standing **up to the last class held** in a column between the name and the
signature box, so the sheet they sign on the way in tells them where they
stand. The map comes from `attendance_report.percent_map(path, sheet)`, which
returns `{roll: {percent, present, held}}` plus a date label.

Two things that are easy to get wrong here:

- **"As of" is the last session with marks in it, not the last dated column.**
  This workbook pre-dates the entire term - 39 dated columns, 9 with marks at
  the time of writing - so a plain `Date.max()` labels the sheet with a class
  that hasn't happened. `percent_map` filters on `Present + Absent > 0`.
- **The batch sheet must match the cohort being printed.** The dashboard
  guesses it from the output folder name (`out/english` → the `English`
  sheet) and shows how many seated students are missing from that sheet; any
  that are print `-` rather than a wrong figure.

Without the argument the sheet is byte-for-byte what it always was - the
column is taken out of the signature box (20 mm of 73 mm), not out of the row
height or the page count, so a plain sheet still prints page-for-page with
every set already filed.

**Worth saying out loud to whoever asks for this:** a signature sheet travels
down a row of students, so every student in a block can read their
neighbours' percentages. The dashboard says so next to the checkbox. If that
isn't wanted, the per-student attendance mail (📧 tab) tells each student the
same thing privately.

## Safety net: snapshots, checks, and the lock

Three things now stand between a mistake and a bad afternoon:

**`history.py`** snapshots whatever a build is about to overwrite into
`out/.history/<stamp>_<label>.zip` before writing, keeps the newest 20 and
prunes the rest. This replaces the abandoned `*.pre-something.xlsx`
convention - same safety, one hidden git-ignored folder instead of clutter in
the working directory. `restore()` snapshots the current state first, so an
undo is itself undoable. Every write path in `dashboard.py` calls
`snapshot_before(...)`, which never raises: failing the build the user asked
for because the *snapshot* failed would be the wrong trade.

**`checks.py`** runs every invariant (`python checks.py` exits 1 on failure,
or use the 🩺 Health check page). Add a check when you find a new way for
things to go wrong - the two bugs that motivated it were a counter reading the
wrong file and a cohort's sheets rebuilt from the other cohort's allocation.

**`lockfile.py`** heartbeats into `.dashboard.lock` and warns when another
machine was active in the last five minutes. Advisory by design: Dropbox
propagation is seconds to minutes, so a real lock would be a lie, and a stale
one would lock you out of your own tool.

**Drift detection.** `history.save_state()` keeps a copy of each
`allocation.csv` as the toolkit last wrote it;
`seating.detect_untracked_moves()` diffs the live file against it to find
block changes made by hand in Excel - which otherwise never reach the mail
page, because `log_transition` only sees moves made through the dashboard.
`adopt_untracked_moves()` writes them to `transitions.csv` and re-baselines.
No baseline yet means "report nothing", never "report everything".

## Rooms the Institute's workbook does not describe

`exam_rooms.EXTRA_ROOMS` defines venues that exist in the building but not in
`LHC Seating Plan.xlsx`. **LHC 105** is there: LHC 206 and LHC 207 merged into
one hall, ten rows of fifteen, 150 seats. Each row is one `ExamBlock`, which is
what keeps the spacing arithmetic honest: alternate spacing skips every second
column of every row, so the pair's 80 spaced places survive the merge exactly.
Defined in code rather than by editing the Institute's own workbook, which this
project only ever reads.

**It is read as rows and columns, never as blocks** (`ROW_COL_ROOMS`). Rows are
keyed `R01` … `R10` so plain sorting keeps them in order, and every label a
person sees goes through `group_label` ("Row 3"), `place_label` ("LHC 105, Row
3, Column 7") or `venue_note` ("Rows 1-10"). LHC 110 is untouched: a block
there is a place an invigilator stands in, with its own grid and signature
sheet. The pack and the notice sheet draw LHC 105 as one grid, columns numbered
across the top and rows down the left (`posters._rowcol_grid`); its signature
sheet has ROW/COL columns instead of BLOCK/SEAT. That grid page prints
**landscape A4** (`posters.LANDSCAPE_COLS`, over 12 columns): portrait shrinks
a fifteen-column row to an 11mm desk and a 5pt roll number, sideways it is 18mm
and 8pt on the same sheet of paper. Every other page in the pack stays portrait,
so the file is still one A4 print job. Its **signature list prints two columns
to a landscape sheet** (`_signature_two_up`): eighty names down a portrait page
is four sheets to keep in order, two lists side by side is two. `posters.report_summary` writes the **one page a student reads**: roll-number
ranges against room and block, no seats and no names, for the noticeboard and
the department group. `roll_ranges` reads the ranges off the *sorted* roll list
rather than the seating order, and prints every stretch a block holds, so the
two students the roster gained at the end show as their own line instead of
being rounded into block B's range. A range also **never crosses a branch**:
"B26AE1943 to B26BB1932" reads as one sweep of numbers but joins AE to BB, so
it breaks at the branch and the stretches are listed comma separated. Where a range would mislead - the seniors'
block spans four batches - it is described in words instead.

A signature list longer than one portrait sheet (`posters.TWO_UP_OVER`, 26
names) prints **two lists to a landscape sheet** instead: LHC 110's back blocks
of 65-70 and LHC 105's 80 each come out at two sheets rather than three or
four, while a front block of twenty keeps its single portrait page.

**Every** signature sheet, LHC 110's block sheets included, is numbered by a
plain serial rather than by block/seat or row/column: the seat is on the
seating plan, and a signature sheet is read down a list of people.
`posters.serial_numbers(alloc)` assigns it once for the whole exam, in the
order the sheets print (room by room, block by block, seat by seat), so the
numbering runs on from one block to the next and from LHC 110 into LHC 105.
The heads collect one pile of sheets at the end, and a serial that restarted
in every hall would make "number 47" a question rather than an answer.

Every seat grid, in every room, draws an **empty desk narrow and short**
(`posters.NARROW_COL`, `_seated_cols` / `_col_widths` / `_draw_seat_row`): with
alternate seating half the columns hold nobody, and drawn full size they took
the width the roll numbers needed. Boxes are also flat rather than square
(`FLAT_CELL`), since the box holds one short line of type; height left over
goes into the aisles between rows, blocks and sections rather than into the
boxes. Together that took LHC 110's block grids and LHC 105's row grid from
about 7pt to about 11pt on the same paper. `min_cell_w` on the notice sheet is
now the width of a *seated* column, so a band splits sooner rather than
printing three blocks at 7pt.

`notice_board_map(..., paper=)` is `"auto"` (A3 for the big halls, A4 for the
rest) or `"a4"` for a printer that has no A3: every sheet is A4, wide rooms
landscape, and a hall that needed A3 is **split across sheets at the same cell
size**, each headed "sheet 2 of 5" so the board can be taped up in order.
Scaling an A3 layout onto A4 was never an option: it shrinks the roll numbers
past reading, which is the one thing the sheet exists to show. It is in `posters.A3_ROOMS`
(150 seats does not read on A4), and `ta_duty` posts TAs to the room rather
than per row (four for LHC 105).

`exam_rooms.SUPERSEDED` marks LHC 206 and 207 unsupported with the reason
attached, so they leave the room picker but stay listed: someone looking for
LHC 206 finds out it became LHC 105 rather than assuming the workbook parser
broke. Rooms are sorted by number (`_room_sort_key`), so LHC 105 reads before
LHC 110 rather than after LHC 308. Both
scheduled exams now run in LHC 110 plus LHC 105: 283 + 80 = 363 spaced seats
for 365 students, with the last two in transition seats.

## The exam side (`exam_plan.py`)

An exam used to live in one Streamlit session and vanish with it. Now it is
saved under `out/exams/<slug>/`: `allocation.csv` (source of truth),
`exam.json` (title, date, venue, spacing, rooms), `seat_map.html` and
`find-my-seat.html`. Re-allocating the same title **replaces** that exam
rather than adding a near-duplicate; the dashboard snapshots first.

**Two pages, two audiences, and the difference is the point.**
`seat_map.html` is for invigilators and the notice board: every room drawn in
its real row geometry (`exam_rooms.ExamBlock.rows`), roll numbers in position,
unused seats still drawn - the empty chairs are what make spaced seating
legible from the front of a hall. `find-my-seat.html` is for students and is
safe to publish: roll numbers appear only as the same one-way hash the
classroom lookup page uses, no names, and the block drawing carries seat
labels only. `tests/test_exam_plan.py` asserts both halves of that.

**Transition seats are held back in exams too.** `ExamBlock.reserved` holds
seats that exist but are never allotted; `_apply_known_geometry` fills it with
the wedge seats of LHC 110's blocks B (8) and C (12) - the taper behind the
last full row, which the classroom side already reserves as the transition
pool (`config.RESERVE_WEDGE`). They stay in `rows` so the plan still draws
them, hatched and labelled "transition seat - never allotted": an unexplained
gap on a seat map reads as a missing student. LHC 110's every-seat capacity is
therefore 566, not 586.

**…except for one or two stragglers.** `ExamBlock.overflow_seats(mode)`
offers the reserved seats - still spaced the same way as the rest of the block
- and the exam page dips into that pool only when at most
`exam_rooms.OVERFLOW_LIMIT` (2) students would otherwise be left over.
Opening a second room, with its own invigilator, for one person is the worse
trade; being short by twenty means the room is genuinely too small and the
allocation should say so. Overflow seats are reported by seat on the page and
drawn hatched-in-the-allotted-colour on the map, so nobody has to wonder why
someone is sitting in the taper.

**Three spacing modes** (`exam_rooms.SPACING_MODES`), chosen per exam and
saved in `exam.json`:

| mode | gap beside | gap behind | LHC 110 |
|---|---|---|---|
| `alternate` | yes | yes | 147 seats |
| `side` | yes | **no** - may sit behind | 283 seats |
| `every` | no | no | 566 seats |

`side` is the one to reach for in a large flat hall: it stops the
over-the-shoulder look, which is the realistic risk, at double the capacity of
full spacing. `_normalise_mode` still accepts the old `spaced=True/False`
boolean, so exams saved before the third mode existed keep loading.

**Room geometry is enriched from `config.ROOMS`.** The exam workbook gives
some blocks only as a *total* - `Block ID - E (140)` - which `exam_rooms`
stored as one flat 140-seat row. That was not merely an ugly drawing:
`spaced_seats` skips alternate rows and alternate seats within a row, and over
a single long row there are no rows to skip, so "alternate seating" meant
every second chair in one continuous line - people seated directly in front of
and behind each other. `_apply_known_geometry` reshapes such blocks using the
hall this project already describes (`LHC 110` ↔ `config.ROOMS["LHC110"]`,
which matches seat-for-seat: A 6×6, B 42+8 wedge, C 42+12, D 6×6, E 14×10,
F 13×10, G 14×10), and only when the totals agree - a mismatch means the two
sources describe different halls, so it falls back to a plain 15-wide wrap.
This changed LHC 110's spaced capacity from 293 to **152**, which is the
honest number.

**`interleave_by_branch` is a seating strategy, not a shuffle.** Dealt
round-robin from per-branch piles, so neighbours are rarely from the same
branch - on the current roster it takes adjacent same-branch pairs from ~100%
to 1%. Deterministic on purpose: the same roster always produces the same
plan, which matters when a query arrives a week later.

**Exam roster: `exam_plan.exam_roster()` - roll-list order, missing students
first.** The roll list is the order the department reads students in, so
seating that follows it can be checked against their own paperwork; but the
export is a snapshot and currently misses B22PH903, a late admission who is on
the register and in class. He is placed at the **front** rather than dropped
(no seat at all) or appended (buried in the last row of the last block), and a
test asserts saved exams seat him first. The Class Attendance workbook remains
available as a roster source in its own right.

**Older note - allocate exams from the register, not the bare roll-list
export.** The export is a snapshot: `student_rollList-<CODE>.xls` has 364
students while the register has 365, and the missing one (B22PH903, a late
admission) had a classroom seat and no exam seat until the source was
changed. The Exam seating page now defaults its roster source to the workbook
and reads both sheets together, since every scheduled exam is sat by both
batches. A test asserts every registered student has a seat in every saved
exam.

**`SCHEDULE`** holds the term's fixed exams - Quiz 1 (31 Aug 2026), Minor
(16 Sep 2026) and Major (21 Nov 2026), all sat by both batches together - and the Exam seating page
prefills title and date/time from it. One mistyped exam time would otherwise
propagate to the roll list, the seat map, the poster and every student mail.

## Course instructors

`config.COURSE_INSTRUCTORS` pairs each name with their address:
Dr. Instructor One (`instructor.one@`) and Dr. Instructor Two
(`instructor.two@`). One list, because the names printed on a pack and the
addresses Cc'd on a mail are the same two people and drift apart the moment
they are typed twice. `instructor_names()` feeds the exam pack cover and every
inner page header, the hall-door poster, the classroom signature sheets, the
seat allotment PDF, both exam HTML pages and the block mail body;
`instructor_emails()` is `mailer.DEFAULT_ATTENDANCE_CC`.

## Invigilators (`ta_duty.py`) and the printable pack

**The posting rule is about sight-lines, not headcount:** one TA per front
block of LHC 110, two per long back block (E/F/G run fourteen rows deep - one
person at the front cannot see the back row's desks), two per smaller room.
`SECTION_TAS` keys off `config.ROOMS`' `Block.section`, and a room this
project has no geometry for is posted whole-room (`WHOLE_ROOM = "-"`) rather
than per block.

Those counts are **defaults, not law**: `requirements(alloc, section_tas=…,
per_room=…)` overrides them per exam, and the Exam seating page exposes all
three as number inputs. Three per long block costs 17 posts against 21 TAs on
the current duty sheet - affordable, and the summary line recalculates live so
the trade is visible before assigning. The override must never mutate
`SECTION_TAS` itself; a test asserts that.

`auto_assign` fills posts in room order from the duty sheet, **Attendance TAs
first** - they stand in these blocks every week for the class register, so
they know the room, the block boundaries and the faces; an invigilator working
out where block F ends is an invigilator not watching anyone - then Evaluation
TAs, heads last since they float. `exclude=` keeps anyone already given an
off-block job out of the block pool, but only while the remaining pool still
covers every post: better a double-booking the sheet shows than an empty post
caused by reserving someone for a twenty-minute job. It **runs out rather than doubling anyone up** - a TA
posted to two rooms is a gap that looks filled - leaving blank rows that print
as "- not assigned -". Manual mode starts from the automatic answer and warns
if a name appears twice or a post is under-filled. The chosen assignment is
saved beside the exam as `ta_assignment.csv`.

**Heads.** They hold no block - they move between rooms - but "no block" is
not "no responsibility", and a pack that never names them reads as though the
exam runs itself. `HEAD_RESPONSIBILITIES` lists the concrete jobs each head
owns on the day. The **Evaluation Head leads every exam, quizzes included**,
and briefs the invigilators: the duty sheet's wording splits quizzes from
Minor/Major, but the course settled on one person answering for the exam
whatever it is called, so an irregularity is never somebody's call to make on
the spot. The page
lets you override the lead for one exam. Heads are still never given a block
or the printing duty. No heads on the sheet → no section, rather than an
invented one. Saved as `ta_head_duties.csv`.

**Off-block duties.** `SUPPORT_ROLES` holds the jobs that belong to the exam
but not to a block - currently question-paper printing, two people.
`assign_support(names=[…])` posts named people and is taken as given -
**including a head**: naming someone is an instruction, not a suggestion, and
the page says so where it happens. Left to itself it never picks a head (they
supervise) and prefers TAs with no block post, because
sending someone who must be standing in block E while the papers are being
printed is how a job silently doesn't happen. If only posted TAs remain it
takes them but flags `AlsoInvigilating`, which the page warns about and the
cover prints as "(also invigilating)". Saved as `ta_other_duties.csv`.

**The grid sizes itself to the content.** Columns: a block the exam sheet
gave as one flat row (LHC 206/207 arrive as a single 15-seat line) is
re-wrapped narrow enough that a whole roll number fits its cell - printing
`B26CI…` on a seat map defeats the map. A block with real geometry is never
re-wrapped: its shape is the hall's. `_wrap_rows`' `width` is a *maximum*, so
passing a smaller one only touches the flat rows. Rows: cell height shrinks to
fit, so a block stays on one page whenever it can - flipping pages mid-block
while standing in front of it is what the pack exists to replace. Two tests
guard both.

**Block-wise or room-wise, by hall.** `posters.BLOCK_WISE_ROOMS` is
`{"LHC 110", "LHC 308"}`: those are big enough that a block is a place someone
stands, each with its own invigilator, so each gets its own grid page and its
own signature sheet. Every other room is one line of sight with two TAs for
the room, where per-block output would mean five grids and five signature
sheets for forty people - five things to lose instead of one to hand over - so
those print **room-wise**, blocks labelled inside the page and named in a
`Block` column on the sheet. Blocks are ordered alphabetically everywhere; the
exam workbook hands the small rooms back E..A, which reads as though the room
starts at the far end.

`posters.notice_board_map` is the other print deliverable: **one sheet per
room** for the board outside the hall, the whole room drawn as it is laid out,
**A3 landscape for the halls in `A3_ROOMS` (LHC 110, LHC 308) and A4 portrait
for the rest** - forty students on A3 is harder to read, not easier. The page
size is set per page, so printing must be at 100%: "fit to page" undoes it. Blocks are grouped by the hall's own sections where `config.ROOMS`
knows them (rear gallery above front section, via `_sections_for`), so the
sheet reads like standing at the door. It is a different object from the
pack's A4 grid, not a scaled copy: that one is held by an invigilator, this
one is read at arm's length with people in front of it, so cells and type are
sized against the sheet and every roll number prints in full. A band of blocks
wider than the paper is broken into further bands rather than shrunk past
legibility (`min_cell_w`), which is what lets the same code lay LHC 110's
seven blocks across A3 and LHC 206's five down A4.

`posters.exam_pack` is the print deliverable: a cover carrying the whole
posting table, then the seat grids in the hall's real geometry (roll numbers
in position, held-back transition seats marked "held"), then **signature
sheets** - seat, roll, name, box, and a
present/absent/scripts-collected tally at the foot. In LHC 110/308 that is per block: the invigilator standing in
block E should not be holding a sheet whose first two pages belong to blocks A
and B, and block-wise sheets can be collected and counted independently. A grid must fit one page; a signature list may run on,
since 70 students cannot sign one sheet. `invigilator_pack` remains as a thin
alias - earlier notes and the Posters page use that name.

## Dashboard look and navigation

**Pages, not tabs.** The eleven sections are `if PAGE == "…":` blocks driven
by a sidebar of buttons (`PAGES`, grouped Seating / Exams / Attendance & mail
/ Maintenance; the choice lives in `st.session_state["phl_page"]`, keyed by
the plain page name). Each page opens with `page_header()`, which reuses that
page's own nav icon. `st.tabs` ran *every*
tab's body on every rerun - nine workbook reads for one visible page - and its
bar overflowed a laptop width, hiding the last tabs behind a scroll arrow. Add
a page by adding a `PAGES` entry and an `if PAGE == …` block with the exact
same label string; the label carries the emoji and two spaces.

**Design tokens, not CSS colours.** `.streamlit/config.toml` carries the full
token set Streamlit 1.62 supports: Inter and JetBrains Mono via `fontFaces`
(CDN - an offline machine falls back to system fonts and simply looks
plainer), separate `[theme.light]` / `[theme.dark]` palettes so the ⋮ menu's
light/dark switch is real, a deliberately dark `[theme.*.sidebar]` in *both*
modes so the nav reads as chrome rather than content, and
`chartCategoricalColors` set to `config.BLOCK_COLOUR` so a chart of per-block
numbers matches the hall plan, the poster and the sheets.

What remains as CSS in `dashboard.py` is layout plus the few components with
no token. Rules to keep in mind when editing it:

- **Text colours use `inherit` + `opacity`, never a baked hex.** The palette
  dict (`_T`) is chosen server-side from `st.context.theme`, which only
  updates on the *next* rerun - a theme switched in the ⋮ menu would otherwise
  leave grey-on-white headings until something else triggered a rerun.
  Surfaces and accents can be baked; text cannot.
- **One `st.markdown` HTML block per injection.** A blank line between
  `<link>` and `<style>` ends the raw-HTML block and the rest of the CSS gets
  rendered on the page as literal text. Use `@import` inside the `<style>`.
- **Icons are Material *Symbols*, not Material Icons.** Streamlit's own
  `:material/…:` widgets draw from Symbols; the older Icons font renders an
  empty box for anything added since 2021 (`stethoscope`, for one), so the
  page-header font must match the nav's.

Three further rules exist because of specific Streamlit behaviour, not taste:

- **Nav buttons take `use_container_width=True` and no `help=`.** A tooltip
  wraps the button in an element that stops it filling the sidebar, which
  leaves every label centred instead of reading as a list. The label sits in a
  nested flex div that centres its own content, so `justify-content` has to be
  set on `button > div`, not just the button.
- **Group headings are `st.caption`, not a styled `<div>`.** A padded div
  inside a markdown block overflows its container and lands underneath the
  first nav button.
- **The sidebar's vertical-block gap is tightened** to `.15rem`; Streamlit's
  default ~1rem turns a nav list into a column of far-apart buttons.

`toolbarMode = "viewer"`, not `"minimal"` - minimal removes the toolbar
outright, taking the viewer's Settings menu (light/dark, print, rerun) with
it. The header strip under the title reads its chips from the same files
everything else writes (`allocation.csv`, the Class Attendance workbook,
`transitions.csv` vs `mail_log.csv`), so a wrong chip means a stale file, not
a display bug. Each chip is wrapped in its own try/except: a status strip must
never be the thing that stops the page rendering.

## Gotcha: Streamlit widget keys and cross-cohort writes

Tab 2's Room and Output-folder widgets are keyed **per allocation**
(`sheets_room_{slug}`, `sheets_out_{slug}`). With one fixed key each,
Streamlit keeps the value set for the previously picked cohort and ignores
the new default - which is how a Hindi allocation once got built into
`out/english/sheets` under room LHC110, overwriting blocks A–D with Hindi
students and **deleting English's E and F sheets** (`build_all` clears block
PDFs the current allocation doesn't use). Room is also guessed from the
cohort folder's own `*_seating.xlsx` rather than defaulting to the first room
in `ROOMS`. On top of that, building into a folder that isn't the
allocation's own now warns and disables the button until an "I mean to write
there" box is ticked. Any new widget whose sensible default depends on the
picked cohort needs the same treatment.

## The three marks: Y, N and E

The workbook carries **three** marks, not two. `E` is an **excused absence** -
the student was away and produced a medical receipt or another accepted
reason. Rules, enforced in `attendance_report.load`:

- **`E` is never an absence.** `Absent` counts `N` only, and every downstream
  consumer (the defaulter list, the shortfall mail, the escalation levels) is
  driven by that column. A warning mail triggered by an excused absence is a
  mail to someone who did everything right.
- **What `E` does to the percentage is a policy choice**, exposed as
  `EXCUSED_MODE` / the `excused=` argument and as a radio on the Attendance
  summary page. `"exclude"` (default) drops that session from the student's
  denominator; `"present"` credits it. Either way the student is not
  penalised - the modes differ only in whether the class counts at all.
- **A session where everyone was excused still happened.** `percent_map`'s
  "as of" date counts `Y + N + E`, so an all-excused class is not treated as
  a class that hasn't taken place yet.
- **A blank is still a gap.** Blank means nobody recorded anything, and the
  health check reports those separately from `E` - before this distinction
  existed it flagged six excused students as data-entry errors.

A fourth marker exists for a different situation: **`N/A` in a student's
cell** means the class does not apply to that student - a late admission whose
first weeks happened before they were on the roll. Counted towards nothing,
exactly like a blank, but *deliberate*: `attendance_report.NOT_APPLICABLE`
recognises it, and the health check reports N/A cells separately so a late
joiner doesn't read as a data-entry gap. The workbook already uses "N/A" as a
row-3 label for spacer columns, so this reuses its own vocabulary rather than
inventing a fifth mark, and its own `%` formula
(`COUNTIF("Y")/(COUNTIF("Y")+COUNTIF("N"))`) ignores it.

At the time of writing: 2 excused marks in English, 4 in Hindi, and 7 N/A
cells for one late admission (B22PH903, joined at C8).

### Editing the master workbook

`attendance_report` only *reads* it, and nothing in the pipeline writes to it.
When a cell genuinely has to be corrected, edit the sheet XML inside the
`.xlsx` rather than round-tripping through openpyxl: openpyxl rewrites the
whole file and silently drops parts it doesn't model (this workbook has two
drawing parts). The safe recipe, used for the B22PH903 fix: snapshot via
`history.snapshot` first, copy every zip part across unchanged except
`xl/worksheets/sheet1.xml` (English) or `sheet2.xml` (Hindi), only ever fill
cells that are *empty* (`<c r="E189" s="18"/>` - refuse if a value is
present), reuse an existing `sharedStrings` index (`N/A` is 26, `Y` 61, `N`
62, `E` 298), then verify part-by-part that nothing else moved.

## Trend, forecast and the per-student marks

`attendance_report.load` gives totals per student; `load_marks` gives every
(student, class) mark - `Y`/`N`/`E`/`NA`/blank - which the Student lookup page
and the trend chart read. `batch_trend` turns marks into turnout per class and
a running (cumulative) percentage per batch; a class counts as held once anyone
has a mark, so the pre-dated future columns stay off the chart. On the combined
sheet the batch comes from its `Batch` column, so one read gives both lines.

The at-risk forecast (`attendance_report.at_risk`) lists students **on or above**
the threshold who would fall below it by missing the next N classes. Its
"can miss" figure is `mailer.misses_allowed`, the mirror of
`classes_to_recover` - same module, same arithmetic, so the forecast and the
shortfall notice cannot disagree. The horizon is capped at the classes left.

## Gotcha: cohort folder ↔ register sheet is matched by name

`out/english` reads the `English` sheet because the names match, ignoring case
(`attendance_report.sheet_for_cohort`, `checks.sheet_for`). `checks.COHORT_SHEET`
used to be a hardcoded `{"english": "English", "hindi": "Hindi"}` table, so any
other course failed "Register sheet present" out of the box; it is now only an
override for a folder that can't be named after its sheet. Likewise
`reorder_block_in_roll_order(mode="register")` no longer guesses the sheet from
the room name: given no `reg_order` it ranks over every batch sheet (rolls are
unique across batches, so the order within a block is the same). The Move page's
reorder defaults to register order - the two roll-number orders it offered
before put blocks out of the order the health check verifies.

## Gotcha: the Class Attendance workbook's columns drift

`Class Attendance <CODE>.xlsx` is maintained by hand
outside this toolkit, and its column layout has changed more than once
during this project (an `Email` column has been added and removed, shifting
every column after it). **`attendance_report.py` finds `Roll Number` /
`Name` / `%` by header text in row 3, not by fixed column index** - this
was a real bug once (hardcoded `NAME_COL`/`FIRST_SESSION_COL` silently read
the wrong cells after the layout changed, and "names" came out as raw
percentage numbers). Don't reintroduce hardcoded column positions here or
in any new code that reads this workbook.

The workbook's own `%` column is a live Excel formula
(`Y/(Y+N)*100` over filled cells, ignoring blanks and `N/A`/holiday
columns) - `attendance_report.py` recomputes this in Python rather than
trusting the cached cell value, so a summary built right after a hand-edit
doesn't need Excel to have reopened-and-resaved the file first.

## Rooms (`config.py`)

- **LHC110** (English) - blocks A–D front (6 wide, B/C have wedge taper
  seats as the transition-pool reserve), E–G rear gallery (10 wide). 586
  total capacity, 20 wedge/transition seats. Branch groups:
  A=BB/CI/MA, B=PH/MT/ME, C=CM/EE/CY, D=AE/CS, F=CH/EC (G has no group,
  catch-all only).
- **LHC2 101** (Hindi) - blocks A–D, 4×11 each, no wedge seats at all. 176
  total capacity. Branch groups: A=CS/ME, B=EC/AE/CH, C=CI/EE/BB,
  D=CM/MA/PH/CY (MT has no group, catch-all only).
- Default seating mode is **block, not seat** - a student is allotted a
  block and can sit anywhere inside it; `Seat`/`SeatRow`/`SeatCol` exist as
  a record and a print/attendance order, not a literal chair assignment.
  This is *why* reshuffling seat numbers within a block for reordering
  purposes is safe.

## What to exclude when transferring to another machine

- `__pycache__/` - Python bytecode cache, regenerates automatically, can
  be stale/incompatible across Python versions.
- `.venv/`, `.venv-Darwin/`, `.venv-Linux/`, `.venv-Windows/` - never copy a
  venv between machines; run `./setup.sh` (or `setup.bat`) on the target
  instead. Deleting another machine's `.venv-*` inside the Dropbox folder
  deletes it *on that machine too* - leave the other OS's dirs alone.
- The Windows-only launcher files, if the target is Linux-only (harmless
  to keep, just inert).

Everything else - `out/`, every roster/config `.xlsx`, the Class Attendance
workbook, `*.py`, `README.md`, this file - should transfer as-is; nothing
in the Python code has a hardcoded Windows path.
