from __future__ import annotations

import csv
import html
import json
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill


BASE = Path("/Users/Home/Documents/GitHub/snapshot")
OUT_HTML = BASE / "HSQE_Ops_Business_Unit_Snapshot.html"
OUT_XLSX = BASE / "HSQE_Ops_Snapshot_Data.xlsx"

EVENTS = BASE / "HSQE Tracker 2 (7).csv"
SSV = BASE / "Safetyculture - Monitoring Visit (SSV) (2).csv"
VEHICLE = BASE / "Safetyculture - Vehicle Inspection.csv"
SKILLKO = BASE / "data (10).xlsx"

DIVISIONS = ["NGED", "SSE", "UKPN", "SWEEPING", "HIGHWAYS", "ESB"]
PERIOD_START = date(2026, 5, 3)
PERIOD_END = date(2026, 5, 10)


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


def parse_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    s = str(value or "").strip()
    for fmt, length in (("%d/%m/%Y %H:%M", 16), ("%d/%m/%Y", 10), ("%Y-%m-%d %H:%M:%S", 19)):
        try:
            return datetime.strptime(s[:length], fmt).date()
        except ValueError:
            pass
    return None


def num(value: object) -> float:
    try:
        return float(str(value or "0").strip())
    except ValueError:
        return 0.0


def norm_div(value: object) -> str:
    s = str(value or "").upper()
    if "NGED" in s or "NATIONAL GRID" in s:
        return "NGED"
    if "SSE" in s:
        return "SSE"
    if "UKPN" in s or "UK POWER" in s:
        return "UKPN"
    if "SWEEP" in s:
        return "SWEEPING"
    if "HIGHWAY" in s:
        return "HIGHWAYS"
    if "ESB" in s:
        return "ESB"
    return str(value or "").strip() or "Unknown"


def clean_team(value: object) -> str:
    s = str(value or "").strip()
    s = re.sub(r"^\d+\s*[:\-]?\s*", "", s)
    return s or "Unassigned"


def count_findings(value: object) -> int:
    s = str(value or "").strip()
    if not s:
        return 0
    parts = [p for p in re.split(r"[\n,;]+", s) if p.strip()]
    return max(1, len(parts))


def load_skillko() -> tuple[list[dict], list[dict]]:
    wb = openpyxl.load_workbook(SKILLKO, data_only=True)
    ws = wb["Export"]
    teams: list[dict] = []
    staff: list[dict] = []
    current_team = ""
    for r in range(2, ws.max_row + 1):
        field_team = ws.cell(r, 2).value
        full_name = ws.cell(r, 3).value
        role = ws.cell(r, 4).value
        compliance = ws.cell(r, 5).value
        required = int(ws.cell(r, 6).value or 0)
        achieved = int(ws.cell(r, 7).value or 0)
        if field_team:
            current_team = str(field_team).strip()
        if not isinstance(compliance, (int, float)):
            continue
        if field_team and full_name == "Total":
            teams.append(
                {
                    "name": str(field_team).strip(),
                    "compliance": float(compliance),
                    "required": required,
                    "achieved": achieved,
                    "gap": required - achieved,
                }
            )
        elif full_name and str(full_name).strip().lower() != "total" and str(role or "").strip() == "Total":
            staff.append(
                {
                    "name": str(full_name).strip(),
                    "team": current_team,
                    "compliance": float(compliance),
                    "required": required,
                    "achieved": achieved,
                    "gap": required - achieved,
                }
            )
    return teams, staff


def build_data() -> dict:
    events = [r for r in read_csv(EVENTS) if norm_div(r.get("GC Division") or r.get("Subdivision")) in DIVISIONS]
    ssv = [r for r in read_csv(SSV) if norm_div(r.get("Division") or r.get("GC Division")) in DIVISIONS]
    vehicle_all = read_csv(VEHICLE)
    vehicle_week = []
    for r in vehicle_all:
        d = parse_date(r.get("Created"))
        div = norm_div(r.get("Division"))
        if d and PERIOD_START <= d <= PERIOD_END and div in DIVISIONS:
            vehicle_week.append(r)

    teams, staff = load_skillko()
    skillko_avg = sum(t["compliance"] for t in teams) / len(teams) if teams else 0
    skillko_bands = Counter("<85%" if t["compliance"] < 0.85 else "85-94%" if t["compliance"] < 0.95 else "95%+" for t in teams)

    division = {
        d: {
            "division": d,
            "ssv": 0,
            "ssv_findings": 0,
            "amber": 0,
            "red": 0,
            "skillko_pct": skillko_avg,
            "vehicle_checks": 0,
            "vehicle_fails": 0,
            "checklist_minutes": 0,
        }
        for d in DIVISIONS
    }

    for r in events:
        div = norm_div(r.get("GC Division") or r.get("Subdivision"))
        impact = str(r.get("Impact") or "").strip().lower()
        if impact == "opportunity":
            division[div]["amber"] += 1
        elif impact in {"minor", "moderate", "major"}:
            division[div]["red"] += 1
        division[div]["checklist_minutes"] += num(r.get("Time Taken"))

    for r in ssv:
        div = norm_div(r.get("Division") or r.get("GC Division"))
        division[div]["ssv"] += 1
        division[div]["ssv_findings"] += count_findings(r.get("NC Questions"))
        division[div]["checklist_minutes"] += num(r.get("Mins taken"))

    for r in vehicle_week:
        div = norm_div(r.get("Division"))
        division[div]["vehicle_checks"] += 1
        if str(r.get("Failed Responses") or "").strip().lower() != "none":
            division[div]["vehicle_fails"] += 1

    top_teams = sorted(teams, key=lambda x: (-x["compliance"], -x["required"], x["name"]))[:5]
    bottom_teams = sorted(teams, key=lambda x: (x["compliance"], -x["gap"], x["name"]))[:5]
    top_staff = sorted(staff, key=lambda x: (-x["compliance"], -x["required"], x["name"]))[:5]
    bottom_staff = sorted(staff, key=lambda x: (x["compliance"], -x["gap"], x["name"]))[:5]

    staff_activity = Counter()
    for r in events:
        staff_activity[str(r.get("Reported By") or "").strip()] += 1
    for r in ssv:
        staff_activity[str(r.get("Completed By") or "").strip()] += 1
    for r in vehicle_week:
        staff_activity[str(r.get("Completed By") or r.get("Driver Name") or "").strip()] += 1

    team_activity = Counter()
    for r in events:
        team_activity[clean_team(r.get("DE/FT Name") or r.get("Reported By"))] += 1
    for r in ssv:
        team_activity[clean_team(r.get("FT On Site") or r.get("Completed By"))] += 1
    for r in vehicle_week:
        team_activity[clean_team(r.get("Completed By") or r.get("Driver Name"))] += 1

    total_minutes = sum(v["checklist_minutes"] for v in division.values())
    data = {
        "period": "3-10 May 2026",
        "divisions": list(division.values()),
        "events_total": len(events),
        "hazards": sum(1 for r in events if str(r.get("Incident Type") or "").lower() == "hazard"),
        "incidents": sum(1 for r in events if str(r.get("Incident Type") or "").lower() == "incident"),
        "open_events": sum(1 for r in events if str(r.get("Resolved") or "").strip().lower() != "yes"),
        "resolved_events": sum(1 for r in events if str(r.get("Resolved") or "").strip().lower() == "yes"),
        "skillko_avg": skillko_avg,
        "skillko_bands": dict(skillko_bands),
        "skillko_team_count": len(teams),
        "vehicle_week": len(vehicle_week),
        "vehicle_fails": sum(v["vehicle_fails"] for v in division.values()),
        "ssv_total": len(ssv),
        "ssv_findings": sum(v["ssv_findings"] for v in division.values()),
        "total_minutes": total_minutes,
        "top_teams": [x | {"activity": team_activity[x["name"]]} for x in top_teams],
        "bottom_teams": [x | {"activity": team_activity[x["name"]]} for x in bottom_teams],
        "top_staff": [x | {"activity": staff_activity[x["name"]]} for x in top_staff],
        "bottom_staff": [x | {"activity": staff_activity[x["name"]]} for x in bottom_staff],
        "priorities": [
            "Close the open hazard tail: 7 records remain open in the event tracker.",
            "Attack Skillko red teams: 42 team totals sit below 85%.",
            "Vehicle inspection follow-up: 2 failed response records need close-out evidence.",
        ],
    }
    return data


def pct(value: float) -> str:
    return f"{round(value * 100)}%"


def write_summary_xlsx(data: dict) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Division Summary"
    headers = ["Division", "SSV", "SSV Findings", "Amber", "Red", "Skillko %", "Vehicle Checks", "Vehicle Fails", "Checklist Minutes"]
    ws.append(headers)
    for d in data["divisions"]:
        ws.append([d["division"], d["ssv"], d["ssv_findings"], d["amber"], d["red"], pct(d["skillko_pct"]), d["vehicle_checks"], d["vehicle_fails"], d["checklist_minutes"]])

    for title, rows in [
        ("Top Teams", data["top_teams"]),
        ("Bottom Teams", data["bottom_teams"]),
        ("Top Staff", data["top_staff"]),
        ("Bottom Staff", data["bottom_staff"]),
    ]:
        sheet = wb.create_sheet(title)
        sheet.append(["Name", "Team", "Compliance", "Achieved", "Required", "Gap", "Sample Activity"])
        for r in rows:
            sheet.append([r["name"], r.get("team", ""), pct(r["compliance"]), r["achieved"], r["required"], r["gap"], r.get("activity", 0)])

    notes = wb.create_sheet("Source Notes")
    notes.append(["Source", "Use"])
    notes.append([EVENTS.name, "Events, amber/red, event time, open/resolved"])
    notes.append([SSV.name, "SSV counts, findings, SSV minutes"])
    notes.append([VEHICLE.name, "Vehicle checks and failed response records in period"])
    notes.append([SKILLKO.name, "Team and staff Skillko compliance rankings"])

    dark = "0B1F17"
    light = "B2D235"
    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor=dark)
            cell.font = Font(color=light, bold=True)
            cell.alignment = Alignment(horizontal="center")
        for col in sheet.columns:
            width = min(42, max(12, max(len(str(c.value or "")) for c in col) + 2))
            sheet.column_dimensions[col[0].column_letter].width = width
        sheet.freeze_panes = "A2"
    wb.save(OUT_XLSX)


def js_data(data: dict) -> str:
    return json.dumps(data, ensure_ascii=False)


def render_html(data: dict) -> str:
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>HSQE Ops Snapshot · Business Unit Performance</title>
<script src="https://unpkg.com/react@18/umd/react.production.min.js"></script>
<script src="https://unpkg.com/react-dom@18/umd/react-dom.production.min.js"></script>
<script src="https://unpkg.com/@babel/standalone/babel.min.js"></script>
<style>
  :root {{
    --dark:#071b14; --ink:#11241b; --lime:#B2D235; --mid:#6BA43A; --orange:#F28C28;
    --paper:#F4F1EA; --grey:#DFDDD3; --white:#fff;
    --display: Georgia, 'Times New Roman', serif; --body: Inter, Arial, sans-serif; --mono: ui-monospace, SFMono-Regular, Menlo, monospace;
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; background:#101510; color:var(--ink); font-family:var(--body); }}
  .deck {{ width:100%; }}
  .slide {{ width:100vw; min-height:56.25vw; max-height:100vh; aspect-ratio:16/9; position:relative; overflow:hidden; margin:0 auto 18px; background:var(--paper); }}
  .dark {{ background:var(--dark); color:var(--white); }}
  .grey {{ background:var(--grey); }}
  .display {{ font-family:var(--display); font-weight:700; letter-spacing:0; }}
  .mono {{ font-family:var(--mono); text-transform:uppercase; letter-spacing:.16em; }}
  .header {{ position:absolute; top:64px; left:80px; right:80px; display:flex; justify-content:space-between; align-items:center; padding-bottom:22px; border-bottom:1px solid rgba(7,27,20,.22); }}
  .dark .header {{ border-bottom-color:rgba(178,210,53,.32); }}
  .label {{ font-size:18px; color:inherit; opacity:.7; }}
  .right {{ font-size:16px; color:var(--mid); }}
  .big {{ font-size:132px; line-height:.92; margin:0; }}
  .cover-title {{ font-size:230px; line-height:.86; margin:0; }}
  .subtitle {{ font-size:28px; line-height:1.35; max-width:1180px; opacity:.8; }}
  .grid6 {{ display:grid; grid-template-columns:repeat(6,1fr); gap:12px; }}
  .card {{ background:var(--dark); color:white; padding:26px; border-top:4px solid var(--lime); }}
  .card.light {{ background:white; color:var(--ink); border-top-color:var(--mid); }}
  .metric {{ font-family:var(--display); font-size:54px; line-height:.95; }}
  .small {{ font-size:14px; opacity:.68; line-height:1.35; }}
  .table-row {{ display:grid; grid-template-columns:48px 1fr 120px 130px 130px 90px; gap:16px; align-items:center; padding:13px 0; border-bottom:1px solid rgba(7,27,20,.16); }}
  .dark .table-row {{ border-bottom-color:rgba(178,210,53,.2); }}
  .rank {{ color:var(--mid); font-size:20px; }}
  .bar {{ height:10px; background:rgba(7,27,20,.14); overflow:hidden; }}
  .bar span {{ display:block; height:100%; background:var(--lime); }}
  .pill {{ display:inline-block; padding:5px 10px; background:var(--dark); color:var(--lime); font-family:var(--mono); font-size:11px; letter-spacing:.12em; }}
  .dark .pill {{ background:var(--lime); color:var(--dark); }}
  .note {{ position:absolute; bottom:46px; left:80px; right:80px; border-top:1px solid currentColor; padding-top:16px; opacity:.7; }}
  [contenteditable] {{ outline:2px dashed transparent; }}
  [contenteditable]:focus {{ outline-color:var(--orange); background:rgba(255,255,255,.14); }}
  .edit-hint {{ position:fixed; right:18px; bottom:18px; z-index:20; background:var(--dark); color:var(--lime); border:1px solid rgba(178,210,53,.45); padding:10px 14px; font:11px var(--mono); letter-spacing:.12em; text-transform:uppercase; opacity:.74; }}
  @media print {{ .edit-hint {{ display:none; }} body {{ background:white; }} .slide {{ margin:0; page-break-after:always; }} }}
</style>
</head>
<body>
<div id="root"></div>
<script type="text/babel">
const DATA = {js_data(data)};
const P = {{ dark:'#071b14', lime:'#B2D235', mid:'#6BA43A', orange:'#F28C28', paper:'#F4F1EA', grey:'#DFDDD3' }};
const pct = v => Math.round(v * 100) + '%';
function Header({{n,label,right,dark=false}}) {{
  return <div className="header"><div className="mono label" style={{{{color:dark?P.lime:P.dark}}}}>{{String(n).padStart(2,'0')}} / {{label}}</div><div className="mono right">{{right}}</div></div>
}}
function DivisionCard({{d}}) {{
  const quality = d.amber + d.red === 0 && d.vehicle_checks === 0 && d.ssv === 0 ? 'NO SAMPLE' : d.red > 0 || d.vehicle_fails > 0 ? 'WATCH' : d.amber > 0 ? 'ACTIVE' : 'QUIET';
  return <div className="card light">
    <div className="mono" style={{{{fontSize:16,color:P.mid}}}}>{{d.division}}</div>
    <div style={{{{marginTop:22,display:'grid',gap:11}}}}>
      {{[['SSV',d.ssv],['FINDINGS',d.ssv_findings],['AMBER',d.amber],['RED',d.red],['SKILLKO',pct(d.skillko_pct)],['VEHICLE',d.vehicle_checks],['MINS',d.checklist_minutes]].map(([k,v]) =>
        <div key={{k}} style={{{{display:'flex',justifyContent:'space-between',alignItems:'baseline',gap:8}}}}>
          <span className="mono small">{{k}}</span><span className="display" style={{{{fontSize:28,color:k==='RED'&&v>0?P.orange:P.dark}}}}>{{v}}</span>
        </div>
      )}}
    </div>
    <div className="pill" style={{{{marginTop:20}}}}>{{quality}}</div>
  </div>
}}
function Ranking({{rows,type='team'}}) {{
  return <div>
    {{rows.map((r,i)=><div className="table-row" key={{r.name}}>
      <div className="display rank">{{String(i+1).padStart(2,'0')}}</div>
      <div><div className="display" style={{{{fontSize:25,lineHeight:1.05}}}}>{{r.name}}</div>{{type==='staff'&&<div className="small">{{r.team}}</div>}}</div>
      <div className="display" style={{{{fontSize:34,color:r.compliance<.85?P.orange:P.mid}}}}>{{pct(r.compliance)}}</div>
      <div><div className="mono small">Achieved</div><b>{{r.achieved}}/{{r.required}}</b></div>
      <div><div className="mono small">Gap</div><b>{{r.gap}}</b></div>
      <div><div className="mono small">Sample</div><b>{{r.activity}}</b></div>
    </div>)}}
  </div>
}}
function App() {{
  return <main className="deck">
    <section className="slide dark">
      <div style={{{{position:'absolute',top:64,left:80,right:80,display:'flex',justifyContent:'space-between',alignItems:'center',borderBottom:'1px solid rgba(178,210,53,.32)',paddingBottom:22}}}}>
        <div className="mono" style={{{{fontSize:16,color:P.lime}}}}>HSQE / OPS DIRECTOR WEEKLY BRIEF</div>
        <div className="mono" style={{{{fontSize:16,color:P.lime}}}}>{{DATA.period}}</div>
      </div>
      <div style={{{{position:'absolute',top:185,left:80,right:80}}}}>
        <h1 className="display cover-title">Contract<br/>Snapshot<span style={{{{color:P.lime}}}}>.</span></h1>
        <p className="subtitle" style={{{{color:P.lime,marginTop:28}}}}>NGED, SSE, UKPN, Sweeping, Highways and ESB. A fresh build from the tracker exports and FT matrix: performance, assurance, competence and closure in one brief. Click any text to edit.</p>
      </div>
      <div className="grid6" style={{{{position:'absolute',left:80,right:80,bottom:70}}}}>
        {{DATA.divisions.map(d=><div key={{d.division}} style={{{{borderTop:'1px solid rgba(178,210,53,.35)',paddingTop:14}}}}><div className="mono" style={{{{color:P.lime,fontSize:13}}}}>{{d.division}}</div><div className="display" style={{{{fontSize:48}}}}>{{d.amber}}<span style={{{{opacity:.35}}}}> / {{d.red}}</span></div><div className="small">amber / red</div></div>)}}
      </div>
    </section>

    <section className="slide grey">
      <Header n="2" label="Business Unit Performance" right="CONTRACT COMPETITION" />
      <div style={{{{position:'absolute',top:150,left:80,right:80}}}}>
        <h2 className="display big">Who is visible. Who is quiet.</h2>
        <p className="subtitle">NGED leads hazard visibility, SSE carries the SSV sample, UKPN dominates vehicle-check volume. Zero rows for ESB, Sweeping and Highways should be treated as a reporting/exposure question, not an automatic clean week.</p>
      </div>
      <div className="grid6" style={{{{position:'absolute',left:80,right:80,bottom:90}}}}>{{DATA.divisions.map(d=><DivisionCard key={{d.division}} d={{d}} />)}}</div>
    </section>

    <section className="slide">
      <Header n="3" label="Overall Figures" right="SIX-COLUMN DIVISION VIEW" />
      <div style={{{{position:'absolute',top:145,left:80,right:80}}}}>
        <h2 className="display" style={{{{fontSize:78,margin:0}}}}>Sampled assurance by division.</h2>
        <div className="grid6" style={{{{marginTop:32}}}}>{{DATA.divisions.map(d=><DivisionCard key={{d.division}} d={{d}} />)}}</div>
      </div>
      <div className="note mono">Total sampled checklist time: {{DATA.total_minutes}} minutes · Events: {{DATA.events_total}} · SSVs: {{DATA.ssv_total}} · Vehicle checks: {{DATA.vehicle_week}}</div>
    </section>

    <section className="slide grey">
      <Header n="4" label="Top Performers" right="TEAMS + STAFF" />
      <div style={{{{position:'absolute',top:145,left:80,right:80,display:'grid',gridTemplateColumns:'1fr 1fr',gap:54}}}}>
        <div><h2 className="display" style={{{{fontSize:58,margin:'0 0 22px'}}}}>Top 5 teams.</h2><Ranking rows={{DATA.top_teams}} /></div>
        <div><h2 className="display" style={{{{fontSize:58,margin:'0 0 22px'}}}}>Top 5 staff.</h2><Ranking rows={{DATA.top_staff}} type="staff" /></div>
      </div>
    </section>

    <section className="slide">
      <Header n="5" label="Bottom Performers" right="FIRST CONTACT LIST" />
      <div style={{{{position:'absolute',top:145,left:80,right:80,display:'grid',gridTemplateColumns:'1fr 1fr',gap:54}}}}>
        <div><h2 className="display" style={{{{fontSize:58,margin:'0 0 22px'}}}}>Bottom 5 teams.</h2><Ranking rows={{DATA.bottom_teams}} /></div>
        <div><h2 className="display" style={{{{fontSize:58,margin:'0 0 22px'}}}}>Bottom 5 staff.</h2><Ranking rows={{DATA.bottom_staff}} type="staff" /></div>
      </div>
    </section>

    <section className="slide grey">
      <Header n="6" label="Editable News Brief" right="ROSPA-STYLE LIST SLIDE" />
      <div style={{{{position:'absolute',top:150,left:80,width:'58%'}}}}>
        <div className="mono" style={{{{fontSize:18,color:P.mid}}}} contentEditable suppressContentEditableWarning>OPS DIRECTOR · NEXT MOVES</div>
        <h2 className="display" style={{{{fontSize:142,lineHeight:.9,margin:'18px 0 0'}}}} contentEditable suppressContentEditableWarning>Three things<br/><span style={{{{color:P.mid}}}}>to move.</span></h2>
        <p className="subtitle" contentEditable suppressContentEditableWarning>Editable slide: click into the headline or list text and type over it before presenting.</p>
      </div>
      <div style={{{{position:'absolute',top:180,right:80,width:'34%',background:P.dark,color:'white',padding:44,bottom:90}}}}>
        <div className="mono" style={{{{fontSize:13,color:P.lime}}}}>THE LIST</div>
        <div style={{{{marginTop:28,display:'grid',gap:24}}}}>
          {{DATA.priorities.map((p,i)=><div key={{p}} style={{{{borderTop:'1px solid rgba(178,210,53,.35)',paddingTop:18}}}}>
            <div className="display" style={{{{fontSize:36,color:P.lime}}}}>0{{i+1}}</div>
            <div style={{{{fontSize:22,lineHeight:1.35,marginTop:6}}}} contentEditable suppressContentEditableWarning>{{p}}</div>
          </div>)}}
        </div>
      </div>
    </section>
    <div className="edit-hint">Inline editable · autosaves in browser</div>
  </main>
}}
ReactDOM.createRoot(document.getElementById('root')).render(<App/>);
setTimeout(() => {{
  const selectors = [
    '.slide h1', '.slide h2', '.slide p', '.slide .mono', '.slide .small',
    '.slide .display', '.slide b', '.slide .pill', '.slide .metric'
  ].join(',');
  const nodes = Array.from(document.querySelectorAll(selectors))
    .filter(el => !el.closest('.edit-hint') && !el.querySelector('.bar'));
  nodes.forEach((el, i) => {{
    const key = 'ops-snapshot-edit-' + i;
    el.setAttribute('contenteditable', 'true');
    el.setAttribute('spellcheck', 'true');
    if (localStorage.getItem(key) !== null) el.innerHTML = localStorage.getItem(key);
    el.addEventListener('input', () => localStorage.setItem(key, el.innerHTML));
  }});
}}, 250);
</script>
</body>
</html>"""


def main() -> None:
    data = build_data()
    write_summary_xlsx(data)
    OUT_HTML.write_text(render_html(data), encoding="utf-8")
    print(OUT_HTML)
    print(OUT_XLSX)
    print("division rows", len(data["divisions"]))
    print("top team", data["top_teams"][0]["name"], pct(data["top_teams"][0]["compliance"]))
    print("bottom team", data["bottom_teams"][0]["name"], pct(data["bottom_teams"][0]["compliance"]))


if __name__ == "__main__":
    main()
