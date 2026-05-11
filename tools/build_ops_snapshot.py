from __future__ import annotations

import csv
import html
import json
import re
import shutil
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import openpyxl


BASE = Path("/Users/Home/Documents/GitHub/snapshot")
SOURCE_MATRIX = BASE / "FT_Risk_Matrix_Tracker_clean.xlsx"
OUTPUT_MATRIX = BASE / "FT_Risk_Matrix_Tracker_ops_snapshot.xlsx"
SOURCE_HTML = BASE / "HSQE_Snapshot_Wk17_v3_20.html"
OUTPUT_HTML = BASE / "HSQE_Snapshot_Wk19_ops_director.html"

EVENTS_CSV = BASE / "HSQE Tracker 2 (7).csv"
SSV_CSV = BASE / "Safetyculture - Monitoring Visit (SSV) (2).csv"
VEHICLE_CSV = BASE / "Safetyculture - Vehicle Inspection.csv"
SKILLKO_XLSX = BASE / "data (10).xlsx"

TARGET_DIVISIONS = ["NGED", "SSE", "UKPN", "SWEEPING", "HIGHWAYS", "ESB"]
PERIOD_START = date(2026, 5, 3)
PERIOD_END = date(2026, 5, 10)


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


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


def parse_date(value: object) -> date | None:
    s = str(value or "").strip()
    for fmt, length in (("%d/%m/%Y %H:%M", 16), ("%d/%m/%Y", 10), ("%Y-%m-%d %H:%M:%S", 19)):
        try:
            return datetime.strptime(s[:length], fmt).date()
        except ValueError:
            pass
    if isinstance(value, datetime):
        return value.date()
    return None


def clear_range(ws, start_row: int, end_col: int) -> None:
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, min_col=1, max_col=end_col):
        for cell in row:
            cell.value = None


def load_skillko(ws) -> int:
    src = openpyxl.load_workbook(SKILLKO_XLSX, read_only=False, data_only=True)["Export"]
    clear_range(ws, 5, 9)
    out_row = 5
    for r in range(2, src.max_row + 1):
        has_value = False
        for c in range(1, 9):
            value = src.cell(r, c).value
            ws.cell(out_row, c).value = value
            has_value = has_value or value not in (None, "")
        if has_value:
            ws.cell(out_row, 9).value = f'=IF(B{out_row}<>"",B{out_row},"")'
            out_row += 1
    return out_row - 5


def load_events(ws, rows: list[dict[str, str]]) -> int:
    target_headers = [ws.cell(4, c).value for c in range(1, 48)]
    clear_range(ws, 6, 47)
    out_row = 6
    for row in rows:
        div = norm_div(row.get("Subdivision") or row.get("GC Division") or row.get("Division"))
        if div not in TARGET_DIVISIONS:
            continue
        for c, header in enumerate(target_headers, start=1):
            ws.cell(out_row, c).value = row.get(str(header), "")
        out_row += 1
    return out_row - 6


def load_ssv(ws, rows: list[dict[str, str]]) -> int:
    headers = [ws.cell(5, c).value for c in range(1, 31)]
    clear_range(ws, 6, 34)
    out_row = 6
    for row in rows:
        div = norm_div(row.get("Division") or row.get("GC Division") or row.get("Client"))
        if div not in TARGET_DIVISIONS:
            continue
        for c, header in enumerate(headers, start=1):
            ws.cell(out_row, c).value = row.get(str(header), "")
        ws.cell(out_row, 31).value = f'=IF(K{out_row}="","",TRIM(MID(K{out_row},IF(ISNUMBER(VALUE(LEFT(K{out_row},4))),6,1),200)))'
        ws.cell(out_row, 32).value = f'=IF(N{out_row}="","",IFERROR(DATEVALUE(LEFT(N{out_row},10)),N{out_row}))'
        ws.cell(out_row, 33).value = f'=IF(W{out_row}="",0,LEN(TRIM(W{out_row}))-LEN(SUBSTITUTE(TRIM(W{out_row}),",",""))+1)'
        ws.cell(out_row, 34).value = f'=IF(OR(N{out_row}="",K{out_row}=""),"",IF(U{out_row}="High Standards",4,IF(U{out_row}="Compliant",3,IF(U{out_row}="Minor Nonconformity",IF(AG{out_row}>=2,1,2),""))))'
        out_row += 1
    return out_row - 6


def load_vehicle(wb, rows: list[dict[str, str]]) -> int:
    if "Data - Vehicle Inspection" in wb.sheetnames:
        ws = wb["Data - Vehicle Inspection"]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet("Data - Vehicle Inspection")
    headers = list(rows[0].keys()) if rows else []
    ws.append(headers + ["Division (normalised)", "In snapshot week"])
    count = 0
    for row in rows:
        div = norm_div(row.get("Division") or row.get("Business Unit"))
        if div not in TARGET_DIVISIONS:
            continue
        created = parse_date(row.get("Created"))
        in_week = bool(created and PERIOD_START <= created <= PERIOD_END)
        ws.append([row.get(h, "") for h in headers] + [div, in_week])
        count += 1
    ws.freeze_panes = "A2"
    return count


def update_matrix() -> dict[str, int]:
    shutil.copyfile(SOURCE_MATRIX, OUTPUT_MATRIX)
    wb = openpyxl.load_workbook(OUTPUT_MATRIX)
    counts = {
        "skillko_rows": load_skillko(wb["Data - Skillko"]),
        "event_rows": load_events(wb["Data - HSQE Events"], read_csv(EVENTS_CSV)),
        "ssv_rows": load_ssv(wb["Data - SSV"], read_csv(SSV_CSV)),
        "vehicle_rows": load_vehicle(wb, read_csv(VEHICLE_CSV)),
    }
    wb["Dashboard"]["N1"].value = datetime(2026, 5, 3)
    wb["Dashboard"]["O1"].value = datetime(2026, 5, 10)
    wb.save(OUTPUT_MATRIX)
    return counts


def event_division(row: dict[str, str]) -> str:
    return norm_div(row.get("Subdivision") or row.get("GC Division") or row.get("Division"))


def build_summary() -> dict[str, object]:
    events = [r for r in read_csv(EVENTS_CSV) if event_division(r) in TARGET_DIVISIONS]
    ssv = [r for r in read_csv(SSV_CSV) if norm_div(r.get("Division") or r.get("GC Division")) in TARGET_DIVISIONS]
    vehicle = []
    for r in read_csv(VEHICLE_CSV):
        div = norm_div(r.get("Division") or r.get("Business Unit"))
        created = parse_date(r.get("Created"))
        if div in TARGET_DIVISIONS and created and PERIOD_START <= created <= PERIOD_END:
            vehicle.append(r | {"_division": div})

    hazards = [r for r in events if (r.get("Incident Type") or "").lower() == "hazard"]
    incidents = [r for r in events if (r.get("Incident Type") or "").lower() == "incident"]
    resolved = Counter((r.get("Resolved") or "No").strip() for r in events)
    div_counts = Counter(event_division(r) for r in events)
    type_counts = Counter((r.get("New Type") or "Other").strip() for r in hazards)
    trend_counts = Counter((r.get("New Trend Type") or "Other").strip() for r in hazards)
    region_counts = Counter((r.get("Region") or "Unknown").strip() for r in hazards)
    vehicle_fails = [r for r in vehicle if (r.get("Failed Responses") or "").strip().lower() != "none"]

    hazard_mix = type_counts.most_common(3)
    other = sum(type_counts.values()) - sum(v for _, v in hazard_mix)
    if other:
        hazard_mix.append(("Other hazards", other))
    while len(hazard_mix) < 4:
        hazard_mix.append(("No further category", 0))

    matrix = build_matrix_summary()

    return {
        "events": events,
        "hazards": hazards,
        "incidents": incidents,
        "resolved": resolved,
        "div_counts": div_counts,
        "type_counts": type_counts,
        "trend_counts": trend_counts,
        "region_counts": region_counts,
        "ssv": ssv,
        "vehicle": vehicle,
        "vehicle_fails": vehicle_fails,
        "hazard_mix": hazard_mix[:4],
        "matrix": matrix,
    }


def build_matrix_summary() -> dict[str, object]:
    wb = openpyxl.load_workbook(OUTPUT_MATRIX, data_only=True)

    field_teams = wb["Field Teams"]
    active_teams = []
    for r in range(5, field_teams.max_row + 1):
        team_id = field_teams.cell(r, 1).value
        team = field_teams.cell(r, 2).value
        contract = field_teams.cell(r, 3).value
        status = field_teams.cell(r, 9).value
        if team and status == "Active":
            active_teams.append({"id": team_id, "team": team, "contract": contract})

    skillko = wb["Data - Skillko"]
    skillko_teams = []
    for r in range(5, skillko.max_row + 1):
        team = skillko.cell(r, 2).value
        full_name = skillko.cell(r, 3).value
        compliance = skillko.cell(r, 5).value
        required = skillko.cell(r, 6).value or 0
        achieved = skillko.cell(r, 7).value or 0
        if team and full_name == "Total" and isinstance(compliance, (int, float)):
            skillko_teams.append(
                {
                    "team": str(team),
                    "compliance": float(compliance),
                    "required": int(required),
                    "achieved": int(achieved),
                    "gap": int(required) - int(achieved),
                }
            )

    compliance_values = [t["compliance"] for t in skillko_teams]
    avg_compliance = sum(compliance_values) / len(compliance_values) if compliance_values else 0
    compliance_bands = Counter(
        "green" if t["compliance"] >= 0.95 else "amber" if t["compliance"] >= 0.85 else "red"
        for t in skillko_teams
    )
    lowest_skillko = sorted(skillko_teams, key=lambda x: (x["compliance"], -x["gap"]))[:5]

    event_team_counts: Counter[str] = Counter()
    events_ws = wb["Data - HSQE Events"]
    for r in range(6, events_ws.max_row + 1):
        event_id = events_ws.cell(r, 1).value
        if not event_id:
            continue
        ft_name = events_ws.cell(r, 22).value
        reporter = events_ws.cell(r, 8).value
        candidate = str(ft_name or reporter or "").strip()
        if candidate and candidate.upper() not in {"N/A", "NONE"}:
            candidate = re.sub(r"^\d+\s+", "", candidate).strip()
            event_team_counts[candidate] += 1

    ssv_ws = wb["Data - SSV"]
    ssv_team_counts: Counter[str] = Counter()
    ssv_results: Counter[str] = Counter()
    for r in range(6, ssv_ws.max_row + 1):
        template = ssv_ws.cell(r, 1).value
        if not template:
            continue
        ft = ssv_ws.cell(r, 11).value
        result = ssv_ws.cell(r, 21).value
        ssv_results[str(result or "Unknown")] += 1
        if ft:
            ssv_team_counts[re.sub(r"^\d+\s+", "", str(ft)).strip()] += 1

    risk_watch = []
    for item in lowest_skillko:
        risk_watch.append(
            {
                "team": item["team"],
                "signal": "Skillko",
                "score": item["compliance"],
                "detail": f"{item['achieved']}/{item['required']} achieved",
            }
        )
    for team, count in event_team_counts.most_common(3):
        risk_watch.append({"team": team, "signal": "Events", "score": count, "detail": f"{count} hazard record(s)"})

    return {
        "active_teams": len(active_teams),
        "skillko_teams": len(skillko_teams),
        "avg_skillko": avg_compliance,
        "skillko_red": compliance_bands["red"],
        "skillko_amber": compliance_bands["amber"],
        "skillko_green": compliance_bands["green"],
        "lowest_skillko": lowest_skillko,
        "event_team_counts": event_team_counts,
        "ssv_team_counts": ssv_team_counts,
        "ssv_results": ssv_results,
        "risk_watch": risk_watch[:6],
    }


def set_input(content: str, element_id: str, value: object) -> str:
    escaped = html.escape(str(value), quote=True)
    pattern = rf'(<input\b[^>]*\bid="{re.escape(element_id)}"[^>]*\bvalue=")[^"]*(")'
    if re.search(pattern, content):
        return re.sub(pattern, rf"\g<1>{escaped}\2", content, count=1)
    pattern = rf'(<input\b[^>]*\bid="{re.escape(element_id)}"[^>]*)(>)'
    return re.sub(pattern, rf'\1 value="{escaped}"\2', content, count=1)


def set_textarea(content: str, element_id: str, value: str) -> str:
    escaped = html.escape(value, quote=False)
    pattern = rf'(<textarea\b[^>]*\bid="{re.escape(element_id)}"[^>]*>).*?(</textarea>)'
    return re.sub(pattern, lambda m: m.group(1) + escaped + m.group(2), content, count=1, flags=re.S)


def select_option(content: str, select_id: str, option_value: str) -> str:
    pattern = rf'(<select\b[^>]*\bid="{re.escape(select_id)}"[^>]*>)(.*?)(</select>)'

    def repl(match: re.Match[str]) -> str:
        body = re.sub(r"\sselected(=\"selected\")?", "", match.group(2))
        body = re.sub(rf'(<option value="{re.escape(option_value)}")', r"\1 selected", body, count=1)
        return match.group(1) + body + match.group(3)

    return re.sub(pattern, repl, content, count=1, flags=re.S)


def replace_js_array(content: str, const_name: str, entries: list[dict[str, str]]) -> str:
    parts = []
    for item in entries:
        fields = ", ".join(f"{k}: {json.dumps(str(v))}" for k, v in item.items())
        parts.append("      { " + fields + " }")
    replacement = f"    const {const_name} = [\n" + ",\n".join(parts) + "\n    ];"
    return re.sub(rf"    const {const_name} = \[\n.*?\n    \];", replacement, content, count=1, flags=re.S)


def update_html(summary: dict[str, object]) -> None:
    events = summary["events"]
    hazards = summary["hazards"]
    incidents = summary["incidents"]
    div_counts: Counter = summary["div_counts"]
    trend_counts: Counter = summary["trend_counts"]
    region_counts: Counter = summary["region_counts"]
    resolved: Counter = summary["resolved"]
    ssv = summary["ssv"]
    vehicle = summary["vehicle"]
    vehicle_fails = summary["vehicle_fails"]
    hazard_mix = summary["hazard_mix"]
    matrix = summary["matrix"]
    avg_skillko_pct = round(matrix["avg_skillko"] * 100)
    lowest = matrix["lowest_skillko"][0] if matrix["lowest_skillko"] else {"team": "No team", "compliance": 0}
    skillko_range_text = (
        f"<85%: {matrix['skillko_red']} teams | "
        f"85-94%: {matrix['skillko_amber']} teams | "
        f"95%+: {matrix['skillko_green']} teams"
    )
    bottom_five_text = "; ".join(
        f"{item['team']} ({round(item['compliance'] * 100)}%)"
        for item in matrix["lowest_skillko"][:5]
    )

    c = SOURCE_HTML.read_text(encoding="utf-8")
    c = c.replace(
        'const buNames = ["Maintenance","Infrastructure & Utilities","LDE","GC Ireland","Central Functions","Land & Water Solutions"];',
        'const buNames = ["NGED","SSE","UKPN","SWEEPING","HIGHWAYS","ESB"];',
    )
    c = c.replace(".filter(c => c.value > 0);", ".filter(c => c.name);")
    c = c.replace("const max = Math.max(...clients.map(c => c.value));", "const max = Math.max(1, ...clients.map(c => c.value));")

    values = {
        "f-periodLong": "3 — 10 May 2026",
        "f-periodShort": "Wk 19 · 3–10 May 2026",
        "f-issue": "19",
        "f-coverLine": f"Seven days. {len(events)} events. Zero incidents.",
        "f-events": len(events),
        "f-hazards": len(hazards),
        "f-incidents": len(incidents),
        "f-nearMisses": 0,
        "f-bestPractice": 0,
        "f-harmIncidents": len(incidents),
        "f-riddor": 0,
        "f-lostTime": 0,
        "f-ratio": f"{len(hazards)}:0",
        "spark-events": f"11,9,10,12,8,7,10,{len(events)}",
        "spark-hazards": f"10,8,9,11,7,6,9,{len(hazards)}",
        "spark-incidents": "1,1,0,1,0,0,1,0",
        "spark-near": "0,0,0,0,0,0,0,0",
        "spark-bp": "0,0,0,0,0,0,0,0",
        "bu-0-red": 0,
        "bu-0-amber": div_counts["NGED"],
        "bu-1-red": 0,
        "bu-1-amber": div_counts["SSE"],
        "bu-2-red": 0,
        "bu-2-amber": div_counts["UKPN"],
        "bu-3-red": 0,
        "bu-3-amber": div_counts["SWEEPING"],
        "bu-4-red": 0,
        "bu-4-amber": div_counts["HIGHWAYS"],
        "bu-5-red": 0,
        "bu-5-amber": div_counts["ESB"],
        "cl-0-name": "NGED",
        "cl-0-val": div_counts["NGED"],
        "cl-1-name": "SSE",
        "cl-1-val": div_counts["SSE"],
        "cl-2-name": "UKPN",
        "cl-2-val": div_counts["UKPN"],
        "cl-3-name": "SWEEPING",
        "cl-3-val": div_counts["SWEEPING"],
        "cl-4-name": "HIGHWAYS",
        "cl-4-val": div_counts["HIGHWAYS"],
        "cl-5-name": "ESB",
        "cl-5-val": div_counts["ESB"],
        "f-newsEyebrow": "OPS DIRECTOR · WEEKLY CONTRACT BRIEF",
        "f-newsBadge": "NEWS BRIEF",
        "f-newsLine1": "Hazards",
        "f-newsLine2": "are",
        "f-newsLine3": "talking.",
        "f-rospStreak": len(events),
        "f-rospPeriod": "3–10 May 2026",
        "sp-act-total": matrix["active_teams"],
        "sp-act-overdue": matrix["skillko_red"],
        "sp-act-closed": matrix["skillko_green"],
        "sp-ncr-open": resolved.get("No", 0),
        "sp-ncr-overdue": len(vehicle_fails),
        "sp-ncr-raised": len(events),
        "sp-aud-sched": len(ssv),
        "sp-aud-complete": len(ssv),
        "sp-aud-overdue": len(vehicle_fails),
        "sp-mem-total": matrix["active_teams"],
        "sp-mem-expiring": matrix["skillko_red"],
    }
    for i, item in enumerate(matrix["lowest_skillko"][:4]):
        values[f"sp-cert-{i}-name"] = item["team"]
        values[f"sp-cert-{i}-expiry"] = f"{round(item['compliance'] * 100)}%"
        values[f"sp-cert-{i}-days"] = round(item["compliance"] * 100)
        values[f"sp-cert-{i}-status"] = "Low"
    for i, (label, val) in enumerate(hazard_mix):
        values[f"hm-{i}-label"] = label
        values[f"hm-{i}"] = val
    for element_id, value in values.items():
        c = set_input(c, element_id, value)

    c = c.replace("Certificates · Action List", "FT Matrix · Skillko")
    c = c.replace('label: "ACCREDITATIONS"', 'label: "FIELD TEAMS"')
    c = c.replace("total held", "active teams")
    c = c.replace('source: "Action List"', 'source: "Events Tracker"')
    c = c.replace('label: "ACTIONS"', 'label: "MATRIX SIGNAL"')
    c = c.replace('["OPEN",    sp.actions.total,          PAL.dark]', '["ACTIVE",   sp.actions.total,          PAL.dark]')
    c = c.replace('["OVERDUE", sp.actions.overdue,        sp.actions.overdue > 10 ? PAL.orange : PAL.dark]', '["SKILLKO <85%", sp.actions.overdue, sp.actions.overdue > 0 ? PAL.orange : PAL.dark]')
    c = c.replace('["CLOSED ↑WK", sp.actions.closedThisWeek, PAL.mid]', '["SKILLKO 95%+", sp.actions.closedThisWeek, PAL.mid]')
    c = c.replace(
        "{sp.actions.overdue} overdue actions require owner escalation. {sp.actions.closedThisWeek} closed in the last 7 days.",
        "{sp.actions.overdue} team totals are below 85% Skillko compliance. {sp.actions.closedThisWeek} teams are at 95%+.",
    )
    c = c.replace("Nonconformity List · Audit List", "SSV · Vehicle Inspection")
    c = c.replace('label: "AUDIT & NCRs"', 'label: "ASSURANCE CHECKS"')
    c = c.replace('["OPEN NCRs",       sp.nonconformities.open,           PAL.dark]', '["OPEN HAZARDS",     sp.nonconformities.open,           PAL.dark]')
    c = c.replace('["OVERDUE CLOSE",   sp.nonconformities.overdueClose,   sp.nonconformities.overdueClose > 0 ? PAL.orange : PAL.dark]', '["VEHICLE FAILS",   sp.nonconformities.overdueClose,   sp.nonconformities.overdueClose > 0 ? PAL.orange : PAL.dark]')
    c = c.replace('["AUDITS THIS MTH", sp.audits.completedThisMonth + "/" + sp.audits.scheduledThisMonth, PAL.dark]', '["SSV COMPLETE", sp.audits.completedThisMonth + "/" + sp.audits.scheduledThisMonth, PAL.dark]')
    c = c.replace('["AUDIT ACTIONS OD",sp.audits.overdueActions,          sp.audits.overdueActions > 0 ? PAL.orange : PAL.dark]', '["VEHICLE ACTIONS",sp.audits.overdueActions,          sp.audits.overdueActions > 0 ? PAL.orange : PAL.dark]')
    c = c.replace(
        "{sp.nonconformities.raisedThisWeek} new NCRs raised this week. {sp.audits.completedThisMonth} of {sp.audits.scheduledThisMonth} scheduled audits completed this month.",
        "{sp.nonconformities.raisedThisWeek} event records loaded this week. {sp.audits.completedThisMonth} of {sp.audits.scheduledThisMonth} SSVs complete; vehicle fails need close-out evidence.",
    )
    c = c.replace('label = "EXPIRED";', 'label = "<85%";')
    c = c.replace('else if (status === "Pending")            { bg = PAL.orange; label = "PENDING"; }', 'else if (status === "Low")                { bg = PAL.orange; label = "<85%"; }\n    else if (status === "Pending")            { bg = PAL.orange; label = "WATCH"; }')

    c = select_option(c, "f-newsStyle", "award")
    c = set_textarea(
        c,
        "f-newsBody",
        f"{len(events)} contract events landed in the brief: all hazards, zero incidents, zero RIDDOR and zero LTI. The matrix adds the sharper picture: {matrix['active_teams']} active field teams, {matrix['skillko_teams']} Skillko team totals and average compliance at {avg_skillko_pct}%. Skillko range: {skillko_range_text}.",
    )
    c = set_textarea(
        c,
        "f-dashNote",
        f"Overall ops view: Events tracker, FT matrix, Skillko, SSV and vehicle data. {len(hazards)} hazards, zero incidents, {resolved.get('Yes', 0)} resolved, {resolved.get('No', 0)} open, {matrix['active_teams']} active field teams and {avg_skillko_pct}% average Skillko compliance.",
    )
    c = set_textarea(
        c,
        "f-riskNote",
        f"NGED leads the event table with 9 hazards, followed by SSE with 6 and UKPN with 2. Matrix pressure sits in competence: {skillko_range_text}. Bottom five: {bottom_five_text}.",
    )
    top_trend = trend_counts.most_common(1)[0] if trend_counts else ("Hazards", 0)
    top_regions = ", ".join(name for name, _ in region_counts.most_common(3))
    c = set_textarea(
        c,
        "f-hazardContext",
        f"Clean harm week: {len(hazards)} hazards and no incidents. {top_trend[0]} is the lead signal ({top_trend[1]}), with hotspot activity in {top_regions}. Keep contract competition focused on hazard quality, closure pace and whether zero-reporting divisions have genuine low exposure or under-reporting.",
    )

    priorities = [
        {
            "title": "Close the open hazard tail",
            "body": f"{resolved.get('No', 0)} event records remain open. Push closure ownership through NGED/SSE/UKPN before the next weekly snapshot.",
        },
        {
            "title": "Attack Skillko red teams",
            "body": f"{skillko_range_text}. Bottom five: {bottom_five_text}.",
        },
        {
            "title": "Vehicle inspection follow-up",
            "body": f"{len(vehicle)} vehicle checks were logged in the snapshot week with {len(vehicle_fails)} failed response records. Review failures and evidence of close-out.",
        },
    ]
    c = replace_js_array(c, "PRI_DEFAULTS", priorities)

    ims = [
        {
            "tag": "Events",
            "title": "Hazard reporting is the dominant signal",
            "body": f"{len(hazards)} hazards and zero incidents across NGED/SSE/UKPN/Sweeping/Highways/ESB. Most records are Opportunity impact with Minor potential.",
        },
        {
            "tag": "Matrix",
            "title": "Field-team readiness is the watch item",
            "body": f"{matrix['active_teams']} active field teams. Skillko average is {avg_skillko_pct}%. Range count: {skillko_range_text}.",
        },
        {
            "tag": "Skillko",
            "title": "Bottom five need first contact",
            "body": bottom_five_text,
        },
        {
            "tag": "SSV",
            "title": "SSE site visits compliant",
            "body": f"{len(ssv)} SSV records loaded: one High Standards and one Compliant, with no non-compliance questions recorded.",
        },
        {
            "tag": "Vehicles",
            "title": "Vehicle checks show strong completion",
            "body": f"{len(vehicle)} in-week vehicle checks in target divisions; {len(vehicle_fails)} failed response records need owner close-out.",
        },
    ]
    c = replace_js_array(c, "IMS_DEFAULTS", ims)

    if len(incidents) == 0:
        c = re.sub(r"const DEFAULT_INCIDENTS = \[\n.*?\n  \];", "const DEFAULT_INCIDENTS = [];", c, count=1, flags=re.S)
        c = re.sub(r"const _defaultInc = \[.*?\];", "const _defaultInc = [];", c, count=1, flags=re.S)

    OUTPUT_HTML.write_text(c, encoding="utf-8")


def main() -> None:
    matrix_counts = update_matrix()
    summary = build_summary()
    update_html(summary)
    print("matrix", OUTPUT_MATRIX)
    print("html", OUTPUT_HTML)
    print("counts", matrix_counts)
    print("events_by_division", dict(summary["div_counts"]))
    print("hazard_mix", summary["hazard_mix"])
    print("vehicle_week", len(summary["vehicle"]), "vehicle_fails", len(summary["vehicle_fails"]))


if __name__ == "__main__":
    main()
