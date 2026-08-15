import json
from pathlib import Path

import pandas as pd
import webview
import keyboard
import threading

SWITCHLIST_FILE = Path("current_switchlist.xlsx")
STATUS_FILE = Path("switchlist_status.json")
COMPLETED_FILE = Path("completed_tasks.xlsx")
SCORE_FILE = Path("score_summary.xlsx")


def load_status():
    if STATUS_FILE.exists():
        try:
            return json.loads(STATUS_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return {}
    return {}


def save_status(status):
    STATUS_FILE.write_text(json.dumps(status, indent=2), encoding="utf-8")


class Api:
    def complete_task(self, row_key, day=None, time=None, partial=None):
        status = load_status()
        row_key = str(row_key)

        current = status.get(row_key, {
            "completed": False,
            "partial": "",
            "assigned": "",
            "active": False,
            "day": "",
            "time": ""
        })

        day = str(day or current.get("day", "")).strip()
        time = str(time or current.get("time", "")).strip()
        partial = str(partial or current.get("partial", "")).strip()

        currently_done = current.get("completed", False)

        if currently_done:
            current["completed"] = False
        else:
            current["completed"] = True

        current["day"] = day
        current["time"] = time
        current["partial"] = partial

        status[row_key] = current
        save_status(status)

        return render_html()
    
    def update_field(self, row_key, field, value):
        status = load_status()
        row_key = str(row_key)

        current = status.get(row_key, {
            "completed": False,
            "partial": "",
            "assigned": "",
            "active": False,
            "day": "",
            "time": ""
        })

        if field == "active":
            current[field] = bool(value)
        else:
            current[field] = str(value).strip()

        status[row_key] = current
        save_status(status)

        if field == "active":
            return render_html()

        return True
    
    def change_counter(self, counter_name, delta):
        status = load_status()

        meta = status.get("_meta", {})

        current_value = int(meta.get(counter_name, 0))
        current_value += int(delta)
        current_value = max(0, current_value)

        meta[counter_name] = current_value
        status["_meta"] = meta

        save_status(status)

        return render_html()

def task_color(task):
    task = str(task).lower()

    if "spot empties" in task:
        return "#ffe48a"

    if "pickup load" in task:
        return "#8fc7ff"

    if "deliver" in task:
        return "#9be89b"

    return "#d8d8d8"

def load_excel_safe(path):
    if path.exists():
        try:
            return pd.read_excel(path)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def build_footer_stats(active_df, status):
    completed_df = load_excel_safe(COMPLETED_FILE)
    score_df = load_excel_safe(SCORE_FILE)

    active_count = len(active_df)

    # Completed but not yet archived — checked in viewer JSON
    active_keys = set()

    if not active_df.empty:
        for _, row in active_df.iterrows():
            shipment_id = str(row.get("shipment_id", ""))
            task_number = str(row.get("task_number", ""))
            active_keys.add(f"{shipment_id}|{task_number}")

    pending_completed = sum(
        1 for key, item in status.items()
        if key in active_keys and item.get("completed", False)
    )

    archived_completed = len(completed_df)

    total_assigned = active_count + archived_completed
    outstanding = max(0, active_count - pending_completed)

    latest_score = "-"

    if not score_df.empty and "total_score" in score_df.columns:
        latest_score = str(score_df.iloc[0]["total_score"])
    
    meta = status.get("_meta", {})
    major_derailments = int(meta.get("major_derailments", 0))
    other_incidents = int(meta.get("other_incidents", 0))

    return f"""
<BR>
<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="1" WIDTH="100%">
<TR BGCOLOR="#222222">
<TD><B>Total tasks assigned:</B> {total_assigned}</TD>
<TD><B>Outstanding:</B> {outstanding}</TD>
<TD><B>Completed pending archive:</B> {pending_completed}</TD>
<TD><B>Archived completed:</B> {archived_completed}</TD>
<TD><B>Latest score:</B> {latest_score}</TD>
</TR>
</TABLE>

<TABLE BORDER="0" CELLPADDING="3" CELLSPACING="1" WIDTH="100%">
<TR BGCOLOR="#2f2f2f">
<TD>
<B>Major derailments:</B> {major_derailments}
&nbsp;
<A HREF="#" onclick="changeCounter('major_derailments', 1); return false;"><FONT COLOR="#d8d8d8">▲</FONT></A>
&nbsp;
<A HREF="#" onclick="changeCounter('major_derailments', -1); return false;"><FONT COLOR="#d8d8d8">▼</FONT></A>
&nbsp;&nbsp;&nbsp;&nbsp;
<B>Other incidents:</B> {other_incidents}
&nbsp;
<A HREF="#" onclick="changeCounter('other_incidents', 1); return false;"><FONT COLOR="#d8d8d8">▲</FONT></A>
&nbsp;
<A HREF="#" onclick="changeCounter('other_incidents', -1); return false;"><FONT COLOR="#d8d8d8">▼</FONT></A>
</TD>
</TR>
</TABLE>
"""

def render_html():
    status = load_status()

    if SWITCHLIST_FILE.exists():
        df = pd.read_excel(SWITCHLIST_FILE)
    else:
        df = pd.DataFrame()

    if not df.empty:
        task_order = {
            "Spot Empties": 1,
            "Pickup Load": 2,
            "Deliver": 3,
        }

        df["task_sort"] = df["task_type"].map(task_order).fillna(9)
        df = df.sort_values(
            by=["industry_name", "commodity", "task_sort"],
            kind="stable"
        )

        df = df.drop(columns=["task_sort"], errors="ignore")

        df.loc[
            df["origin"].astype(str).str.strip()
            == df["industry_name"].astype(str).str.strip(),
            "origin"
        ] = "-"

        df.loc[
            df["destination"].astype(str).str.strip()
            == df["industry_name"].astype(str).str.strip(),
            "destination"
        ] = "-"

    rows = []

    for i, (_, row) in enumerate(df.iterrows()):

        ready_value = row.get("ready", True)

        if pd.isna(ready_value):
            ready_flag = True
        elif isinstance(ready_value, str):
            ready_flag = ready_value.strip().lower() not in ["false", "0", "no"]
        else:
            ready_flag = bool(ready_value)
        bg = "#383838" if i % 2 == 0 else "#2f2f2f"

        shipment_id = str(row.get("shipment_id", ""))
        task_number = str(row.get("task_number", ""))
        row_key = f"{shipment_id}|{task_number}"
        safe_key = row_key.replace("|", "_")

        industry = str(row.get("industry_name", ""))
        task = str(row.get("task_type", ""))
        commodity = str(row.get("commodity", ""))
        vehicle = str(row.get("car_type", ""))
        cars = str(row.get("number_of_cars", ""))
        origin = str(row.get("origin", ""))
        destination = str(row.get("destination", ""))

        
        row_status = status.get(row_key, {})
        active_flag = row_status.get("active", False)

        effective_ready = ready_flag or active_flag

        if task.strip().lower() == "spot empties":
            commodity = "-"

        completed = status.get(row_key, {}).get("completed", False)

        row_style = ""
        done_text = "☑" if completed else "☐"
        done_color = "#515151" if completed else "#999999"

        if completed:
            row_style = ' STYLE="color:#515151;"'
            color = "#515151"

        elif not effective_ready:
            row_style = ' STYLE="color:#3A5864;"'
            color = "#3A5864"

        else:
            color = task_color(task)


        day_value = row_status.get("day", "")
        time_value = row_status.get("time", "")
        partial_value = row_status.get("partial", "")
        assigned_value = row_status.get("assigned", "")
        active_flag = row_status.get("active", False)

        rows.append(f"""
<TR BGCOLOR="{bg}" DATA-SHIPMENT="{shipment_id}" DATA-TASK="{task_number}"{row_style}>
<TD WIDTH="190" NOWRAP>{industry}</TD>
<TD WIDTH="110" NOWRAP><FONT COLOR="{color}">{task}</FONT></TD>
<TD WIDTH="130" NOWRAP>{commodity}</TD>
<TD WIDTH="100" NOWRAP>{vehicle}</TD>
<TD WIDTH="55" ALIGN="CENTER" NOWRAP>{cars}</TD>
<TD WIDTH="120" NOWRAP>{origin}</TD>
<TD WIDTH="120" NOWRAP>{destination}</TD>

<TD WIDTH="55" ALIGN="CENTER" NOWRAP>
    <INPUT TYPE="CHECKBOX"
           {"CHECKED" if active_flag else ""}
           onchange="saveField('{row_key}', 'active', this.checked)">
</TD>
<TD WIDTH="80" ALIGN="CENTER" NOWRAP>
    <A HREF="#" onclick="completeTask('{row_key}', '{safe_key}'); return false;">
        <FONT COLOR="{done_color}">{done_text}</FONT>
    </A>
</TD>
<TD WIDTH="45" ALIGN="CENTER" NOWRAP>
    <INPUT TYPE="TEXT" ID="day_{safe_key}" VALUE="{day_value}" SIZE="1"
       onchange="saveField('{row_key}', 'day', this.value)">
</TD>
<TD WIDTH="60" ALIGN="CENTER" NOWRAP>
    <INPUT TYPE="TEXT" ID="time_{safe_key}" VALUE="{time_value}" SIZE="4"
       onchange="saveField('{row_key}', 'time', this.value)">
</TD>
<TD WIDTH="55" ALIGN="CENTER" NOWRAP>
    <INPUT TYPE="TEXT" ID="partial_{safe_key}" VALUE="{partial_value}" SIZE="3"
           onchange="saveField('{row_key}', 'partial', this.value)">
</TD>
<TD WIDTH="85" ALIGN="CENTER" NOWRAP>
<INPUT TYPE="TEXT"
       ID="assigned_{safe_key}"
       VALUE="{assigned_value}"
       SIZE="8"
       onchange="saveField('{row_key}','assigned',this.value)">
</TD>
</TR>
""")

    footer_html = build_footer_stats(df, status)
    return f"""<HTML>
<HEAD>
<TITLE>Current Switch List</TITLE>
<STYLE>
INPUT {{
    background-color: #2f2f2f;
    color: #d8d8d8;
    border: 1px solid #555555;
    font-family: Arial, Helvetica;
    font-size: 10px;
}}
A {{ text-decoration: none; }}
</STYLE>
<SCRIPT>
function saveField(rowKey, field, value) {{
    pywebview.api.update_field(rowKey, field, value).then(function(result) {{
        if (field == "active" && typeof result === "string") {{
            document.open();
            document.write(result);
            document.close();
        }}
    }});
}}
function completeTask(rowKey, safeKey) {{
    var day = document.getElementById("day_" + safeKey).value;
    var time = document.getElementById("time_" + safeKey).value;
    var partial = document.getElementById("partial_" + safeKey).value;

    pywebview.api.complete_task(rowKey, day, time, partial).then(function(newHtml) {{
        document.open();
        document.write(newHtml);
        document.close();
    }});
}}
function changeCounter(counterName, delta) {{
    pywebview.api.change_counter(counterName, delta).then(function(newHtml) {{
        document.open();
        document.write(newHtml);
        document.close();
    }});
}}
</SCRIPT>
</HEAD>


<BODY BGCOLOR="#2f2f2f" TEXT="#d8d8d8">
<DIV STYLE="background:#292929; height:20px; -webkit-app-region: drag;"></DIV>

<FONT FACE="Arial,Helvetica" SIZE="3"><B>Current Switch List</B></FONT>
<BR><BR>

<FONT FACE="Arial,Helvetica" SIZE="0">
<TABLE BORDER="0" CELLPADDING="1" CELLSPACING="0" WIDTH="100%">
<TR BGCOLOR="#222222">
<TD WIDTH="190"><B>Industry</B></TD>
<TD WIDTH="110"><B>Task</B></TD>
<TD WIDTH="130"><B>Commodity</B></TD>
<TD WIDTH="100"><B>Vehicle</B></TD>
<TD WIDTH="55" ALIGN="CENTER"><B># Cars</B></TD>
<TD WIDTH="120"><B>Origin</B></TD>
<TD WIDTH="120"><B>Destination</B></TD>
<TD WIDTH="55" ALIGN="CENTER"><B>Active</B></TD>
<TD WIDTH="80" ALIGN="CENTER"><B>Done</B></TD>
<TD WIDTH="45" ALIGN="CENTER"><B>Day</B></TD>
<TD WIDTH="60" ALIGN="CENTER"><B>Time</B></TD>
<TD WIDTH="55" ALIGN="CENTER"><B>Partial</B></TD>
<TD WIDTH="85" ALIGN="CENTER"><B>Assigned</B></TD>
</TR>

{''.join(rows)}

</TABLE>
{footer_html}
</FONT>
</BODY>
</HTML>
"""
def hotkey_loop(window):
    visible = True

    def toggle_window():
        nonlocal visible

        if visible:
            window.minimize()
            visible = False
        else:
            window.restore()
            window.show()
            visible = True

    keyboard.add_hotkey("F9", toggle_window)
    keyboard.wait()

if __name__ == "__main__":
    api = Api()

    df = pd.read_excel(SWITCHLIST_FILE)
    task_count = len(df)

    if task_count <= 10:
        window_height = 450
    elif task_count <= 20:
        window_height = 620
    else:
        window_height = 900

    window = webview.create_window(
        "",
        html=render_html(),
        js_api=api,
        width=1535,
        height=window_height,
        resizable=False,
        frameless=True,
        on_top=True,
    )

    threading.Thread(
        target=hotkey_loop,
        args=(window,),
        daemon=True
    ).start()

    webview.start()