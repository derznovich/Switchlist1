import pandas as pd
import random
from pathlib import Path

INDUSTRY_FILE = Path("pb-industry-list.xlsx")
SWITCHLIST_FILE = Path("current_switchlist.xlsx")
COMPLETED_FILE = Path("completed_tasks.xlsx")
SHIPMENT_HISTORY_FILE = Path("shipment_history.xlsx")
STATE_FILE = Path("game_state.json")
SCORE_FILE = Path("score_summary.xlsx")
HTML_FILE = Path("current_switchlist.html")
STATUS_FILE = Path("switchlist_status.json")
import json

PORTALS = ["Juniata Portal", "Grafton Portal"]

# load state
def load_state():
    default_state = {
        "last_generated_day": None,
        "playthrough_active": False,
        "last_run_time": "",
        "last_game_day": None
    }

    if not STATE_FILE.exists():
        return default_state

    try:
        with open(STATE_FILE, "r") as f:
            state = json.load(f)
            default_state.update(state)
            return default_state
    except json.JSONDecodeError:
        return default_state

def save_state(state):
    with open(STATE_FILE, "w") as f:
        json.dump(state, f, indent=2)


def is_new_playthrough(gtime):
    return str(gtime).strip().upper() == "D1-0830"

def reset_completion_status():
    save_completion_status({})

def game_day(gtime):
    day, _ = parse_game_time(gtime)
    return day

# scoring
def elapsed_days_score(assigned, completed):
    assigned_mins = game_time_minutes(assigned)
    completed_mins = game_time_minutes(completed)

    if assigned_mins is None or completed_mins is None:
        return 0

    elapsed_days = (completed_mins - assigned_mins) / (24 * 60)

    if elapsed_days < 1:
        return 4
    elif elapsed_days < 2:
        return 3
    elif elapsed_days < 3:
        return 2
    elif elapsed_days < 4:
        return 1
    else:
        return 0
    
def calculate_score(completed_log):
    if completed_log.empty:
        return pd.DataFrame([{
            "completed_tasks": 0,
            "completed_carloads": 0,
            "total_score": 0,
            "average_score_per_carload": 0
        }])

    scored = completed_log.copy()

    scored["time_score"] = scored.apply(
        lambda r: elapsed_days_score(
            r.get("assigned_datetime"),
            r.get("completed_datetime")
        ),
        axis=1
    )

    scored["carload_score"] = scored["number_of_cars"] * scored["time_score"]

    total_carloads = scored["number_of_cars"].sum()
    total_score = scored["carload_score"].sum()

    summary = pd.DataFrame([{
        "completed_tasks": len(scored),
        "completed_carloads": total_carloads,
        "total_score": total_score,
        "average_score_per_carload": round(total_score / total_carloads, 2) if total_carloads else 0
    }])

    with pd.ExcelWriter(SCORE_FILE, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="summary", index=False)
        scored.to_excel(writer, sheet_name="scored_tasks", index=False)

    return summary    

def calculate_incident_penalty():
    status = load_completion_status()

    meta = status.get("_meta", {})

    major_derailments = int(meta.get("major_derailments", 0))
    other_incidents = int(meta.get("other_incidents", 0))

    penalty = (
        major_derailments * -250
        + other_incidents * -75
    )

    return major_derailments, other_incidents, penalty

# switch list creation
def norm(value):
    return str(value).strip().lower()


def is_portal(location):
    return norm(location) in {norm(p) for p in PORTALS}


def bounded_normal_int(mu=4, sigma=1.5, low=2, high=7):
    return max(low, min(high, round(random.normalvariate(mu, sigma))))


def split_cars(total_cars, probability=0.25):
    if total_cars <= 1 or random.random() > probability:
        return total_cars, 0
    first = random.randint(1, total_cars - 1)
    return first, total_cars - first


def load_excel(path):
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame()


def task_block_key(industry, commodity, produce_consume):
    return (
        norm(industry),
        norm(commodity),
        str(produce_consume).strip().upper()
    )


def get_active_blocked_keys(active_tasks):
    blocked = set()

    if active_tasks.empty:
        return blocked

    for _, row in active_tasks.iterrows():
        if norm(row.get("status", "")) != "active":
            continue

        industry = row.get("industry_name")
        commodity = row.get("commodity")
        produce_consume = row.get("produce_consume")

        if pd.notna(industry) and pd.notna(commodity) and pd.notna(produce_consume):
            blocked.add(task_block_key(industry, commodity, produce_consume))

    return blocked

def parse_game_time(gtime):
    """
    Converts D2-1340 -> (2, 1340)
    """
    if pd.isna(gtime):
        return None, None

    text = str(gtime).strip().upper()

    if "-" not in text or not text.startswith("D"):
        return None, None

    try:
        day_part, time_part = text.split("-")
        day = int(day_part[1:])
        hhmm = int(time_part)
        return day, hhmm
    except:
        return None, None

def game_time_minutes(gtime):
    """
    Convert D2-1340 into minutes since Day 1 00:00
    """
    day, hhmm = parse_game_time(gtime)

    if day is None:
        return None

    hours = hhmm // 100
    mins = hhmm % 100

    return ((day - 1) * 24 * 60) + hours * 60 + mins

def archive_completed_tasks(existing, completed_log):
    if existing.empty:
        return existing, completed_log

    completed = existing[existing["status"].astype(str).str.lower().eq("completed")].copy()
    active = existing[~existing["status"].astype(str).str.lower().eq("completed")].copy()

    if not completed.empty:
        completed_log = pd.concat([completed_log, completed], ignore_index=True)

    return active, completed_log


def next_number(df_list, column, prefix=None):
    numbers = []

    for df in df_list:
        if df.empty or column not in df.columns:
            continue

        for value in df[column].dropna():
            text = str(value)
            if prefix:
                text = text.replace(prefix, "")
            try:
                numbers.append(int(text))
            except ValueError:
                pass

    nxt = max(numbers, default=0) + 1
    return f"{prefix}{nxt:04d}" if prefix else nxt

def find_flexible_partners(master, industry, commodity, produce_consume):
    direction = str(produce_consume).strip().upper()

    if direction.startswith("P"):
        opposite = "C"
    else:
        opposite = "P"

    matches = master[
        (master["commodity"].astype(str).str.strip().str.lower() == str(commodity).strip().lower()) &
        (master["produce_consume"].astype(str).str.strip().str.upper().str.startswith(opposite)) &
        (master["industry_name"].astype(str).str.strip().str.lower() != str(industry).strip().lower())
    ]

    partners = matches["industry_name"].dropna().astype(str).str.strip().unique().tolist()
    random.shuffle(partners)

    if len(partners) == 0:
        return None, None
    if len(partners) == 1:
        return partners[0], None
    return partners[0], partners[1]

def choose_portal():
    return random.choice(PORTALS)


def choose_partner(row):
    partners = [
        row.get("fixed_partner_1"),
        row.get("fixed_partner_2"),
    ]
    partners = [p for p in partners if pd.notna(p) and str(p).strip()]

    if not partners:
        return choose_portal(), None

    if len(partners) == 1:
        return partners[0], None

    return partners[0], partners[1]


def create_task(task_number, shipment_id, task_type, assigned_datetime, industry_name,
                commodity, produce_consume, load_status, car_type, cars, origin, destination):
    if str(task_type).strip().lower() == "spot empties":
        origin = "-"
    return {
        "shipment_id": shipment_id,
        "task_number": task_number,
        "task_type": task_type,
        "assigned_datetime": assigned_datetime,
        "industry_name": industry_name,
        "commodity": commodity,
        "produce_consume": produce_consume,
        "load_status": load_status,
        "car_type": car_type,
        "number_of_cars": cars,
        "origin": origin,
        "destination": destination,
        "status": "active",
        "completed_datetime": None,
    }


def expand_row_to_tasks(row, shipment_id, first_task_number, assigned_datetime, master):
    industry = row["industry_name"]
    commodity = row["commodity"]
    produce_consume = str(row["produce_consume"]).strip().upper()
    car_type = row["car_type"]
    cars = random.randint(int(row["min_carloads"]), int(row["max_carloads"]))

    source_mode = norm(row.get("source_mode", "off_map_ok"))
    split_probability = float(row.get("split_probability", 0.30) or 0.30)
    allow_split = norm(row.get("allow_split", "yes")) in ["yes", "true", "1"]

    first_cars, second_cars = split_cars(cars, split_probability if allow_split else 0)

    partner_1, partner_2 = choose_partner(row)
    if not partner_1 or is_portal(partner_1):
        flex_1, flex_2 = find_flexible_partners(
            master, industry, commodity, produce_consume
        )
        if flex_1:
            partner_1, partner_2 = flex_1, flex_2

    use_on_map_partner = False

    if source_mode in ["on_map_only", "fixed_on_map_only"]:
        use_on_map_partner = True

    elif source_mode == "off_map_ok":
        # Flexible traffic: about half the time use an on-map partner if one exists.
        if partner_1 and random.random() < 0.50:
            use_on_map_partner = True

    if use_on_map_partner:
        external_1 = partner_1

        # Only allow a split if a real second on-map partner exists.
        if second_cars and partner_2:
            external_2 = partner_2
        else:
            external_2 = None
            first_cars = cars
            second_cars = 0

    else:
        # Off-map traffic should go to/from one portal only.
        external_1 = choose_portal()
        external_2 = None
        first_cars = cars
        second_cars = 0

    tasks = []
    task_number = first_task_number

    if produce_consume.startswith("C"):
        load_status = "loaded"

        tasks.append(create_task(
            task_number, shipment_id, "Deliver", assigned_datetime,
            industry, commodity, "C", load_status, car_type,
            first_cars, external_1, industry
        ))
        task_number += 1

        if second_cars:
            tasks.append(create_task(
                task_number, shipment_id, "Deliver", assigned_datetime,
                industry, commodity, "C", load_status, car_type,
                second_cars, external_2, industry
            ))
            task_number += 1

        if not is_portal(external_1):
            tasks.append(create_task(
                task_number, shipment_id, "Pickup Load", assigned_datetime,
                external_1, commodity, "P", load_status, car_type,
                first_cars, external_1, industry
            ))
            task_number += 1

            # Supply empty cars to the on-map producer
            tasks.append(create_task(
                task_number, shipment_id, "Spot Empties", assigned_datetime,
                external_1, commodity, "P", "empty", car_type,
                first_cars, choose_portal(), external_1
            ))
            task_number += 1

        if second_cars and external_2 and not is_portal(external_2):
            tasks.append(create_task(
                task_number, shipment_id, "Pickup Load", assigned_datetime,
                external_2, commodity, "P", load_status, car_type,
                second_cars, external_2, industry
            ))
            task_number += 1

            # Supply empty cars to the second on-map producer
            tasks.append(create_task(
                task_number, shipment_id, "Spot Empties", assigned_datetime,
                external_2, commodity, "P", "empty", car_type,
                second_cars, choose_portal(), external_2
            ))
            task_number += 1

    else:
        # Producer logic:
        # 1. Bring empty cars from a portal to the producing industry.
        # 2. Move loaded cars from the producing industry to either:
        #    - an on-map destination, or
        #    - an off-map portal.

        empty_origin = choose_portal()

        tasks.append(create_task(
            task_number, shipment_id, "Spot Empties", assigned_datetime,
            industry, commodity, "P", "empty", car_type,
            cars, empty_origin, industry
        ))
        task_number += 1

        # Loaded outbound movement from producer
        tasks.append(create_task(
            task_number, shipment_id, "Pickup Load", assigned_datetime,
            industry, commodity, "P", "loaded", car_type,
            first_cars, industry, external_1
        ))
        task_number += 1


        tasks.append(create_task(
            task_number, shipment_id, "Deliver", assigned_datetime,
            external_1, commodity, "C", "loaded", car_type,
            first_cars, industry, external_1
        ))
        task_number += 1

        if second_cars and external_2:
            tasks.append(create_task(
                task_number, shipment_id, "Pickup Load", assigned_datetime,
                industry, commodity, "P", "loaded", car_type,
                second_cars, industry, external_2
            ))
            task_number += 1

            tasks.append(create_task(
                task_number, shipment_id, "Deliver", assigned_datetime,
                external_2, commodity, "C", "loaded", car_type,
                second_cars, industry, external_2
            ))
            task_number += 1

    return tasks


def row_block_keys(row):
    industry = row["industry_name"]
    commodity = row["commodity"]
    produce_consume = row["produce_consume"]

    keys = {task_block_key(industry, commodity, produce_consume)}

    source_mode = norm(row.get("source_mode", "off_map_ok"))

    if source_mode in ["on_map_only", "fixed_on_map_only"]:
        partner_1 = row.get("fixed_partner_1")
        partner_2 = row.get("fixed_partner_2")

        if pd.notna(partner_1) and str(partner_1).strip():
            direction = str(produce_consume).strip().upper()
            opposite = "P" if direction.startswith("C") else "C"
            keys.add(task_block_key(partner_1, commodity, opposite))

        if pd.notna(partner_2) and str(partner_2).strip():
            direction = str(produce_consume).strip().upper()
            opposite = "P" if direction.startswith("C") else "C"
            keys.add(task_block_key(partner_2, commodity, opposite))

    return keys

def generate_tasks(
    master,
    active,
    completed,
    assigned_datetime,
    first_run,
    current_day
):
    blocked = get_active_blocked_keys(active)

    # Day 6+: no new work
    if current_day >= 6:
        return pd.DataFrame()

    if first_run:
        mandatory = master[
            master["mandatory_start"]
            .astype(str)
            .str.lower()
            .isin(["yes", "true", "1"])
        ]

        target_total = random.randint(10, 15)
        selected = mandatory.to_dict("records")

        remaining = master[~master.index.isin(mandatory.index)]
        needed = max(0, target_total - len(selected))

        if needed:
            selected.extend(
                remaining.sample(
                    min(needed, len(remaining))
                ).to_dict("records")
            )

        target_shipments = target_total

    elif current_day == 5:
        # Final intake day:
        # only small, one-and-done deliveries from portals.

        late_candidates = master[
            master["produce_consume"]
            .astype(str)
            .str.strip()
            .str.upper()
            .str.startswith("C")
        ].copy()

        # Exclude consumers that must source from an on-map producer.
        late_candidates = late_candidates[
            ~late_candidates["source_mode"]
            .astype(str)
            .str.strip()
            .str.lower()
            .isin(["on_map_only", "fixed_on_map_only"])
        ]

        # Small amount of new work on Day 5.
        target_shipments = random.randint(3, 6)

        selected = late_candidates.sample(
            min(target_shipments * 3, len(late_candidates))
        ).to_dict("records")

    else:
        # Days 2-4: normal daily generation
        target_shipments = bounded_normal_int()

        selected = master.sample(
            min(target_shipments * 3, len(master))
        ).to_dict("records")

    task_number = next_number(
        [active, completed],
        "task_number"
    )

    shipment_id = next_number(
        [active, completed],
        "shipment_id",
        "S"
    )

    new_tasks = []

    for row in selected:
        keys = row_block_keys(row)

        if keys & blocked:
            continue

        if current_day == 5:
            # One-and-done portal delivery.
            industry = row["industry_name"]
            commodity = row["commodity"]
            car_type = row["car_type"]

            # Keep final-day jobs deliberately small.
            normal_min = int(row["min_carloads"])
            normal_max = int(row["max_carloads"])

            day5_max = min(normal_max, 5)
            day5_min = min(normal_min, day5_max)

            cars = random.randint(day5_min, day5_max)

            shipment_tasks = [
                create_task(
                    task_number,
                    shipment_id,
                    "Deliver",
                    assigned_datetime,
                    industry,
                    commodity,
                    "C",
                    "loaded",
                    car_type,
                    cars,
                    choose_portal(),
                    industry
                )
            ]

        else:
            shipment_tasks = expand_row_to_tasks(
                row,
                shipment_id,
                task_number,
                assigned_datetime,
                master
            )

        # New chains with Spot Empties start with follow-on work not ready.
        has_spot_empties = any(
            task.get("task_type") == "Spot Empties"
            for task in shipment_tasks
        )

        for task in shipment_tasks:
            if has_spot_empties:
                task["ready"] = (
                    task.get("task_type") == "Spot Empties"
                )
            else:
                task["ready"] = True

        new_tasks.extend(shipment_tasks)

        for task in shipment_tasks:
            blocked.add(
                task_block_key(
                    task["industry_name"],
                    task["commodity"],
                    task["produce_consume"]
                )
            )

        task_number += len(shipment_tasks)

        shipment_id = (
            f"S{int(shipment_id.replace('S', '')) + 1:04d}"
        )

        # Stop once this day's target number of shipments is reached.
        if (
            not first_run
            and len({t["shipment_id"] for t in new_tasks})
            >= target_shipments
        ):
            break

    return pd.DataFrame(new_tasks)

# importing json
def load_completion_status():
    if not STATUS_FILE.exists():
        return {}

    try:
        with open(STATUS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

# sanitizing the day and time

def clean_day(value, fallback_day=1):
    try:
        day = int(str(value).strip().replace("D", "").replace("d", ""))
        return max(1, day)
    except:
        return fallback_day
    
def clean_time(value):
    text = str(value).strip().lower().replace(":", "")

    try:
        n = int(text)
    except:
        return "1200"

    # 1-5 means afternoon railroad hours: 1300-1500?
    if 1 <= n <= 5:
        n = n + 12
        return f"{n:02d}00"

    # 8-9 means morning
    if 8 <= n <= 9:
        return f"0{n}00"

    # 10-12 means 1000-1200
    if 10 <= n <= 12:
        return f"{n}00"

    # 3 digits: 830 -> 0830, 230 -> 1430
    if 100 <= n <= 999:
        hour = n // 100
        mins = n % 100

        if 1 <= hour <= 5:
            hour += 12

        return clamp_time(hour, mins)

    # 4 digits
    if 0 <= n <= 2359:
        hour = n // 100
        mins = n % 100
        return clamp_time(hour, mins)

    return "1200"

def clamp_time(hour, mins):
    if mins > 59:
        mins = 0

    total = hour * 60 + mins

    start = 8 * 60 + 30
    end = 16 * 60

    total = max(start, min(end, total))

    hour = total // 60
    mins = total % 60

    return f"{hour:02d}{mins:02d}"

def clean_completed_datetime(row_status, fallback_day=1):
    day = clean_day(row_status.get("day", ""), fallback_day)
    time = clean_time(row_status.get("time", ""))
    return f"D{day}-{time}"

def latest_completion_time_from_status():
    latest = None
    latest_minutes = None

    # 1. Check pending viewer completions in switchlist_status.json
    status = load_completion_status()

    for key, item in status.items():
        if key == "_meta":
            continue

        if not item.get("completed", False):
            continue

        day_raw = item.get("day", "")
        time_raw = item.get("time", "")

        try:
            day = int(str(day_raw).strip())
        except:
            continue

        if day < 1 or day > 10:
            continue

        time = clean_time(time_raw)
        gtime = f"D{day}-{time}"
        mins = game_time_minutes(gtime)

        if mins is not None and (latest_minutes is None or mins > latest_minutes):
            latest_minutes = mins
            latest = gtime

    # 2. Also check already archived completed tasks
    completed_log = load_excel(COMPLETED_FILE)

    if not completed_log.empty and "completed_datetime" in completed_log.columns:
        for value in completed_log["completed_datetime"].dropna():
            gtime = str(value).strip().upper()
            mins = game_time_minutes(gtime)

            if mins is not None and (latest_minutes is None or mins > latest_minutes):
                latest_minutes = mins
                latest = gtime

    return latest

def valid_game_day(value, fallback=1):
    try:
        day = int(str(value).strip())
    except:
        return fallback

    if 1 <= day <= 10:
        return day

    return fallback

def make_game_time_for_current_day(state):
    current_day = valid_game_day(
        state.get("last_game_day"),
        fallback=valid_game_day(state.get("last_generated_day"), fallback=1)
    )

    latest_completed = latest_completion_time_from_status()
    last_run = state.get("last_run_time") or "unknown"

    reference_raw = latest_completed or last_run
    reference = reference_raw

    if reference_raw and str(reference_raw).startswith("D") and "-" in str(reference_raw):
        d, t = str(reference_raw).split("-", 1)
        reference = f"Day {d.replace('D','')}, at {t}"

    entered = input(
        f"Enter time for Day {current_day}. Last recorded game-time was {reference}: "
    ).strip()

    cleaned_time = clean_time(entered)
    gtime = f"D{current_day}-{cleaned_time}"

    if not valid_game_time_entry(gtime):
        print("Incorrect time entry.")
        return None

    return gtime

def reset_score_summary():
    if SCORE_FILE.exists():
        SCORE_FILE.unlink()

# generate html file
def task_color(task):
    task = str(task).lower()

    if "spot empties" in task:
        return "#ffe48a"   # light yellow

    if "pickup load" in task:
        return "#8fc7ff"   # light blue

    if "deliver" in task:
        return "#9be89b"   # light green

    return "#d8d8d8"


def export_trainz_files(active):
    if active.empty:
        trainz = pd.DataFrame(columns=[
            "shipment_id", "task_number", "industry_name", "task_type",
            "commodity", "car_type", "number_of_cars", "origin", "destination"
        ])
    else:
        trainz = active.copy()

        task_order = {
            "Spot Empties": 1,
            "Pickup Load": 2,
            "Deliver": 3,
            "Deliver": 3,
            "Deliver": 3,
        }

        trainz["task_sort"] = trainz["task_type"].map(task_order).fillna(9)

        trainz = trainz.sort_values(
            by=["industry_name", "commodity", "task_sort"],
            kind="stable"
        )

        trainz = trainz.drop(columns=["task_sort"], errors="ignore")

        # Blank origin/destination if they match the industry
        trainz.loc[
            trainz["origin"].astype(str).str.strip()
            == trainz["industry_name"].astype(str).str.strip(),
            "origin"
        ] = "-"

        trainz.loc[
            trainz["destination"].astype(str).str.strip()
            == trainz["industry_name"].astype(str).str.strip(),
            "destination"
        ] = "-"

    rows = []

    for i, (_, row) in enumerate(trainz.iterrows()):
        bg = "#383838" if i % 2 == 0 else "#2f2f2f"
        shipment_id = row.get("shipment_id", "")
        task_number = row.get("task_number", "")

        industry = row.get("industry_name", "")
        task = row.get("task_type", "")
        commodity = row.get("commodity", "")
        if str(task).strip().lower() == "spot empties":
            commodity = "-" 
        vehicle = row.get("car_type", "")
        cars = row.get("number_of_cars", "")
        origin = row.get("origin", "")
        destination = row.get("destination", "")

        color = task_color(task)

        rows.append(f"""
<TR BGCOLOR="{bg}" DATA-SHIPMENT="{shipment_id}" DATA-TASK="{task_number}">
<TD>{industry}</TD>
<TD><FONT COLOR="{color}">{task}</FONT></TD>
<TD>{commodity}</TD>
<TD>{vehicle}</TD>
<TD ALIGN="CENTER">{cars}</TD>
<TD>{origin}</TD>
<TD>{destination}</TD>
<TD ALIGN="CENTER"><FONT COLOR="#999999">☐</FONT></TD>
</TR>
""")

    html_page = f"""<HTML>
<BODY BGCOLOR="#2f2f2f" TEXT="#d8d8d8">
<FONT FACE="Arial,Helvetica" SIZE="1">

<FONT FACE="Arial,Helvetica" SIZE="3"><B>Current Switch List</B></FONT>
<BR><BR>
<FONT FACE="Arial,Helvetica" SIZE="1">

<TABLE CELLPADDING="3" CELLSPACING="1" WIDTH="100%">
<TR BGCOLOR="#222222">
<TD WIDTH="190"><B>Industry</B></TD>
<TD WIDTH="110"><B>Task</B></TD>
<TD WIDTH="130"><B>Commodity</B></TD>
<TD WIDTH="100"><B>Vehicle</B></TD>
<TD WIDTH="55" ALIGN="CENTER"><B># Cars</B></TD>
<TD WIDTH="120"><B>Origin</B></TD>
<TD WIDTH="120"><B>Destination</B></TD>
<TD WIDTH="80" ALIGN="CENTER"><B>Done</B></TD>
</TR>

{''.join(rows)}

</TABLE>
</FONT>
</BODY>
</HTML>
"""

    HTML_FILE.write_text(html_page, encoding="utf-8")

    print(f"HTML switchlist exported to: {HTML_FILE.resolve()}")


def get_startup_game_time(state):
    print("Starting switch-list.")
    print("Press 1 to begin a new session.")
    print("Press 2 to begin a new day.")
    print("Press 3 to continue current day.")

    choice = input("Choice: ").strip()

    if choice == "1":
        confirm = input("Are you sure you want to start a new session? This resets active/completed records. (y/n): ").strip().lower()

        if confirm == "y":
            return "D1-0830"
        else:
            print("Cancelled.")
            return None

    if choice == "2":
        last_day = state.get("last_game_day") or state.get("last_generated_day") or 1
        next_day = int(last_day) + 1
        return f"D{next_day}-0830"

    if choice == "3":
        return make_game_time_for_current_day(state)

    print("Incorrect choice.")
    return None

def valid_game_time_entry(gtime):
    day, hhmm = parse_game_time(gtime)

    if day is None or hhmm is None:
        return False

    hour = hhmm // 100
    minute = hhmm % 100

    if minute > 59:
        return False

    minutes = hour * 60 + minute

    start = 8 * 60 + 30
    end = 16 * 60

    return start <= minutes <= end

def save_completion_status(status):
    STATUS_FILE.write_text(json.dumps(status, indent=2), encoding="utf-8")


def unlock_ready_shipments(active, completed_log):
    """
    Unlock follow-on tasks when:
    1. At least one Spot Empties task for the shipment was completed.
    2. No Spot Empties tasks for that shipment remain active.

    This function should only be called when beginning a new game day.
    """
    if "ready" in active.columns:
        active["ready"] = (
            active["ready"]
            .fillna(True)
            .astype(bool)
        )
    else:
        active["ready"] = True


    if "ready" not in active.columns:
        active["ready"] = True

    completed_spots = completed_log[
        completed_log["task_type"].astype(str).str.strip().eq("Spot Empties")
    ]

    completed_shipments = set(
        completed_spots["shipment_id"].dropna().astype(str)
    )

    active_spots = active[
        active["task_type"].astype(str).str.strip().eq("Spot Empties")
    ]

    shipments_still_waiting = set(
        active_spots["shipment_id"].dropna().astype(str)
    )

    shipments_to_unlock = (
        completed_shipments - shipments_still_waiting
    )

    if shipments_to_unlock:
        mask = (
            active["shipment_id"].astype(str).isin(shipments_to_unlock)
            & ~active["task_type"].astype(str).str.strip().eq("Spot Empties")
        )

        active.loc[mask, "ready"] = True

    return active


def main():
    state = load_state()

    assigned_datetime = get_startup_game_time(state)

    if assigned_datetime is None:
        return

    current_day = game_day(assigned_datetime)

    # Full reset for new playthrough
    if is_new_playthrough(assigned_datetime):
        active = pd.DataFrame()
        completed_log = pd.DataFrame()
        reset_completion_status()
        reset_score_summary()
        state = {
            "last_generated_day": None,
            "playthrough_active": True,
            "last_run_time": "",
            "last_game_day": None
        }
    else:
        existing = load_excel(SWITCHLIST_FILE)
        completed_log = load_excel(COMPLETED_FILE)

        # First archive anything manually marked completed in Excel
        active, completed_log = archive_completed_tasks(existing, completed_log)

        # Then import completions from switchlist_status.json
        completion_status = load_completion_status()
        completed_rows = []
        completed_indices = []
        completed_status_keys = []

        for idx, row in active.iterrows():
            row_key = f"{row['shipment_id']}|{row['task_number']}"
            row_status = completion_status.get(row_key, {})

            if row_status.get("completed", False):
                completed_status_keys.append(row_key)
                row = row.copy()
                row["status"] = "completed"
                row["completed_datetime"] = clean_completed_datetime(
                    row_status,
                    fallback_day=current_day or 1
                )

                completed_rows.append(row)
                completed_indices.append(idx)

        if completed_rows:
            completed_df = pd.DataFrame(completed_rows)
            completed_log = pd.concat(
                [completed_log, completed_df],
                ignore_index=True
            )

            active = active.drop(index=completed_indices).reset_index(drop=True)
            for key in completed_status_keys:
                completion_status.pop(key, None)

            save_completion_status(completion_status)

    master = pd.read_excel(INDUSTRY_FILE)

    last_generated_day = state.get("last_generated_day")
    first_run = is_new_playthrough(assigned_datetime)

    should_generate = False

    if first_run:
        should_generate = True
    elif current_day is not None and current_day != last_generated_day:
        should_generate = True

    if should_generate and not first_run:
        active = unlock_ready_shipments(active, completed_log)

    if should_generate:
        new_tasks = generate_tasks(
            master=master,
            active=active,
            completed=completed_log,
            assigned_datetime=assigned_datetime,
            first_run=first_run,
            current_day=current_day
        )

        active = pd.concat([active, new_tasks], ignore_index=True)
        state["last_generated_day"] = current_day
    else:
        new_tasks = pd.DataFrame()

    # Save active switchlist
    with pd.ExcelWriter(SWITCHLIST_FILE, engine="openpyxl") as writer:
        active.to_excel(writer, index=False)
        ws = writer.sheets["Sheet1"]

        for col in ws.iter_cols():
            header = col[0].value
            if header in ["assigned_datetime", "completed_datetime"]:
                for cell in col[1:]:
                    cell.number_format = "@"

    # export_trainz_files(active)

    completed_log.to_excel(COMPLETED_FILE, index=False)

    # Calculate base score from completed tasks
    if not completed_log.empty:
        score = calculate_score(completed_log)

        if active.empty:
            # Final playthrough score includes incidents
            major_derailments, other_incidents, incident_penalty = (
                calculate_incident_penalty()
            )

            base_score = int(score.iloc[0]["total_score"])
            final_score = base_score + incident_penalty

            final_summary = pd.DataFrame([{
                "base_score": base_score,
                "major_derailments": major_derailments,
                "other_incidents": other_incidents,
                "incident_penalty": incident_penalty,
                "total_score": final_score
            }])

            final_summary.to_excel(SCORE_FILE, index=False)

            print("All tasks completed. Final score generated.")
            print(f"Base score: {base_score}")
            print(f"Major derailments: {major_derailments} x -250")
            print(f"Other incidents: {other_incidents} x -75")
            print(f"Incident penalty: {incident_penalty}")
            print(f"Final score: {final_score}")

        else:
            print("Score updated for completed tasks so far.")
            print(score.to_string(index=False))

    state["last_run_time"] = assigned_datetime
    state["last_game_day"] = current_day 
       
    save_state(state)


if __name__ == "__main__":
    main()