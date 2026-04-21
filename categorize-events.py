import warnings
warnings.filterwarnings('ignore', 'urllib3 v2 only supports OpenSSL')
import re
import os
import json
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import date
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import pytz

from collections import Counter
from collections import defaultdict
from tabulate import tabulate
import os
import yaml
from datetime import datetime, date

# If modifying the calendar, the required scopes are 'https://www.googleapis.com/auth/calendar'
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/spreadsheets']
with open('credentials.json', 'r') as _f:
    SHEET_ID = json.load(_f)['sheet_id']
UNCATEGORIZED_EVENTS = []

# -------- config loaders --------
def load_categories(path="categories.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f) or {}
    # ensure lists & lowercase keywords
    normalized = {}
    for k, v in data.items():
        if isinstance(v, list):
            normalized[k] = [str(x).lower() for x in v]
    return normalized

def _parse_date_any(s: str) -> date:
    for fmt in ("%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unrecognized date format: {s}")

def load_blacklist_dates(path="blacklist_dates"):
    if not os.path.exists(path):
        return set()
    out = set()
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            t = line.strip()
            if not t or t.startswith("#"):
                continue
            out.add(_parse_date_any(t))
    return out


# -------- categories and blacklist_dates from different files --------
categories = load_categories(os.getenv("CATEGORIES_FILE", "categories.yaml"))
blacklist_dates = load_blacklist_dates(os.getenv("BLACKLIST_FILE", "blacklist_dates"))

def authenticate_google_account():
    creds = None

    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def build_service():
    creds = authenticate_google_account()
    service = build('calendar', 'v3', credentials=creds)
    return service


def get_user_selected_calendars():
    service = build_service()

    # Get the list of all calendars
    calendar_list = service.calendarList().list().execute()

    print("\nAvailable calendars:")
    calendars = calendar_list['items']

    # Display calendar options
    for i, calendar in enumerate(calendars, 1):
        print(f"{i}. {calendar['summary']}")

    # Prompt the user once
    selected = input("\nEnter the number of the calendar you want to query (press Enter to select primary): ").strip()

    # Default to 'primary' if no input
    if selected == "":
        return ['primary']

    try:
        selected_index = int(selected) - 1
        if 0 <= selected_index < len(calendars):
            return [calendars[selected_index]['id']]
        else:
            print("Invalid selection. Defaulting to primary.")
            return ['primary']
    except ValueError:
        print("Invalid input. Defaulting to primary.")
        return ['primary']


# Function to get the date range with default to last 7 days if no input is provided
def _parse_day_month_or_full(s: str, default_year: int) -> date:
    s = s.strip()
    if not s:
        return None
    if s.count("-") == 1:
        day_s, month_s = s.split("-")
        return date(default_year, int(month_s), int(day_s))
    # Allow full date like DD-MM-YYYY or YYYY-MM-DD
    return _parse_date_any(s)

def _parse_bool_flag(s: str):
    s = s.strip().lower()
    if not s:
        return None
    if s in {"y", "yes", "true", "1", "on"}:
        return True
    if s in {"n", "no", "false", "0", "off"}:
        return False
    return None

def get_date_range():
    today = datetime.today().date()
    current_year = today.year

    start_in = input(
        f"Enter start date (DD-MM or DD-MM-YYYY) or last X days (incl. today) [default {today - timedelta(days=7):%d-%m}]: "
    ).strip()

    if start_in.isdigit():
        days = max(1, int(start_in))
        # inclusive X-day window ending today
        start_date = today - timedelta(days=days - 1)
        end_date = today
        colorize_flag = None
    else:
        end_in = input(
            f"Enter end date (DD-MM or DD-MM-YYYY) and optional colorize flag [y/N] "
            f"(e.g. 05-02-2026 y) [default {today:%d-%m}]: "
        ).strip()
        parts = end_in.split()
        end_date_part = parts[0] if parts else ""
        colorize_flag = _parse_bool_flag(parts[1]) if len(parts) > 1 else None
        start_date = _parse_day_month_or_full(start_in, current_year) if start_in else today - timedelta(days=7)
        end_year_default = start_date.year if start_in else current_year
        end_date = _parse_day_month_or_full(end_date_part, end_year_default) if end_date_part else today

    print("Range:", start_date, "to", end_date, "(inclusive)")

    # timeMax is exclusive → +1 day to include the last day fully
    return localizeTime(start_date, end_date + timedelta(days=1)), colorize_flag

def localizeTime(start_date, end_date):
    tz = pytz.timezone('Europe/Berlin')
    return (
        tz.localize(datetime.combine(start_date, datetime.min.time())),
        tz.localize(datetime.combine(end_date, datetime.min.time()))
    )


def get_events_in_date_range(start_date, end_date):
    service = build_service()
    selected_calendars = get_user_selected_calendars()

    time_min = start_date.isoformat()
    time_max = end_date.isoformat()  # exclusive

    all_events = []
    for calendar_id in selected_calendars:
        page_token = None
        while True:
            resp = service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,          # exclusive end
                singleEvents=True,         # expand recurring
                orderBy='startTime',
                maxResults=250,            # explicit; still paginate
                pageToken=page_token
            ).execute()
            all_events.extend(resp.get('items', []))
            page_token = resp.get('nextPageToken')
            if not page_token:
                break
    return all_events


def remove_all_day(events):
    return [e for e in events if "dateTime" in e.get("start", {})]


def remove_long_events(events):
    result = []
    for e in events:
        s_iso = e.get("start", {}).get("dateTime")
        e_iso = e.get("end", {}).get("dateTime", s_iso)
        if s_iso and e_iso:
            s = datetime.fromisoformat(s_iso.replace('Z', '+00:00'))
            t = datetime.fromisoformat(e_iso.replace('Z', '+00:00'))
            if (t - s).total_seconds() < 86400:
                result.append(e)
    return result


# Get the date range from the user
(start_date, end_date), colorize_flag = get_date_range()

# Fetch the events from Google Calendar in the given date range
events_from_calendar = get_events_in_date_range(start_date, end_date)

# Now you can proceed with filtering, categorization, and plotting as before
events_in_timeframe = remove_all_day(events_from_calendar)
events_in_timeframe = remove_long_events(events_in_timeframe)

# Now you can proceed with the categorization and plotting as before
category_times = {category: 0 for category in categories}
category_times["other"] = 0  # Initialize "other" category


def parse_blacklist_dates(s: str):
    # input like: "2025-01-01, 2025-02-14"
    out = set()
    for part in (p.strip() for p in s.split(",") if p.strip()):
        out.add(datetime.strptime(part, "%Y-%m-%d").date())
    return out


def event_start_dt_local(event, tz):
    if "dateTime" not in event.get("start", {}):
        return None
    return datetime.fromisoformat(event["start"]["dateTime"].replace('Z', '+00:00')).astimezone(tz)


def exclude_blacklisted_events(events, blacklist_dates, tz_name='Europe/Berlin'):
    tz = pytz.timezone(tz_name)
    kept = []
    for e in events:
        dt = event_start_dt_local(e, tz)
        if dt is None:
            continue  # still skipping all-day events; remove to include them
        if dt.date() in blacklist_dates:
            continue
        kept.append(e)
    return kept


events_in_timeframe = exclude_blacklisted_events(events_in_timeframe, blacklist_dates)


# Function to categorize events
def categorize_event(event):
    summary = (event.get("summary") or "").lower()
    for category, keywords in categories.items():
        for keyword in keywords:
            if re.search(r'\b' + re.escape(keyword.lower()) + r'\b', summary):
                return category
    # collect instead of printing
    UNCATEGORIZED_EVENTS.append(event)
    return "other"


ALL_CATEGORIES = list(categories.keys()) + ["other"]

# --- Google Calendar color IDs (strings "1".."11")
# 1 Lavender, 2 Sage, 3 Grape, 4 Flamingo, 5 Banana, 6 Tangerine,
# 7 Peacock, 8 Graphite, 9 Blueberry, 10 Basil, 11 Tomato
DEFAULT_COLOR_ID = "10"  # Basil

CATEGORY_COLOR_ID = {
    "work": "7",            # Peacock
    "uni": "5",             # Banana
    "practical": "2",       # Sage
    "food": "6",            # Tangerine
    "time_together": "9",   # Blueberry
    "sleep": "1",           # Lavender
    "chill": "11",          # Tomato
    "fun": "4",             # Flamingo
    "sport": "3",           # Grape
    "other": "8"            # Graphite
}

def color_id_for(category: str) -> str:
    return CATEGORY_COLOR_ID.get(category, DEFAULT_COLOR_ID)


def ensure_event_color(service, calendar_id: str, event: dict, category: str) -> bool:
    """
    Set Google Calendar event color based on category.
    Returns True if an update was made, False if already correct or not applicable.
    """
    # Skip all-day or cancelled
    if event.get("status") == "cancelled" or "dateTime" not in event.get("start", {}):
        return False

    desired = color_id_for(category)  # uses your CATEGORY_COLOR_ID map
    current = event.get("colorId")

    if current == desired:
        return False  # no change needed

    service.events().patch(
        calendarId=calendar_id,
        eventId=event["id"],
        body={"colorId": desired},
    ).execute()
    return True

def event_local_start_end(e, tz):
    s = datetime.fromisoformat(e["start"]["dateTime"].replace('Z', '+00:00')).astimezone(tz)
    e_iso = e["end"].get("dateTime", e["start"]["dateTime"])
    t = datetime.fromisoformat(e_iso.replace('Z', '+00:00')).astimezone(tz)
    return s, t


def calculate_duration(event):
    tz = pytz.timezone('Europe/Berlin')
    s, e = event_local_start_end(event, tz)
    return (e - s).total_seconds() / 60


def _env_colorize_default():
    env = os.getenv("COLORIZE_EVENTS")
    return False if env is None else (_parse_bool_flag(env) or False)

COLORIZE_EVENTS = colorize_flag if colorize_flag is not None else _env_colorize_default()
service = build_service() if COLORIZE_EVENTS else None

for event in events_in_timeframe:
    category = categorize_event(event)
    duration = calculate_duration(event)
    category_times[category] += duration
    # --- update event color in Google Calendar (optional) ---
    if COLORIZE_EVENTS:
        calendar_id = event.get("organizer", {}).get("email", "primary")
        try:
            changed = ensure_event_color(service, calendar_id, event, category)
            if changed:
                print(f"Updated color for: {event.get('summary','<no title>')} → {category}")
        except Exception as e:
            print(f"Failed to update color for {event.get('summary','<no title>')}: {e}")


def reclassify_small_categories(category_times, total_time, threshold=0.02):
    # Create a new dictionary for the updated categories
    updated_category_times = {"other": 0}  # Start with "other" category

    for category, time in category_times.items():
        # Calculate the percentage of total time
        if time / total_time < threshold:
            updated_category_times["other"] += time  # Add to "other" if below the threshold
        else:
            updated_category_times[category] = time  # Keep the category as is if above the threshold

    if updated_category_times["other"] == 0:
        del updated_category_times["other"] #don't show other cat if other is no time

    return updated_category_times


# Reclassify small categories as "other"
total_time = sum(category_times.values())  # Total time spent
updated_category_times = reclassify_small_categories(category_times, total_time)

# Convert minutes to hours for display
def minutes_to_hours(minutes):
    return round(minutes / 60, 2)


# --- Google Calendar event colorId → hex (classic Google palette)
GCAL_COLOR_HEX = {
    "1":  "#a4bdfc",  # Lavender
    "2":  "#7ae7bf",  # Sage
    "3":  "#dbadff",  # Grape
    "4":  "#ff887c",  # Flamingo
    "5":  "#fbd75b",  # Banana
    "6":  "#ffb878",  # Tangerine
    "7":  "#46d6db",  # Peacock
    "8":  "#e1e1e1",  # Graphite
    "9":  "#5484ed",  # Blueberry
    "10": "#51b749",  # Basil
    "11": "#dc2127",  # Tomato
}

def hex_to_rgb_tuple(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16)/255 for i in (0, 2, 4))

# Rebuild COLOR_MAP to follow your CATEGORY_COLOR_ID / DEFAULT_COLOR_ID
def color_for_category(cat: str):
    cid = CATEGORY_COLOR_ID.get(cat, DEFAULT_COLOR_ID)
    hexv = GCAL_COLOR_HEX.get(cid, GCAL_COLOR_HEX.get(DEFAULT_COLOR_ID))
    return hex_to_rgb_tuple(hexv)

ALL_CATEGORIES = list(categories.keys()) + ["other"]
COLOR_MAP = {cat: color_for_category(cat) for cat in ALL_CATEGORIES}
ordered_cats = [c for c in ALL_CATEGORIES if c in updated_category_times and updated_category_times[c] > 0]


def plot_total_pie():
    sizes = [updated_category_times[c] for c in ordered_cats]
    labels = [f"{c} ({minutes_to_hours(updated_category_times[c])} hrs)" for c in ordered_cats]
    colors = [COLOR_MAP[c] for c in ordered_cats]
    plt.figure(figsize=(8, 6))
    plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.title('Time Distribution by Activity Category')
    timestamp = datetime.now().strftime("%H:%M:%S")
    plt.figtext(0.5, 0.02, f"Generated at {timestamp}", ha="center", fontsize=9, color="gray")
    plt.axis('equal')
    plt.show()


def compute_day_stats(events, start_dt, end_dt, tz_name='Europe/Berlin'):
    tz = pytz.timezone(tz_name)
    start_day = start_dt.astimezone(tz).date()
    end_day   = end_dt.astimezone(tz).date()      # exclusive

    # exact set of valid days in the query window
    valid_days = {start_day + timedelta(days=i)
                  for i in range((end_day - start_day).days)}

    event_days = set()
    for e in events:
        dt = event_start_dt_local(e, tz)
        if dt is None:
            continue
        d = dt.date()
        if d in valid_days:           # clamp to window
            event_days.add(d)

    total_event_days = len(event_days)
    working_days = sum(1 for d in event_days if d.weekday() < 5)

    wd_counter = Counter(d.weekday() for d in event_days)
    names = ['Mondays','Tuesdays','Wednesdays','Thursdays','Fridays','Saturdays','Sundays']
    weekday_lines = {names[i]: wd_counter.get(i, 0) for i in range(7)}
    return total_event_days, working_days, weekday_lines

# ---- call after you have `events_in_timeframe` ----
total_days, working_days, weekday_lines = compute_day_stats(
    events_in_timeframe, start_date, end_date
)

rows = [(name, weekday_lines[name]) for name in ['Mondays','Tuesdays','Wednesdays','Thursdays','Fridays','Saturdays','Sundays']]
print(f"\nTotal Days: {total_days} | Working Days: {working_days} | Weekend Days: {total_days - working_days}\n")
print(tabulate(rows, headers=["Weekday", "Count"], tablefmt="pretty"))
print(f"\nBlacklist days: {', '.join(str(d) for d in sorted(blacklist_dates))}")


def keep_timed_and_active(evts):
    out = []
    for e in evts:
        if e.get("status") == "cancelled":
            continue
        if "dateTime" not in e.get("start", {}):
            continue
        out.append(e)
    return out


print(f"Events fetched: {len(keep_timed_and_active(events_from_calendar))}")
print(f"After blacklist: {len(keep_timed_and_active(events_in_timeframe))}")


def minutes(e, tz):
    s, t = event_local_start_end(e, tz)
    return (t - s).total_seconds() / 60, s.date()


def per_day_category_minutes(events, tz_name='Europe/Berlin'):
    tz = pytz.timezone(tz_name)
    per_day = defaultdict(lambda: defaultdict(float))  # {date: {category: minutes}}
    for e in events:
        if "dateTime" not in e.get("start", {}):
            continue  # keep skip rule; remove to include all-day
        dur_min, day = minutes(e, tz)
        cat = categorize_event(e)
        per_day[day][cat] += dur_min
    return per_day


def avg_minutes_by_daytype(per_day):
    # Split days
    work_days = [d for d in per_day if d.weekday() < 5]
    weekend_days = [d for d in per_day if d.weekday() >= 5]

    def avg_for(days):
        if not days:
            return {}
        total = defaultdict(float)
        for d in days:
            for cat, mins in per_day[d].items():
                total[cat] += mins
        # minutes per day -> average
        n = len(days)
        return {cat: m / n for cat, m in total.items()}

    return avg_for(work_days), avg_for(weekend_days), len(work_days), len(weekend_days)


# --- compute ---
per_day = per_day_category_minutes(events_in_timeframe)
avg_work, avg_weekend, n_work, n_weekend = avg_minutes_by_daytype(per_day)


def plot_pie_from_minutes_map(title, cat_min_map):
    if not cat_min_map:
        print(f"{title}: no data")
        return

    # optional: collapse tiny categories into "other" (2% threshold)
    total = sum(cat_min_map.values())
    data = dict(cat_min_map)
    if total > 0:
        other = 0.0
        for k in list(data.keys()):
            if data[k] / total < 0.02:
                other += data.pop(k)
        if other > 0:
            data["other"] = data.get("other", 0) + other

    # stable ordering + stable colors
    cats = [c for c in ALL_CATEGORIES if c in data and data[c] > 0]
    sizes = [data[c] for c in cats]
    labels = [f"{c} ({round(data[c]/60, 2)} hrs)" for c in cats]
    colors = [COLOR_MAP[c] for c in cats]

    plt.figure(figsize=(8, 6))
    plt.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90)
    plt.title(title)

    # Add timestamp label below chart
    timestamp = datetime.now().strftime("%H:%M:%S")
    plt.figtext(0.5, 0.02, f"Generated at {timestamp}", ha="center", fontsize=9, color="gray")

    plt.axis('equal')
    plt.show()


def write_to_sheet(per_day, all_categories):
    creds = authenticate_google_account()
    sheets_service = build('sheets', 'v4', credentials=creds)
    sheet = sheets_service.spreadsheets()

    # Read all existing data
    result = sheet.values().get(spreadsheetId=SHEET_ID, range='DataDay').execute()
    existing_rows = result.get('values', [])

    # Determine column order: preserve existing header, append any new categories
    if existing_rows and existing_rows[0] and existing_rows[0][0] == 'Date':
        existing_header = existing_rows[0]
        new_cats = [c for c in all_categories if c not in existing_header]
        header = existing_header + new_cats
    else:
        header = ['Date'] + list(all_categories)
        existing_rows = []

    cat_columns = header[1:]

    # Write header (in case it's new or gained columns)
    sheet.values().update(
        spreadsheetId=SHEET_ID,
        range='DataDay!A1',
        valueInputOption='USER_ENTERED',
        body={'values': [header]}
    ).execute()

    # Build date → row number map (1-based; row 1 is the header)
    date_to_row = {}
    for i, row in enumerate(existing_rows[1:], start=2):
        if row and row[0]:
            date_to_row[row[0]] = i

    batch_updates = []
    appends = []

    for day in sorted(per_day.keys()):
        date_str = day.strftime('%Y-%m-%d')
        cat_minutes = per_day[day]
        row_values = [date_str] + [round(cat_minutes.get(cat, 0) / 60, 2) for cat in cat_columns]

        if date_str in date_to_row:
            batch_updates.append({
                'range': f'DataDay!A{date_to_row[date_str]}',
                'values': [row_values]
            })
        else:
            appends.append(row_values)

    if batch_updates:
        sheet.values().batchUpdate(
            spreadsheetId=SHEET_ID,
            body={'valueInputOption': 'USER_ENTERED', 'data': batch_updates}
        ).execute()

    if appends:
        sheet.values().append(
            spreadsheetId=SHEET_ID,
            range='DataDay!A1',
            valueInputOption='USER_ENTERED',
            insertDataOption='INSERT_ROWS',
            body={'values': appends}
        ).execute()

    print(f"\nSheet updated: {len(batch_updates)} row(s) updated, {len(appends)} row(s) added.")


write_to_sheet(per_day, ALL_CATEGORIES)

if UNCATEGORIZED_EVENTS:
    print("\nUncategorized events (fix these before pie charts will show):")
    seen = set()
    for e in UNCATEGORIZED_EVENTS:
        s = e.get("summary", "<no summary>")
        if s not in seen:
            seen.add(s)
            print("-", s)
else:
    plot_total_pie()
    plot_pie_from_minutes_map(f'Average per Working Day (n={n_work})', avg_work)
    plot_pie_from_minutes_map(f'Average per Weekend Day (n={n_weekend})', avg_weekend)
