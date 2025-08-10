import os
import pickle
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from datetime import date
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import pytz
import matplotlib.cm as cm
from collections import Counter
from collections import defaultdict
from tabulate import tabulate
import os
import yaml
from datetime import datetime, date

# If modifying the calendar, the required scopes are 'https://www.googleapis.com/auth/calendar'
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
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
def get_date_range():
    today = datetime.today().date()
    current_year = today.year

    start_in = input(f"Enter start date (DD-MM) or last X days [default {today - timedelta(days=7):%d-%m}]: ").strip()
    if start_in.isdigit():
        start_date = today - timedelta(days=int(start_in))
        end_date = today
    else:
        end_in = input(f"Enter end date (DD-MM) [default {today:%d-%m}]: ").strip()

        if start_in:
            start_date = datetime.strptime(f"{start_in}-{current_year}", "%d-%m-%Y").date()
        else:
            start_date = today - timedelta(days=7)

        if end_in:
            end_date = datetime.strptime(f"{end_in}-{current_year}", "%d-%m-%Y").date()
        else:
            end_date = today

    return localizeTime(start_date, end_date + timedelta(days=1))  # NOTE: +1 day


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


# Get the date range from the user
start_date, end_date = get_date_range()

# Fetch the events from Google Calendar in the given date range
events_from_calendar = get_events_in_date_range(start_date, end_date)

# Now you can proceed with filtering, categorization, and plotting as before
events_in_timeframe = remove_all_day(events_from_calendar)

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
            if keyword.lower() in summary:
                return category
    # collect instead of printing
    UNCATEGORIZED_EVENTS.append(event)
    return "other"


ALL_CATEGORIES = list(categories.keys()) + ["other"]
_PALETTE = cm.get_cmap('tab20', len(ALL_CATEGORIES))
COLOR_MAP = {cat: _PALETTE(i) for i, cat in enumerate(ALL_CATEGORIES)}


# Function to calculate the duration of an event in minutes
def calculate_duration(event):
    tz = pytz.timezone('Europe/Berlin')
    s = datetime.fromisoformat(event["start"]["dateTime"].replace('Z', '+00:00')).astimezone(tz)
    e_iso = event["end"].get("dateTime", event["start"]["dateTime"])
    e = datetime.fromisoformat(e_iso.replace('Z', '+00:00')).astimezone(tz)
    return max(0, (e - s).total_seconds() / 60)


for event in events_in_timeframe:
    category = categorize_event(event)
    duration = calculate_duration(event)
    category_times[category] += duration


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

# Plotting the results
labels = list(updated_category_times.keys())
sizes = list(updated_category_times.values())


# Convert minutes to hours for display
def minutes_to_hours(minutes):
    return round(minutes / 60, 2)


# Prepare labels with hours spent
labels_with_hours = [f"{label} ({minutes_to_hours(time)} hrs)" for label, time in updated_category_times.items()]

# Plotting the results with hours spent in the labels (TOTAL)
plt.figure(figsize=(8, 6))
# keep category order stable based on ALL_CATEGORIES
ordered_cats = [c for c in ALL_CATEGORIES if c in updated_category_times and updated_category_times[c] > 0]
ordered_sizes = [updated_category_times[c] for c in ordered_cats]
ordered_labels = [f"{c} ({minutes_to_hours(updated_category_times[c])} hrs)" for c in ordered_cats]
ordered_colors = [COLOR_MAP[c] for c in ordered_cats]

plt.pie(ordered_sizes, labels=ordered_labels, colors=ordered_colors, autopct='%1.1f%%', startangle=90)
plt.title('Time Distribution by Activity Category')

# Add timestamp label below chart
timestamp = datetime.now().strftime("%H:%M:%S")
plt.figtext(0.5, 0.02, f"Generated at {timestamp}", ha="center", fontsize=9, color="gray")

plt.axis('equal')
plt.show()


def event_start_dt(e, tz):
    if "dateTime" not in e.get("start", {}):  # skip all-day
        return None
    iso = e["start"]["dateTime"].replace('Z', '+00:00')
    return datetime.fromisoformat(iso).astimezone(tz)


def compute_day_stats(events, tz_name='Europe/Berlin'):
    tz = pytz.timezone(tz_name)
    # unique dates that have >=1 event
    event_days = set()
    for e in events:
        dt = event_start_dt(e, tz)
        if dt is not None:
            event_days.add(dt.date())

    total_event_days = len(event_days)
    working_days = sum(1 for d in event_days if d.weekday() < 5)

    # weekday distribution (0=Mon ... 6=Sun)
    wd_counter = Counter(d.weekday() for d in event_days)
    names = ['Mondays','Tuesdays','Wednesdays','Thursdays','Fridays','Saturdays','Sundays']
    weekday_lines = {names[i]: wd_counter.get(i, 0) for i in range(7)}

    return total_event_days, working_days, weekday_lines


# ---- call after you have `events_in_timeframe` ----
total_days, working_days, weekday_lines = compute_day_stats(events_in_timeframe)


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


def event_local_start_end(e, tz):
    s = datetime.fromisoformat(e["start"]["dateTime"].replace('Z', '+00:00')).astimezone(tz)
    e_iso = e["end"].get("dateTime", e["start"]["dateTime"])
    t = datetime.fromisoformat(e_iso.replace('Z', '+00:00')).astimezone(tz)
    return s, t


def minutes(e, tz):
    s, t = event_local_start_end(e, tz)
    return max(0, (t - s).total_seconds() / 60), s.date()  # duration, start-date


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


# New pies (averages):
plot_pie_from_minutes_map(f'Average per Working Day (n={n_work})', avg_work)
plot_pie_from_minutes_map(f'Average per Weekend Day (n={n_weekend})', avg_weekend)

# log all events that didn't fit a category and got put into 'other'
if UNCATEGORIZED_EVENTS:
    print("\nUncategorized events:")
    seen = set()
    for e in UNCATEGORIZED_EVENTS:
        s = e.get("summary", "<no summary>")
        if s not in seen:           # dedupe summaries; remove if you want all
            seen.add(s)
            print("-", s)

