"""Flatten Alvys Loads into per-stop rows and compute detention time.

Field names below are taken from a live /loads/search response (verified
2026-08-26), not guessed from docs. `hours_from_appointment` is the full
span from appointment to departure (so it includes the free 2-hour window).
`detention_hours` is the billable part -- only the time *past* that free
window (0 if the stop wasn't held that long). A stop is flagged when
`hours_from_appointment` exceeds `DETENTION_THRESHOLD_HOURS`, which is the
same thing as `detention_hours` being positive.

For FCFS stops (no fixed appointment, just a window) the "appointment" used
is the end of that window -- the time by which the stop should have been
done -- since there's no single appointment timestamp to measure from.
"""

import pandas as pd

DETENTION_THRESHOLD_HOURS = 2


def _appointment(stop):
    if stop.get("AppointmentDate"):
        return stop["AppointmentDate"]
    window = stop.get("StopWindow") or {}
    return window.get("End")


def _location(stop):
    address = stop.get("Address") or {}
    city = address.get("City")
    state = address.get("State")
    city_state = ", ".join(part for part in (city, state) if part)
    street = address.get("Street")
    if street and city_state:
        return f"{street}, {city_state}"
    return street or city_state or ""


def flatten_loads(loads):
    """Turn a list of raw Alvys Load dicts into one row per stop."""
    rows = []
    for load in loads:
        load_number = load.get("LoadNumber", "")
        customer = load.get("CustomerName") or "Unknown Customer"
        stops = load.get("Stops") or []

        for i, stop in enumerate(stops):
            appointment = pd.to_datetime(_appointment(stop), errors="coerce", utc=False)
            arrival = pd.to_datetime(stop.get("ArrivedAt"), errors="coerce", utc=False)
            departure = pd.to_datetime(stop.get("DepartedAt"), errors="coerce", utc=False)

            hours_from_appt = None
            detention_hours = None
            if pd.notna(departure) and pd.notna(appointment):
                hours_from_appt = round((departure - appointment).total_seconds() / 3600, 2)
                detention_hours = round(max(0.0, hours_from_appt - DETENTION_THRESHOLD_HOURS), 2)

            dwell_hours = None
            if pd.notna(departure) and pd.notna(arrival):
                dwell_hours = round((departure - arrival).total_seconds() / 3600, 2)

            rows.append(
                {
                    "load_number": load_number,
                    "customer": customer,
                    "stop_sequence": i + 1,
                    "stop_type": stop.get("StopType", ""),
                    "location": _location(stop),
                    "appointment": appointment.isoformat() if pd.notna(appointment) else None,
                    "arrival": arrival.isoformat() if pd.notna(arrival) else None,
                    "departure": departure.isoformat() if pd.notna(departure) else None,
                    "hours_from_appointment": hours_from_appt,
                    "detention_hours": detention_hours,
                    "dwell_hours": dwell_hours,
                    "detention_flag": hours_from_appt is not None and hours_from_appt > DETENTION_THRESHOLD_HOURS,
                }
            )
    return rows


def filter_rows_by_date(rows, start_date, end_date):
    """Keep rows whose stop activity (departure, else arrival, else
    appointment) falls within [start_date, end_date] (inclusive, date-only
    bounds). Rows with none of those three timestamps are dropped.

    Needed because the Alvys /loads/search API doesn't filter by date
    server-side -- see alvys_client.py.
    """
    start = pd.Timestamp(start_date).tz_localize(None)
    end = pd.Timestamp(end_date).tz_localize(None) + pd.Timedelta(days=1)

    def _in_range(row):
        for key in ("departure", "arrival", "appointment"):
            value = row.get(key)
            if not value:
                continue
            ts = pd.Timestamp(value)
            if ts.tzinfo is not None:
                ts = ts.tz_localize(None)
            return start <= ts < end
        return False

    return [row for row in rows if _in_range(row)]
