"""Pure helpers for shared date picker widgets."""

import calendar
import tkinter as tk


CALENDAR_MONTHS_FR = (
    "Janvier",
    "Février",
    "Mars",
    "Avril",
    "Mai",
    "Juin",
    "Juillet",
    "Août",
    "Septembre",
    "Octobre",
    "Novembre",
    "Décembre",
)


def get_calendar_day_headers():
    """Return the weekday headers used by the shared calendar dialog."""
    return ("Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim")


def get_calendar_weeks(year, month):
    """Return the month grid with Monday as the first weekday."""
    return calendar.Calendar(firstweekday=calendar.MONDAY).monthdayscalendar(
        year, month
    )


def get_calendar_year_options(current_year, span=10):
    """Return a centered year range for the year selector."""
    start_year = current_year - span
    end_year = current_year + span
    return list(range(start_year, end_year + 1))


def apply_modal_grab(window):
    """Wait until a dialog is viewable before applying a modal grab."""
    try:
        window.wait_visibility()
        window.grab_set()
        return True
    except tk.TclError:
        return False
