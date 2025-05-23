from __future__ import annotations

from flogin import Query, Result, SearchHandler

from ..plugin import OutlookAgendaPlugin

import win32com.client
import pywintypes
import re
import pytz
import locale
from datetime import datetime, timedelta, time


class GetOutlookAgenda(SearchHandler[OutlookAgendaPlugin]):
    async def callback(self, query: Query):
        assert self.plugin
        query_regex = re.compile(
            r"^(today|t|tomorrow|tm|week|w|month|m|custom\s+\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])(?:\s+([01]\d|2[0-3]):[0-5]\d)?\s+\d{4}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])(?:\s+([01]\d|2[0-3]):[0-5]\d)?)$"
        )
        query_find = query_regex.match(query.text)
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            if query_find is None:
                yield Result(
                    title="today(t), tomorrow(tm), week(w), month(m), custom (from date, end date in YYYY-MM-DD [HH:MM] format)",
                    sub="Shortcuts in brackets. For custom two dates are required but times optional",
                    icon="assets/app.png",
                )
            else:
                # Get the current UTC time for use below
                utc_now = datetime.now().astimezone(pytz.timezone("UTC"))
                match query_find[0]:
                    # Capture the custom date range. Regex has ensured the format is correct
                    case "custom":
                        # Either custom date can have a time component so some work to do to get to valid datetimes
                        parts = x.split()
                        # Remove "custom" so we can just deal with the dates
                        parts = parts[1:]

                        # Deal with good path - 4 parts mean date and time for both
                        if len(parts) == 4:
                            date1_str = f"{parts[0]} {parts[1]}"
                            date2_str = f"{parts[2]} {parts[3]}"

                        # 2 partsx means 2 dates with no time
                        elif len(parts) == 2:
                            date1_str = parts[0]
                            date2_str = parts[1]

                        # Bad path - 3 parts means date and time for one and date for the other
                        elif len(parts) == 3:
                            # Check if first part has time
                            if ":" in parts[1]:
                                date1_str = f"{parts[0]} {parts[1]}"
                                date2_str = parts[2]
                            else:
                                date1_str = parts[0]
                                date2_str = f"{parts[1]} {parts[2]}"

                        # Parse dates (your existing parsing logic is fine here)
                        if " " in date1_str:
                            start_date = datetime.strptime(date1_str, "%Y-%m-%d %H:%M")
                        else:
                            date_part = datetime.strptime(date1_str, "%Y-%m-%d")
                            start_date = datetime.combine(date_part.date(), time(0, 1))

                        if " " in date2_str:
                            end_date = datetime.strptime(date2_str, "%Y-%m-%d %H:%M")
                        else:
                            date_part = datetime.strptime(date2_str, "%Y-%m-%d")
                            end_date = datetime.combine(date_part.date(), time(23, 59))

                        meetings = self._get_meetings(start_date, end_date)
                        if not meetings:
                            yield Result(
                                title="No meetings found",
                                sub=f"No meetings found for {start_date} to {end_date}",
                                icon="assets/app.png",
                            )
                        else:
                            for meeting in meetings:
                                subj = self._meeting_result_sub(
                                    meeting["start"],
                                    meeting["end"],
                                    meeting["location"],
                                )
                                yield Result(
                                    title=f"{meeting['subject']} {'(recurring)' if meeting['is_recurring'] else ''}",
                                    sub=subj,
                                    icon="assets/app.png",
                                )

                    case "today" | "t":
                        start_date, end_date = self._get_date_range("today")
                        meetings = self._get_meetings(start_date, end_date) or []
                        if not meetings:
                            yield Result(
                                title="No meetings found",
                                sub="No meetings found for today",
                                icon="assets/app.png",
                            )
                        else:
                            meeting_count = 0
                            for meeting in meetings:
                                # Filter out past meetings. We can use UTC as it's just time math not local TZ specific
                                if meeting["end_utc"] >= utc_now:
                                    meeting_count += 1
                                    subj = self._meeting_result_sub(
                                        meeting["start"],
                                        meeting["end"],
                                        meeting["location"],
                                    )
                                    yield Result(
                                        title=f"{meeting['subject']} {'(recurring)' if meeting['is_recurring'] else ''}",
                                        sub=subj,
                                        icon="assets/app.png",
                                    )
                            if meeting_count == 0:
                                # No meetings found for the remainder of today
                                yield Result(
                                    title="No meetings found",
                                    sub="No meetings found for the remainder of today",
                                    icon="assets/app.png",
                                )

                    case "tomorrow" | "tm":
                        start_date, end_date = self._get_date_range("tomorrow")
                        meetings = self._get_meetings(start_date, end_date)
                        if not meetings:
                            yield Result(
                                title="No meetings found",
                                sub="No meetings found for tomorrow",
                                icon="assets/app.png",
                            )
                        else:
                            for meeting in meetings:
                                subj = self._meeting_result_sub(
                                    meeting["start"],
                                    meeting["end"],
                                    meeting["location"],
                                )
                                yield Result(
                                    title=f"{meeting['subject']} {'(recurring)' if meeting['is_recurring'] else ''}",
                                    sub=subj,
                                    icon="assets/app.png",
                                )

                    case "week" | "w":
                        start_date, end_date = self._get_date_range("week")
                        meetings = self._get_meetings(start_date, end_date)
                        if not meetings:
                            yield Result(
                                title="No meetings found",
                                sub="No meetings found for the week",
                                icon="assets/app.png",
                            )
                        else:
                            for meeting in meetings:
                                # Filter out past meetings. We can use UTC as it's just time math not local TZ specific
                                if meeting["end_utc"] >= utc_now:
                                    subj = self._meeting_result_sub(
                                        meeting["start"],
                                        meeting["end"],
                                        meeting["location"],
                                    )
                                    yield Result(
                                        title=f"{meeting['subject']} {'(recurring)' if meeting['is_recurring'] else ''}",
                                        sub=subj,
                                        icon="assets/app.png",
                                    )
                    case "month" | "m":
                        start_date, end_date = self._get_date_range("month")
                        meetings = self._get_meetings(start_date, end_date)
                        if not meetings:
                            yield Result(
                                title="No meetings found",
                                sub="No meetings found for the month",
                                icon="assets/app.png",
                            )
                        else:
                            for meeting in meetings:
                                # Filter out past meetings. We can use UTC as it's just time math not local TZ specific
                                if meeting["end_utc"] >= utc_now:
                                    subj = self._meeting_result_sub(
                                        meeting["start"],
                                        meeting["end"],
                                        meeting["location"],
                                    )
                                    yield Result(
                                        title=f"{meeting['subject']} {'(recurring)' if meeting['is_recurring'] else ''}",
                                        sub=subj,
                                        icon="assets/app.png",
                                    )
        except pywintypes.com_error as e:
            yield Result(
                title=f"Error - {str(e)}",
                sub="This plugin won't work (most likely reason is Outlook not installed)",
                icon="assets/app_err.png",
            )

    def _get_date_range(self, period):
        """Calculate start and end dates based on period"""
        now = datetime.now()

        period = period.lower()
        if period == "today":
            start = now.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=1) - timedelta(seconds=1)
        elif period == "tomorrow":
            start = (now + timedelta(days=1)).replace(
                hour=0, minute=0, second=0, microsecond=0
            )
            end = start + timedelta(days=1) - timedelta(seconds=1)
        elif period == "week":
            start = now - timedelta(days=now.weekday())
            start = start.replace(hour=0, minute=0, second=0, microsecond=0)
            end = start + timedelta(days=6, hours=23, minutes=59, seconds=59)
        elif period == "month":
            start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
            if now.month == 12:
                end = start.replace(year=start.year + 1, month=1)
            else:
                end = start.replace(month=start.month + 1)
            end = end - timedelta(seconds=1)

        return start, end

    def _parse_datetime(self, dt_str, default_hour, default_minute):
        """Parse datetime string with flexible format handling"""
        try:
            return datetime.strptime(dt_str, "%Y-%m-%d %H:%M")
        except ValueError:
            dt = datetime.strptime(dt_str, "%Y-%m-%d")
            return dt.replace(hour=default_hour, minute=default_minute)

    def _meeting_result_sub(self, appt_s, appt_e, location):
        appt_s_date = appt_s.strftime("%a %d %b")
        appt_s_time = appt_s.strftime("%I:%M %p")
        appt_e_date = appt_e.strftime("%a %d %b")
        appt_e_time = appt_e.strftime("%I:%M %p")
        duration = appt_e - appt_s
        minutes, seconds = divmod(duration.seconds, 60)
        hours, minutes = divmod(minutes, 60)
        if hours == 0:
            return f"{appt_s_date} {appt_s_time} to {appt_e_time} ({minutes} m), Location: {location if location else 'None specified'}"
        elif minutes == 0:
            return f"{appt_s_date} {appt_s_time} to {appt_e_time} ({hours} h), Location: {location if location else 'None specified'}"
        else:
            return f"{appt_s_date} {appt_s_time} to {appt_e_time} ({hours} h {minutes} m), Location: {location if location else 'None specified'}"

    def _get_meetings(
        self,
        start_date,
        end_date,
        subject_filter=None,
        organizer_filter=None,
        attendee_filter=None,
    ):
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        calendar = outlook.GetDefaultFolder(9).Items
        calendar.Sort("[Start]")
        calendar.IncludeRecurrences = True

        # Format dates for Outlook filter
        locale.setlocale(locale.LC_TIME, "")  # Use system default
        start_str = start_date.strftime("%x %I:%M %p")
        end_str = end_date.strftime("%x %I:%M %p")

        filter_str = f"[Start] >= '{start_str}' AND [End] <= '{end_str}'"
        appointments = calendar.Restrict(filter_str)

        meetings = []

        for appointment in appointments:
            try:
                # Apply filters
                if (
                    subject_filter
                    and subject_filter.lower() not in appointment.Subject.lower()
                ):
                    continue

                if (
                    organizer_filter
                    and organizer_filter.lower() not in appointment.Organizer.lower()
                ):
                    continue

                if attendee_filter:
                    attendees = appointment.RequiredAttendees or ""
                    if attendee_filter.lower() not in attendees.lower():
                        continue

                meetings.append(
                    {
                        "subject": appointment.Subject,
                        "start": appointment.start,
                        "start_utc": appointment.StartUTC,
                        "end": appointment.end,
                        "end_utc": appointment.EndUTC,
                        "organizer": appointment.Organizer,
                        "required_attendees": appointment.RequiredAttendees,
                        "location": appointment.Location,
                        "body": appointment.Body,
                        "is_recurring": appointment.IsRecurring,
                    }
                )
            except Exception:
                continue

        return meetings
