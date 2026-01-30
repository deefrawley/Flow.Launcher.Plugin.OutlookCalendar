# Outlook Agenda (Flow.Launcher.Plugin.OutlookCalendar)

Outlook calendar/agenda viewer for the [Flow Launcher](https://github.com/Flow-Launcher/Flow.Launcher)

### About

Requires a local installation of Outlook to be present. Does not work with new 365 accounts or web based Outlook.

### Usage

Default keyword is 'ocal'

| Keyword                                            | Description                                                              |
| -------------------------------------------------- | ------------------------------------------------------------------------ |
| `olcal today` or `olcal t`                         | Show meetings for today. Past meetings marked as (COMPLETED)             |
| `olcal tomorrow` or `olcal tm`                     | Show meetings for tomorrow                                               |
| `olcal week` or `olcal w`                          | Show meetings for the current week. Past meetings marked as (COMPLETED)  |
| `olcal nextweek` or `olcal nw`                     | Show meetings for next week                                              |
| `olcal month` or `olcal m`                         | Show meetings for the current month. Past meetings marked as (COMPLETED) |
| `olcal YYYY-MM-DD`                                 | Show meetings for a custom day                                           |

Result will show the subject of the meeting and whether it is part of a recurring series as the title, with the result
sub-title showing the start date and time and end time, duration of the meeting, and the location (if given)
