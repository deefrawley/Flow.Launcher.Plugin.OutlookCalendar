# Outlook Agendar (Flow.Launcher.Plugin.OutlookCalendar)

Outlook calendar/agenda viewer for the [Flow Launcher](https://github.com/Flow-Launcher/Flow.Launcher)

### About

Requires a local installation of Outlook to be present. Does not work with new 365 accounts or web based Outlook.

### Usage

Default keyword is 'ocal'

| Keyword                                            | Description                                                       |
| -------------------------------------------------- | ----------------------------------------------------------------- |
| `olcal today` or `olcal t`                         | Show meetings for the remainder of the day (past events not shown)|
| `olcal tomorrow` or `olcal tm`                     | Show meetings for tomorrow                                        |
| `olcal week` or `olcal w`                          | Show meetings for the remainder of the week                       |
| `olcal month` or `olcal m`                         | Show meetings for the remainder of the calendar month             |
| `olcal custom YYYY-MM-DD (HH:MM) YYY-MM-DD (HH:MM)`| Show meetings for a custom range. Both dates are required but times are optional. If no time given then default is from midnight to midnight |

Result will show the subject of the meeting and whether it is part of a recurring series as the title, with the result
sub-title showing the start date and time and end time, duration of the meeting, and the location (if given)
