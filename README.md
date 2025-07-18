# CreateSoundRoomSchedule

[![.NET](https://github.com/markjulmar/CreateSoundRoomSchedule/actions/workflows/dotnet.yml/badge.svg)](https://github.com/markjulmar/CreateSoundRoomSchedule/actions/workflows/dotnet.yml)

## Overview
The `CreateSoundRoomSchedule` project is a C# application to generate an Excel spreadsheet with the planned services for a church. It interacts with the [Planning Center](https://www.planningcenter.com/) API to create a three-month calendar set for a given quarter. This .NET Core application reads data from Planning Center Online for all planned services in a specific quarter. It generates an Excel spreadsheet containing a 3-month calendar to identify the people assigned to manage A/V weekly. It also pulls public holidays and includes them in the calendar data.

![Sample calendar in Excel](./images/calendar-example.png)

### Dependencies
- **Language**: C#
- **Framework**: .NET 8.0
- **Main Libraries**: [EPPlus](https://www.epplussoftware.com/)

## Requirements
Add a .NET User Secret with a Client Id and Client Secret obtained from the Planning Center development dashboard.

| Key | Description |
|-----|-------------|
| `PlanningCenter:clientId` | Client Id for a PAT assigned by Planning Center. |
| `PlanningCenter:clientSecret` | Corresponding PAT secret for the assigned client id. |

## Running the app
Run the app from the console:

```console
dotnet run
```

By default, it generates a calendar for the _next_ quarter. However, you can give it a date as an optional parameter, and it will generate the calendar for whatever quarter the date falls in. The resulting file will be placed onto the desktop and named **TFC_SoundRoom_Schedule_YYYY-QX** where `YYYY` will be the year, and `QX` will be the quarter. If a file with that name exists, it will be replaced.

```console
dotnet run 2/1/2025
```

This command would create a Q1 2025 calendar spanning January - March.