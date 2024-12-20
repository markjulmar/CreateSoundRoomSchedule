# CreateSoundRoomSchedule

[![.NET](https://github.com/markjulmar/CreateSoundRoomSchedule/actions/workflows/dotnet.yml/badge.svg)](https://github.com/markjulmar/CreateSoundRoomSchedule/actions/workflows/dotnet.yml)

This .NET Core application reads data from Planning Center Online for all planned services in a specific quarter. It generates an Excel spreadsheet containing a 3-month calendar to identify the specific people assigned to manage A/V each week. It also pulls public holidays and includes them in the calendar data.

## Running the app

Run the app from the console:

```console
dotnet run
```

By default, it generates a calendar for the _next_ quarter, however you can give it a date as an optional parameter and it will generate the calendar for whatever quarter the date falls in.

```console
dotnet run 2/1/2025
```

Would create a Q1 2025 calendar spanning January - March.
