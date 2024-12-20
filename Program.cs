﻿using System.Diagnostics;
using System.Drawing;
using System.Text;
using CreateSoundRoomSchedule;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;

var configuration = new ConfigurationBuilder()
    .AddUserSecrets<Program>() // Link to user secrets
    .Build();

var clientId = configuration["PlanningCenter:clientId"];
var clientSecret = configuration["PlanningCenter:clientSecret"];

var planningCenter = new PlanningCenter(clientId, clientSecret);

// Non commercial use.
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

// Missing date - use next quarter
if (args.Length != 1 
    || !DateOnly.TryParse(args[0], out var date))
{
    var now = DateTime.Now;
    var month = now.Month <= 3 ? 4 : now.Month <= 6 ? 7 : now.Month <= 9 ? 10 : 1;
    date = new DateOnly(month==1?now.Year+1:now.Year, month, 1);
}

Console.WriteLine($"CreateSoundRoomSchedule for {date.ToShortDateString()}");
Console.WriteLine();
Console.WriteLine("Retrieving data from Planning Center Online...");

var services = await planningCenter.GetServicesAsync(date, date.AddMonths(3)).ToListAsync();
Console.WriteLine($"... found {services.Count} services.");
Console.WriteLine();
Console.WriteLine("Creating Excel file...");

var quarter = $"{date:yyyy}-Q{date.Month / 4 + 1}";
var filename = $"TFC_SoundRoom_Schedule_{quarter}.xlsx";
var fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), filename);
if (File.Exists(fullPath))
    File.Delete(fullPath);

// Get the holidays
var holidays = (await Holiday.GetAllAsync(date.Year)).ToList();

using var package = new ExcelPackage(new FileInfo(fullPath));

// Legacy flag to use 1-based collections (porting from desktop .net)
package.Compatibility.IsWorksheets1Based = true;

// Set worksheet properties.
package.Workbook.Properties.Title = $"TFC Sound Room Schedule - {quarter}";
package.Workbook.Properties.Author = "Julie Smith";
package.Workbook.Properties.Comments = $"This is the sound room schedule for A/V at Trinity Fellowship Church for {quarter}.";
package.Workbook.Properties.Company = "Trinity Fellowship Church";
package.Workbook.Properties.Created = DateTime.Now;

// Add a sheet.
var worksheet = package.Workbook.Worksheets.Add(quarter);
worksheet.PrinterSettings.TopMargin = .7M;
worksheet.PrinterSettings.LeftMargin = .75M;
worksheet.PrinterSettings.BottomMargin = .7M;
worksheet.PrinterSettings.RightMargin = .75M;
worksheet.PrinterSettings.HeaderMargin = .5M;
worksheet.PrinterSettings.FooterMargin = .5M;
worksheet.Cells.Style.Font.Name = "Arial";

// Set pagebreak after column 7.
worksheet.Columns[7].PageBreak = true;

// Add each month.
int row = 1;
for (int i = 0; i < 3; i++)
{
    var month = new DateOnly(date.Year, date.Month + i, 1);
    AddMonth(worksheet, month, holidays, services, ref row);
}

// Autofit columns.
worksheet.Cells.AutoFitColumns(22);

worksheet.PrinterSettings.PrintArea = worksheet.Cells[1, 1, row, 7];

// Save the package
package.Save();

Console.WriteLine($"File saved to {fullPath}");

return;

static void AddMonth(ExcelWorksheet ws, DateOnly month, List<Holiday> holidays, List<Service> services, ref int row)
{
    var title = $"{month.ToString("MMMM yyyy")} Trinity Fellowship Church A/V Schedule";

    ws.Rows[row].Height = 55;
    
    // First row has the title - merged across seven columns, centered.
    using (var range = ws.Cells[row, 1, row, 7])
    {
        range.Merge = true;
        range.Style.Font.Bold = true;
        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        range.Style.Font.Color.SetColor(Color.Red);
        range.Style.Font.Size = 20;
    }
    
    ws.Cells[row,1].Value = title;
    row++;
    
    // Second row has the days of the week.
    AddDaysOfWeekHeader(ws, ref row);
    AddCalendar(ws, month, holidays, services, ref row);
    
    // Set a pagebreak.
    ws.Row(row-1).PageBreak = true;
}

static void AddDaysOfWeekHeader(ExcelWorksheet ws, ref int row)
{
    var daysOfWeek = Enum.GetValues<DayOfWeek>();

    ws.Rows[row].Height = 40;
    
    for (int i = 0; i < daysOfWeek.Length; i++)
    {
        var cell = ws.Cells[row, i + 1];
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 12;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.SetBackground(i == 0 ? Color.LightBlue : Color.LightGray);
        cell.Style.Locked = true;

        cell.Value = daysOfWeek[i].ToString();
    }
    
    row++;
}

static void AddCalendar(ExcelWorksheet ws, DateOnly month, List<Holiday> holidays, List<Service> services, ref int row)
{
    int daysInMonth = month.AddMonths(1).AddDays(-1).Day;
    int column = (int)month.DayOfWeek + 1;
    Debug.Assert(column >= 0);
    
    // Calculate the total number of rows needed for the calendar
    int totalRows = (daysInMonth + column - 1) / 7 + 1;
    int startRow = row;
    
    for (int i = 1; i <= daysInMonth; i++)
    {
        ws.Rows[row].Height = 85;
        
        var day = new DateOnly(month.Year, month.Month, i);
        var holidayName = holidays.FirstOrDefault(h => h.Date == day)?.Name ?? "";
        if (holidayName.Length > 20)
            holidayName = holidayName[..20] + "..";

        var service = services.FirstOrDefault(s => s.Date == day);
        
        var cell = ws.Cells[row, column];
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        cell.Style.Font.Size = 10;
        cell.Style.WrapText = true;
        cell.Style.Indent = 1;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.SetBackground(day.DayOfWeek switch {
                DayOfWeek.Sunday => Color.LightYellow,
                DayOfWeek.Wednesday => Color.LightCyan,
                _ => Color.White
            });

        if (service != null)
        {
            var sb = new StringBuilder(i.ToString()).AppendLine()
                .AppendLine(holidayName)
                .AppendLine($"(sound) {service.FindRole("sound")}")
                .AppendLine($"(stream) {service.FindRole("stream")}")
                .AppendLine($"(slides) {service.FindRole("slides")}");
            cell.Value = sb.ToString();
        }
        else if (!string.IsNullOrEmpty(holidayName))
        {
            cell.Value = $"{i}{Environment.NewLine}{holidayName}";
        }
        else
        {
            cell.Value = i;
        }

        if (++column > 7)
        {
            column = 1;
            row++;
        }
    }

    // Apply border to all cells in the calendar range
    for (int r = startRow; r < startRow + totalRows; r++)
    {
        for (int c = 1; c <= 7; c++)
        {
            var cell = ws.Cells[r, c];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        }
    }

    row++;
}