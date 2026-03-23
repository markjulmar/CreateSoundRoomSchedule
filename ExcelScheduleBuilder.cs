using System.Diagnostics;
using System.Drawing;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CreateSoundRoomSchedule;

public sealed class ExcelScheduleBuilder
{
    public void Build(string fullPath, string quarter, DateOnly quarterStart, IReadOnlyList<Holiday> holidays, IReadOnlyList<Service> services)
    {
        using var package = new ExcelPackage(new FileInfo(fullPath));
        package.Compatibility.IsWorksheets1Based = true;

        package.Workbook.Properties.Title = $"TFC Sound Room Schedule - {quarter}";
        package.Workbook.Properties.Author = "Sound Room Scheduler";
        package.Workbook.Properties.Comments = $"This is the sound room schedule for A/V at Trinity Fellowship Church for {quarter}.";
        package.Workbook.Properties.Company = "Trinity Fellowship Church";
        package.Workbook.Properties.Created = DateTime.Now;

        for (int i = 0; i < 3; i++)
        {
            var month = new DateOnly(quarterStart.Year, quarterStart.Month + i, 1);
            var worksheet = CreateWorksheet(package, month.ToString("MMM-yy"));
            var row = Constants.FirstRow;
            AddMonth(worksheet, month, holidays, services, ref row);
            worksheet.PrinterSettings.PrintArea = worksheet.Cells[Constants.FirstRow, Constants.FirstColumn, row - 1, Constants.MaxColumns];
        }

        package.Save();
    }

    private static ExcelWorksheet CreateWorksheet(ExcelPackage package, string title)
    {
        var worksheet = package.Workbook.Worksheets.Add(title);
        worksheet.PrinterSettings.TopMargin = Constants.TopBottomMargin;
        worksheet.PrinterSettings.LeftMargin = Constants.LeftRightMargin;
        worksheet.PrinterSettings.BottomMargin = Constants.TopBottomMargin;
        worksheet.PrinterSettings.RightMargin = Constants.LeftRightMargin;
        worksheet.PrinterSettings.HeaderMargin = Constants.HeaderFooterMargin;
        worksheet.PrinterSettings.FooterMargin = Constants.HeaderFooterMargin;

        worksheet.PrinterSettings.FitToWidth = 1;
        worksheet.PrinterSettings.FitToHeight = 0;
        worksheet.PrinterSettings.FitToPage = true;
        worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
        worksheet.PrinterSettings.HorizontalCentered = true;
        worksheet.PrinterSettings.VerticalCentered = false;

        worksheet.Cells.Style.Font.Name = Constants.FontFamily;
        worksheet.Columns[Constants.MaxColumns].PageBreak = true;

        return worksheet;
    }

    private static void AddMonth(ExcelWorksheet worksheet, DateOnly month, IReadOnlyList<Holiday> holidays, IReadOnlyList<Service> services, ref int row)
    {
        var title = $"{month:MMMM yyyy} Trinity Fellowship Church A/V Schedule";

        worksheet.Rows[row].Height = Constants.HeaderRowHeight;

        using (var range = worksheet.Cells[row, Constants.FirstColumn, row, Constants.MaxColumns])
        {
            range.Merge = true;
            range.Style.Font.Bold = true;
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            range.Style.Font.Color.SetColor(Color.Red);
            range.Style.Font.Size = Constants.HeaderFontSize;
        }

        worksheet.Cells[row, Constants.FirstColumn].Value = title;
        row++;

        AddDaysOfWeekHeader(worksheet, ref row);
        AddCalendar(worksheet, month, holidays, services, ref row);
        worksheet.Row(row - 1).PageBreak = true;
    }

    private static void AddDaysOfWeekHeader(ExcelWorksheet worksheet, ref int row)
    {
        var daysOfWeek = Enum.GetValues<DayOfWeek>();
        worksheet.Rows[row].Height = Constants.DaysOfWeekHeaderRowHeight;

        for (int i = 0; i < daysOfWeek.Length; i++)
        {
            var cell = worksheet.Cells[row, i + 1];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = Constants.DaysOfWeekFontSize;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(i == 0 ? Color.LightBlue : Color.LightGray);
            cell.Style.Locked = true;
            cell.Value = daysOfWeek[i].ToString();
        }

        row++;
    }

    private static void AddCalendar(ExcelWorksheet worksheet, DateOnly month, IReadOnlyList<Holiday> holidays, IReadOnlyList<Service> services, ref int row)
    {
        var daysInMonth = month.AddMonths(1).AddDays(-1).Day;
        var column = (int)month.DayOfWeek + 1;
        Debug.Assert(column >= 0);

        var totalRows = (daysInMonth + column - 1 + Constants.MaxColumns - 1) / Constants.MaxColumns;
        var startRow = row;

        for (int dayNumber = 1; dayNumber <= daysInMonth; dayNumber++)
        {
            worksheet.Rows[row].Height = Constants.DayCellHeight;
            worksheet.Columns[column].Width = Constants.MinimumCellWidth;

            var day = new DateOnly(month.Year, month.Month, dayNumber);
            var holidayName = GetHolidayName(holidays, day);
            var service = services.FirstOrDefault(item => item.Date == day);

            var cell = worksheet.Cells[row, column];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            cell.Style.Font.Size = Constants.DayCellsFontSize;
            cell.Style.WrapText = true;
            cell.Style.Indent = 1;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(day.DayOfWeek switch
            {
                DayOfWeek.Sunday => Color.LightYellow,
                DayOfWeek.Wednesday => Color.LightCyan,
                _ => Color.White
            });

            cell.Value = BuildCellValue(dayNumber, holidayName, service);

            if (++column > Constants.MaxColumns)
            {
                column = Constants.FirstColumn;
                row++;
            }
        }

        if (column == Constants.FirstColumn)
            row--;

        for (int currentRow = startRow; currentRow < startRow + totalRows; currentRow++)
        {
            for (int currentColumn = Constants.FirstColumn; currentColumn <= Constants.MaxColumns; currentColumn++)
            {
                worksheet.Cells[currentRow, currentColumn].Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        row++;
    }

    internal static object BuildCellValue(int dayNumber, string holidayName, Service? service)
    {
        if (service != null)
        {
            return new StringBuilder(dayNumber.ToString())
                .AppendLine()
                .AppendLine(holidayName)
                .AppendLine($"{Constants.SoundPrefix}{service.FindRole(Constants.SoundRoleName)}")
                .AppendLine($"{Constants.SlidesPrefix}{service.FindRole(Constants.SlidesRoleName)}")
                .ToString();
        }

        return string.IsNullOrEmpty(holidayName)
            ? dayNumber
            : $"{dayNumber}{Environment.NewLine}{holidayName}";
    }

    private static string GetHolidayName(IReadOnlyList<Holiday> holidays, DateOnly day)
    {
        var holidayName = (holidays.FirstOrDefault(holiday => holiday.Date == day)?.Name ?? string.Empty).Trim();
        if (holidayName.Length > Constants.MaxHolidayLength)
            return holidayName[..Constants.MaxHolidayLength] + "..";

        return holidayName;
    }
}
