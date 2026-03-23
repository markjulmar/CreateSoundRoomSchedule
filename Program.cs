using System.Globalization;
using CreateSoundRoomSchedule;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

var configuration = new ConfigurationBuilder()
    .AddUserSecrets<Program>() // Link to user secrets
    .Build();

var clientId = configuration["PlanningCenter:clientId"];
var clientSecret = configuration["PlanningCenter:clientSecret"];

using var planningCenter = new PlanningCenter(clientId, clientSecret);

// Non commercial use.
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var date = Program.ResolveQuarterStart(args, DateOnly.FromDateTime(DateTime.Today), CultureInfo.InvariantCulture);
var quarter = Program.CalculateQuarter(date);

if (args.Length > 1)
{
    Console.Error.WriteLine("Expected zero or one date argument in M/d/yyyy format. Using next quarter.");
}

Console.WriteLine($"CreateSoundRoomSchedule for {quarter}");
Console.WriteLine();
Console.WriteLine("Retrieving data from Planning Center Online...");

var services = await planningCenter.GetServicesAsync(date, date.AddMonths(3));
Console.WriteLine($"... found {services.Count} services.");
Console.WriteLine();
Console.WriteLine("Creating Excel file...");

var filename = $"TFC_SoundRoom_Schedule_{quarter}.xlsx";
var fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), filename);
if (File.Exists(fullPath))
    File.Delete(fullPath);

// Get the holidays
var holidays = (await Holiday.GetAllAsync(date.Year)).ToList();
var builder = new ExcelScheduleBuilder();
builder.Build(fullPath, quarter, date, holidays, services);

Console.WriteLine($"File saved to {fullPath}");

return;
