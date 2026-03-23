using System.Globalization;

internal partial class Program
{
    internal static string CalculateQuarter(DateOnly date)
    {
        date = GetQuarterStart(date);
        return $"{date:yyyy}-Q{(date.Month - 1) / 3 + 1}";
    }

    internal static DateOnly ResolveQuarterStart(string[] args, DateOnly today, IFormatProvider formatProvider)
    {
        if (args.Length == 0)
            return GetNextQuarterStart(today);

        if (args.Length == 1
            && DateOnly.TryParseExact(args[0], "M/d/yyyy", formatProvider, DateTimeStyles.None, out var parsedDate))
        {
            return GetQuarterStart(parsedDate);
        }

        Console.Error.WriteLine("Invalid date format. Expected M/d/yyyy. Using next quarter.");
        return GetNextQuarterStart(today);
    }

    internal static DateOnly GetQuarterStart(DateOnly date)
    {
        var startOfQuarterMonth = date.Month <= 3 ? 1
            : date.Month <= 6 ? 4
            : date.Month <= 9 ? 7
            : 10;
        return new DateOnly(date.Year, startOfQuarterMonth, 1);
    }

    internal static DateOnly GetNextQuarterStart(DateOnly date)
    {
        var currentQuarterStart = GetQuarterStart(date);
        return currentQuarterStart.AddMonths(3);
    }
}
