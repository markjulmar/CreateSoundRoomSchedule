internal partial class Program
{
    internal static string CalculateQuarter(DateOnly date)
    {
        var startOfQuarterMonth = date.Month <= 3 ? 1
            : date.Month <= 6 ? 4
            : date.Month <= 9 ? 7
            : 10;
        date = new DateOnly(date.Year, startOfQuarterMonth, 1);
        return $"{date:yyyy}-Q{(date.Month - 1) / 3 + 1}";
    }
}
