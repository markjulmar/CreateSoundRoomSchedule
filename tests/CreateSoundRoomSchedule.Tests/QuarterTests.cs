using System.Globalization;
using Xunit;

namespace CreateSoundRoomSchedule.Tests;

public class QuarterTests
{
    [Fact]
    public void CalculateQuarter_Returns_Q1_For_February_2025()
    {
        var date = new DateOnly(2025, 2, 1);
        var quarter = Program.CalculateQuarter(date);
        Assert.Equal("2025-Q1", quarter);
    }

    [Fact]
    public void ResolveQuarterStart_UsesNextQuarter_WhenNoArgumentsAreProvided()
    {
        var quarterStart = Program.ResolveQuarterStart([], new DateOnly(2025, 2, 15), null!);

        Assert.Equal(new DateOnly(2025, 4, 1), quarterStart);
    }

    [Fact]
    public void ResolveQuarterStart_UsesQuarterContainingProvidedDate()
    {
        var quarterStart = Program.ResolveQuarterStart(["5/20/2025"], new DateOnly(2025, 2, 15), CultureInfo.InvariantCulture);

        Assert.Equal(new DateOnly(2025, 4, 1), quarterStart);
    }

    [Fact]
    public void GetNextQuarterStart_RollsIntoNextYear()
    {
        var quarterStart = Program.GetNextQuarterStart(new DateOnly(2025, 12, 15));

        Assert.Equal(new DateOnly(2026, 1, 1), quarterStart);
    }
}
