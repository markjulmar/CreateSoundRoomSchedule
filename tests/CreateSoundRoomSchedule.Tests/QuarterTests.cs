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
}
