using Xunit;

namespace CreateSoundRoomSchedule.Tests;

public class ExcelScheduleBuilderTests
{
    [Fact]
    public void BuildCellValue_ReturnsDayNumber_WhenNoHolidayOrServiceExists()
    {
        var value = ExcelScheduleBuilder.BuildCellValue(7, string.Empty, null);

        Assert.Equal(7, value);
    }

    [Fact]
    public void BuildCellValue_IncludesAssignedRoles_WhenServiceExists()
    {
        var service = new Service("svc-1", new DateOnly(2025, 2, 9))
        {
            Team =
            [
                new TeamMember("Alice", Constants.SoundRoleName),
                new TeamMember("Bob", Constants.SlidesRoleName)
            ]
        };

        var value = ExcelScheduleBuilder.BuildCellValue(9, "Holiday", service);

        var text = Assert.IsType<string>(value);
        Assert.Contains("(sound) Alice", text);
        Assert.Contains("(slides) Bob", text);
        Assert.Contains("Holiday", text);
    }
}
