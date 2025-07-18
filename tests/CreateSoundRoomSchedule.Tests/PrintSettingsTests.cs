using Xunit;

namespace CreateSoundRoomSchedule.Tests;

public class PrintSettingsTests
{
    [Fact]
    public void Constants_TopBottomMargin_IsSetForMaximizedPrinting()
    {
        // The top/bottom margin should be reduced from 0.75 to 0.5 for maximized printing
        Assert.Equal(0.5m, Constants.TopBottomMargin);
    }

    [Fact]
    public void Constants_LeftRightMargin_IsSetForLandscapePrinting()
    {
        // Left/right margins should remain small for landscape printing
        Assert.Equal(0.25m, Constants.LeftRightMargin);
    }
}