using System.Text.Json;
using Xunit;

namespace CreateSoundRoomSchedule.Tests;

public class PlanningCenterTests
{
    [Fact]
    public void GetNextPageSegment_ReturnsNull_WhenLinksPropertyIsMissing()
    {
        using var doc = JsonDocument.Parse("""
        {
          "data": []
        }
        """);

        var nextPage = PlanningCenter.GetNextPageSegment(doc.RootElement);

        Assert.Null(nextPage);
    }

    [Fact]
    public void GetNextPageSegment_ReturnsRelativeSegment_WhenNextLinkExists()
    {
        using var doc = JsonDocument.Parse("""
        {
          "links": {
            "next": "https://api.planningcenteronline.com/services/v2/service_types/12/plans?offset=100"
          }
        }
        """);

        var nextPage = PlanningCenter.GetNextPageSegment(doc.RootElement);

        Assert.Equal("services/v2/service_types/12/plans?offset=100", nextPage);
    }

    [Fact]
    public void TryGetPlan_ReturnsFalse_WhenSortDateIsMissing()
    {
        using var doc = JsonDocument.Parse("""
        {
          "type": "Plan",
          "id": "123",
          "attributes": {
          }
        }
        """);

        var success = PlanningCenter.TryGetPlan(doc.RootElement, out _);

        Assert.False(success);
    }

    [Fact]
    public void TryGetPlan_ReturnsService_WhenRequiredFieldsExist()
    {
        using var doc = JsonDocument.Parse("""
        {
          "type": "Plan",
          "id": "123",
          "attributes": {
            "sort_date": "2025-02-09T00:00:00Z"
          }
        }
        """);

        var success = PlanningCenter.TryGetPlan(doc.RootElement, out var service);

        Assert.True(success);
        Assert.Equal("123", service.Id);
        Assert.Equal(new DateOnly(2025, 2, 9), service.Date);
    }
}
