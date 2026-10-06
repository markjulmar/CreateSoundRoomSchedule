using System.Net;
using System.Text;
using System.Text.Json;
using Xunit;

namespace CreateSoundRoomSchedule.Tests;

public class PlanningCenterTests
{
    [Fact]
    public async Task GetServicesAsync_SelectsReplacements_WhenOriginalAssignmentsAreDeclined()
    {
        using var httpClient = new HttpClient(new ReplacedAssignmentsHandler());
        using var planningCenter = new PlanningCenter("test-client", "test-secret", httpClient);

        var services = await planningCenter.GetServicesAsync(
            new DateOnly(2026, 10, 1), new DateOnly(2027, 1, 1));

        var service = Assert.Single(services);
        Assert.Equal("Annette Bergsagel", service.FindRole("slides"));
        Assert.Equal("Carissa Bergsagel", service.FindRole("sound"));
        Assert.Equal(2, service.Team!.Count);
    }

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

    private sealed class ReplacedAssignmentsHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var uri = request.RequestUri!;
            string response;
            switch (uri.AbsolutePath)
            {
                case "/services/v2/service_types":
                    response = """
                        {"data":[{"id":"12","attributes":{"name":"TFC Worship Service"}}]}
                        """;
                    break;
                case "/services/v2/service_types/12/plans":
                    response = """
                        {"data":[{"type":"Plan","id":"123","attributes":{"sort_date":"2026-10-11T10:00:00Z"}}]}
                        """;
                    break;
                case "/services/v2/service_types/12/plans/123/team_members":
                    // Model the API's not_declined filter; unfiltered results put the originals first.
                    var assignments = new[]
                    {
                        new { name = "Julie Smith", team_position_name = "Slides", status = "D" },
                        new { name = "Mark Smith", team_position_name = "Sound", status = "D" },
                        new { name = "Annette Bergsagel", team_position_name = "Slides", status = "C" },
                        new { name = "Carissa Bergsagel", team_position_name = "Sound", status = "U" }
                    }.AsEnumerable();
                    if (uri.Query == "?filter=not_declined")
                        assignments = assignments.Where(assignment => assignment.status != "D");
                    response = JsonSerializer.Serialize(new
                    {
                        data = assignments.Select(attributes => new { type = "PlanPerson", attributes })
                    });
                    break;
                default:
                    throw new InvalidOperationException($"Unexpected API request: {uri}");
            }

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(response, Encoding.UTF8, "application/json")
            });
        }
    }
}
