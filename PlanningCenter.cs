using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace CreateSoundRoomSchedule;

public sealed class TeamMember(string name, string role)
{
    public string Role => role.ToLowerInvariant();
    public string Name => name;
}

public sealed class Service(string id, DateOnly date)
{
    public string Id => id;
    public DateOnly Date => date;
    public List<TeamMember>? Team { get; set; }
    public string FindRole(string roleName) => Team?.FirstOrDefault(member => member.Role == roleName)?.Name ?? "unassigned";
}

public sealed class PlanningCenter : IDisposable
{
    private const string LookForServiceTypeName = "TFC Worship Service";
    private const string BaseUrl = "https://api.planningcenteronline.com/";
    private readonly string? clientId;
    private readonly string? clientSecret;
    private readonly HttpClient httpClient;
    private readonly bool ownsHttpClient;

    public PlanningCenter(string? clientId, string? clientSecret, HttpClient? httpClient = null)
    {
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.httpClient = httpClient ?? CreateHttpClient();
        ownsHttpClient = httpClient is null;
    }

    public async Task<List<Service>> GetServicesAsync(DateOnly startDate, DateOnly endDate)
    {
        if (string.IsNullOrEmpty(clientId)
            || string.IsNullOrEmpty(clientSecret))
            throw new InvalidOperationException("Missing Client Id or Client Secret.");

        int serviceTypeId = await GetWorshipServiceType(LookForServiceTypeName);
        if (serviceTypeId == 0)
        {
            Console.Error.WriteLine($"Service type '{LookForServiceTypeName}' not found.");
            return [];
        }

        var services = new List<Service>();
        var url = $"services/v2/service_types/{serviceTypeId}/plans?offset=0&limit=100&order=-sort_date";
        bool done = false;
        while (!done)
        {
            if (string.IsNullOrEmpty(url))
                break;

            var jsonResponse = await CallAsync(url);
            using var doc = JsonDocument.Parse(jsonResponse);

            url = GetNextPageSegment(doc.RootElement);
            if (!doc.RootElement.TryGetProperty("data", out var data)
                || data.ValueKind != JsonValueKind.Array)
            {
                break;
            }

            if (!data.EnumerateArray().Any())
                break;

            foreach (var element in data.EnumerateArray())
            {
                if (!IsPlan(element)
                    || !TryGetPlan(element, out var service))
                {
                    continue;
                }

                if (service.Date >= startDate && service.Date < endDate)
                {
                    services.Add(service);
                }
            }
        }

        if (services.Count == 0)
        {
            Console.Error.WriteLine($"No services found for service type '{LookForServiceTypeName}' between {startDate} and {endDate}.");
            return [];
        }

        Console.WriteLine($"... found {services.Count} services of type '{LookForServiceTypeName}' between {startDate} and {endDate}.");

        var tasks = services.Select(async service =>
        {
            service.Team = await GetTeamAsync(serviceTypeId, service.Id);
            return service;
        });

        return (await Task.WhenAll(tasks)).ToList();
    }

    private async Task<List<TeamMember>?> GetTeamAsync(int serviceTypeId, string id)
    {
        var url = $"services/v2/service_types/{serviceTypeId}/plans/{id}/team_members";
        var jsonResponse = await CallAsync(url);
        using var doc = JsonDocument.Parse(jsonResponse);
        if (!doc.RootElement.TryGetProperty("data", out var data)
            || data.ValueKind != JsonValueKind.Array
            || !data.EnumerateArray().Any())
        {
            return [];
        }

        return data.EnumerateArray()
            .Select(TryGetTeamMember)
            .OfType<TeamMember>()
            .ToList();
    }

    private async Task<int> GetWorshipServiceType(string serviceTypeName)
    {
        var jsonResponse = await CallAsync("services/v2/service_types");
        using var doc = JsonDocument.Parse(jsonResponse);
        if (!doc.RootElement.TryGetProperty("data", out var data)
            || data.ValueKind != JsonValueKind.Array)
        {
            return 0;
        }

        // Find the item with the name "TFC Worship Service"
        return (from item in data.EnumerateArray()
            let attributes = TryGetObject(item, "attributes")
            where attributes.HasValue && TryGetString(attributes.Value, "name") == serviceTypeName
            select int.TryParse(TryGetString(item, "id"), out var result) ? result : 0)
        .FirstOrDefault();
    }

    private async Task<string> CallAsync(string segment)
    {
        var url = BaseUrl + segment;
        var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{clientId}:{clientSecret}"));

        try
        {
            using var request = new HttpRequestMessage(HttpMethod.Get, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Basic", credentials);
            using var response = await httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
        catch (HttpRequestException ex)
        {
            throw new InvalidOperationException($"Error calling Planning Center API endpoint '{segment}'.", ex);
        }
    }

    internal static string? GetNextPageSegment(JsonElement root)
    {
        if (!root.TryGetProperty("links", out var links)
            || links.ValueKind != JsonValueKind.Object
            || !links.TryGetProperty("next", out var next)
            || next.ValueKind == JsonValueKind.Null)
        {
            return null;
        }

        var nextUrl = next.GetString();
        if (string.IsNullOrWhiteSpace(nextUrl))
            return null;

        return nextUrl.StartsWith(BaseUrl, StringComparison.OrdinalIgnoreCase)
            ? nextUrl[BaseUrl.Length..]
            : nextUrl;
    }

    internal static bool TryGetPlan(JsonElement element, out Service service)
    {
        service = default!;

        var id = TryGetString(element, "id");
        var attributes = TryGetObject(element, "attributes");
        var sortDate = attributes.HasValue ? TryGetDateTime(attributes.Value, "sort_date") : null;

        if (string.IsNullOrWhiteSpace(id) || sortDate is null)
            return false;

        service = new Service(id, DateOnly.FromDateTime(sortDate.Value));
        return true;
    }

    private static TeamMember? TryGetTeamMember(JsonElement element)
    {
        var attributes = TryGetObject(element, "attributes");
        if (!attributes.HasValue)
            return null;

        return new TeamMember(
            TryGetString(attributes.Value, "name") ?? string.Empty,
            TryGetString(attributes.Value, "team_position_name") ?? string.Empty);
    }

    private static bool IsPlan(JsonElement element) => TryGetString(element, "type") == "Plan";

    private static JsonElement? TryGetObject(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.Object
            ? property
            : null;
    }

    private static string? TryGetString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var property) || property.ValueKind == JsonValueKind.Null)
            return null;

        return property.GetString();
    }

    private static DateTime? TryGetDateTime(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.TryGetDateTime(out var value)
            ? value
            : null;
    }

    private static HttpClient CreateHttpClient()
    {
        return new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(30)
        };
    }

    public void Dispose()
    {
        if (ownsHttpClient)
            httpClient.Dispose();
    }
}
