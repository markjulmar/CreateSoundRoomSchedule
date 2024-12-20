using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

namespace CreateSoundRoomSchedule;

public sealed class TeamMember(string name, string role)
{
    public string Role => role.ToLower();
    public string Name => name;
}

public sealed class Service
{
    public DateOnly Date { get; init; }
    public List<TeamMember>? Team { get; init; }
    public string FindRole(string sound) => Team?.FirstOrDefault(member => member.Role == sound)?.Name ?? "";
}

public sealed class PlanningCenter(string? clientId, string? clientSecret)
{
    private const string BaseUrl = "https://api.planningcenteronline.com/";

    public async IAsyncEnumerable<Service> GetServicesAsync(DateOnly startDate, DateOnly endDate)
    {
        if (string.IsNullOrEmpty(clientId)
            || string.IsNullOrEmpty(clientSecret))
            throw new InvalidOperationException("Missing Client Id or Client Secret.");
        
        int serviceTypeId = await GetWorshipServiceType();
        if (serviceTypeId == 0) 
            yield break;

        var url = $"services/v2/service_types/{serviceTypeId}/plans?offset=0&limit=100&order=-sort_date";
        while (true)
        {
            if (string.IsNullOrEmpty(url)) yield break;
            
            var jsonResponse = await CallAsync(url);
            using var doc = JsonDocument.Parse(jsonResponse);
            
            url = doc.RootElement.GetProperty("links").GetProperty("next").GetString()?[BaseUrl.Length..];            
            var data = doc.RootElement.GetProperty("data");
            if (!data.EnumerateArray().Any())
                yield break;
            
            foreach (var element in data.EnumerateArray().Where(e => e.GetProperty("type").GetString() == "Plan"))
            {
                var id = element.GetProperty("id").GetString();
                if (string.IsNullOrEmpty(id))
                    continue;
                
                var date = DateOnly.FromDateTime(element.GetProperty("attributes").GetProperty("sort_date").GetDateTime());
                if (date > endDate)
                    continue;
                if (date < startDate) 
                    yield break;
                yield return new Service
                {
                    Date = date,
                    Team = await GetTeamAsync(serviceTypeId, id)
                };
            }
        }
    }

    private async Task<List<TeamMember>?> GetTeamAsync(int serviceTypeId, string id)
    {
        var url = $"services/v2/service_types/{serviceTypeId}/plans/{id}/team_members";
        var jsonResponse = await CallAsync(url);
        using var doc = JsonDocument.Parse(jsonResponse);
        var data = doc.RootElement.GetProperty("data");
        if (!data.EnumerateArray().Any())
            return [];

        return data.EnumerateArray().Select(e =>
            new TeamMember(e.GetProperty("attributes").GetProperty("name").GetString() ?? "",
                e.GetProperty("attributes").GetProperty("team_position_name").GetString() ?? ""))
            .ToList();
    }

    private async Task<int> GetWorshipServiceType()
    {
        var jsonResponse = await CallAsync("services/v2/service_types");
        using var doc = JsonDocument.Parse(jsonResponse);
        var data = doc.RootElement.GetProperty("data");

        // Find the item with the name "TFC Worship Service"
        return (from item in data.EnumerateArray() 
            let attributes = item.GetProperty("attributes") 
            where attributes.GetProperty("name").GetString() == "TFC Worship Service" 
            select int.TryParse(item.GetProperty("id").GetString(), out var result) ? result : 0)
        .FirstOrDefault();
    }
    
    private async Task<string> CallAsync(string segment)
    {
        var url = BaseUrl + segment;
        var credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{clientId}:{clientSecret}"));
        using var client = new HttpClient();
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);
        return await client.GetStringAsync(url);
    }
}