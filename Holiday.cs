using System.Text.Json;
using System.Text.Json.Serialization;

namespace CreateSoundRoomSchedule;

public sealed class Holiday
{
    [JsonPropertyName("date")] public DateOnly Date { get; set; }
    [JsonPropertyName("name")] public string Name { get; set; } = "";

    public static async Task<IEnumerable<Holiday>> GetAllAsync(int year)
    {
        var url = $"https://date.nager.at/api/v3/PublicHolidays/{year}/US";
        using var client = new HttpClient();
        var response = await client.GetStringAsync(url);
        return JsonSerializer.Deserialize<IEnumerable<Holiday>>(response) ?? Array.Empty<Holiday>();
    }
}