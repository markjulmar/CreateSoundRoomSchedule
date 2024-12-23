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
        var results = JsonSerializer.Deserialize<List<Holiday>>(response) ?? [];

        // Add a few additional holidays.
        var goodFriday = results.FirstOrDefault(h => h.Name == "Good Friday");
        if (goodFriday != null)
        {
            var easter = goodFriday.Date.AddDays(2);
            results.Add(new Holiday { Date = goodFriday.Date.AddDays(-1), Name = "Maundy Thursday" });
            results.Add(new Holiday { Date = easter, Name = "Easter Sunday" });

            var ashWednesday = easter.AddDays((-1 * 7 * 6) - 4);
            results.Add(new Holiday { Date = ashWednesday, Name = "Ash Wednesday" });
        }

        return results;
    }
}