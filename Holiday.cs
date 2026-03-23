using System.Text.Json;
using System.Text.Json.Serialization;

namespace CreateSoundRoomSchedule;

public sealed class Holiday
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromSeconds(15)
    };

    [JsonPropertyName("date")] public DateOnly Date { get; set; }
    [JsonPropertyName("name")] public string Name { get; set; } = "";

    public static async Task<IEnumerable<Holiday>> GetAllAsync(int year)
    {
        var url = $"https://date.nager.at/api/v3/PublicHolidays/{year}/US";
        List<Holiday> results;

        try
        {
            var response = await HttpClient.GetStringAsync(url);
            results = JsonSerializer.Deserialize<List<Holiday>>(response) ?? [];
        }
        catch (HttpRequestException)
        {
            Console.Error.WriteLine("Unable to retrieve holidays from the public holiday API. Continuing without them.");
            results = [];
        }

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
