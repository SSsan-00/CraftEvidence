using System.Text.Json;

namespace EvidenceCrafter.App;

internal sealed class AppSettingsStore
{
  private static readonly JsonSerializerOptions JsonOptions = new()
  {
    WriteIndented = true,
  };

  private readonly string settingsPath;

  internal AppSettingsStore(string? settingsPath = null)
  {
    this.settingsPath = settingsPath ?? Path.Combine(
      Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
      "EvidenceCrafter",
      "settings.json");
  }

  internal string SettingsPath => settingsPath;

  internal EvidenceCrafterSettings Load()
  {
    try
    {
      if (!File.Exists(settingsPath))
      {
        return new EvidenceCrafterSettings();
      }

      using var stream = File.OpenRead(settingsPath);
      return (JsonSerializer.Deserialize<EvidenceCrafterSettings>(stream, JsonOptions) ??
        new EvidenceCrafterSettings()).Normalize();
    }
    catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException)
    {
      return new EvidenceCrafterSettings();
    }
  }

  internal void Save(EvidenceCrafterSettings settings)
  {
    ArgumentNullException.ThrowIfNull(settings);

    var directory = Path.GetDirectoryName(settingsPath) ??
      throw new InvalidOperationException("The settings path must include a directory.");
    Directory.CreateDirectory(directory);

    var temporaryPath = settingsPath + ".tmp";
    try
    {
      using (var stream = new FileStream(temporaryPath, FileMode.Create, FileAccess.Write, FileShare.None))
      {
        JsonSerializer.Serialize(stream, settings.Normalize(), JsonOptions);
        stream.Flush(flushToDisk: true);
      }

      File.Move(temporaryPath, settingsPath, overwrite: true);
    }
    finally
    {
      if (File.Exists(temporaryPath))
      {
        File.Delete(temporaryPath);
      }
    }
  }
}
