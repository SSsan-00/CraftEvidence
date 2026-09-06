using System.Text.Json;
using CraftEvidence.App;

namespace CraftEvidence.Tests;

[TestClass]
public sealed class AppInfrastructureTests
{
  [TestMethod]
  public void SettingsStore_RoundTripsNormalizedValuesWithoutRuntimeSelection()
  {
    using var directory = new TemporaryDirectory();
    var path = Path.Combine(directory.Path, "settings.json");
    var store = new AppSettingsStore(path);

    store.Save(new CraftEvidenceSettings
    {
      HorizontalMarginPoints = 200,
      DiagnosticLoggingEnabled = false,
      GlobalShortcutEnabled = false,
    });

    var loaded = store.Load();
    Assert.AreEqual(72, loaded.HorizontalMarginPoints);
    Assert.IsFalse(loaded.DiagnosticLoggingEnabled);
    Assert.IsFalse(loaded.GlobalShortcutEnabled);
    var savedJson = File.ReadAllText(path);
    Assert.IsFalse(savedJson.Contains("Workbook", StringComparison.Ordinal));
    Assert.IsFalse(savedJson.Contains("Side", StringComparison.Ordinal));
  }

  [TestMethod]
  public void SettingsStore_InvalidJson_ReturnsDefaults()
  {
    using var directory = new TemporaryDirectory();
    var path = Path.Combine(directory.Path, "settings.json");
    File.WriteAllText(path, "not-json");

    var loaded = new AppSettingsStore(path).Load();

    Assert.AreEqual(CraftEvidenceSettings.DefaultHorizontalMarginPoints, loaded.HorizontalMarginPoints);
  }

  [TestMethod]
  public void DiagnosticLog_WritesOnlyStructuralFields()
  {
    using var directory = new TemporaryDirectory();
    var log = new DiagnosticLog(directory.Path);
    var exception = new InvalidOperationException("secret cell text and image path");

    log.Write(
      DiagnosticEventKind.MutationResult,
      DiagnosticOutcome.Failed,
      processId: 42,
      itemCount: 3,
      duration: TimeSpan.FromMilliseconds(12),
      exception);

    var text = File.ReadAllText(Path.Combine(directory.Path, "diagnostic.jsonl"));
    using var document = JsonDocument.Parse(text);
    Assert.AreEqual("MutationResult", document.RootElement.GetProperty("event").GetString());
    Assert.AreEqual(42, document.RootElement.GetProperty("processId").GetInt32());
    StringAssert.Contains(text, nameof(InvalidOperationException));
    Assert.IsFalse(text.Contains(exception.Message, StringComparison.Ordinal));
  }

  [TestMethod]
  public void GlobalShortcut_InvalidRegistration_DoesNotCallWindowsOrThrow()
  {
    var succeeded = GlobalShortcutRegistration.TryRegister(
      windowHandle: 0,
      id: 1,
      ShortcutModifiers.Control,
      Keys.V,
      out var registration,
      out var error);

    Assert.IsFalse(succeeded);
    Assert.IsNull(registration);
    Assert.IsNotNull(error);
  }

  private sealed class TemporaryDirectory : IDisposable
  {
    internal TemporaryDirectory()
    {
      Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"CraftEvidenceTests-{Guid.NewGuid():N}");
      Directory.CreateDirectory(Path);
    }

    internal string Path { get; }

    public void Dispose() => Directory.Delete(Path, recursive: true);
  }
}
