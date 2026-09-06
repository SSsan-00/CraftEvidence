using System.Reflection;
using System.Text.Json;
using CraftEvidence.Core.Models;

namespace CraftEvidence.Tests;

internal static class FixtureLoader
{
  private static readonly JsonSerializerOptions Options = new()
  {
    PropertyNameCaseInsensitive = true,
  };

  public static SheetLayoutSignals LoadLayout(string fileName)
  {
    var assembly = Assembly.GetExecutingAssembly();
    var resourceName = $"CraftEvidence.Tests.Fixtures.{fileName}";
    using var stream = assembly.GetManifestResourceStream(resourceName) ??
      throw new InvalidOperationException($"Missing embedded fixture: {resourceName}");

    return JsonSerializer.Deserialize<SheetLayoutSignals>(stream, Options) ??
      throw new InvalidOperationException($"Invalid fixture: {resourceName}");
  }
}
