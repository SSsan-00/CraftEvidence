using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Core.Services;

/// <summary>Converts captured cells and Shapes into the row spans consumed by placement planning.</summary>
public sealed class ContentOccupancyAnalyzer
{
  public IReadOnlyList<ContentSpan> Analyze(SheetSnapshot snapshot, EvidenceCaseLayout layout)
  {
    ArgumentNullException.ThrowIfNull(snapshot);
    ArgumentNullException.ThrowIfNull(layout);

    var spans = snapshot.Cells
      .Where(cell =>
        cell.IsOccupied &&
        cell.Row >= layout.StartRow &&
        cell.Row <= layout.EndRow)
      .Select(cell => new ContentSpan(
        ResolveSide(cell.Column, cell.Column, layout),
        cell.Row,
        cell.Row,
        cell.IntersectsMerge ? ContentKind.Merge : ContentKind.Cell))
      .Where(span => span.Side is not null)
      .Concat(snapshot.Shapes
        .Where(shape =>
          shape.EndRow >= layout.StartRow &&
          shape.StartRow <= layout.EndRow &&
          shape.EndColumn >= layout.NewRegion.FirstColumn &&
          shape.StartColumn <= layout.OldRegion.LastColumn)
        .Select(shape => new ContentSpan(
          ResolveSide(shape.StartColumn, shape.EndColumn, layout),
          Math.Max(shape.StartRow, layout.StartRow),
          Math.Min(shape.EndRow, layout.EndRow),
          shape.IsManagedImage ? ContentKind.ManagedImage : ContentKind.Shape)))
      .ToArray();

    return MergeSpans(spans);
  }

  private static EvidenceSide? ResolveSide(
    int firstColumn,
    int lastColumn,
    EvidenceCaseLayout layout)
  {
    if (firstColumn >= layout.NewRegion.FirstColumn && lastColumn <= layout.NewRegion.LastColumn)
    {
      return EvidenceSide.New;
    }

    if (firstColumn >= layout.OldRegion.FirstColumn && lastColumn <= layout.OldRegion.LastColumn)
    {
      return EvidenceSide.Old;
    }

    return null;
  }

  private static IReadOnlyList<ContentSpan> MergeSpans(IEnumerable<ContentSpan> source)
  {
    var merged = new List<ContentSpan>();
    foreach (var group in source
      .OrderBy(span => span.Side)
      .ThenBy(span => span.Kind)
      .ThenBy(span => span.StartRow)
      .ThenBy(span => span.EndRow)
      .GroupBy(span => new { span.Side, span.Kind }))
    {
      ContentSpan? current = null;
      foreach (var span in group)
      {
        if (current is null || (long)span.StartRow > current.EndRow + 1L)
        {
          if (current is not null)
          {
            merged.Add(current);
          }

          current = span;
          continue;
        }

        current = current with { EndRow = Math.Max(current.EndRow, span.EndRow) };
      }

      if (current is not null)
      {
        merged.Add(current);
      }
    }

    return merged
      .OrderBy(span => span.StartRow)
      .ThenBy(span => span.EndRow)
      .ToArray();
  }
}
