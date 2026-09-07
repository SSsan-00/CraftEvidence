using System.Runtime.InteropServices;

namespace EvidenceCrafter.Excel;

internal static class ComRelease
{
  public static void Release(object? value)
  {
    if (value is not null && Marshal.IsComObject(value))
    {
      _ = Marshal.ReleaseComObject(value);
    }
  }
}
