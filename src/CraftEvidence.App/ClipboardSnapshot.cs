using System.Collections.Specialized;
using System.Runtime.InteropServices;

namespace CraftEvidence.App;

internal sealed class ClipboardSnapshot : IDisposable
{
  private readonly DataObject data = new();
  private readonly List<IDisposable> owned = [];

  private ClipboardSnapshot()
  {
  }

  internal static ClipboardSnapshot Capture()
  {
    var snapshot = new ClipboardSnapshot();
    var source = Clipboard.GetDataObject();
    if (source is null)
    {
      return snapshot;
    }

    foreach (var format in source.GetFormats(autoConvert: false))
    {
      var value = source.GetData(format, autoConvert: false);
      switch (value)
      {
        case Image image:
          var copy = new Bitmap(image);
          snapshot.owned.Add(copy);
          snapshot.data.SetData(format, false, copy);
          break;
        case MemoryStream stream:
          var streamCopy = new MemoryStream(stream.ToArray(), writable: false);
          snapshot.owned.Add(streamCopy);
          snapshot.data.SetData(format, false, streamCopy);
          break;
        case byte[] bytes:
          snapshot.data.SetData(format, false, bytes.ToArray());
          break;
        case StringCollection files:
          snapshot.data.SetData(format, false, files.Cast<string>().ToArray());
          break;
        case string text:
          snapshot.data.SetData(format, false, text);
          break;
      }
    }

    return snapshot;
  }

  internal bool Restore(uint expectedSequence)
  {
    if (NativeClipboard.GetClipboardSequenceNumber() != expectedSequence)
    {
      return false;
    }

    try
    {
      Clipboard.SetDataObject(data, copy: true);
      return true;
    }
    catch (ExternalException)
    {
      // Clipboard restoration is best effort; the Excel mutation remains safe.
      return false;
    }
  }

  public void Dispose()
  {
    foreach (var item in owned)
    {
      item.Dispose();
    }
  }
}
