using CraftEvidence.App;

namespace CraftEvidence.Tests;

[TestClass]
public sealed class ImageEditDocumentTests
{
  [TestMethod]
  public void CropUndoRedo_RestoresEachRasterState()
  {
    using var source = CreateQuadrantImage();
    using var document = new ImageEditDocument(source);

    Assert.IsTrue(document.Crop(new Rectangle(4, 0, 4, 4)));
    Assert.AreEqual(4, document.Width);
    Assert.AreEqual(Color.Green.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());

    Assert.IsTrue(document.Undo());
    Assert.AreEqual(8, document.Width);
    Assert.AreEqual(Color.Red.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());

    Assert.IsTrue(document.Redo());
    Assert.AreEqual(4, document.Width);
    Assert.AreEqual(Color.Green.ToArgb(), document.GetImageCopy().GetPixel(1, 1).ToArgb());
  }

  [TestMethod]
  public void ResetThenUndo_RestoresEditedImage()
  {
    using var source = new Bitmap(20, 20);
    using (var graphics = Graphics.FromImage(source))
    {
      graphics.Clear(Color.White);
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.DrawRedRectangle(new Rectangle(2, 2, 12, 12)));
    using var edited = document.GetImageCopy();
    Assert.IsTrue(document.HasChanges);

    Assert.IsTrue(document.Reset());
    Assert.IsFalse(document.HasChanges);
    using var reset = document.GetImageCopy();
    Assert.AreEqual(Color.White.ToArgb(), reset.GetPixel(2, 2).ToArgb());

    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.HasChanges);
    using var restored = document.GetImageCopy();
    Assert.AreEqual(edited.GetPixel(2, 2).ToArgb(), restored.GetPixel(2, 2).ToArgb());
  }

  [TestMethod]
  public void NewEditAfterUndo_DiscardsRedoState()
  {
    using var source = new Bitmap(30, 30);
    using var document = new ImageEditDocument(source);

    Assert.IsTrue(document.DrawArrow(new Point(1, 1), new Point(20, 20)));
    Assert.IsTrue(document.Undo());
    Assert.IsTrue(document.CanRedo);

    Assert.IsTrue(document.DrawText("test", new Point(2, 2)));
    Assert.IsFalse(document.CanRedo);
  }

  [TestMethod]
  public void Mosaic_ProducesUniformBlocksWithinSelection()
  {
    using var source = new Bitmap(8, 8);
    for (var y = 0; y < source.Height; y++)
    {
      for (var x = 0; x < source.Width; x++)
      {
        source.SetPixel(x, y, Color.FromArgb(255, x * 20, y * 20, (x + y) * 10));
      }
    }

    using var document = new ImageEditDocument(source);
    Assert.IsTrue(document.Mosaic(new Rectangle(0, 0, 8, 8), blockSize: 4));
    using var result = document.GetImageCopy();

    Assert.AreEqual(result.GetPixel(0, 0).ToArgb(), result.GetPixel(3, 3).ToArgb());
    Assert.AreEqual(result.GetPixel(4, 4).ToArgb(), result.GetPixel(7, 7).ToArgb());
  }

  [TestMethod]
  public void InvalidSelections_DoNotCreateHistory()
  {
    using var source = new Bitmap(10, 10);
    using var document = new ImageEditDocument(source);

    Assert.IsFalse(document.DrawArrow(new Point(2, 2), new Point(2, 2)));
    Assert.IsFalse(document.Crop(new Rectangle(1, 1, 1, 1)));
    Assert.IsFalse(document.DrawText("   ", Point.Empty));
    Assert.IsFalse(document.CanUndo);
  }

  private static Bitmap CreateQuadrantImage()
  {
    var image = new Bitmap(8, 8);
    using var graphics = Graphics.FromImage(image);
    graphics.Clear(Color.White);
    graphics.FillRectangle(Brushes.Red, 0, 0, 4, 4);
    graphics.FillRectangle(Brushes.Green, 4, 0, 4, 4);
    graphics.FillRectangle(Brushes.Blue, 0, 4, 4, 4);
    graphics.FillRectangle(Brushes.Yellow, 4, 4, 4, 4);
    return image;
  }
}
