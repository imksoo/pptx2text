using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

foreach (string filename in args)
{
  if (File.Exists(filename))
  {
    ExtractPptxFileText(filename);
  }
  else
  {
    Console.Error.WriteLine("File not exists: {0}", filename);
  }
}

void ExtractPptxFileText(string filename)
{
  using (PresentationDocument? presentationDocument = PresentationDocument.Open(filename, true))
  {
    PresentationPart? presentationPart = presentationDocument.PresentationPart;
    if (presentationPart?.Presentation?.SlideIdList != null)
    {
      foreach (SlideId slideId in presentationPart.Presentation.SlideIdList.Cast<SlideId>())
      {
        string? relationshipId = slideId?.RelationshipId;
        if (!string.IsNullOrEmpty(relationshipId))
        {
          SlidePart slidePart = (SlidePart)presentationPart.GetPartById(relationshipId);
          ExtractSlideText(slidePart);
        }
      }
    }
  }
}

void ExtractSlideText(SlidePart slidePart)
{
  var paragraphTextList = new List<string>();

  if (slidePart?.Slide == null) return;

  foreach (Drawing.Paragraph paragraph in slidePart.Slide.Descendants<Drawing.Paragraph>())
  {
    foreach (Drawing.Run run in paragraph.Elements<Drawing.Run>())
    {
      if (run.Text != null)
      {
        string innterText = run.Text.InnerText;
        if (!string.IsNullOrWhiteSpace(innterText))
        {
          paragraphTextList.Add(normalizeText(innterText));
        }
      }
    }
  }

  var paragraphText = normalizeText(string.Join(" ", paragraphTextList));
  Console.WriteLine(paragraphText);
}

string normalizeText(string text)
{
  return Regex.Replace(text, "\\s+", " ").Trim();
}