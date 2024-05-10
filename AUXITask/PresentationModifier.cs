using DocumentFormat.OpenXml.Office2013.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

public class PresentationModifier
{
    private const string InputFileName = "InputSlide.pptx";
    private const string OutputFileName = "OutputSlide.pptx";
    private const string OutputText = "Output Slide";
    private const string FontTypeface = "Beirut";
    
    public void ModifyPresentation()
    {

        // Relative path to your file
        string relativePath = "../../../";
        var inputFilePath = Path.Combine(relativePath, InputFileName);
        var outputFilePath = Path.Combine(relativePath, OutputFileName); // Define the output file path

        // Create a copy of the original file to modify
        File.Copy(inputFilePath, outputFilePath, true);


        using (PresentationDocument presentationDocument = PresentationDocument.Open(outputFilePath, true))
        {
            SlidePart slidePart = presentationDocument.PresentationPart.SlideParts.FirstOrDefault();

            if (slidePart != null)
            {
                ModifyTitle(slidePart);
                ModifyShapeTexts(slidePart);
                AdjustShapes(slidePart);

                slidePart.Slide.Save();
            }
        }
    }

    private void ModifyTitle(SlidePart slidePart)
    {
        Shape titleShape = slidePart.Slide.Descendants<Shape>().FirstOrDefault();
        if (titleShape != null)
        {
            var paragraph = titleShape.TextBody.Descendants<A.Paragraph>().FirstOrDefault();
            if (paragraph != null)
            {
                ModifyTextAndAlignment(paragraph, OutputText, A.TextAlignmentTypeValues.Center);
                SetFontStyle(paragraph, FontTypeface);
            }
        }
    }

    private void ModifyShapeTexts(SlidePart slidePart)
    {
        var shapes = slidePart.Slide.Descendants<Shape>().ToList();
        var workingShapes = shapes.Skip(1).Take(8).ToArray();
        for (int i = 0; i < workingShapes.Length; i++)
        {
            var shapeParagraph = workingShapes[i].TextBody.Descendants<A.Paragraph>().FirstOrDefault();
            if (shapeParagraph != null)
            {
                var shapeText = shapeParagraph.Descendants<A.Text>().FirstOrDefault();
                if (shapeText == null)
                {
                    string textContent = workingShapes[i + 4].InnerText;
                    ReplaceTextInShape(workingShapes[i], textContent);
                }
            }
        }
        foreach (var item in shapes.Skip(5).Take(4))
        {
            item.Remove();
        }
    }

    private void AdjustShapes(SlidePart slidePart)
    {
        var shapes = slidePart.Slide.Descendants<Shape>().ToList();
        var firstRectangle = shapes.Skip(1).FirstOrDefault();
        long firstWidth = 2555204;
        long firstHeight = 1446936;
        long firstY = firstRectangle.ShapeProperties.Transform2D.Offset.Y;
        foreach (var shape in shapes.Skip(1).Take(4))
        {
            shape.ShapeProperties.Transform2D.Extents.Cx = firstWidth;
            shape.ShapeProperties.Transform2D.Extents.Cy = firstHeight;
            long offsetY = firstY - shape.ShapeProperties.Transform2D.Offset.Y;
            shape.ShapeProperties.Transform2D.Offset.Y += offsetY;
        }
    }

    private void ModifyTextAndAlignment(A.Paragraph paragraph, string newText, A.TextAlignmentTypeValues alignment)
    {
        var text = paragraph.Descendants<A.Text>().FirstOrDefault();
        if (text != null)
        {
            text.Text = newText;
        }

        var paragraphProperties = paragraph.ParagraphProperties;
        if (paragraphProperties != null)
        {
            paragraphProperties.Alignment = alignment;
        }
        else
        {
            paragraphProperties = new A.ParagraphProperties() { Alignment = alignment };
            paragraph.InsertAt(paragraphProperties, 0);
        }
    }

    private void SetFontStyle(A.Paragraph paragraph, string typeface)
    {
        var run = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault();
        if (run != null)
        {
            var runProperties = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.RunProperties>();
            if (runProperties == null)
            {
                runProperties = new DocumentFormat.OpenXml.Drawing.RunProperties();
                run.Append(runProperties);
            }

            var font = runProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>();
            if (font == null)
            {
                font = new DocumentFormat.OpenXml.Drawing.LatinFont() { Typeface = typeface };
                runProperties.Append(font);
            }
            else
            {
                font.Typeface = typeface;
            }
        }
    }

    private void ReplaceTextInShape(Shape shape, string newText)
    {
        P.TextBody textBody = shape.TextBody;
        textBody.RemoveAllChildren<A.Paragraph>();
        A.Paragraph paragraph = new A.Paragraph();
        A.Run run = new A.Run();
        A.Text text = new A.Text() { Text = newText };
        run.Append(text);
        paragraph.Append(run);
        textBody.Append(paragraph);
    }
}

