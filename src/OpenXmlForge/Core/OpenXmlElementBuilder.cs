using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlForge.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlForge.Core;

internal static class OpenXmlElementBuilder
{
    public static Hyperlink CreateHyperlink(OpenXmlPartContainer partContainer, WordHyperLink linkInfo)
    {
        var rel = partContainer switch
        {
            MainDocumentPart main => main.AddHyperlinkRelationship(new Uri(linkInfo.Url), true),
            HeaderPart header => header.AddHyperlinkRelationship(new Uri(linkInfo.Url), true),
            FooterPart footer => footer.AddHyperlinkRelationship(new Uri(linkInfo.Url), true),
            _ => throw new InvalidOperationException("Unsupported OpenXmlPartContainer for Hyperlink."),
        };

        return new Hyperlink(
            new Run(
                new RunProperties(
                    new RunStyle { Val = "Hyperlink" },
                    new Underline { Val = linkInfo.UnderLineValue },
                    new Color { ThemeColor = ThemeColorValues.Hyperlink }),
                new Text(linkInfo.Text)))
        {
            DocLocation = linkInfo.Url,
            Id = rel.Id,
            TargetFrame = linkInfo.GetTargetFrameString(),
        };
    }

    public static void AppendColorTexts(Run run, ReadOnlySpan<WordColorText> items)
    {
        foreach (var item in items)
        {
            if (item is null)
            {
                continue;
            }

            var runPro = new RunProperties();
            if (!string.IsNullOrEmpty(item.HighlightColor))
            {
                runPro.Append(new Shading { Fill = item.HighlightColor.Replace("#", "", StringComparison.Ordinal) });
            }

            if (!string.IsNullOrEmpty(item.FontColor))
            {
                runPro.Append(new Color { Val = item.FontColor.Replace("#", "", StringComparison.Ordinal) });
            }

            var text = new Text(item.Text ?? string.Empty);
            if (text.Text.Length > 0 && (text.Text[0] == ' ' || text.Text[^1] == ' '))
            {
                text.Space = SpaceProcessingModeValues.Preserve;
            }

            run.Append(runPro);
            run.Append(text);
        }
    }

    public static Drawing CreateInlinePictureDrawing(
        string relationshipId,
        long cxEmus,
        long cyEmus,
        string extension)
    {
        return new Drawing(new DW.Inline(
                new DW.Extent { Cx = cxEmus, Cy = cyEmus },
                new DW.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L,
                },
                new DW.DocProperties
                {
                    Id = (UInt32Value)1U,
                    Name = $"Picture {Guid.NewGuid()}",
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties
                                {
                                    Id = (UInt32Value)0U,
                                    Name = $"Image {Guid.NewGuid()}.{extension}",
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension
                                        {
                                            Uri = $"{{{Guid.NewGuid():n}}}",
                                        }))
                                {
                                    Embed = relationshipId,
                                    CompressionState = A.BlipCompressionValues.Print,
                                },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = cxEmus, Cy = cyEmus }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle,
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                    }))
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946",
            });
    }

    public static Drawing CreateAnchorPictureDrawing(
        string relationshipId,
        long cxEmus,
        long cyEmus,
        WordPicture pic)
    {
        var anchor = new DW.Anchor
        {
            DistanceFromTop = (UInt32Value)0U,
            DistanceFromBottom = (UInt32Value)0U,
            DistanceFromLeft = (UInt32Value)114300U,
            DistanceFromRight = (UInt32Value)114300U,
            SimplePos = false,
            RelativeHeight = (UInt32Value)0U,
            BehindDoc = pic.BehindDoc,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = pic.AllowOverlap,
            EditId = "1EACBECC",
            AnchorId = "5AF073F3",
        };

        var simplePosition = new DW.SimplePosition { X = 0L, Y = 0L };

        var horizontalPosition = new DW.HorizontalPosition
        {
            RelativeFrom = DW.HorizontalRelativePositionValues.Column,
        };
        horizontalPosition.Append(
            new DW.PositionOffset(
                $"{pic.HorizontalPositionOffset * WordPicture.EmusPerPixel}"));

        var verticalPosition = new DW.VerticalPosition
        {
            RelativeFrom = DW.VerticalRelativePositionValues.Paragraph,
        };
        verticalPosition.Append(
            new DW.PositionOffset(
                $"{pic.VerticalPositionOffset * WordPicture.EmusPerPixel}"));

        var extent = new DW.Extent { Cx = cxEmus, Cy = cyEmus };
        var effectExtent = new DW.EffectExtent
        {
            LeftEdge = 0L,
            TopEdge = 0L,
            RightEdge = 0L,
            BottomEdge = 0L,
        };
        var wrapNone = new DW.WrapNone();
        var docProperties = new DW.DocProperties
        {
            Id = (UInt32Value)774829591U,
            Name = $"Picture {Guid.NewGuid()}",
        };

        var nonVisualGraphicFrameDrawingProperties = new DW.NonVisualGraphicFrameDrawingProperties();
        var graphicFrameLocks = new A.GraphicFrameLocks { NoChangeAspect = true };
        graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        nonVisualGraphicFrameDrawingProperties.Append(graphicFrameLocks);

        var graphic = new A.Graphic();
        graphic.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        var graphicData = new A.GraphicData
        {
            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
        };

        var picture = new PIC.Picture();
        picture.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

        var nonVisualPictureProperties = new PIC.NonVisualPictureProperties(
            new PIC.NonVisualDrawingProperties
            {
                Id = (UInt32Value)0U,
                Name = $"Image {Guid.NewGuid()}.{pic.Extension}",
            },
            new PIC.NonVisualPictureDrawingProperties());

        var blipFill = new PIC.BlipFill(
            new A.Blip(
                new A.BlipExtensionList(
                    new A.BlipExtension { Uri = $"{{{Guid.NewGuid():n}}}" }))
            {
                Embed = relationshipId,
                CompressionState = A.BlipCompressionValues.Print,
            },
            new A.Stretch(new A.FillRectangle()));

        var shapeProperties = new PIC.ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = cxEmus, Cy = cyEmus }),
            new A.PresetGeometry(new A.AdjustValueList())
            {
                Preset = A.ShapeTypeValues.Rectangle,
            });

        picture.Append(nonVisualPictureProperties, blipFill, shapeProperties);
        graphicData.Append(picture);
        graphic.Append(graphicData);

        anchor.Append(
            simplePosition,
            horizontalPosition,
            verticalPosition,
            extent,
            effectExtent,
            wrapNone,
            docProperties,
            nonVisualGraphicFrameDrawingProperties,
            graphic);

        return new Drawing(anchor);
    }
}
