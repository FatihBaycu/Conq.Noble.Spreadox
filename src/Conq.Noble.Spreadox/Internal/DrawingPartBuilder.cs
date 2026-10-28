using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace Conq.Noble.Spreadox.Internal
{
    /// <summary>
    /// Worksheet drawings: classic charts, ChartEx (box-and-whisker), and pictures share one DrawingsPart.
    /// </summary>
    internal static class DrawingPartBuilder
    {
        public static void Attach(WorksheetPart worksheetPart, ExcelWorksheet model)
        {
            if (model.Charts.Count == 0 && model.Pictures.Count == 0)
            {
                return;
            }

            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            var worksheetDrawing = new Xdr.WorksheetDrawing();
            drawingsPart.WorksheetDrawing = worksheetDrawing;

            uint shapeId = 2;
            foreach (ExcelChart chart in model.Charts)
            {
                if (chart.Type == ExcelChartType.BoxAndWhisker)
                {
                    ExtendedChartPart exPart = drawingsPart.AddNewPart<ExtendedChartPart>();
                    exPart.ChartSpace = ChartExPartBuilder.BuildBoxAndWhisker(chart, model);
                    exPart.ChartSpace.Save();
                    string relId = drawingsPart.GetIdOfPart(exPart);
                    worksheetDrawing.AppendChild(BuildChartExAnchor(chart, relId, shapeId));
                }
                else
                {
                    ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                    chartPart.ChartSpace = ChartPartBuilder.BuildChartSpace(chart, model);
                    chartPart.ChartSpace.Save();
                    string relId = drawingsPart.GetIdOfPart(chartPart);
                    worksheetDrawing.AppendChild(ChartPartBuilder.BuildTwoCellAnchor(chart, relId, shapeId));
                }

                shapeId++;
            }

            foreach (ExcelPicture picture in model.Pictures)
            {
                PartTypeInfo imageType = picture.Format == ExcelPictureFormat.Jpeg
                    ? ImagePartType.Jpeg
                    : ImagePartType.Png;
                ImagePart imagePart = drawingsPart.AddImagePart(imageType);
                using (var ms = new MemoryStream(picture.ImageBytes, writable: false))
                {
                    imagePart.FeedData(ms);
                }

                string relId = drawingsPart.GetIdOfPart(imagePart);
                worksheetDrawing.AppendChild(BuildPictureAnchor(picture, relId, shapeId));
                shapeId++;
            }

            drawingsPart.WorksheetDrawing.Save();

            DocumentFormat.OpenXml.Spreadsheet.Drawing? existing =
                worksheetPart.Worksheet!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
            existing?.Remove();

            worksheetPart.Worksheet.AppendChild(
                new DocumentFormat.OpenXml.Spreadsheet.Drawing
                {
                    Id = worksheetPart.GetIdOfPart(drawingsPart)
                });
        }

        private static Xdr.TwoCellAnchor BuildChartExAnchor(ExcelChart chart, string relId, uint shapeId)
        {
            var anchor = ChartPartBuilder.CreateRangeAnchor(chart.FromRow, chart.FromColumn, chart.ToRow, chart.ToColumn);
            var graphicFrame = new Xdr.GraphicFrame { Macro = string.Empty };
            graphicFrame.AppendChild(new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = shapeId, Name = "Chart " + shapeId },
                new Xdr.NonVisualGraphicFrameDrawingProperties()));
            graphicFrame.AppendChild(new Xdr.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }));
            graphicFrame.AppendChild(new A.Graphic(
                new A.GraphicData(
                    new Cx.RelId { Id = relId })
                {
                    Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex"
                }));
            anchor.AppendChild(graphicFrame);
            anchor.AppendChild(new Xdr.ClientData());
            return anchor;
        }

        private static Xdr.TwoCellAnchor BuildPictureAnchor(ExcelPicture picture, string relId, uint shapeId)
        {
            ExcelRange range = picture.AnchorRange;
            var anchor = ChartPartBuilder.CreateRangeAnchor(
                range.StartRow, range.StartColumn, range.EndRow, range.EndColumn);

            string name = string.IsNullOrWhiteSpace(picture.Name) ? "Picture " + shapeId : picture.Name!;
            var pic = new Xdr.Picture();
            pic.AppendChild(new Xdr.NonVisualPictureProperties(
                new Xdr.NonVisualDrawingProperties { Id = shapeId, Name = name },
                new Xdr.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = picture.LockAspectRatio })));
            pic.AppendChild(new Xdr.BlipFill(
                new A.Blip { Embed = relId },
                new A.Stretch(new A.FillRectangle())));
            pic.AppendChild(new Xdr.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

            anchor.AppendChild(pic);
            anchor.AppendChild(new Xdr.ClientData());
            return anchor;
        }
    }
}
