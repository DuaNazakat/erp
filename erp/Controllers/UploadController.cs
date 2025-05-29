using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace PptxExtractorApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UploadController : ControllerBase
    {
        public enum SlideTag
        {
            City,
            Medium,
            SiteStatus,
            DurationDays,
            CostDuration,
            MoveID,
            ImpressionID,
            QTY,
            Location,
            Size,
            TrafficDirection
        }
        private static readonly Dictionary<string, SlideTag> PlaceholderTagMap = new()
        {
            { "Text Placeholder 1", SlideTag.City },
            { "Text Placeholder 3", SlideTag.Medium },
            { "Text Placeholder 7", SlideTag.SiteStatus },
            { "Text Placeholder 8", SlideTag.DurationDays },
            { "Text Placeholder 9", SlideTag.CostDuration },
            { "Text Placeholder 10", SlideTag.MoveID },
            { "Text Placeholder 11", SlideTag.ImpressionID },
            { "Text Placeholder 5", SlideTag.QTY },
            { "Text Placeholder 6", SlideTag.Location },
            { "Text Placeholder 4", SlideTag.Size },
            { "Text Placeholder 22", SlideTag.TrafficDirection }
        };

        [HttpPost("pptx-to-json")]
        public async Task<IActionResult> UploadPptxAndExtractText(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var extension = Path.GetExtension(file.FileName).ToLowerInvariant();
            if (extension != ".pptx")
                return BadRequest("Only .pptx files are supported.");

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            stream.Position = 0;

            var slideTexts = ExtractPptxText(stream);
            return Ok(slideTexts);
        }

        private List<object> ExtractPptxText(Stream pptxStream)
        {
            var slides = new List<object>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(pptxStream, false))
            {
                var presentationPart = presentationDocument.PresentationPart;
                if (presentationPart?.Presentation?.SlideIdList == null)
                    return slides;

                int slideIndex = 1;

                foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
                {
                    var slidePart = presentationPart.GetPartById(slideId.RelationshipId) as SlidePart;
                    if (slidePart == null) continue;

                    var shapesWithText = new List<object>();

                    foreach (var shape in slidePart.Slide.Descendants<Shape>())
                    {
                        var textElements = shape.Descendants<A.Text>().Select(t => t.Text).ToList();
                        var content = string.Join(" ", textElements).Trim();

                        if (string.IsNullOrWhiteSpace(content))
                            continue;

                        var shapeName = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                        var altText = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Description?.Value
                                    ?? shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Title?.Value;

                        string originalTag = !string.IsNullOrWhiteSpace(altText) ? altText :
                       !string.IsNullOrWhiteSpace(shapeName) ? shapeName : "Shape";

                        string finalTag = PlaceholderTagMap.TryGetValue(originalTag, out var mappedTag)
                            ? mappedTag.ToString()
                            : originalTag;

                        shapesWithText.Add(new
                        {
                            tag = finalTag,
                            content = content
                        });
                    }

                    if (shapesWithText.Count > 0)
                    {
                        slides.Add(new
                        {
                            slideNumber = slideIndex,
                            shapes = shapesWithText
                        });
                    }

                    slideIndex++;
                }
            }

            return slides;
        }
        [HttpGet("download-pptx")]
        public IActionResult DownloadPptx()
        {
            var pptxPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files", "sample.pptx");

            if (!System.IO.File.Exists(pptxPath))
                return NotFound("The PPTX file was not found.");

            var pptxBytes = System.IO.File.ReadAllBytes(pptxPath);
            var fileName = "PresentationDownload.pptx";
            var contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

            return File(pptxBytes, contentType, fileName);
        }

    }
}
