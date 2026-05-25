using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocxTemplater.Images;
using ImageMagick;

namespace DocxTemplater.Test
{
    internal class ImageFormatterTest
    {
        [TestCase("jpg", "testImage")]
        [TestCase("tiff", "testImage")]
        [TestCase("png", "testImage")]
        [TestCase("png", "testImage_rot")]
        [TestCase("bmp", "testImage")]
        [TestCase("gif", "testImage")]
        public void ProcessTemplateWithDifferentImageTypes(string extension, string image)
        {
            var imageBytes = File.ReadAllBytes($"Resources/{image}.jpg");
            using var img = new MagickImage(imageBytes);
            img.Format = GetFormatFromExtension(extension);
            imageBytes = img.ToByteArray();

            using var fileStream = File.OpenRead("Resources/ImageFormatterTest.docx");
            var docTemplate = new DocxTemplate(fileStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", new { MyLogo = imageBytes, EmptyArray = Array.Empty<byte>(), NullValue = (byte[])null });

            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        [Test]
        public void InsertSVGAndScaleAndRotate()
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.svg");
            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(h:1cm, r:90)}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }


        [TestCase("w:14cm,h:3cm")]
        [TestCase("w:14cm")]
        [TestCase("h:1cm, r:90")]
        [TestCase("w:1cm")]
        [TestCase("h:1cm")]
        [TestCase("h:15mm")]
        public void InsertHugeImageInsertWithoutContainerFitsToPage(string argument)
        {
            var imageBytes = File.ReadAllBytes("Resources/testImage.jpg");

            // change the size to be bigger than the page
            using var img = new MagickImage(imageBytes);
            img.Resize(img.Width * 10, img.Height * 10);
            img.Format = MagickFormat.Jpeg;
            imageBytes = img.ToByteArray();

            using var memStream = new MemoryStream();
            using var wpDocument = WordprocessingDocument.Create(memStream, WordprocessingDocumentType.Document);
            MainDocumentPart mainPart = wpDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("{{ds}:img(" + argument + ")}")))));
            wpDocument.Save();
            memStream.Position = 0;

            var docTemplate = new DocxTemplate(memStream);
            docTemplate.RegisterFormatter(new ImageFormatter());
            docTemplate.BindModel("ds", imageBytes);
            var result = docTemplate.Process();
            docTemplate.Validate();
            result.SaveAsFileAndOpenInWord();
        }

        private static MagickFormat GetFormatFromExtension(string extension)
        {
            return extension.ToLowerInvariant() switch
            {
                "jpg" or "jpeg" => MagickFormat.Jpeg,
                "png" => MagickFormat.Png,
                "gif" => MagickFormat.Gif,
                "bmp" => MagickFormat.Bmp,
                "tiff" or "tif" => MagickFormat.Tiff,
                _ => throw new ArgumentOutOfRangeException(nameof(extension), extension, "Unsupported image format extension")
            };
        }
    }
}
