// Wacom Ink sdk
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Quality;
using System.Globalization;
using System.Web;
using System.Windows;

namespace TestPDFSharpSignatures
{
    class SignatureAppearanceHandler : IAnnotationAppearanceHandler
    {
        private string _imagePath;
        private string _signer;
        private string _place;
        private string _font;


        public SignatureAppearanceHandler(string inImgPath = "", string inSigner = "", string inPlace = "", string inFont = "")
        {
            _imagePath = inImgPath;
            _signer = inSigner;
            _place = inPlace;
            _font = inFont;
        }

        public void DrawAppearance(XGraphics gfx, XRect rect)
        {
            var image = XImage.FromFile(_imagePath);

            string text = $"{_signer}\n{_place}, " + DateTime.Now.ToString(CultureInfo.GetCultureInfo("EN-US"));
            var font = new XFont(_font, 7.0);
            var textFormatter = new XTextFormatter(gfx);
            double num = (double)image.PixelWidth / image.PixelHeight;
            double signatureHeight = rect.Height;
            var point = new XPoint(rect.Width / 10, rect.Height / 10);
            // Draw image.
            gfx.DrawImage(image, point.X, point.Y, signatureHeight * num, signatureHeight);
            // Adjust position for text. We draw it below image.
            point = new XPoint(point.X, rect.Height * 0.8f);
            textFormatter.DrawString(text, font, new XSolidBrush(XColor.FromKnownColor(XKnownColor.Black)), new XRect(point.X, point.Y, rect.Width, rect.Height - point.Y), XStringFormats.TopLeft);
        }
    }
}