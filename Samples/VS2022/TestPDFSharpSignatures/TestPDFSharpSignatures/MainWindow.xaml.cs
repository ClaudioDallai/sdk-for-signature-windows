using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.Signatures;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;

using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace TestPDFSharpSignatures
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var document = PdfReader.Open("file.pdf", "password");
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);

            var pdfPosition = gfx.Transformer.WorldToDefaultPage(new XPoint(144, 600));
            var options = new DigitalSignatureOptions
            {
                ContactInfo = "John Doe",
                Location = "Seattle",
                Reason = "License Agreement",
                Rectangle = new XRect(pdfPosition.X, pdfPosition.Y, 200, 50),
                //AppearanceHandler = new SignatureAppearanceHandler()
            };
        }
    }
}