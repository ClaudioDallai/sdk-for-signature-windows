// Wacom Ink sdk
using FlSigCaptLib;
using FLSIGCTLLib;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.Signatures;
using PdfSharp.Snippets.Pdf;
using System.IO;
using System.Reflection.Metadata;
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
using UglyToad.PdfPig.AcroForms;
using UglyToad.PdfPig.AcroForms.Fields;
using UglyToad.PdfPig.Content;


/*
 Libraries and technologies used:

    -> Wacom INK SDK (free license if only used to sign using Wacom STU);
    -> PigPDF (pdf data reading, MIT);
    -> PDFSharp (pdf data creation and signing, MIT);
    -> BouncyCastle (legal signing, MIT);
 */


namespace TestPDFSharpSignatures
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string _sdkLicenseKey = "eyJhbGciOiJSUzUxMiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImV4cCI6MjE0NzQ4MzY0NywiaWF0IjoxNTYwOTUwMjcyLCJyaWdodHMiOlsiU0lHX1NES19DT1JFIiwiU0lHQ0FQVFhfQUNDRVNTIl0sImRldmljZXMiOlsiV0FDT01fQU5ZIl0sInR5cGUiOiJwcm9kIiwibGljX25hbWUiOiJTaWduYXR1cmUgU0RLIiwid2Fjb21faWQiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImxpY191aWQiOiJiODUyM2ViYi0xOGI3LTQ3OGEtYTlkZS04NDlmZTIyNmIwMDIiLCJhcHBzX3dpbmRvd3MiOltdLCJhcHBzX2lvcyI6W10sImFwcHNfYW5kcm9pZCI6W10sIm1hY2hpbmVfaWRzIjpbXX0.ONy3iYQ7lC6rQhou7rz4iJT_OJ20087gWz7GtCgYX3uNtKjmnEaNuP3QkjgxOK_vgOrTdwzD-nm-ysiTDs2GcPlOdUPErSp_bcX8kFBZVmGLyJtmeInAW6HuSp2-57ngoGFivTH_l1kkQ1KMvzDKHJbRglsPpd4nVHhx9WkvqczXyogldygvl0LRidyPOsS5H2GYmaPiyIp9In6meqeNQ1n9zkxSHo7B11mp_WXJXl0k1pek7py8XYCedCNW5qnLi4UCNlfTd6Mk9qz31arsiWsesPeR9PN121LBJtiPi023yQU8mgb9piw_a-ccciviJuNsEuRDN3sGnqONG3dMSA";
        private readonly string _font = "Arial";

        public MainWindow()
        {

            InitializeComponent();

            String signTargetPath = "";

            //print("btnSign was pressed");
            SigCtl sigCtl = new SigCtl();
            sigCtl.Licence = _sdkLicenseKey;
            //DynamicCapture dc = new DynamicCaptureClass();

            string inputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\LoremIpsumLeft.pdf";
            string outputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\ResultPDFSharpForms.pdf";

            string marker_sign = "{{{SIGN_HERE}}}";
            string marker_name = "{{{NAME_HERE}}}";
            string marker_surname = "{{{SURNAME_HERE}}}";

            string reason = "Firma di prova";
            string location = "Italia";
            string signer = "Mario Rossi";


            // Get signature requested positions in every page. Using PDFSharp works only for a single signature.
            #region PdfPig


            //// Use PdfPig to interrogate the PDF and rtreive infos about sign slot locations
            //int signTargetPage = -1;

            Dictionary<int, List<System.Windows.Vector>> markerPositions = new Dictionary<int, List<System.Windows.Vector>>();

            //using (UglyToad.PdfPig.PdfDocument documentPig = UglyToad.PdfPig.PdfDocument.Open(inputPdf))
            //{
            //    foreach (var page in documentPig.GetPages())
            //    {
            //        foreach (var word in page.GetWords())
            //        {
            //            if (word.Text.Contains(marker_sign))
            //            {
            //                signTargetPage = page.Number;

            //                if (markerPositions.ContainsKey(signTargetPage))
            //                {
            //                    if (word.BoundingBox.Left > page.Width * 0.5)
            //                    {
            //                        markerPositions[signTargetPage].Add(new System.Windows.Vector(word.BoundingBox.Left - 25, word.BoundingBox.Top));
            //                    }
            //                    else
            //                    {
            //                        markerPositions[signTargetPage].Add(new System.Windows.Vector(word.BoundingBox.Left, word.BoundingBox.Top));
            //                    }
            //                }
            //                else
            //                {
            //                    if (word.BoundingBox.Left > page.Width * 0.5)
            //                    {
            //                        markerPositions.Add(signTargetPage, new List<System.Windows.Vector> { new System.Windows.Vector(word.BoundingBox.Left - 25, word.BoundingBox.Top) });
            //                    }
            //                    else
            //                    {
            //                        markerPositions.Add(signTargetPage, new List<System.Windows.Vector> { new System.Windows.Vector(word.BoundingBox.Left, word.BoundingBox.Top) });
            //                    }
            //                }
            //            }
            //        }
            //    }

            //    // PDFPig counts from page 1.
            //    HashSet<int> invalidKeys = new HashSet<int>();
            //    foreach (KeyValuePair<int, List<System.Windows.Vector>> pair in markerPositions)
            //    {
            //        if (pair.Key <= 0 || pair.Key > documentPig.NumberOfPages)
            //        {
            //            invalidKeys.Add(pair.Key);
            //        }
            //    }

            //    foreach (int invKey in invalidKeys)
            //    {
            //        if (markerPositions.ContainsKey(invKey))
            //        {
            //            markerPositions.Remove(invKey);
            //        }
            //    }
            //}

            using (UglyToad.PdfPig.PdfDocument documentPig = UglyToad.PdfPig.PdfDocument.Open(inputPdf))
            {
                documentPig.TryGetForm(out var form);
                if (form == null)
                {
                    Console.WriteLine("Nessun AcroForm trovato.");
                    return;
                }

                foreach (var field in form.Fields)
                {
                    if (field.Information.PartialName == "Text2")
                    {

                        markerPositions.Add(field.PageNumber.Value, new List<System.Windows.Vector> { new System.Windows.Vector(field.Bounds.Value.Left, field.Bounds.Value.Right - 50d) });
                    }
                }
            }


            #endregion


            #region Certificate


            // Hardoded as example

            // Not necessary if using Windows Certificate Store
            //string pfxPath = @"C:\temp\firma-demo.pfx"; 
            //string pfxPassword = "password123";


            // Certificate subject name
            string subjectName = "FirmaDemo";

            // Open current user store
            var store = new System.Security.Cryptography.X509Certificates.X509Store(
                    System.Security.Cryptography.X509Certificates.StoreName.My,
                    System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser);

            store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);


            // Cerca certificati con il subject indicato (anche quelli scaduti con validOnly:false)
            var certs = store.Certificates.Find(
                System.Security.Cryptography.X509Certificates.X509FindType.FindBySubjectName,
                subjectName,
                validOnly: false);

            if (certs.Count == 0)
                throw new System.Exception($"Certificato con subject '{subjectName}' non trovato.");

            // Prendi il primo certificato trovato
            // ATTENZIONE: In caso di certificati con lo stesso nome, con magari alcuni scaduti, ne va preso uno valido.
            var cert = certs[0];


            // I campi NotBefore e NotAfter nei certificati sono sempre espressi in UTC secondo lo standard X.509
            DateTime now = DateTime.UtcNow;
            if (now > cert.NotAfter.ToUniversalTime())
            {
                throw new System.Exception($"Certificato con subject '{subjectName}' scaduto.");
            }

            if (!cert.HasPrivateKey)
                throw new System.Exception("Il certificato non contiene la chiave privata.");

            var chain = new X509Chain();
            chain.Build(cert);
            var chainCerts = new X509Certificate2Collection();
            foreach (var element in chain.ChainElements)
            {
                chainCerts.Add(element.Certificate);
            }


            #endregion


            PdfDocument document = PdfReader.Open(inputPdf, PdfDocumentOpenMode.Modify);


            // TODO: Undersant forms
            //PdfAcroForm formfield = document.AcroForm;

            //if (formfield.Elements.ContainsKey("/NeedAppearances"))
            //{
            //    formfield.Elements["/NeedAppearances"] = new PdfSharp.Pdf.PdfBoolean(true);
            //}
            //else
            //{
            //    formfield.Elements.Add("/NeedAppearances", new PdfSharp.Pdf.PdfBoolean(true));
            //}

            //PdfTextField testField = (PdfTextField)(formfield.Fields["Text1"]);
            //testField.Text = signer;


            ActivateWacom(sigCtl, ref signTargetPath);
            var options = new DigitalSignatureOptions
            {
                ContactInfo = signer,
                Location = location,
                Reason = reason,
                Rectangle = new XRect((double)markerPositions[1][0].X, (double)markerPositions[1][0].Y, 150d, 100d),
                AppearanceHandler = new SignatureAppearanceHandler(signTargetPath, signer, location, _font)
            };


            // Different alghoritms are available. This one is required to work on Edge PDF Viewer (Microsoft issue)
            document.SecurityHandler.SetEncryption(PdfDefaultEncryption.V2With128Bits);

            document.SecurityHandler.OwnerPassword = "Admin";
            //document.SecurityHandler.UserPassword = "";

            document.SecuritySettings.PermitPrint = true;
            document.SecuritySettings.PermitFullQualityPrint = true;
            document.SecuritySettings.PermitExtractContent = true;

            document.SecuritySettings.PermitAssembleDocument = false;
            document.SecuritySettings.PermitModifyDocument = false;
            document.SecuritySettings.PermitFormsFill = false;
            document.SecuritySettings.PermitAnnotations = false;


            var pdfSignatureHandler = DigitalSignatureHandler.ForDocument(document, new BouncyCastleSigner((cert, chainCerts), PdfMessageDigestType.SHA256), options);
            document.Save(outputPdf);
            //document.Close(); Save already do this

            store.Close();
            Application.Current.Shutdown();
        }

        public void ActivateWacom(SigCtl sigCtl, ref string signTargetPath)
        {
            DynamicCapture dc = new FlSigCaptLib.DynamicCapture();
            DynamicCaptureResult res = dc.Capture(sigCtl, "Mario Rossi", "Presentazione Esempio 1", null, null);
            if (res == DynamicCaptureResult.DynCaptOK)
            {
                //print("signature captured successfully");
                SigObj sigObj = (SigObj)sigCtl.Signature;
                sigObj.set_ExtraData("AdditionalData", "C# test: Additional data");

                //var testRead = sigObj.ExtraData["AdditionalData"]; // Works

                String dateStr = DateTime.Now.ToString("hhmmss");

                string folderPath = @"C:\temp";
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                signTargetPath = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Signs\" + dateStr + ".png";
                //print("Outputting to file " + signTargetPath);
                try
                {
                    //print("Saving signature to file " + filename);
                    // Need to understand if DimensionX and DimensionY alter the image metadata. These values (200, 150) were already here in the example.
                    sigObj.RenderBitmap(signTargetPath, 200, 150, "image/png", 0.5f, 0xff0000, 0xffffff, 10.0f, 10.0f, RBFlags.RenderOutputFilename | RBFlags.RenderColor32BPP | RBFlags.RenderEncodeData);

                    //print("Loading image from " + signTargetPath);
                    BitmapImage src = new BitmapImage();
                    src.BeginInit();
                    src.UriSource = new Uri(signTargetPath, UriKind.Absolute);
                    src.EndInit();

                    //imgSig.Source = src;
                }
                catch (Exception ex)
                {
                    //System.Windows.Forms.MessageBox.Show(ex.Message);
                }

            }
            else
            {

                // IF WACOM STU IS NOT ATTACHED: Returns message that license is invalid!
                //print("Signature capture error res=" + (int)res + "  ( " + res + " )");
                switch (res)
                {
                    case DynamicCaptureResult.DynCaptCancel: /*print("signature cancelled")*/; break;
                    case DynamicCaptureResult.DynCaptError: /*print("no capture service available");*/ break;
                    case DynamicCaptureResult.DynCaptPadError: /*print("signing device error");*/ break;
                    case DynamicCaptureResult.DynCaptNotLicensed: /*print("license error or device is not connected");*/ break;
                    default: /*print("Unexpected error code ");*/ break;

                }
            }
        }
    }
}