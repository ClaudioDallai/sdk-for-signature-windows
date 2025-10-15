
// Wacom Ink sdk
//using Interop.FlSigCOM; DO NOT NEED THIS
using FLSIGCTLLib;
using Interop.FlSigCapt;

// PDF related
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.AcroForms;
using PdfSharp.Pdf.Annotations;
using PdfSharp.Pdf.IO;
using UtilityWordLib;

// Crypto and signing
using PdfSharp.Pdf.Security;
using PdfSharp.Pdf.Signatures;
using PdfSharp.Snippets.Pdf;
// System
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Windows;
using System.Windows.Media.Imaging;

// provare a creare interop da .COM dll di Word 16. Altrimenti non builda l'exe in NET8. Il nuget non è più supportato!!!

/*
 Libraries and technologies used:
    -> Wacom INK SDK (free license if only used to sign using Wacom STU);
    -> PDFSharp (pdf data creation, forms, signing, MIT);
    -> BouncyCastle (legal signing, MIT);
 */

// NOT USED ANYMORE -> PigPDF (pdf data reading, forms, MIT);


namespace TestPDFSharpSignatures
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly string _sdkLicenseKey = "eyJhbGciOiJSUzUxMiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImV4cCI6MjE0NzQ4MzY0NywiaWF0IjoxNTYwOTUwMjcyLCJyaWdodHMiOlsiU0lHX1NES19DT1JFIiwiU0lHQ0FQVFhfQUNDRVNTIl0sImRldmljZXMiOlsiV0FDT01fQU5ZIl0sInR5cGUiOiJwcm9kIiwibGljX25hbWUiOiJTaWduYXR1cmUgU0RLIiwid2Fjb21faWQiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImxpY191aWQiOiJiODUyM2ViYi0xOGI3LTQ3OGEtYTlkZS04NDlmZTIyNmIwMDIiLCJhcHBzX3dpbmRvd3MiOltdLCJhcHBzX2lvcyI6W10sImFwcHNfYW5kcm9pZCI6W10sIm1hY2hpbmVfaWRzIjpbXX0.ONy3iYQ7lC6rQhou7rz4iJT_OJ20087gWz7GtCgYX3uNtKjmnEaNuP3QkjgxOK_vgOrTdwzD-nm-ysiTDs2GcPlOdUPErSp_bcX8kFBZVmGLyJtmeInAW6HuSp2-57ngoGFivTH_l1kkQ1KMvzDKHJbRglsPpd4nVHhx9WkvqczXyogldygvl0LRidyPOsS5H2GYmaPiyIp9In6meqeNQ1n9zkxSHo7B11mp_WXJXl0k1pek7py8XYCedCNW5qnLi4UCNlfTd6Mk9qz31arsiWsesPeR9PN121LBJtiPi023yQU8mgb9piw_a-ccciviJuNsEuRDN3sGnqONG3dMSA";
        private readonly string _font = "Arial";
        //private readonly string _nameSurnameForm = "Name_Surname_Form";
        private readonly string _signPlaceholderForm = "Sign_Placeholder_Form";
        //private readonly string _currentDateForm = "Current_Date_Form";
        private readonly string _appName = " [Sign_It+] -> Wacom STU-430";

        private readonly string _nameMarker = "{{{NOME_COGNOME}}}";
        private readonly string _placeMarker = "{{{LUOGO_DATA}}}";
        private readonly string _signMarker = "{{{FIRMA}}}";

        // Certificate subject name
        private readonly string _subjectName = "FirmaDemo";


        public MainWindow()
        {

            InitializeComponent();

            try
            {
                string[] args = Environment.GetCommandLineArgs();
                String signTargetPath = "";

                //print("btnSign was pressed");
                SigCtl sigCtl = new SigCtl();
                sigCtl.Licence = _sdkLicenseKey;


                string inputTemplateWord = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok.docx";
                string inputFilledTemplateWord = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok_Filled.docx";

                string outputFilledTemplatePathNoSign = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok_Filled_NoSign.pdf";


                string inputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\ModelloAutorizzazioneTrattamentoDatiForm.pdf";
                //string inputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\LoremIpsumRightForms.pdf";
                //string inputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\LoremIpsumMultiForms.pdf";
                string outputPath = @"C:\Users\visio\Desktop\PDF_ManipulationTests\";

                // These will be paramethric
                string reason = "";
                string location = "";
                string signerName = "";
                string signerSurname = "";

                if (args.Length <= 1)
                {
                    reason = "Sign Reason Placeholder";
                    location = "Italia, Firenze";
                    signerName = "Luca";
                    signerSurname = "Bianchi";
                }
                else
                {
                    // First [0] is exe path
                    reason = args[1];
                    location = args[2];
                    signerName = args[3];
                    signerSurname = args[4];
                }

                string signer = signerName + " " + signerSurname;
                SignLocationInfo signInfo = new SignLocationInfo();

                if (File.Exists(inputTemplateWord))
                {
                    WordProcessor utWord = new WordProcessor();
                    signInfo = utWord.FillTemplate(inputTemplateWord, inputFilledTemplateWord, new Dictionary<string, string> { { _nameMarker, signer }, { _signMarker, "" } });
                    utWord.ConvertToPdf(inputFilledTemplateWord, outputFilledTemplatePathNoSign);
                }


                #region Certificate


                // Hardoded as example

                // Not necessary if using Windows Certificate Store
                //string pfxPath = @"C:\temp\firma-demo.pfx"; 
                //string pfxPassword = "password123";


                // Open current user store
                var store = new System.Security.Cryptography.X509Certificates.X509Store(
                        System.Security.Cryptography.X509Certificates.StoreName.My,
                        System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser);

                store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.ReadOnly);


                // Cerca certificati con il subject indicato (anche quelli scaduti con validOnly:false)
                var certs = store.Certificates.Find(
                    System.Security.Cryptography.X509Certificates.X509FindType.FindBySubjectName,
                    _subjectName,
                    validOnly: false);

                if (certs.Count == 0)
                    throw new System.Exception($"Certificato con subject '{_subjectName}' non trovato.");


                X509Certificate2? validCertFound = null;
                DateTime now = DateTime.UtcNow;
                foreach (X509Certificate2? certificate in certs)
                {
                    if (certificate == null) continue;

                    // I campi NotBefore e NotAfter nei certificati sono sempre espressi in UTC secondo lo standard X.509
                    if (now >= certificate.NotBefore.ToUniversalTime() && now <= certificate.NotAfter.ToUniversalTime() && certificate.HasPrivateKey)
                    {
                        validCertFound = certificate;
                        break;
                    }
                }

                if (validCertFound == null)
                {
                    throw new System.Exception("No valid certificates found. Check expiration data or key");
                }

                var chain = new X509Chain();
                chain.Build(validCertFound);
                var chainCerts = new X509Certificate2Collection();
                foreach (var element in chain.ChainElements)
                {
                    chainCerts.Add(element.Certificate);
                }


                #endregion


                // PDFSharp has lots of problem in recognizing metadata of acroforms, better to use texts and use AcroForms just as placeholders to get location.
                PdfDocument document = PdfReader.Open(outputFilledTemplatePathNoSign, PdfDocumentOpenMode.Modify);
                //PdfAcroForm formFields = document.AcroForm;


                //// Name and Surname
                //int? pageName_index = GetPageIndexOfField(document, _nameSurnameForm);
                //if (pageName_index.HasValue)
                //{
                //    InsertTextWhereFormIs(document, formFields, pageName_index.Value, _nameSurnameForm, signer);
                //}

                //// Date and place
                //string dateAndPlace = location + "," + $" {DateTime.Now.ToString("dd/MM/yyyy")}";
                //int? pageDate_index = GetPageIndexOfField(document, _currentDateForm);
                //if (pageDate_index.HasValue)
                //{
                //    InsertTextWhereFormIs(document, formFields, pageDate_index.Value, _currentDateForm, dateAndPlace);
                //}

                // Sign
                //int? pageSign_index = GetPageIndexOfField(document, _signPlaceholderForm);
                //if (pageSign_index.HasValue)
                if (true)
                {
                    //InsertSignature(out DigitalSignatureOptions? options, formFields, sigCtl, pageSign_index.Value, ref signTargetPath, signer, location, reason);
                    //InsertSignature(out DigitalSignatureOptions? options, formFields, sigCtl, 1, ref signTargetPath, signer, location, reason);

                    InsertSignatureWordMarker(out DigitalSignatureOptions? options, signInfo, sigCtl, document, ref signTargetPath, signer, location, reason);

                    if (options != null)
                    {
                        // Different alghoritms are available. This one is required to work on Edge PDF Viewer (Microsoft issue)
                        document.SecurityHandler.SetEncryption(PdfDefaultEncryption.V2With128Bits);

                        document.SecurityHandler.OwnerPassword = "Admin_Visio";
                        //document.SecurityHandler.UserPassword = "";

                        document.SecuritySettings.PermitPrint = true;
                        document.SecuritySettings.PermitFullQualityPrint = true;
                        document.SecuritySettings.PermitExtractContent = true;

                        document.SecuritySettings.PermitAssembleDocument = false;
                        document.SecuritySettings.PermitModifyDocument = false;
                        document.SecuritySettings.PermitFormsFill = false;
                        document.SecuritySettings.PermitAnnotations = false;


                        var pdfSignatureHandler = DigitalSignatureHandler.ForDocument(document, new BouncyCastleSigner((validCertFound, chainCerts), PdfMessageDigestType.SHA256), options);
                        string resultPDF = outputPath + $"{signer.Replace(" ", "")}_{reason.Replace(" ", "")}.pdf";

                        document.Save(resultPDF);
                        //document.Close(); Save already do this
                    }
                }

                store.Close();
            }
            catch
            {
            }


            System.Windows.Application.Current.Shutdown();
        }

        public void ActivateWacom(SigCtl sigCtl, ref string signTargetPath, string signer, string reason)
        {
            DynamicCapture dc = new Interop.FlSigCapt.DynamicCapture();
            DynamicCaptureResult res = dc.Capture(sigCtl, signer, reason, null, null);
            if (res == DynamicCaptureResult.DynCaptOK)
            {
                //print("signature captured successfully");
                SigObj sigObj = (SigObj)sigCtl.Signature;
                sigObj.set_ExtraData("AdditionalData", "C# test: Additional data");

                //var testRead = sigObj.ExtraData["AdditionalData"]; // Works

                String dateStr = DateTime.Now.ToString("hhmmss");

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

        int? GetPageIndexOfField(PdfDocument document, string fieldName)
        {
            for (int i = 0; i < document.Pages.Count; i++)
            {
                PdfPage page = document.Pages[i];
                if (page.Annotations == null)
                    continue;

                foreach (PdfAnnotation annotation in page.Annotations)
                {
                    // Controlla se l'annotazione è di tipo /Widget
                    var subtype = annotation.Elements.GetName("/Subtype");
                    if (subtype == "/Widget")
                    {
                        // Leggi il nome del campo /T
                        string t = annotation.Elements.GetString("/T");
                        if (t == fieldName)
                        {
                            return i;
                        }
                    }
                }
            }

            return null;
        }

        void InsertTextWhereFormIs(PdfDocument document, PdfAcroForm formFields, int page_index, string field, string text)
        {
            try
            {

                PdfAcroField? foundField = formFields.Fields[field];
                if (foundField == null) return;

                var fieldFormPage = document.Pages[page_index];

                var font = new XFont(_font, 14d);
                XGraphics gfx = XGraphics.FromPdfPage(fieldFormPage);
                var textFormatter = new XTextFormatter(gfx);

                PdfRectangle rectForm = foundField.Elements.GetRectangle(PdfAnnotation.Keys.Rect);
                XRect locationFormField = rectForm.ToXRect();

                // Text has pivot inverted
                double invertedY = fieldFormPage.Height.Point - locationFormField.Y - locationFormField.Height;
                XRect adjustedFormFieldRect = new XRect(locationFormField.X, invertedY, locationFormField.Width, locationFormField.Height);

                //gfx.DrawRectangle(XPens.Red, adjustedNameRect);
                textFormatter.DrawString(text, font, XBrushes.Black, adjustedFormFieldRect, XStringFormats.TopLeft);
                gfx.Dispose();
            }
            catch
            {

            }
        }

        void InsertSignature(out DigitalSignatureOptions? options, PdfAcroForm formFields, SigCtl sigCtl, int pageSign_index, ref string signTargetPath, string signer, string location, string reason)
        {
            options = null;

            try
            {
                PdfAcroField? testSignField = formFields.Fields[_signPlaceholderForm];
                if (testSignField == null)
                {
                    return;
                }

                PdfRectangle rectSign = testSignField.Elements.GetRectangle(PdfAnnotation.Keys.Rect);
                XRect locationFinalSign = rectSign.ToXRect();

                ActivateWacom(sigCtl, ref signTargetPath, signer, reason);
                options = new DigitalSignatureOptions
                {
                    ContactInfo = signer,
                    Location = location,
                    Reason = reason,
                    Rectangle = new XRect(locationFinalSign.Location.X, locationFinalSign.Location.Y, 150d, 100d),
                    AppearanceHandler = new SignatureAppearanceHandler(signTargetPath, signer, location, _font),
                    AppName = _appName,
                    PageIndex = pageSign_index
                };
            }
            catch
            {

            }
        }

        void InsertSignatureWordMarker(out DigitalSignatureOptions? options, SignLocationInfo signInfo, SigCtl sigCtl, PdfDocument pdfDoc, ref string signTargetPath, string signer, string location, string reason)
        {
            options = null;

            try
            {
                if (signInfo.Page <= 0 || pdfDoc == null) return;

                PdfPage pdfPage = pdfDoc.Pages[signInfo.Page - 1];

                double pageHeight = pdfPage.Height.Point;
                double pdfX = signInfo.Left;
                double pdfY = pageHeight - signInfo.Top;

                double width = 150;
                double height = 100;

                XRect signatureRect = new XRect(pdfX, pdfY - height, width, height);

                ActivateWacom(sigCtl, ref signTargetPath, signer, reason);

                options = new DigitalSignatureOptions
                {
                    ContactInfo = signer,
                    Location = location,
                    Reason = reason,
                    Rectangle = signatureRect,
                    AppearanceHandler = new SignatureAppearanceHandler(signTargetPath, signer, location, _font),
                    AppName = _appName,
                    PageIndex = signInfo.Page - 1
                };
            }
            catch
            {
            }
        }

    }
}