
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
using System.Windows.Media.Imaging;

/*
 Libraries and technologies used:
    -> Wacom INK SDK (free license if only used to sign using Wacom STU);
    -> PDFSharp (pdf signing, MIT);
    -> BouncyCastle (legal signing, MIT);
    -> Using .dll of a 4.8 Framework project using Microsoft.Office.Word.Interop to modify word and create pdf;
 */


namespace TestPDFSharpSignatures
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly string _sdkLicenseKey = "eyJhbGciOiJSUzUxMiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImV4cCI6MjE0NzQ4MzY0NywiaWF0IjoxNTYwOTUwMjcyLCJyaWdodHMiOlsiU0lHX1NES19DT1JFIiwiU0lHQ0FQVFhfQUNDRVNTIl0sImRldmljZXMiOlsiV0FDT01fQU5ZIl0sInR5cGUiOiJwcm9kIiwibGljX25hbWUiOiJTaWduYXR1cmUgU0RLIiwid2Fjb21faWQiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImxpY191aWQiOiJiODUyM2ViYi0xOGI3LTQ3OGEtYTlkZS04NDlmZTIyNmIwMDIiLCJhcHBzX3dpbmRvd3MiOltdLCJhcHBzX2lvcyI6W10sImFwcHNfYW5kcm9pZCI6W10sIm1hY2hpbmVfaWRzIjpbXX0.ONy3iYQ7lC6rQhou7rz4iJT_OJ20087gWz7GtCgYX3uNtKjmnEaNuP3QkjgxOK_vgOrTdwzD-nm-ysiTDs2GcPlOdUPErSp_bcX8kFBZVmGLyJtmeInAW6HuSp2-57ngoGFivTH_l1kkQ1KMvzDKHJbRglsPpd4nVHhx9WkvqczXyogldygvl0LRidyPOsS5H2GYmaPiyIp9In6meqeNQ1n9zkxSHo7B11mp_WXJXl0k1pek7py8XYCedCNW5qnLi4UCNlfTd6Mk9qz31arsiWsesPeR9PN121LBJtiPi023yQU8mgb9piw_a-ccciviJuNsEuRDN3sGnqONG3dMSA";
        private readonly string _font = "Arial";
        private readonly string _appName = " [Sign_It+] -> Wacom STU-430";

        // known text markers. Maybe a Dict is better in a real implementation
        private readonly string _nameMarker = "{{{NOME_COGNOME}}}";
        private readonly string _placeMarker = "{{{LUOGO_DATA}}}";
        private readonly string _signMarker = "{{{FIRMA}}}";

        // Certificate subject name
        private readonly string _subjectName = "FirmaDemo";
        private readonly string _adminPassword = "Visio_Admin";

        // Paths
        private readonly string _inputTemplateWord = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok.docx";
        private readonly string _inputFilledTemplateWord = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok_Filled.docx";
        private readonly string _outputFilledTemplatePathNoSign = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Liberatoria Slide_Provider_Ok_Filled_NoSign.pdf";
        private readonly string _outputPath = @"C:\Users\visio\Desktop\PDF_ManipulationTests\";


        public MainWindow()
        {

            InitializeComponent();

            try
            {
                // Setup
                string[] args = Environment.GetCommandLineArgs();
                String signTargetPath = "";

                //print("btnSign was pressed");
                SigCtl sigCtl = new SigCtl();
                sigCtl.Licence = _sdkLicenseKey;

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


                // Setup module to sign
                if (!File.Exists(_inputTemplateWord))
                {
                    throw new Exception("Input docx template not found");
                }

                WordProcessor.FillTemplate(out Dictionary<string, MarkerFoundLocationinfo> signLocationInfos,
                                           new Dictionary<string, string> { { _nameMarker, signer }, { _signMarker, "" } },
                                           _inputTemplateWord,
                                           _inputFilledTemplateWord);

                WordProcessor.ConvertToPdf(_inputFilledTemplateWord, _outputFilledTemplatePathNoSign);


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


                // PDFSharp has lots of problem in recognizing metadata of acroforms, that's why we use .dll of another project
                PdfDocument document = PdfReader.Open(_outputFilledTemplatePathNoSign, PdfDocumentOpenMode.Modify);
                if (document == null)
                {
                    throw new Exception("Cannot open PDF to sign");
                }

                if (!signLocationInfos.ContainsKey(_signMarker))
                {
                    throw new Exception("Sign marker was not found. No position available to insert sign");
                }

                InsertSignatureWordMarker(out DigitalSignatureOptions? options, ref signTargetPath, signLocationInfos[_signMarker], sigCtl, document, signer, location, reason);

                if (options != null)
                {
                    // Different alghoritms are available. This one is required to work on Edge PDF Viewer (Microsoft issue)
                    document.SecurityHandler.SetEncryption(PdfDefaultEncryption.V2With128Bits);

                    document.SecurityHandler.OwnerPassword = _adminPassword;
                    //document.SecurityHandler.UserPassword = "";

                    document.SecuritySettings.PermitPrint = true;
                    document.SecuritySettings.PermitFullQualityPrint = true;
                    document.SecuritySettings.PermitExtractContent = true;

                    document.SecuritySettings.PermitAssembleDocument = false;
                    document.SecuritySettings.PermitModifyDocument = false;
                    document.SecuritySettings.PermitFormsFill = false;
                    document.SecuritySettings.PermitAnnotations = false;


                    var pdfSignatureHandler = DigitalSignatureHandler.ForDocument(document, new BouncyCastleSigner((validCertFound, chainCerts), PdfMessageDigestType.SHA256), options);
                    if (pdfSignatureHandler == null)
                    {
                        throw new Exception("Sign failed");
                    }

                    string resultPDF = _outputPath + $"{signer.Replace(" ", "")}_{reason.Replace(" ", "")}.pdf";
                    document.Save(resultPDF);
                    //document.Close(); Save already do this
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
            try
            {

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
                        // If Sign is too big like a painting (literally), an exception will be fired:
                        // $exceptio {"Too much data to encode in image"}	System.Runtime.InteropServices.COMException

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
            catch
            {

            }

        }

        void InsertSignatureWordMarker(out DigitalSignatureOptions? options, ref string signTargetPath, MarkerFoundLocationinfo signInfo, SigCtl sigCtl, PdfDocument pdfDoc, string signer, string location, string reason)
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