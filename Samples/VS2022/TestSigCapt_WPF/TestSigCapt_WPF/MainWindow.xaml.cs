/******************************************************* 

  MainWindow.xaml.cs
  
  Displays a form with a button to start signature capture
  The captured signature is encoded in an image file which is displayed on the form
  
  Copyright (c) 2020 Wacom Co. Ltd. All rights reserved.
  
********************************************************/

// Wacom Ink sdk
using FlSigCaptLib;
using FLSIGCTLLib;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf.security;
using Microsoft.Win32;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Pkcs;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
// C# framework
using System;
using System.Collections.Generic;
using System.Data.Odbc;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Imaging;
// PDF manipulation
using PdfPig = UglyToad.PdfPig;



namespace TestSigCapt_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string _sdkLicenseKey = "eyJhbGciOiJSUzUxMiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImV4cCI6MjE0NzQ4MzY0NywiaWF0IjoxNTYwOTUwMjcyLCJyaWdodHMiOlsiU0lHX1NES19DT1JFIiwiU0lHQ0FQVFhfQUNDRVNTIl0sImRldmljZXMiOlsiV0FDT01fQU5ZIl0sInR5cGUiOiJwcm9kIiwibGljX25hbWUiOiJTaWduYXR1cmUgU0RLIiwid2Fjb21faWQiOiI3YmM5Y2IxYWIxMGE0NmUxODI2N2E5MTJkYTA2ZTI3NiIsImxpY191aWQiOiJiODUyM2ViYi0xOGI3LTQ3OGEtYTlkZS04NDlmZTIyNmIwMDIiLCJhcHBzX3dpbmRvd3MiOltdLCJhcHBzX2lvcyI6W10sImFwcHNfYW5kcm9pZCI6W10sIm1hY2hpbmVfaWRzIjpbXX0.ONy3iYQ7lC6rQhou7rz4iJT_OJ20087gWz7GtCgYX3uNtKjmnEaNuP3QkjgxOK_vgOrTdwzD-nm-ysiTDs2GcPlOdUPErSp_bcX8kFBZVmGLyJtmeInAW6HuSp2-57ngoGFivTH_l1kkQ1KMvzDKHJbRglsPpd4nVHhx9WkvqczXyogldygvl0LRidyPOsS5H2GYmaPiyIp9In6meqeNQ1n9zkxSHo7B11mp_WXJXl0k1pek7py8XYCedCNW5qnLi4UCNlfTd6Mk9qz31arsiWsesPeR9PN121LBJtiPi023yQU8mgb9piw_a-ccciviJuNsEuRDN3sGnqONG3dMSA";
        public MainWindow()
        {
            InitializeComponent();
        }

        public void ActivateWacom(SigCtl sigCtl, ref string signTargetPath)
        {
            DynamicCapture dc = new FlSigCaptLib.DynamicCapture();
            DynamicCaptureResult res = dc.Capture(sigCtl, "Mario Rossi", "Presentazione Esempio 1", null, null);
            if (res == DynamicCaptureResult.DynCaptOK)
            {
                print("signature captured successfully");
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
                print("Outputting to file " + signTargetPath);
                try
                {
                    //print("Saving signature to file " + filename);
                    // Need to understand if DimensionX and DimensionY alter the image metadata. These values (200, 150) were already here in the example.
                    sigObj.RenderBitmap(signTargetPath, 150, 100, "image/png", 0.5f, 0xff0000, 0xffffff, 10.0f, 10.0f, RBFlags.RenderOutputFilename | RBFlags.RenderColor32BPP | RBFlags.RenderEncodeData);

                    print("Loading image from " + signTargetPath);
                    BitmapImage src = new BitmapImage();
                    src.BeginInit();
                    src.UriSource = new Uri(signTargetPath, UriKind.Absolute);
                    src.EndInit();

                    imgSig.Source = src;
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }

            }
            else
            {

                // IF WACOM STU IS NOT ATTACHED: Returns message that license is invalid!
                print("Signature capture error res=" + (int)res + "  ( " + res + " )");
                switch (res)
                {
                    case DynamicCaptureResult.DynCaptCancel: print("signature cancelled"); break;
                    case DynamicCaptureResult.DynCaptError: print("no capture service available"); break;
                    case DynamicCaptureResult.DynCaptPadError: print("signing device error"); break;
                    case DynamicCaptureResult.DynCaptNotLicensed: print("license error or device is not connected"); break;
                    default: print("Unexpected error code "); break;
                }
            }
        }

        private void btnSign_Click(object sender, RoutedEventArgs e)
        {


            #region Signature INK sdk and PDF Manipulation


            String signTargetPath = "";

            print("btnSign was pressed");
            SigCtl sigCtl = new SigCtl();
            sigCtl.Licence = _sdkLicenseKey;
            //DynamicCapture dc = new DynamicCaptureClass();




            string inputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\LoremIpsumMulti.pdf";
            string outputPdf = @"C:\Users\visio\Desktop\PDF_ManipulationTests\Result.pdf";
            string marker = "{{{SIGN_HERE}}}";


            #region MIT Libraries


            #region PdfPig


            // Use PdfPig to interrogate the PDF and rtreive infos about sign slot locations
            int signTargetPage = -1;

            Dictionary<int, List<System.Windows.Vector>> markerPositions = new Dictionary<int, List<System.Windows.Vector>>();

            using (PdfPig.PdfDocument document = PdfPig.PdfDocument.Open(inputPdf))
            {
                foreach (var page in document.GetPages())
                {
                    foreach (var word in page.GetWords())
                    {
                        if (word.Text.Contains(marker))
                        {
                            signTargetPage = page.Number;

                            if (markerPositions.ContainsKey(signTargetPage))
                            {
                                markerPositions[signTargetPage].Add(new System.Windows.Vector(word.BoundingBox.Left, word.BoundingBox.Top));
                            }
                            else
                            {
                                markerPositions.Add(signTargetPage, new List<System.Windows.Vector> { new System.Windows.Vector(word.BoundingBox.Left, word.BoundingBox.Top) });
                            }
                        }
                    }
                }

                // PDFPig counts from page 1.
                HashSet<int> invalidKeys = new HashSet<int>();
                foreach (KeyValuePair<int, List<System.Windows.Vector>> pair in markerPositions)
                {
                    if (pair.Key <= 0 || pair.Key > document.NumberOfPages)
                    {
                        invalidKeys.Add(pair.Key);
                    }
                }

                foreach (int invKey in invalidKeys)
                {
                    if (markerPositions.ContainsKey(invKey))
                    {
                        markerPositions.Remove(invKey);
                    }
                }
            }


            #endregion


            #region PDFSharp

            //    using (PdfSharp.Pdf.PdfDocument pdf = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify))
            //    {

            //        // Need to add all null checks. Even if we're inside Try.
            //        foreach (KeyValuePair<int, List<Vector>> pair in markerPositions)
            //        {
            //            PdfSharp.Pdf.PdfPage pageToSign = pdf.Pages[pair.Key - 1];

            //            PdfSharp.Drawing.XGraphics gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(pageToSign);
            //            PdfSharp.Drawing.XImage signatureImage = PdfSharp.Drawing.XImage.FromFile(signTargetPath);

            //            // PdfPig counts Y from top, PDFsharp from bottom.
            //            // X uses same coordinates.
            //            foreach(Vector markerInstance in pair.Value)
            //            {
            //                double posX = markerInstance.X;
            //                double posY = pageToSign.Height.Point - markerInstance.Y - signatureImage.PixelHeight * 0.5f;

            //                gfx.DrawImage(signatureImage, posX, posY, signatureImage.PixelWidth, signatureImage.PixelHeight);
            //            }
            //        }

            //        // Encryption.
            //        pdf.SecurityHandler.SetEncryptionToV2With128Bits(); // Different alghoritms are available.
            //        pdf.SecuritySettings.UserPassword = "1234";
            //        pdf.SecuritySettings.OwnerPassword = "Admin";

            //        // No permit.
            //        pdf.SecuritySettings.PermitModifyDocument = false;
            //        pdf.SecuritySettings.PermitAssembleDocument = false;
            //        pdf.SecuritySettings.PermitFormsFill = false;
            //        pdf.SecuritySettings.PermitAnnotations = false;

            //        // Permit.
            //        pdf.SecuritySettings.PermitExtractContent = true;
            //        pdf.SecuritySettings.PermitPrint = true;
            //        pdf.SecuritySettings.PermitFullQualityPrint = true;

            //        pdf.Save(outputPath);

            //        // ATTENTION: Wacom SignatureMiniscope cannot open encrypted files. Also, copy-paste of biomethric signature does not work (it works on the real Wacoom SignatureScope full version).
            //        // It does not work even when creating a pdf from world copy-pasting an image in it.
            //        // A .png intead could be loaded, so in theory we can save both pdf (encrypted) and the sign's png in a protected ZIP. The use MiniScope on the png inside given necessary password.
            //        // https://developer-support.wacom.com/hc/en-us/articles/9354488252311-How-do-I-copy-paste-a-signature-from-sign-pro-PDF-into-another-Wacom-application

            //        // !!! PdfSharp supports signature: Check complete documentation https://docs.pdfsharp.net/PDFsharp/Overview/About.html !!!
            //        // Maybe an authorized Sign with certificate is really needed to provide legal value:https://github.com/empira/PDFsharp/tree/master/src/foundation/src/PDFsharp/src/PdfSharp.Cryptography
            //        // PDFSharp.Cryptography needs .NET v6.0
            //    }
            //}
            //catch
            //{
            //    Console.WriteLine();
            //}

            #endregion


            #endregion


            #region iTextSharp


            // Hardoded as example

            // Not necessary if using Windows Certificate Store
            //string pfxPath = @"C:\temp\firma-demo.pfx"; 
            //string pfxPassword = "password123";

            string reason = "Firma di prova";
            string location = "Italia";


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
            var cert = certs[0];


            // I campi NotBefore e NotAfter nei certificati sono sempre espressi in UTC secondo lo standard X.509
            DateTime now = DateTime.UtcNow;
            if (now > cert.NotAfter.ToUniversalTime())
            {
                throw new System.Exception($"Certificato con subject '{subjectName}' scaduto.");
            }

            if (!cert.HasPrivateKey)
                throw new System.Exception("Il certificato non contiene la chiave privata.");

            // Converte il certificato .NET in BouncyCastle X509Certificate
            var parser = new Org.BouncyCastle.X509.X509CertificateParser();
            Org.BouncyCastle.X509.X509Certificate bcCert = parser.ReadCertificate(cert.RawData);

            // Estrai la chiave privata RSA da .NET e convertila in BouncyCastle.
            // La chiave RSA è una particolare implementazione di questa coppia di chiavi (RSA è uno degli algoritmi più usati).
            Org.BouncyCastle.Crypto.AsymmetricKeyParameter bcPrivateKey;

            using (var rsa = cert.GetRSAPrivateKey())
            {
                if (rsa == null)
                    throw new System.Exception("Impossibile ottenere la chiave RSA privata.");

                bcPrivateKey = Org.BouncyCastle.Security.DotNetUtilities.GetRsaKeyPair(rsa).Private;
            }

            // Crea la chain BouncyCastle (qui metto solo il certificato principale, aggiungi intermedi se vuoi)
            System.Collections.Generic.IList<Org.BouncyCastle.X509.X509Certificate> chainBC = new System.Collections.Generic.List<Org.BouncyCastle.X509.X509Certificate> { bcCert };

            IExternalSignature externalSignature = new PrivateKeySignature(bcPrivateKey, "SHA-256");

            System.Console.WriteLine("Certificato caricato correttamente e pronto per firmare.");


            // Now we can sign using externalSignature (certificated) and chainBC


            try
            {

                #region Load pfx directly from file (not really ok)


                //// Load the certificate with BouncyCastle from .pfx file
                //Pkcs12Store pk12;
                //using (var fs = new FileStream(pfxPath, FileMode.Open, FileAccess.Read))
                //{
                //    // Use builder to create the store (deprecated factory. CTOR in iText will do)
                //    pk12 = new Pkcs12StoreBuilder().Build();
                //    pk12.Load(fs, pfxPassword.ToCharArray());
                //}

                //// Find the alias that corresponds to a private key entry in the certificate
                //string alias = null;
                //foreach (string tAlias in pk12.Aliases)
                //{
                //    if (pk12.IsKeyEntry(tAlias))
                //    {
                //        alias = tAlias;
                //        break;
                //    }
                //}

                //// Get private key and chain of authentication entities
                //var pk = pk12.GetKey(alias).Key;
                //var chain = pk12.GetCertificateChain(alias);

                //// Create an iTextSharp compatible chain of authentication entities (literally a cast-like)
                //ICollection<Org.BouncyCastle.X509.X509Certificate> chainBC = new List<Org.BouncyCastle.X509.X509Certificate>();
                //foreach (var entry in chain)
                //{
                //    chainBC.Add(entry.Certificate);
                //}

                //// Create the external signature object, specifying the private key and hashing algorithm(SHA-256)
                //IExternalSignature externalSignature = new PrivateKeySignature(pk, "SHA-256");


                #endregion


                string processedPdf = ProcessPDF(inputPdf);

                // To Multi-Sign, we need an incremental-sign-method. Using temp pdf (maybe not necessary but it is advised)
                // Multi-Sign is a difficult topic. In theory a digital Sign validated the entire doc. Maybe an approach "png - png - ... - Legal Sign" could be tested (?)
                string signedFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "signed_working_copy.pdf");
                int iterator = 0;

                // Start by copying pdf to the starting file
                File.Copy(processedPdf, signedFilePath, true);

                // Cycle all markers in every page
                foreach (var page in markerPositions)
                {
                    foreach (var markerFound in page.Value)
                    {
                        // Read PDF as bytes to not cause UnauthorizedAccessException because file is already used by FileStream
                        byte[] pdfBytes = File.ReadAllBytes(signedFilePath);
                        iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(pdfBytes, Encoding.ASCII.GetBytes("Admin"));

                        // Open the file.
                        using (FileStream os = new FileStream(signedFilePath, FileMode.Open, FileAccess.ReadWrite))
                        {
                            // Stamper in Append mode (not FileStream!!!) does not invalidate the PDF
                            // Stamper from iTextSharp is used to create a Signature
                            // Being in append mode, and a legally-sign, PDF modify functions are already kinda disabled
                            PdfStamper stamper = PdfStamper.CreateSignature(reader, os, '\0', null, true);

                            // Setup signature slot extension (kinda hardcoded)
                            float rectWidth = 225f;
                            float rectHeight = 75f;
                            var pageSize = reader.GetPageSize(page.Key);
                            float pageWidth = pageSize.Width;
                            float offset = (markerFound.X > pageWidth / 2) ? 100f : 10f;
                            float rectX = (float)markerFound.X - offset;
                            float rectY = (float)markerFound.Y - 85f;

                            // Configure signature appearance properties (reason, location, visible signature rectangle, etc.)
                            PdfSignatureAppearance appearance = stamper.SignatureAppearance;
                            appearance.Reason = reason;
                            appearance.Location = location;

                            // Add a visible signature slot. Names of different signatures (fieldName MUST be different). Name of the metadata inside the signature itself can be the same instead
                            appearance.SetVisibleSignature(new iTextSharp.text.Rectangle(rectX, rectY, rectX + rectWidth, rectY + rectHeight), page.Key, $"Signature_{iterator}");


                            // This explicitly manage permissions on the PDF result:
                            //public const int NOT_CERTIFIED = 0;
                            //public const int CERTIFIED_NO_CHANGES_ALLOWED = 1;
                            //public const int CERTIFIED_FORM_FILLING = 2;
                            //public const int CERTIFIED_FORM_FILLING_AND_ANNOTATIONS = 3;

                            // Any kind of Certificationlevel must be specified in the first sign (PDF standard)
                            // While CERTIFIED_NO_CHANGES_ALLOWED exists and it is the safest, it invalidates multi-signs
                            // Using CERTIFIED_FORM_FILLING instead, allows for different signature slots, still protecting from unwanted changes
                            // Because other signatures, done in other applications are allowed (even if they miss the specific certificate and so they're marked as suspicious),
                            // another try is to let the pdf be immutable AFTER all our signs are done
                            if (iterator == 0)
                            {
                                appearance.CertificationLevel = PdfSignatureAppearance.CERTIFIED_FORM_FILLING;
                            }

                            // Appearance object actually needs a PNG to take the image from. We use the one that was taken using Wacom INK SDK
                            // In theory, we can call here Wacom INK SDK to get signature (better and legally-valid method). An example is using ActivateWacom custom method
                            ActivateWacom(sigCtl, ref signTargetPath);
                            if (File.Exists(signTargetPath))
                            {
                                var signatureImg = iTextSharp.text.Image.GetInstance(signTargetPath);
                                appearance.SignatureGraphic = signatureImg;
                                appearance.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.GRAPHIC_AND_DESCRIPTION;
                            }

                            // Apply the sign
                            MakeSignature.SignDetached(appearance, externalSignature, chainBC, null, null, null, 0, CryptoStandard.CADES);
                        }

                        iterator++;
                    }
                }

                // Copy temp file inside output path
                File.Copy(signedFilePath, outputPdf, true);

                if (File.Exists(signedFilePath))
                {
                    File.Delete(signedFilePath);
                }

                Console.WriteLine("PDF was signed!");
            }
            catch
            {
                Console.WriteLine();
            }
            finally
            {
                store.Close();
            }



            #endregion iTextSharp Signature cryptho (deprecated, use iText instead under licensing)


            #endregion Signature INK sdk and PDF Manipulation



            // ------------------------------------------------------------------------------------------------------ //



            #region Wacom SignPro pdf API call tests (premium license is needed)

            //// Invocazione di Wacom SignPRO API. Richiede licenza premium.
            //// Leggi il contenuto del file JSON
            //string jsonPath = @"C:\Users\visio\Downloads\Wacom\api-demo-v4\scripts\API Command Demo\demo api autosave.txt";
            //string json = File.ReadAllText(jsonPath);

            //// Converti in Base64
            //string base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

            //// Esegui Sign Pro PDF con il parametro API
            //Process.Start(
            //    @"C:\Program Files (x86)\Wacom sign pro PDF\Sign Pro PDF.exe",
            //    $"-api signpro:{base64}"
            //);

            #endregion Wacom SignPro pdf API call tests (premium license is needed)


        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void print(string txt)
        {
            txtInfo.Text += txt + "\r\n";
        }

        private string ProcessPDF(string src)
        {
            string preProcessedPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "pdf_with_permissions.pdf");

            using (iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(src))
            using (FileStream fs = new FileStream(preProcessedPath, FileMode.Create))
            {
                PdfStamper stamper = new PdfStamper(reader, fs);

                // Set encryption and permissions
                string userPassword = "test"; // No password to open
                string ownerPassword = "Admin"; // Needed to change permissions

                int permissions =
                    PdfWriter.ALLOW_PRINTING |
                    PdfWriter.ALLOW_FILL_IN;

                stamper.SetEncryption(
                    Encoding.ASCII.GetBytes(userPassword),
                    Encoding.ASCII.GetBytes(ownerPassword),
                    permissions,
                    PdfWriter.ENCRYPTION_AES_128
                );

                stamper.Close();
            }

            return preProcessedPath;
        }

    }
}

