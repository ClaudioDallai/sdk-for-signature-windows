/******************************************************* 

  MainWindow.xaml.cs
  
  Displays a form with a button to start signature capture
  The captured signature is encoded in an image file which is displayed on the form
  
  Copyright (c) 2020 Wacom Co. Ltd. All rights reserved.
  
********************************************************/

// Wacom Ink sdk
using FlSigCaptLib;
using FLSIGCTLLib;

// C# framework
using System;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

// PDF manipulation
using PdfPig = UglyToad.PdfPig;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using System.Windows.Controls;
using System.Data.Odbc;


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
        private void btnSign_Click(object sender, RoutedEventArgs e)
        {
            String signTargetPath = "";

            print("btnSign was pressed");
            SigCtl sigCtl = new SigCtl();
            sigCtl.Licence = _sdkLicenseKey;
            //DynamicCapture dc = new DynamicCaptureClass();
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

                signTargetPath = "C:\\temp\\sig" + dateStr + ".png";
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
                    MessageBox.Show(ex.Message);
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
                    default: print("Unexpected error code "); break;
                }
            }


            if (!File.Exists(signTargetPath))
            {
                return;
            }

            try
            {

                // All these vars need to be managed through MVC.
                string marker = "{{{SIGN_HERE}}}";
                string inputPath = @"C:\Users\visio\Desktop\LoremIpsumRight.pdf";
                //string inputPath = @"C:\Users\visio\Desktop\LoremIpsumLeft.pdf";
                string imgPath = signTargetPath;
                string outputPath = @"C:\Users\visio\Desktop\Result.pdf";

                int signTargetPage = -1;
                double markerFoundX = 0, markerFoundY = 0;

                using (PdfPig.PdfDocument document = PdfPig.PdfDocument.Open(inputPath))
                {
                    foreach (var page in document.GetPages())
                    {
                        foreach (var word in page.GetWords())
                        {
                            if ((word.Text.Contains(marker)))
                            {
                                signTargetPage = page.Number;
                                markerFoundX = word.BoundingBox.Left;
                                markerFoundY = word.BoundingBox.Top;
                                break;
                            }
                        }
                    }

                    // PDFPig counts from page 1.
                    if (signTargetPage <= 0)
                    {
                        return;
                    }


                }

                using (PdfSharp.Pdf.PdfDocument pdf = PdfSharp.Pdf.IO.PdfReader.Open(inputPath, PdfDocumentOpenMode.Modify))
                {

                    // PdfPig counts from 1, PdfSharp from 0. 
                    if (signTargetPage <= 0 || signTargetPage > pdf.Pages.Count)
                    {
                        return;
                    }

                    PdfSharp.Pdf.PdfPage pageToSign = pdf.Pages[signTargetPage - 1];

                    PdfSharp.Drawing.XGraphics gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(pageToSign);
                    PdfSharp.Drawing.XImage signatureImage = PdfSharp.Drawing.XImage.FromFile(signTargetPath);

                    // PdfPig counts Y from top, PDFsharp from bottom
                    // X uses same coordinates.
                    double posX = markerFoundX;
                    double posY = pageToSign.Height.Point - markerFoundY - signatureImage.PixelHeight * 0.5f;

                    gfx.DrawImage(signatureImage, posX, posY, signatureImage.PixelWidth, signatureImage.PixelHeight);

                    pdf.Save(outputPath);
                }



                #region No more used iTextSharp (incompatible license)

                //PdfReader reader = new PdfReader(@"C:\Users\visio\Desktop\LoremIpsumRight.pdf");
                ////PdfReader reader = new PdfReader(@"C:\Users\visio\Desktop\LoremIpsumLeft.pdf");
                //using (FileStream fs = new FileStream(@"C:\Users\visio\Desktop\Result.pdf", FileMode.Create))
                //using (PdfStamper stamper = new PdfStamper(reader, fs))
                //{
                //    int pageNum = 1;

                //    var strategy = new TextLocationStrategyChar(marker);
                //    var processor = new PdfContentStreamProcessor(strategy);

                //    var pageDic = reader.GetPageN(pageNum);
                //    var resourcesDic = pageDic.GetAsDict(PdfName.RESOURCES);
                //    processor.ProcessContent(ContentByteUtils.GetContentBytesForPage(reader, pageNum), resourcesDic);

                //    strategy.FinalizeSearch();

                //    if (strategy.MarkerPositions.Count > 0)
                //    {
                //        PdfContentByte canvas = stamper.GetOverContent(pageNum);
                //        iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(signTargetPath);

                //        foreach (var pos in strategy.MarkerPositions)
                //        {
                //            // Move of half height and width (it depends on marker position found) to center the image (high left corner) on the marker.
                //            // To check X we consider an offset of 70% of the PDF itself.
                //            // Values are kinda hadcoded for now, and they depend on the size of the sign canvas.
                //            float resultAbsX = pos[Vector.I1] - img.Width * 0.25f;
                //            float resultAbsY = pos[Vector.I2] - img.Height * 0.5f;

                //            if (pos[Vector.I1] > reader.GetPageSize(pageNum).Width * 0.7f)
                //            {
                //                resultAbsX = pos[Vector.I1] - img.Width * 0.5f;
                //            }

                //            img.SetAbsolutePosition(resultAbsX, resultAbsY);
                //            canvas.AddImage(img);
                //        }
                //    }
                //    else
                //    {
                //        Console.WriteLine("Marker non trovato");
                //    }
                //}

                #endregion



            }
            catch
            {
                Console.WriteLine();

            }



            // ------------------------------------------------------------------------------------------------------ //



            #region Wacom SignPro pdf API call tests (premium license is needed)

            //// Invocazione di Wacom SignPRO API. Richiede licenza premium.
            //// Leggi il contenuto del file JSON
            //string jsonPath = @"C:\Users\visio\Downloads\api-demo-v4\scripts\API Command Demo\demo api autosave.txt";
            //string json = File.ReadAllText(jsonPath);

            //// Converti in Base64
            //string base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(json));

            //// Esegui Sign Pro PDF con il parametro API
            //Process.Start(
            //    @"C:\Program Files (x86)\Wacom sign pro PDF\Sign Pro PDF.exe",
            //    $"-api signpro:{base64}"
            //);

            #endregion


        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void print(string txt)
        {
            txtInfo.Text += txt + "\r\n";
        }
    }
}
