using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
using Document = Aspose.Words.Document;
using Font = System.Drawing.Font;

namespace DocumentReview
{
    static class Program
    {
        static void Main(string[] args)
        {
            //PrintInfo(@"D:\test.docx");
            //Check(@"D:\test1.docx");
            DocxToPng(@"D:\test1.docx",2);
            Aspose.Words.License l = new License();
            l.SetLicense("demo");
            //var bw = BitmapToBlackWhite2(new Bitmap(bmp));

            //var imp = Image.FromHbitmap(bw.GetHbitmap());
            //bmp.Save(@"D:\BW.bmp", ImageFormat.Bmp);

           // DrawBorders(new Bitmap(@"D:\pdf2.bmp")).Save(@"D:\2.bmp");
           // DrawBorders(new Bitmap(@"D:\pdf3.bmp")).Save(@"D:\3.bmp"); ;
           // DrawBorders(new Bitmap(@"D:\pdf4.bmp")).Save(@"D:\4.bmp"); ;
           // DrawBorders(new Bitmap(@"D:\pdf5.bmp")).Save(@"D:\5.bmp"); ;
        }

        public static Bitmap DrawBorders(Bitmap bitmap)
        {
            var left = FindLeftBorder(bitmap);
            left--;
            //Делаем линию
            DrawHorizontalLine(bitmap,left);

            var right = FindRightBorder(bitmap);
            right++;
            //Делаем линию
            DrawHorizontalLine(bitmap, right);

            DrawText(bitmap, $"{left}", new Point(10,10));


            DrawText(bitmap, $"{bitmap.Width - right}", new Point(right, 10));

            return bitmap;
        }

        public static void DrawText(Bitmap bitmap, string str, Point point)
        {
            Graphics g = Graphics.FromImage(bitmap);

            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.DrawString(str, new Font("Tahoma", 8), Brushes.Black,point);

            g.Flush();
        }

        public static void DrawHorizontalLine(Bitmap bitmap, int x)
        {
            for (int i = 0; i < bitmap.Height; i++)
            {
                bitmap.SetPixel(x, i, System.Drawing.Color.Chartreuse);
            }
        }

        public static int FindLeftBorder(Bitmap bitmap)
        {
            for (int x = 0; x < bitmap.Width; x++)
            {
                for (int y = 0; y < bitmap.Height; y++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.R + pixel.G + pixel.B != 255+255+255)
                    {
                        return x;
                    }
                }
            }

            return 1;
        }

        public static int FindRightBorder(Bitmap bitmap)
        {
            for (int x = bitmap.Width-1; x >= 0; x--)
            {
                for (int y = 0; y < bitmap.Height; y++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.R + pixel.G + pixel.B != 255 + 255 + 255)
                    {
                        return x;
                    }
                }
            }

            return 1;
        }


        static void DocxToPng(string filename, int page)
        {
            // Open the document.
            Document doc = new Document(filename);

            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.Resolution = 160;

            // Save each page of the document as Png.
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(string.Format(@"D:\test" + i + "SaveAsPNG out.Png", i), options);
            }
        }

        
        static void Check(string fileName)
        {
            using (var document =
        WordprocessingDocument.Open(fileName, true))
            {
                PageViewCheckers.CheckMargins(document, new PageMargin()
                {
                    Top = 1134,
                    Bottom = 1134,
                    Left = 1701,
                    Right = 567
                });
                PageViewCheckers.CheckPageSizes(document, new PageSize()
                {
                    Width = 11906,
                    Height = 16838
                });

                document.Save();
            }
        }

        
        

    }
}
