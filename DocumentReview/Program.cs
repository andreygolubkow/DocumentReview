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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;

namespace DocumentReview
{
    static class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 1)
            {
                Check(args[0]);
            }
            //Check(@"C:\Users\andreygolubkow\Desktop\Лабы\Проверка\Zurbaev_Vlasova_Sadalova_588-2.docx");
            //Check(@"C:\Users\andreygolubkow\Desktop\Лабы\Проверка\lab_5.docx");
            Check(@"D:\5_laba.docx");
        }

        

        
        static void Check(string fileName)
        {
            using (var document =
        WordprocessingDocument.Open(fileName, true))
            {
                /*var styles = AllTextCheckStrategy.ExtractStylesPart(document, false);
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

                PageViewCheckers.CheckFont(document, new Fonts()
                {
                    Ascii = "Times New Roman",
                    ComplexScript = "Times New Roman",
                    HighAnsi = "Times New Roman"
                },
                    new FontParams()
                    {
                        Size = "28"
                    });
                    */

                var author = new Author("Andrey Golubkov", "GAA");


                var strategy1 = new AllTextCheckStrategy(author);
                strategy1.DoCheck(document);

                document.Save();
            }
        }

        
        

    }
}
