using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    class Program
    {
        static void Main(string[] args)
        {
            //PrintInfo(@"D:\test.docx");
            Check(@"D:\test1.docx");

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
