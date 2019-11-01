using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;

namespace DocumentReview
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var doc = WordprocessingDocument.Open(@"D:\test.docx", false))
            {
                foreach (var el in doc.MainDocumentPart.Document.Body.Elements().OfType<Paragraph>())
                {
                    Console.WriteLine(el.InnerText);
                }
            }
        }
    }
}
