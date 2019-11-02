using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public class PageStructureTools
    {

        public static Paragraph FindNearParagraphWithRun(OpenXmlElement element)
        {
            var e = element;
            while (e != null)
            {
                //Если сейчас это параграф берем его
                if (e is Paragraph paragraph && e.Descendants<Run>().Any())
                {
                    return paragraph;
                }

                // Ищем ближайший параграф
                var firstParagraph = e.ElementsBefore().OfType<Paragraph>().LastOrDefault();
                if (firstParagraph != null && firstParagraph.Descendants<Run>().Any())
                {
                    return firstParagraph;
                }


                e = e.Parent;
            }
            return null;
        }



    }
}