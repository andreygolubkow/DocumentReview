using System.Collections.Generic;
using System.Runtime.Versioning;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public static class PageViewCheckers
    {
        public static void CheckFont(WordprocessingDocument document, Fonts fonts, FontParams fontParams = null)
        {
            foreach (var p in document.MainDocumentPart.Document.Body.Descendants<Paragraph>())
            {
                var comment = new StringBuilder();

                var runProperties = p.Descendants<RunProperties>();

                foreach (var rp in runProperties)
                {
                    if (rp.RunFonts != null)
                    {
                        if ((fonts.Ascii!= null && rp.RunFonts.Ascii != null &&  rp.RunFonts.Ascii.Value != fonts.Ascii)||
                            (fonts.ComplexScript != null && rp.RunFonts.ComplexScript != null && rp.RunFonts.ComplexScript != fonts.ComplexScript) ||
                            (fonts.HighAnsi != null && rp.RunFonts.HighAnsi != null && rp.RunFonts.HighAnsi != fonts.HighAnsi))
                        {
                            comment.AppendLine(Resources.CheckFont);
                        }
                    }

                    //TOOD: Нужно сделать функцию получения стиля по умолчанию
                    if (fontParams?.Size != null &&
                        rp.FontSize != null && rp.FontSize.Val != fontParams.Size)
                    {
                        comment.AppendLine(Resources.CheckFontSize);
                    }

                    /*if (fontParams?.Color != null &&
                        rp.Color != null && rp.Color.Val != fontParams.Color)
                    {
                        comment.AppendLine(Resources.CheckFontColor);
                    }*/

                }
                if (comment.Length > 0)
                {
                    //CommentTools.AddCommentNearElement(document, p, "Andrey", "GAA", comment.ToString());
                }
            }
        }

        public static void CheckMargins(WordprocessingDocument document, PageMargin margin)
        {
            var margins = GetMargins(document);
            foreach (var m in margins)
            {
                var paragraph = PageStructureTools.FindNearParagraphWithRun(m);
                if (paragraph == null)
                {
                    paragraph = new Paragraph(new Run(new Text("")));
                    m.Parent.InsertBeforeSelf(paragraph);
                }

                var comment = new StringBuilder();

                if (margin.Left != m.Left)
                {
                    comment.AppendLine(Resources.CheckLeftMargin);
                }
                if (margin.Right != m.Right)
                {
                    comment.AppendLine(Resources.CheckRightMargin);
                }
                if (margin.Top != m.Top)
                {
                    comment.AppendLine(Resources.CheckTopMargin);
                }
                if (margin.Bottom != m.Bottom)
                {
                    comment.AppendLine(Resources.CheckBottomMargin);
                }

                if (comment.Length > 0)
                {
                    //CommentTools.AddCommentToElement(document, paragraph, "Andrey", "GAA", comment.ToString());
                }
            }
        }

        public static void CheckPageSizes(WordprocessingDocument document, PageSize size)
        {
            var sizes = GetPageSizes(document);
            foreach (var s in sizes)
            {
                var paragraph = PageStructureTools.FindNearParagraphWithRun(s);
                if (paragraph == null)
                {
                    paragraph = new Paragraph(new Run(new Text("")));
                    s.Parent.InsertBeforeSelf(paragraph);
                }
                var comment = new StringBuilder();

                if (size.Height != s.Height)
                {
                    comment.AppendLine(Resources.CheckPageHeight);
                }
                if (size.Width != s.Width)
                {
                    comment.AppendLine(Resources.CheckPageWidth);
                }
                if (s.Orient!= null && size.Orient != null && s.Orient.HasValue && size.Orient.HasValue && (s.Orient.Value != size.Orient.Value))
                {
                    comment.AppendLine(Resources.CheckPageOrientation);
                }

                if (s.Code != null && size.Code != null &&  s.Code.HasValue && size.Code.HasValue && (s.Code.Value != size.Code.Value))
                {
                    comment.AppendLine(Resources.CheckPageCode);
                }


                if (comment.Length > 0)
                {
                    //CommentTools.AddCommentNearElement(document, paragraph, "Andrey", "GAA", comment.ToString());
                }
            }
        }


        static IEnumerable<PageSize> GetPageSizes(WordprocessingDocument document)
        {
            return document.MainDocumentPart.Document.Descendants<PageSize>();
        }

        /// <summary>
        /// Получает поля документа.
        /// </summary>
        /// <param name="document">Открытый документ.</param>
        /// <returns>Список полей.</returns
        static IEnumerable<PageMargin> GetMargins(WordprocessingDocument document)
        {
            return document.MainDocumentPart.Document.Descendants<PageMargin>();
        }

    }
}