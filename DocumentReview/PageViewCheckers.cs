using System.Collections.Generic;
using System.Runtime.Versioning;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public static class PageViewCheckers
    {
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
                    CommentTools.AddCommentToParagraph(document, paragraph, "Andrey", "GAA", comment.ToString());
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
                    CommentTools.AddCommentToParagraph(document, paragraph, "Andrey", "GAA", comment.ToString());
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