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
            PrintInfo(@"D:\WithRows.docx");

        }

        static void PrintInfo(string fileName)
        {
            using (var document =
        WordprocessingDocument.Open(fileName, true))
            {
                CheckMargins(document, new PageMargin()
                {
                    Top = 0,
                    Bottom = 0,
                    Left = 0,
                    Right = 0
                });


                document.Save();
                
            }
        }

        static Paragraph FindNearParagraphWithRun(OpenXmlElement element)
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

        static void CheckMargins(WordprocessingDocument document, PageMargin margin)
        {
            var margins = GetMargins(document);
            foreach (var m in margins)
            {
                var paragraph = FindNearParagraphWithRun(m);
                if (paragraph == null)
                {
                    paragraph = new Paragraph(new Run(new Text("ИНФОРМАЦИЯ")));
                    m.Parent.InsertBeforeSelf(paragraph);
                }

                var comment = new StringBuilder();

                if (margin.Left != m.Left)
                {
                    comment.AppendLine("Проверьте левое поле.");
                }
                if (margin.Right != m.Right )
                {
                    comment.AppendLine("Проверьте правое поле.");
                }
                if (margin.Top != m.Top)
                {
                    comment.AppendLine("Проверьте верхнее поле.");
                }
                if (margin.Bottom != m.Bottom)
                {
                    comment.AppendLine("Проверьте нижнее поле.");
                }

                if (comment.Length > 0)
                {
                    AddCommentToParagraph(document, paragraph, "Andrey", "GAA", comment.ToString());
                }
            }
        }

        public static void AddCommentToParagraph(WordprocessingDocument document, OpenXmlElement paragraph,
            string author, string initials, string comment)
        {
                Comments comments = null;
                string id = "0";

                if (document.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0)
                {
                    comments =
                        document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                    if (comments.HasChildren)
                    {
                        // Obtain an unused ID.
                        id = comments.Descendants<Comment>().Select(e => e.Id.Value).Max();
                    }
                }
                else
                {
                    WordprocessingCommentsPart commentPart =
                        document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    comments = commentPart.Comments;
                }

                // Compose a new Comment and add it to the Comments part.
                Paragraph p = new Paragraph(new Run(new Text($"{comment} ID: {id}")));
                Comment cmt =
                    new Comment()
                    {
                        Id = id,
                        Author = author,
                        Initials = initials,
                        Date = DateTime.Now
                    };
                cmt.AppendChild(p);
                comments.AppendChild(cmt);
                comments.Save();

                // Specify the text range for the Comment. 
                // Insert the new CommentRangeStart before the first run of paragraph.
                paragraph.InsertBefore(new CommentRangeStart()
                { Id = id }, paragraph.GetFirstChild<Run>());

                var commentEnd = new CommentRangeEnd()
                    {Id = id};
                var lastRun = paragraph.Descendants<Run>();
                // Insert the new CommentRangeEnd after last run of paragraph.
                var cmtEnd = paragraph.InsertAfter(commentEnd, lastRun.Last());

                // Compose a run with CommentReference and insert it.
                paragraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }

        /// <summary>
        /// Получает поля документа.
        /// </summary>
        /// <param name="document">Открытый документ.</param>
        /// <returns>Список полей.</returns>
        static IEnumerable<PageMargin> GetMargins(WordprocessingDocument document)
        {
            return document.MainDocumentPart.Document.Descendants<PageMargin>();
        }

        static IEnumerable<PageSize> GetPageSizes(WordprocessingDocument document)
        {
            return document.MainDocumentPart.Document.Descendants<PageSize>();
        }


    }
}
