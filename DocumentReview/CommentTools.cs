using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public class CommentTools
    {
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
                    var i = int.Parse(id);
                    id = (++i).ToString();
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
            { Id = id };
            var lastRun = paragraph.Descendants<Run>();
            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = paragraph.InsertAfter(commentEnd, lastRun.Last());

            // Compose a run with CommentReference and insert it.
            paragraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }

    }
}