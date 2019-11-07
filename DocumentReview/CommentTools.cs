using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public static class CommentTools
    {
        public static void AddCommentNearElement( this WordprocessingDocument document, OpenXmlElement element,
            Author author, string comment, string fontName = "Times New Roman", string fontSize = "28")
        {
            Comments comments = null;
            int id = 0;
            //var lastRun = paragraph.Descendants<Run>().ToList().LastOrDefault();
            if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any())
            {
                comments =
                    document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                if (comments.HasChildren)
                {
                    // Obtain an unused ID.
                    id = comments.Descendants<Comment>().Select(e => Int32.Parse(e.Id.Value)).Max();
                    id++;
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
            Paragraph p = new Paragraph(new Run(new Text($"{comment} ID: {id}"), new RunProperties()
            {
                FontSize = new FontSize()
                {
                    Val = fontSize
                },
                RunFonts = new RunFonts()
                {
                    Ascii = fontName
                }
            }));

            Comment cmt =
                new Comment()
                {
                    Id = id.ToString(),
                    Author = author.Name,
                    Initials = author.Initials,
                    Date = DateTime.Now
                };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            // Specify the text range for the Comment. 
            // Insert the new CommentRangeStart before the first run of paragraph.
            element.InsertBeforeSelf(new CommentRangeStart()
                { Id = id.ToString() });


            var commentEnd = new CommentRangeEnd()
            { Id = id.ToString() };
            
            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = element.InsertAfterSelf(commentEnd);

            // Compose a run with CommentReference and insert it.
            commentEnd.InsertAfterSelf(new Run(new CommentReference() { Id = id.ToString() }));
        }

    }
}