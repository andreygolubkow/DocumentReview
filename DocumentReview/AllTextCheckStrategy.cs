using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    /// <summary>
    /// Общая проверка шрифта, цвета шрифта, интервалов.
    /// </summary>
    public class AllTextCheckStrategy : ICheckStrategy
    {
        private readonly Author _author;

        public AllTextCheckStrategy(Author author)
        {
            _author = author ?? throw new ArgumentException("Укажите автора комментариев.");
        }

        public string FontName { get; set; }
        public string[] FontSizes { get; set; } = {"28", "24"};

        public void DoCheck(WordprocessingDocument document)
        {
            //Получить текущие стили
            var defaultStyles = document.GetDefaultStyles();
            var allStyles = document.GetStyles();

            //Собрать все параграфы
            var paragraphs = document.MainDocumentPart.Document.Body.Descendants<Paragraph>();
            foreach (var paragraph in paragraphs)
            {
                foreach (var run in paragraph.Descendants<Run>())
                {
                    var comment = new StringBuilder();
                    var runProperties = run.Descendants<RunProperties>();

                    bool fontSizeErr = false;
                    foreach (var rp in runProperties)
                    {
                        if (!fontSizeErr && rp.FontSize != null)
                        {
                            if (!FontSizes.Contains(rp.FontSize.Val.ToString()))
                            {
                                comment.AppendLine(Resources.CheckFontSize);
                                fontSizeErr = true;
                            }
                        }
                        else if (!fontSizeErr && rp.RunStyle != null)
                        {
                            var currentStyle = allStyles.FindById(rp.RunStyle.Val);
                            if (currentStyle == null)
                            {
                                continue;
                            }

                            var fontSize = FindFirstFontSize(currentStyle, allStyles);
                            if (fontSize != null &&  !FontSizes.Contains(fontSize))
                            {
                                comment.AppendLine(Resources.CheckFontSize);
                                fontSizeErr = true;
                            }
                            else
                            {
                                fontSizeErr = CheckDefaultFontSize(defaultStyles, comment);
                            }
                        }
                        else if (!fontSizeErr)
                        {
                            fontSizeErr = CheckDefaultFontSize(defaultStyles, comment);
                        }
                    }
                    //Проверяем шрифт

                    //Добавим коммент
                    if (comment.Length > 0)
                    {
                        document.AddCommentNearElement(run, _author, comment.ToString());
                    }
                }
            }
            
            //Проверить шрифт, размер, цвет
            //Проверить интервалы
        }

        private bool CheckDefaultFontSize(DocDefaults defaultStyles, StringBuilder comment)
        {
            if (defaultStyles.RunPropertiesDefault?.RunPropertiesBaseStyle?.FontSize == null ||
                !FontSizes.Contains(defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize.Val.ToString()))
                return false;
            comment.AppendLine(Resources.CheckFontSize);
            return true;
        }

        private string FindFirstFontSize(Style style, Styles styles)
        {
            while (true)
            {
                if (style.StyleRunProperties?.FontSize != null)
                {
                    return style.StyleRunProperties.FontSize.Val;
                }

                if (style.BasedOn != null)
                {
                    var s = styles.FindById(style.BasedOn.Val);
                    if (s != null)
                    {
                        style = s;
                        continue;
                    }

                    return null;
                }

                return null;
            }
        }
    }

    
}