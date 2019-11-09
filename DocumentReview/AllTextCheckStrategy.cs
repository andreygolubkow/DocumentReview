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

        public string[] RunFontsNames { get; set; } = {"Times New Roman"};
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

                    bool fontSizeErr = !(FontSizes.Length>0);
                    bool runFontsErr = !(RunFontsNames.Length>0);
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

                        
                        
                        if (!runFontsErr && rp.RunFonts != null)
                        {
                            if (( rp.RunFonts.Ascii != null && rp.RunFonts.Ascii.HasValue  && !RunFontsNames.Contains(rp.RunFonts.Ascii.Value)) || 
                                ( rp.RunFonts.ComplexScript != null && rp.RunFonts.ComplexScript.HasValue  
                                                                    && !RunFontsNames.Contains(rp.RunFonts.ComplexScript.Value)) ||
                                (rp.RunFonts.HighAnsi != null && rp.RunFonts.HighAnsi.HasValue && !RunFontsNames.Contains(rp.RunFonts.HighAnsi.Value)) )
                            {
                                comment.AppendLine(Resources.CheckFont);
                                fontSizeErr = true;
                            }
                        }
                        else if (!runFontsErr && rp.RunStyle != null)
                        {
                            var currentStyle = allStyles.FindById(rp.RunStyle.Val);
                            if (currentStyle == null)
                            {
                                continue;
                            }

                            var runFonts = FindFonts(currentStyle, allStyles);
                            if (runFonts != null && 
                                (
                                    (runFonts.Ascii != null && runFonts.Ascii.HasValue && !RunFontsNames.Contains(runFonts.Ascii.Value)) ||
                                    (runFonts.ComplexScript != null && runFonts.ComplexScript.HasValue
                                                                       && !RunFontsNames.Contains(runFonts.ComplexScript.Value))
                                ))
                            {
                                comment.AppendLine(Resources.CheckFont);
                                runFontsErr = true;
                            }
                            else
                            {
                                runFontsErr = CheckDefaultRunFonts(defaultStyles, comment);
                            }
                        }
                        else if (!runFontsErr)
                        {
                            runFontsErr = CheckDefaultRunFonts(defaultStyles, comment);
                        }
                    }

                    //Добавим коммент
                    if (comment.Length > 0)
                    {
                        document.AddCommentNearElement(run, _author, comment.ToString());
                    }
                }
            }
            
            
            //Проверить интервалы
        }

        private RunFonts FindFonts(Style style, Styles styles)
        {
            var cs = FindFirstCsFont(style, styles);
            var ascii = FindFirstAsciiFont(style, styles);

            var fonts = new RunFonts();
            if (cs != null)
            {
                fonts.ComplexScript = cs;
            }

            if (ascii != null)
            {
                fonts.Ascii = ascii;
            }

            return fonts;
        }

        private string FindFirstAsciiFont(Style style, Styles styles)
        {
            while (true)
            {
                if (style.StyleRunProperties?.RunFonts != null && style.StyleRunProperties.RunFonts.Ascii != null)
                {
                    return style.StyleRunProperties.RunFonts.Ascii.Value;
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

        private string FindFirstCsFont(Style style, Styles styles)
        {
            while (true)
            {
                if (style.StyleRunProperties?.RunFonts != null && style.StyleRunProperties.RunFonts.ComplexScript != null)
                {
                    return style.StyleRunProperties.RunFonts.ComplexScript.Value;
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


        private bool CheckDefaultFontSize(DocDefaults defaultStyles, StringBuilder comment)
        {
            if (defaultStyles.RunPropertiesDefault?.RunPropertiesBaseStyle?.FontSize == null ||
                !FontSizes.Contains(defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.FontSize.Val.ToString()))
                return false;
            comment.AppendLine(Resources.CheckFontSize);
            return true;
        }

        private bool CheckDefaultRunFonts(DocDefaults defaultStyles, StringBuilder comment)
        {
            
            if (defaultStyles.RunPropertiesDefault?.RunPropertiesBaseStyle?.RunFonts == null ||
                (
                                    (defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii == null ||
                                     !defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii.HasValue || 
                                     RunFontsNames.Contains(defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.Ascii.Value)) ||
                                    (defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.ComplexScript != null &&
                                     defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.ComplexScript.HasValue
                                    && RunFontsNames.Contains(defaultStyles.RunPropertiesDefault.RunPropertiesBaseStyle.RunFonts.ComplexScript.Value))
                                ))
                return false;
            comment.AppendLine(Resources.CheckFont);
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