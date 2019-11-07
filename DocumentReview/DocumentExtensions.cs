using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentReview
{
    public static class DocumentExtensions
    {
        /// <summary>
        /// Возвращает стандратные стили документа.
        /// </summary>
        /// <param name="document">Открытый документ.</param>
        /// <returns>Стандартные стили.</returns>
        public static DocDefaults GetDefaultStyles(this WordprocessingDocument document)
        {
            var docPart = document.MainDocumentPart;
            return docPart.StylesWithEffectsPart != null ? docPart.StylesWithEffectsPart.Styles.DocDefaults : docPart.StyleDefinitionsPart.Styles.DocDefaults;
        }

        /// <summary>
        /// Возвращает стандратные стили документа.
        /// </summary>
        /// <param name="document">Открытый документ.</param>
        /// <returns>Стандартные стили.</returns>
        public static Styles GetStyles(this WordprocessingDocument document)
        {
            var docPart = document.MainDocumentPart;
            return docPart.StylesWithEffectsPart != null ? docPart.StylesWithEffectsPart.Styles : docPart.StyleDefinitionsPart.Styles;
        }

        public static Style FindById(this Styles styles, string id)
        {
            return styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId.HasValue && s.StyleId.Value == id);
        }
    }
}