using DocumentFormat.OpenXml.Packaging;

namespace DocumentReview
{
    public interface ICheckStrategy
    {
        void DoCheck(WordprocessingDocument document);
    }
}