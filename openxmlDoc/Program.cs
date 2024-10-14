using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace openxmlDoc
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");
            string mainDoc = @"C:\Users\Meghana Jakata\source\repos\openxmlDoc\main.docx";
            string subDoc = @"C:\Users\Meghana Jakata\source\repos\openxmlDoc\sub.docx";
            RPTSectionMergeToMainAndDelete(mainDoc, subDoc);
        }

        public static void RPTSectionMergeToMainAndDelete(string temppathMain, string temppathSub)
        {
            using (WordprocessingDocument maindocTextwordDoc = WordprocessingDocument.Open(temppathMain, true))
            {
                // Open the source document
                using (WordprocessingDocument subTextwordDoc = WordprocessingDocument.Open(temppathSub, false))
                {
                    // Get the body of the source document
                    Body sourceBody = subTextwordDoc.MainDocumentPart.Document.Body;

                    // Clone the contents to avoid issues with OpenXML structure
                    Body clonedBody = (Body)sourceBody.CloneNode(true);

                    // Get the body of the destination document
                    Body mainDocBody = maindocTextwordDoc.MainDocumentPart.Document.Body;

                    // Append the cloned content
                    foreach (var element in clonedBody.Elements())
                    {
                        mainDocBody.Append(element.CloneNode(true));
                    }
                    maindocTextwordDoc.MainDocumentPart.Document.Save();
                }
            }
        }
    }
}
