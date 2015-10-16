using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace Groupdocs.Comparison.Samples.Words.Components
{
    public static class CompareParagraphs
    {
        public static void CompareParagraphsFromDifferentDocuments()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareParagraphsFromDifferentDocuments.source.docx";
            string targetPath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareParagraphsFromDifferentDocuments.target.docx";

            // Create to streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            // Opening two documents
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
            IComparisonDocument targetDocument = new ComparisonDocument(targetStream);
            Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

            // Getting first Paragraph from source document
            IComparisonParagraph sourceParagraph = sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Paragraph,false)[0] as IComparisonParagraph;
            // Getting first Paragraph from target document
            IComparisonParagraph targetParagraph = targetDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Paragraph, false)[0] as IComparisonParagraph;

            // Creating settings for comparison of Paragraphs
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Paragraphs
            IWordsCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, settings);
            Console.WriteLine("Paragraphs was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareParagraphsFromDifferentDocuments/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareParagraphFromDocumentWithCreatingParagraph()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareParagraphFromDocumentWithCreatingParagraph.source.docx";

            // Create to stream of document
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

            // Opening source document
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");

            // Getting first Paragraph from source document
            IComparisonParagraph sourceParagraph = sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Paragraph, false)[0] as IComparisonParagraph;
            
            // Creating Paragraph
            IComparisonParagraph targetParagraph = new ComparisonParagraph();
            targetParagraph.AddRun("This paragraph was created.");
            Console.WriteLine("New Paragraph was created.");

            // Creating settings for comparison of Paragraphs
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Paragraphs
            IWordsCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, settings);
            Console.WriteLine("Paragraphs was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareParagraphFromDocumentWithCreatingParagraph/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTwoCreatingParagraphs()
        {
            // Creating Paragraphs
            IComparisonParagraph sourceParagraph = new ComparisonParagraph();
            sourceParagraph.AddRun("This is source Paragraph.");
            Console.WriteLine("New Paragraph was created.");

            IComparisonParagraph targetParagraph = new ComparisonParagraph();
            targetParagraph.AddRun("This is target Paragraph.");
            Console.WriteLine("New Paragraph was created.");

            // Creating settings for comparison of Paragraphs
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Paragraphs
            IWordsCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, settings);
            Console.WriteLine("Paragraphs was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTwoCreatingParagraphs/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }
    }
}
