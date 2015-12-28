using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace GroupDocs.Comparison.Samples.Words.Components
{
    class CompareCells
    {
        public static void CompareCellsFromDifferentDocuments()
        {
            string sourcePath =
                @"GroupDocs.Comparison.Samples.Words.Components.data.CompareCellsFromDifferentDocuments.source.docx";
            string targetPath =
                @"GroupDocs.Comparison.Samples.Words.Components.data.CompareCellsFromDifferentDocuments.target.docx";

            // Create to streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            // Opening two documents
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
            IComparisonDocument targetDocument = new ComparisonDocument(targetStream);
            Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

            // Getting first cell from source document
            IComparisonCell sourceCell = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0].Cells[0];
            // Getting first cell from target document
            IComparisonCell targetCell = (targetDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0].Cells[0];

            // Creating settings for comparison of cells
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing cells
            IWordsCompareResult compareResult = sourceCell.CompareWith(targetCell, settings);
            Console.WriteLine("Cells was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareCellsFromDifferentDocuments/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareCellFromDocumentWithCreatingCell()
        {
            string sourcePath =
                @"GroupDocs.Comparison.Samples.Words.Components.data.CompareCellFromDocumentWithCreatingCell.source.docx";

            // Create to stream of document
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

            // Opening source document
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");

            // Getting first cell from source document
            IComparisonCell sourceCell = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0].Cells[0];

            // Creating Cell
            IComparisonCell targetCell = new ComparisonCell();
            IComparisonParagraph paragraph = targetCell.AddParagraph();
            paragraph.AddRun("This is cell.");
            Console.WriteLine("New Cell was created.");

            // Creating settings for comparison of Cells
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Cells
            IWordsCompareResult compareResult = sourceCell.CompareWith(targetCell, settings);
            Console.WriteLine("Cells was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareCellFromDocumentWithCreatingCell/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTwoCreatingCells()
        {
            // Creating Cells
            IComparisonCell sourceCell = new ComparisonCell();
            IComparisonParagraph paragraph = sourceCell.AddParagraph();
            paragraph.AddRun("This is Cell of source table.");
            Console.WriteLine("New Column was created.");

            IComparisonCell targetCell = new ComparisonCell();
            paragraph = targetCell.AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            Console.WriteLine("New Column was created.");

            // Creating settings for comparison of Cells
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Cells
            IWordsCompareResult compareResult = sourceCell.CompareWith(targetCell, settings);
            Console.WriteLine("Cells was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTwoCreatingCells/result.docx";
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
