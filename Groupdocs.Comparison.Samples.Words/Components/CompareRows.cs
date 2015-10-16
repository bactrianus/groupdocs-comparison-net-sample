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
    class CompareRows
    {
        public static void CompareRowsFromDifferentDocuments()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareRowsFromDifferentDocuments.source.docx";
            string targetPath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareRowsFromDifferentDocuments.target.docx";

            // Create to streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            // Opening two documents
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
            IComparisonDocument targetDocument = new ComparisonDocument(targetStream);
            Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

            // Getting first Row from source document
            IComparisonRow sourceRow = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0];
            // Getting first Row from target document
            IComparisonRow targetRow = (targetDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0];

            // Creating settings for comparison of Rows
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Rows
            IWordsCompareResult compareResult = sourceRow.CompareWith(targetRow, settings);
            Console.WriteLine("Rows was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareRowsFromDifferentDocuments/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareRowFromDocumentWithCreatingRow()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareRowFromDocumentWithCreatingRow.source.docx";

            // Create to stream of document
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

            // Opening source document
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");

            // Getting first Row from source document
            IComparisonRow sourceRow = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Rows[0];
            
            // Creating Row
            IComparisonRow targetRow = new ComparisonRow(new double[]{100,100,100}, 20);
            IComparisonParagraph paragraph = targetRow.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetRow.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            paragraph = targetRow.Cells[2].AddParagraph();
            paragraph.AddRun("This is Cell.");
            Console.WriteLine("New Row was created.");

            // Creating settings for comparison of Rows
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Rows
            IWordsCompareResult compareResult = sourceRow.CompareWith(targetRow, settings);
            Console.WriteLine("Rows was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareRowFromDocumentWithCreatingRow/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTwoCreatingRows()
        {
            // Creating Rows
            IComparisonRow sourceRow = new ComparisonRow(new double[] { 100, 100 }, 20);
            IComparisonParagraph paragraph = sourceRow.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = sourceRow.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of source table.");
            Console.WriteLine("New Row was created.");

            IComparisonRow targetRow = new ComparisonRow(new double[] { 100, 100 }, 20);
            paragraph = targetRow.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetRow.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            Console.WriteLine("New Row was created.");

            // Creating settings for comparison of Rows
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Rows
            IWordsCompareResult compareResult = sourceRow.CompareWith(targetRow, settings);
            Console.WriteLine("Rows was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTwoCreatingRows/result.docx";
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

