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
    class CompareTables
    {
        public static void CompareTablesFromDifferentDocuments()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareTablesFromDifferentDocuments.source.docx";
            string targetPath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareTablesFromDifferentDocuments.target.docx";

            // Create to streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            // Opening two documents
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
            IComparisonDocument targetDocument = new ComparisonDocument(targetStream);
            Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

            // Getting first Table from source document
            IComparisonTable sourceTable = sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable;
            // Getting first Table from target document
            IComparisonTable targetTable = targetDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable;

            // Creating settings for comparison of Tables
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Tables
            IWordsCompareResult compareResult = sourceTable.CompareWith(targetTable, settings);
            Console.WriteLine("Tables was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTablesFromDifferentDocuments/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTableFromDocumentWithCreatingTable()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareTableFromDocumentWithCreatingTable.source.docx";

            // Create to stream of document
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

            // Opening source document
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");

            // Getting first Table from source document
            IComparisonTable sourceTable = sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable;

            // Creating Table
            IComparisonTable targetTable = new ComparisonTable(new double[]{100,100,100}, new double[] { 20, 20 });
            IComparisonParagraph paragraph = targetTable.Rows[0].Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetTable.Rows[0].Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            paragraph = targetTable.Rows[1].Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell.");
            Console.WriteLine("New Table was created.");

            // Creating settings for comparison of Tables
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Tables
            IWordsCompareResult compareResult = sourceTable.CompareWith(targetTable, settings);
            Console.WriteLine("Tables was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTableFromDocumentWithCreatingTable/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTwoCreatingTables()
        {
            // Creating Tables
            IComparisonTable sourceTable = new ComparisonTable(new double[]{100,100}, new double[] { 20, 20 });
            IComparisonParagraph paragraph = sourceTable.Rows[0].Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = sourceTable.Rows[0].Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of source table.");
            paragraph = sourceTable.Rows[1].Cells[0].AddParagraph();
            paragraph.AddRun("This is Cel of tble.");
            Console.WriteLine("New Table was created.");

            IComparisonTable targetTable = new ComparisonTable(new double[]{100,100}, new double[] { 20, 20 });
            paragraph = targetTable.Rows[0].Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetTable.Rows[0].Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            paragraph = targetTable.Rows[1].Cells[0].AddParagraph();
            paragraph.AddRun("This is Cell of table.");
            Console.WriteLine("New Table was created.");

            // Creating settings for comparison of Tables
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Tables
            IWordsCompareResult compareResult = sourceTable.CompareWith(targetTable, settings);
            Console.WriteLine("Tables was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTwoCreatingTables/result.docx";
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
