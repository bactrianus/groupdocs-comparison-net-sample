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
    class CompareColumns
    {
        public static void CompareColumnsFromDifferentDocuments()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareColumnsFromDifferentDocuments.source.docx";
            string targetPath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareColumnsFromDifferentDocuments.target.docx";

            // Create to streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            // Opening two documents
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
            IComparisonDocument targetDocument = new ComparisonDocument(targetStream);
            Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

            // Getting first Column from source document
            IComparisonColumn sourceColumn = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Columns[0];
            // Getting first Column from target document
            IComparisonColumn targetColumn = (targetDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Columns[0];

            // Creating settings for comparison of Columns
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Columns
            IWordsCompareResult compareResult = sourceColumn.CompareWith(targetColumn, settings);
            Console.WriteLine("Columns was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareColumnsFromDifferentDocuments/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareColumnFromDocumentWithCreatingColumn()
        {
            string sourcePath =
                @"Groupdocs.Comparison.Samples.Words.Components.data.CompareColumnFromDocumentWithCreatingColumn.source.docx";

            // Create to stream of document
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

            // Opening source document
            IComparisonDocument sourceDocument = new ComparisonDocument(sourceStream);
            Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");

            // Getting first Column from source document
            IComparisonColumn sourceColumn = (sourceDocument.Sections[0].Body.GetChildNodes(ComparisonNodeType.Table, false)[0] as IComparisonTable).Columns[0];
            
            // Creating Column
            IComparisonColumn targetColumn = new ComparisonColumn(100, new double[]{20,20,20});
            IComparisonParagraph paragraph = targetColumn.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetColumn.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            paragraph = targetColumn.Cells[2].AddParagraph();
            paragraph.AddRun("This is Cell.");
            Console.WriteLine("New Column was created.");

            // Creating settings for comparison of Columns
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Columns
            IWordsCompareResult compareResult = sourceColumn.CompareWith(targetColumn, settings);
            Console.WriteLine("Columns was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareColumnFromDocumentWithCreatingColumn/result.docx";
            IComparisonDocument result = compareResult.GetDocument();
            Stream resultStream = new FileStream(resultPath, FileMode.Create);
            result.Save(resultStream, ComparisonSaveFormat.Docx);
            resultStream.Close();
            Console.WriteLine("Result of comparison was saved to document with the following source path" +
                              resultPath + ".");
            Console.WriteLine("===============================================");
            Console.WriteLine("");
        }

        public static void CompareTwoCreatingColumns()
        {
            // Creating Columns
            IComparisonColumn sourceColumn = new ComparisonColumn(100, new double[] { 20,20 });
            IComparisonParagraph paragraph = sourceColumn.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = sourceColumn.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of source table.");
            Console.WriteLine("New Column was created.");
            
            IComparisonColumn targetColumn = new ComparisonColumn(100, new double[] { 20,20 });
            paragraph = targetColumn.Cells[0].AddParagraph();
            paragraph.AddRun("This is cell.");
            paragraph = targetColumn.Cells[1].AddParagraph();
            paragraph.AddRun("This is Cell of target table.");
            Console.WriteLine("New Column was created.");

            // Creating settings for comparison of Columns
            WordsComparisonSettings settings = new WordsComparisonSettings();
            // Comparing Columns
            IWordsCompareResult compareResult = sourceColumn.CompareWith(targetColumn, settings);
            Console.WriteLine("Columns was compared.");

            // Saving result of comparison to new document
            string resultPath = @"./../../Components/testresult/CompareTwoCreatingColumns/result.docx";
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

