using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace Groupdocs.Comparison.Samples.Words.Documents
{
	public class CompareDocumentsWithTables
	{
		public static void WithTablesOnAppropriatePages()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesOnAppropriatePages.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesOnAppropriatePages.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTablesOnAppropriatePages/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void WithTablesOnDefferentPages()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesOnDefferentPages.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesOnDefferentPages.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTablesOnDefferentPages/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void WithTablesWitchContainsDifferentCountOfColumns()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfColumns.source.docx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfColumns.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTablesWitchContainsDifferentCountOfColumns/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void WithTablesWitchContainsDifferentCountOfRows()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfRows.source.docx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfRows.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTablesWitchContainsDifferentCountOfRows/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void WithTablesWitchContainsDifferentCountOfColumnsAndDifferentCountOfRows()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfColumnAndDifferentCountOfRows.source.docx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Words.Documents.data.WithTablesWitchContainsDifferentCountOfColumnAndDifferentCountOfRows.target.docx";
			string resultPath =
				@"./../../Documents/testresult/WithTablesWitchContainsDifferentCountOfColumnAndDifferentCountOfRows/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		private static void Compare(string sourcePath, string targetPath, string resultPath)
		{
			// Create to streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);
			// Opening two documents
			IComparisonDocument sourcePresentation = new ComparisonDocument(sourceStream);
			Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
			IComparisonDocument targetPresentation = new ComparisonDocument(targetStream);
			Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

			// Creating settings for comparison of documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();

			// Comparing documents
			IWordsCompareResult compareResult = sourcePresentation.CompareWith(targetPresentation, comparisonSettings);
			Console.WriteLine("Documents was compared.");

			// Saving result of comparison to new document
			IComparisonDocument result = compareResult.GetDocument();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Docx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to document with the following source path" + resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}
	}
}