using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace GroupDocs.Comparison.Samples.Words.Documents
{
	public class CompareDocumentsWithText
	{
		public static void WithTextOnAppropriatePages()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithTextOnAppropriatePages.source.docx";
			string targetPath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithTextOnAppropriatePages.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTextOnAppropriatePages/result.docx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void WithTextOnDifferentPages()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithTextOnDifferentPages.source.docx";
			string targetPath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithTextOnDifferentPages.target.docx";
			string resultPath = @"./../../Documents/testresult/WithTextOnDifferentPages/result.docx";
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