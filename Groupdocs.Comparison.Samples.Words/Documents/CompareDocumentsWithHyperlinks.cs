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
	public class CompareDocumentsWithHyperlinks
	{
		public static void WithIgnoreLinkSetting()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithIgnoreLinkSetting.source.docx";
			string targetPath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithIgnoreLinkSetting.target.docx";
			string resultPath = @"./../../Documents/testresult/WithIgnoreLinkSetting/result.docx";
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.ProcessHyperLinksAsAText = true;
			Compare(sourcePath, targetPath, resultPath, comparisonSettings);
		}

		public static void WithoutIgnoreLinkSetting()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithoutIgnoreLinkSetting.source.docx";
			string targetPath = @"GroupDocs.Comparison.Samples.Words.Documents.data.WithoutIgnoreLinkSetting.target.docx";
			string resultPath = @"./../../Documents/testresult/WithoutIgnoreLinkSetting/result.docx";
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, comparisonSettings);
		}

		private static void Compare(string sourcePath, string targetPath, string resultPath,
			WordsComparisonSettings comparisonSettings)
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