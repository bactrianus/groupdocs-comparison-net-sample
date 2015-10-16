using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace Groupdocs.Comparison.Samples.Words.Settings
{
	internal class ComparisonWithDifferentSettings
	{
		public static void CompareDocumentsWithGenerationSummaryPage()
		{
			string resultPath = @"./../../Settings/testresult/WithGenerationSummaryPage/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.GenerateSummaryPage = true;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithOutGenerationSummaryPage()
		{
			string resultPath = @"./../../Settings/testresult/WithOutGenerationSummaryPage/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
            comparisonSettings.GenerateSummaryPage = false;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithShowDeletedContent()
		{
			string resultPath = @"./../../Settings/testresult/WithShowDeletedContent/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.ShowDeletedContent = true;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithOutShowDeletedContent()
		{
			string resultPath = @"./../../Settings/testresult/WithOutShowDeletedContent/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.ShowDeletedContent = false;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithComparisonDepthSetWords()
		{
			string resultPath = @"./../../Settings/testresult/WithComparisonDepthSetWords/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.ComparisonDepth = ComparisonDepth.Words;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithComparisonDepthSetChars()
		{
			string resultPath = @"./../../Settings/testresult/WithComparisonDepthSetChars/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.ComparisonDepth = ComparisonDepth.Chars;
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithWordsSepCharsSetSpace()
		{
			string resultPath = @"./../../Settings/testresult/WithWordsSepCharsSetSpace/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.WordsSeparatorChars = new char[] {' '};
			Compare(resultPath, comparisonSettings);
		}

		public static void CompareDocumentsWithSettingStylesOnDelInsComponents()
		{
			string resultPath = @"./../../Settings/testresult/WithSettingStylesOnDelInsComponents/result.docx";
			// Creating settings for comparison of Documents
			WordsComparisonSettings comparisonSettings = new WordsComparisonSettings();
			comparisonSettings.DeletedItemsStyle.Color = Color.Brown;
			comparisonSettings.InsertedItemsStyle.Color = Color.DarkOrange;
			Compare(resultPath, comparisonSettings);
		}

		private static void Compare(string resultPath, WordsComparisonSettings comparisonSettings)
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Words.Settings.data.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.Words.Settings.data.target.docx";
			// Create to streams of Documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);
			// Opening two Documents
			IComparisonDocument sourcePresentation = new ComparisonDocument(sourceStream);
			Console.WriteLine("Document with source path: " + sourcePath + " was loaded.");
			IComparisonDocument targetPresentation = new ComparisonDocument(targetStream);
			Console.WriteLine("Document with source path: " + targetPath + " was loaded.");

			// Comparing Documents
			IWordsCompareResult compareResult = sourcePresentation.CompareWith(targetPresentation, comparisonSettings);
			Console.WriteLine("Documents was compared.");

			// Saving result of comparison to new document
			IComparisonDocument result = compareResult.GetDocument();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Docx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to document with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}
	}
}