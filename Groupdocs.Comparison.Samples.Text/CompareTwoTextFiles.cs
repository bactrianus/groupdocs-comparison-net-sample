using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Text;
using GroupDocs.Comparison.Text.Contracts;

namespace Groupdocs.Comparison.Samples.Text
{
	internal class CompareTwoTextFiles
	{
		public static void CompareTwoFilesWithPlainText()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Text.data.source.txt";
			string targetPath = @"Groupdocs.Comparison.Samples.Text.data.target.txt";
			string resultPath = @"./../../testresult/WithPlainText/result.txt";
			Compare(sourcePath, targetPath, resultPath);
		}

		private static void Compare(string sourcePath, string targetPath, string resultPath)
		{
			// Create two streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);
			// Opening two presentations
			IComparisonTextFile sourcePresentation = new ComparisonTextFile(sourceStream);
			Console.WriteLine("TextFile with source path: " + sourcePath + " was loaded.");
			IComparisonTextFile targetPresentation = new ComparisonTextFile(targetStream);
			Console.WriteLine("TextFile with source path: " + targetPath + " was loaded.");

			// Creating settings for comparison of presentations
			TextComparisonSettings SlidesComparisonSettings = new TextComparisonSettings();

			// Comparing presentations
			ITextCompareResult compareResult = sourcePresentation.CompareWith(targetPresentation, SlidesComparisonSettings);
			Console.WriteLine("Presentations was compared.");

			// Saving result of comparison to new presentation
			IComparisonTextFile result = compareResult.GetTextFile();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to TextFile with the following source path" + resultPath + ".");
		}
	}
}