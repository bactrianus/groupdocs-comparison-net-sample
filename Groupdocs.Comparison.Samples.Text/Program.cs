using GroupDocs.Comparison.Samples.Text.Settings;

namespace GroupDocs.Comparison.Samples.Text
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			//Text
			CompareTwoTextFiles.CompareTwoFilesWithPlainText();

			//Settings
			ComparisonWithDifferentSettings.CompareTextFilesWithComparisonDepthSetChars();
			ComparisonWithDifferentSettings.CompareTextFilesWithComparisonDepthSetWords();
			ComparisonWithDifferentSettings.CompareTextFilesWithGenerationSummaryPage();
			ComparisonWithDifferentSettings.CompareTextFilesWithOutGenerationSummaryPage();
			ComparisonWithDifferentSettings.CompareTextFilesWithOutShowDeletedContent();
			ComparisonWithDifferentSettings.CompareTextFilesWithShowDeletedContent();
			ComparisonWithDifferentSettings.CompareTextFilesWithWordsSepCharsSetSpace();
			ComparisonWithDifferentSettings.CompareTextFilesWithSettingStylesOnDelInsComponents();

			//Console.ReadKey();
		}
	}
}