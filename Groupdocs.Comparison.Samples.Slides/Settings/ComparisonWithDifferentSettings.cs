using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Slides;
using GroupDocs.Comparison.Slides.Contracts;
using GroupDocs.Comparison.Slides.Contracts.Comparison;
using GroupDocs.Comparison.Slides.Contracts.Enums;

namespace GroupDocs.Comparison.Samples.Slides.Settings
{
	internal class ComparisonWithDifferentSettings
	{
		public static void ComparePresentationsWithGenerationSummaryPage()
		{
			string resultPath = @"./../../Settings/testresult/WithGenerationSummaryPage/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.GenerateSummaryPage = true;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithOutGenerationSummaryPage()
		{
			string resultPath = @"./../../Settings/testresult/WithOutGenerationSummaryPage/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.GenerateSummaryPage = false;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithShowDeletedContent()
		{
			string resultPath = @"./../../Settings/testresult/WithShowDeletedContent/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.ShowDeletedContent = true;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithOutShowDeletedContent()
		{
			string resultPath = @"./../../Settings/testresult/WithOutShowDeletedContent/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.ShowDeletedContent = false;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithComparisonDepthSetWords()
		{
			string resultPath = @"./../../Settings/testresult/WithComparisonDepthSetWords/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.ComparisonDepth = ComparisonDepth.Words;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithComparisonDepthSetChars()
		{
			string resultPath = @"./../../Settings/testresult/WithComparisonDepthSetChars/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.ComparisonDepth = ComparisonDepth.Chars;
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithWordsSepCharsSetSpace()
		{
			string resultPath = @"./../../Settings/testresult/WithWordsSepCharsSetSpace/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			SlidesComparisonSettings.WordsSeparatorChars = new char[] {' '};
			Compare(resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithSettingStylesOnDelInsComponents()
		{
			string resultPath = @"./../../Settings/testresult/WithSettingStylesOnDelInsComponents/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings settings = new SlidesComparisonSettings();
			settings.InsertedItemsStyle.Color = Color.Aqua;
			settings.DeletedItemsStyle.Color = Color.BlueViolet;
			Compare(resultPath, settings);
		}

		private static void Compare(string resultPath, SlidesComparisonSettings SlidesComparisonSettings)
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Slides.Settings.data.source.pptx";
			string targetPath = @"GroupDocs.Comparison.Samples.Slides.Settings.data.target.pptx";
			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);
			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Comparing presentations
			ISlidesCompareResult compareResult = sourcePresentation.CompareWith(targetPresentation, SlidesComparisonSettings);
			Console.WriteLine("Presentations was compared.");

			// Saving result of comparison to new presentation
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}
	}
}