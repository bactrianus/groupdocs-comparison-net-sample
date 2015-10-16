using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Slides;
using GroupDocs.Comparison.Slides.Contracts;
using GroupDocs.Comparison.Slides.Contracts.Comparison;
using GroupDocs.Comparison.Slides.Contracts.Enums;

namespace Groupdocs.Comparison.Samples.Slides.Presentations
{
	public class ComparePresentationsWithTables
	{
		public static void ComparePresentationsWithTablesOnAppropriateSlides()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesOnAppropriateSlides.old.pptx";
			string targetPath = @"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesOnAppropriateSlides.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithTablesOnAppropriateSlides/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithTablesOnDifferentSlides()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesOnDifferentSlides.old.pptx";
			string targetPath = @"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesOnDifferentSlides.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithTablesOnDifferentSlides/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithTablesWitchConyainsDifferentCountOfRows()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfRows.old.pptx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfRows.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithTablesWitchConyainsDifferentCountOfRows/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithTablesWitchConyainsDifferentCountOfColumns()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfColumns.old.pptx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfColumns.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithTablesWitchConyainsDifferentCountOfColumns/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, SlidesComparisonSettings);
		}

		public static void ComparePresentationsWithTablesWitchConyainsDifferentCountOfColumnsAndDifferentCountOfRows()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfColumnsAndDifferentCountOfRows.old.pptx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Slides.Presentations.data.WithTablesWitchConyainsDifferentCountOfColumnsAndDifferentCountOfRows.new.pptx";
			string resultPath =
				@"./../../Presentations/testresult/WithTablesWitchConyainsDifferentCountOfColumnsAndDifferentCountOfRows/result.pptx";
			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			Compare(sourcePath, targetPath, resultPath, SlidesComparisonSettings);
		}

		private static void Compare(string sourcePath, string targetPath, string resultPath,
			SlidesComparisonSettings SlidesComparisonSettings)
		{
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