using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Slides;
using GroupDocs.Comparison.Slides.Contracts;
using GroupDocs.Comparison.Slides.Contracts.Comparison;
using GroupDocs.Comparison.Slides.Contracts.Enums;

namespace GroupDocs.Comparison.Samples.Slides.Presentations
{
	public class ComparePresentationsWithAutoShapes
	{
		public static void ComparePresentationsWithAutoShapesOnAppropriateSlides()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Presentations.data.WithAutoShapesOnAppropriateSlides.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Presentations.data.WithAutoShapesOnAppropriateSlides.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithAutoShapesOnAppropriateSlides/result.pptx";
			Compare(sourcePath, targetPath, resultPath);
		}

		public static void ComparePresentationsWithAutoShapesOnDifferentSlides()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Presentations.data.WithAutoShapesOnDifferentSlides.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Presentations.data.WithAutoShapesOnDifferentSlides.new.pptx";
			string resultPath = @"./../../Presentations/testresult/WithAutoShapesOnDifferentSlides/result.pptx";
			Compare(sourcePath, targetPath, resultPath);
		}

		private static void Compare(string sourcePath, string targetPath, string resultPath)
		{
			// Create two streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);
			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Creating settings for comparison of presentations
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();

			// Comparing presentations
			ISlidesCompareResult compareResult = sourcePresentation.CompareWith(targetPresentation, SlidesComparisonSettings);
			Console.WriteLine("Presentations was compared.");

			// Saving result of comparison to new presentation
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" + resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}
	}
}