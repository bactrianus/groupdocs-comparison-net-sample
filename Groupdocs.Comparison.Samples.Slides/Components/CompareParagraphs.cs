using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Slides;
using GroupDocs.Comparison.Slides.Contracts;
using GroupDocs.Comparison.Slides.Contracts.Comparison;
using GroupDocs.Comparison.Slides.Contracts.Enums;

namespace GroupDocs.Comparison.Samples.Slides.Components
{
	public static class CompareParagraphs
	{
		public static void CompareParagraphsFromDifferentPresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareParagraphsFromDifferentPresentations.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareParagraphsFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first Paragraph from source presentation
			ComparisonParagraphBase sourceParagraph =
				(sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase).TextFrame.Paragraphs[0];
			// Getting first Paragraph from target presentation
			ComparisonParagraphBase targetParagraph =
				(targetPresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase).TextFrame.Paragraphs[0];

			// Creating settings for comparison of Paragraphs
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Paragraphs
			ISlidesCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, SlidesComparisonSettings);
			Console.WriteLine("Paragraphs was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareParagraphsFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareParagraphsFromOnePresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareParagraphsFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Paragraph from source presentation
			ComparisonParagraphBase sourceParagraph =
				(sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase).TextFrame.Paragraphs[0];
			// Getting second Paragraph from source presentation
			ComparisonParagraphBase targetParagraph =
				(sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase).TextFrame.Paragraphs[1];

			// Creating settings for comparison of Paragraphs
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Paragraphs
			ISlidesCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, SlidesComparisonSettings);
			Console.WriteLine("Paragraphs was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareParagraphsFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareParagraphFromPresentationsWithCreatingParagraph()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareParagraphFromPresentationsWithCreatingParagraph.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Paragraph from source presentation
			ComparisonParagraphBase sourceParagraph =
				(sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase).TextFrame.Paragraphs[0];
			// Creating Paragraph
			ComparisonParagraphBase targetParagraph = new ComparisonParagraph();
			targetParagraph.Text = "This Paragraph was created.";
			Console.WriteLine("New Paragraph was created.");

			// Creating settings for comparison of Paragraphs
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Paragraphs
			ISlidesCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, SlidesComparisonSettings);
			Console.WriteLine("Paragraphs was compared.");

			// Saving result of comparison to new presentation
			string resultPath =
				@"./../../Components/testresult/CompareParagraphFromPresentationsWithCreatingParagraph/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingParagraphs()
		{
			// Creating Paragraphs
			ComparisonParagraphBase sourceParagraph = new ComparisonParagraph();
			sourceParagraph.Text = "This is source Paragraph.";
			Console.WriteLine("New Paragraph was created.");

			ComparisonParagraphBase targetParagraph = new ComparisonParagraph();
			targetParagraph.Text = "This is target Paragraph.";
			Console.WriteLine("New Paragraph was created.");

			// Creating settings for comparison of Paragraphs
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Paragraphs
			ISlidesCompareResult compareResult = sourceParagraph.CompareWith(targetParagraph, SlidesComparisonSettings);
			Console.WriteLine("Paragraphs was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingParagraphs/result.pptx";
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