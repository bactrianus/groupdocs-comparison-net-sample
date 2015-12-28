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
	public static class CompareAutoShapes
	{
		public static void CompareAutoShapesFromDifferentPresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareAutoShapesFromDifferentPresentations.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareAutoShapesFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first AutoShape from source presentation
			ComparisonAutoShapeBase sourceAutoShape = (sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase);
			// Getting first AutoShape from target presentation
			ComparisonAutoShapeBase targetAutoShape = (targetPresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase);

			// Creating settings for comparison of AutoShapes
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing AutoShapes
			ISlidesCompareResult compareResult = sourceAutoShape.CompareWith(targetAutoShape, SlidesComparisonSettings);
			Console.WriteLine("AutoShapes was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareAutoShapesFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareAutoShapesFromOnePresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareAutoShapesFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();

			// Opening source presentation
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first AutoShape from source presentation
			ComparisonAutoShapeBase sourceAutoShape = (sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase);
			// Getting second AutoShape from source presentation
			ComparisonAutoShapeBase targetAutoShape = (sourcePresentation.Slides[1].Shapes[0] as ComparisonAutoShapeBase);

			// Creating settings for comparison of AutoShapes
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing AutoShapes
			ISlidesCompareResult compareResult = sourceAutoShape.CompareWith(targetAutoShape, SlidesComparisonSettings);
			Console.WriteLine("AutoShapes was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareAutoShapesFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareAutoShapeFromPresentationsWithCreatingAutoShape()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareAutoShapeFromPresentationsWithCreatingAutoShape.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first AutoShape from source presentation
			ComparisonAutoShapeBase sourceAutoShape = (sourcePresentation.Slides[0].Shapes[0] as ComparisonAutoShapeBase);
			// Creating AutoShape
			ComparisonAutoShapeBase targetAutoShape = new ComparisonAutoShape(100, 100, 500, 300);
			ComparisonParagraphBase paragraph = new ComparisonParagraph();
			paragraph.Text = "This AutoShape was created.";
			targetAutoShape.AddTextFrame("");
			targetAutoShape.TextFrame.Paragraphs.Add(paragraph);
			Console.WriteLine("New AutoShape was created.");

			// Creating settings for comparison of AutoShapes
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing AutoShapes
			ISlidesCompareResult compareResult = sourceAutoShape.CompareWith(targetAutoShape, SlidesComparisonSettings);
			Console.WriteLine("AutoShapes was compared.");

			// Saving result of comparison to new presentation
			string resultPath =
				@"./../../Components/testresult/CompareAutoShapeFromPresentationsWithCreatingAutoShape/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingAutoShapes()
		{
			// Creating AutoShapes
			ComparisonAutoShapeBase sourceAutoShape = new ComparisonAutoShape(100, 100, 500, 300);
			ComparisonParagraphBase sourceParagraph = new ComparisonParagraph();
			sourceParagraph.Text = "This is source AutoShape.";
			sourceAutoShape.AddTextFrame("");
			sourceAutoShape.TextFrame.Paragraphs.Add(sourceParagraph);
			Console.WriteLine("New AutoShape was created.");

			ComparisonAutoShapeBase targetAutoShape = new ComparisonAutoShape(100, 100, 500, 300);
			ComparisonParagraphBase targetParagraph = new ComparisonParagraph();
			targetParagraph.Text = "This is target AutoShape.";
			targetAutoShape.AddTextFrame("");
			targetAutoShape.TextFrame.Paragraphs.Add(targetParagraph);
			Console.WriteLine("New AutoShape was created.");

			// Creating settings for comparison of AutoShapes
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing AutoShapes
			ISlidesCompareResult compareResult = sourceAutoShape.CompareWith(targetAutoShape, SlidesComparisonSettings);
			Console.WriteLine("AutoShapes was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingAutoShapes/result.pptx";
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