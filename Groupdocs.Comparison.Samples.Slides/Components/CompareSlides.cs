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
	public static class CompareSlides
	{
		public static void CompareSlidesFromDifferentPresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareSlidesFromDifferentPresentations.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareSlidesFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first slide from source presentation
			ComparisonSlideBase sourceSlide = sourcePresentation.Slides[0];
			// Getting second slide from target presentation
			ComparisonSlideBase targetSlide = targetPresentation.Slides[0];

			// Creating settings for comparison of slides
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing slides
			ISlidesCompareResult compareResult = sourceSlide.CompareWith(targetSlide, SlidesComparisonSettings);
			Console.WriteLine("Slides was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareSlidesFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareSlidesFromOnePresentations()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Slides.Components.data.CompareSlidesFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first slide from source presentation
			ComparisonSlideBase sourceSlide = sourcePresentation.Slides[0];
			// Getting second slide from source presentation
			ComparisonSlideBase targetSlide = sourcePresentation.Slides[1];

			// Creating settings for comparison of slides
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing slides
			ISlidesCompareResult compareResult = sourceSlide.CompareWith(targetSlide, SlidesComparisonSettings);
			Console.WriteLine("Slides was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareSlidesFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareSlideFromPresentationsWithCreatingSlide()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareSlideFromPresentationsWithCreatingSlide.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first slide from source presentation
			ComparisonSlideBase sourceSlide = sourcePresentation.Slides[0];
			// Creating Slide
			ComparisonSlideBase targetSlide = new ComparisonSlide();
			ComparisonAutoShapeBase autoShape = targetSlide.Shapes.AddAutoShape(ComparisonShapeType.Rectangle, 100, 100, 500, 300);
			ComparisonParagraphBase paragraph = new ComparisonParagraph();
			paragraph.Text = "Contrary to popular belief, Lorem Ipsum is not simply random text. It has" +
			                 " roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard " +
			                 "McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure";
			autoShape.TextFrame.Paragraphs.Add(paragraph);
			Console.WriteLine("New slide was created.");

			// Creating settings for comparison of slides
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing slides
			ISlidesCompareResult compareResult = sourceSlide.CompareWith(targetSlide, SlidesComparisonSettings);
			Console.WriteLine("Slides was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareSlideFromPresentationsWithCreatingSlide/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingSlides()
		{
			// Creating Slides
			ComparisonSlideBase sourceSlide = new ComparisonSlide();
			ComparisonAutoShapeBase sourceAutoShape = sourceSlide.Shapes.AddAutoShape(ComparisonShapeType.Rectangle, 100, 100, 500,
				300);
			ComparisonParagraphBase sourceParagraph = new ComparisonParagraph();
			sourceParagraph.Text = "Contrary to popular belief, Lorem Ipsum is not simply random text. It has" +
			                       " roots in a piece of classical Latin literature from 45 BC, making it over 2000 years old. Richard " +
			                       "McClintock, a Latin professor at Hampden-Sydney College in Virginia, looked up one of the more obscure";
			sourceAutoShape.TextFrame.Paragraphs.Add(sourceParagraph);
			Console.WriteLine("New slide was created.");

			ComparisonSlideBase targetSlide = new ComparisonSlide();
			ComparisonAutoShapeBase targetAutoShape = targetSlide.Shapes.AddAutoShape(ComparisonShapeType.Rectangle, 100, 100, 500,
				300);
			ComparisonParagraphBase targetParagraph = new ComparisonParagraph();
			targetParagraph.Text = "Contrary to popular belief, the Lorem Ipsum is not simply random text. Richard " +
			                       "McClintock, a Latin professor at Hampden-Sydney in Virginia, looked up one of the more obscure";
			targetAutoShape.TextFrame.Paragraphs.Add(targetParagraph);
			Console.WriteLine("New slide was created.");

			// Creating settings for comparison of slides
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing slides
			ISlidesCompareResult compareResult = sourceSlide.CompareWith(targetSlide, SlidesComparisonSettings);
			Console.WriteLine("Slides was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingSlides/result.pptx";
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