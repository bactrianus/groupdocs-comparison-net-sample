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
	public static class CompareRows
	{
		public static void CompareRowsFromDifferentPresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareRowsFromDifferentPresentations.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareRowsFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first Row from source presentation
			ComparisonRowBase sourceRow = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Rows[0];
			// Getting first Row from target presentation
			ComparisonRowBase targetRow = (targetPresentation.Slides[0].Shapes[0] as ComparisonTableBase).Rows[0];

			// Creating settings for comparison of Rows
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Rows
			ISlidesCompareResult compareResult = sourceRow.CompareWith(targetRow, SlidesComparisonSettings);
			Console.WriteLine("Rows was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareRowsFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareRowsFromOnePresentations()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Slides.Components.data.CompareRowsFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Row from source presentation
			ComparisonRowBase sourceRow = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Rows[0];
			// Getting second Row from source presentation
			ComparisonRowBase targetRow = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Rows[1];

			// Creating settings for comparison of Rows
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Rows
			ISlidesCompareResult compareResult = sourceRow.CompareWith(targetRow, SlidesComparisonSettings);
			Console.WriteLine("Rows was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareRowsFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareRowFromPresentationsWithCreatingRow()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareRowFromPresentationsWithCreatingRow.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Row from source presentation
			ComparisonRowBase sourceRow = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Rows[0];
			// Creating Row
			ComparisonRowBase targetRow = new ComparisonRow(new double[] {200, 200}, 50);
			ComparisonParagraphBase paragraph0 = new ComparisonParagraph();
			paragraph0.Text = "This is first cell in Row that was created.";
			targetRow[0].TextFrame.Paragraphs.Add(paragraph0);
			ComparisonParagraphBase paragraph1 = new ComparisonParagraph();
			paragraph1.Text = "This is second cell in Row that was created.";
			targetRow[1].TextFrame.Paragraphs.Add(paragraph1);
			Console.WriteLine("New Row was created.");

			// Creating settings for comparison of Rows
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Rows
			ISlidesCompareResult compareResult = sourceRow.CompareWith(targetRow, SlidesComparisonSettings);
			Console.WriteLine("Rows was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareRowFromPresentationsWithCreatingRow/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingRows()
		{
			// Creating Rows
			ComparisonRowBase sourceRow = new ComparisonRow(new double[] {200, 200}, 50);
			ComparisonParagraphBase sourceParagraph0 = new ComparisonParagraph();
			sourceParagraph0.Text = "This is first cell in source Row.";
			sourceRow[0].TextFrame.Paragraphs.Add(sourceParagraph0);
			ComparisonParagraphBase sourceParagraph1 = new ComparisonParagraph();
			sourceParagraph1.Text = "This is second cell in source Row.";
			sourceRow[1].TextFrame.Paragraphs.Add(sourceParagraph1);
			Console.WriteLine("New Row was created.");

			ComparisonRowBase targetRow = new ComparisonRow(new double[] {200, 200}, 50);
			ComparisonParagraphBase targetParagraph0 = new ComparisonParagraph();
			targetParagraph0.Text = "This is first cell in target Row.";
			targetRow[0].TextFrame.Paragraphs.Add(targetParagraph0);
			ComparisonParagraphBase targetParagraph1 = new ComparisonParagraph();
			targetParagraph1.Text = "This is second cell in target Row.";
			targetRow[1].TextFrame.Paragraphs.Add(targetParagraph1);
			Console.WriteLine("New Row was created.");

			// Creating settings for comparison of Rows
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Rows
			ISlidesCompareResult compareResult = sourceRow.CompareWith(targetRow, SlidesComparisonSettings);
			Console.WriteLine("Rows was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingRows/result.pptx";
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