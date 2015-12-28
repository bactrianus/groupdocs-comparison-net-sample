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
	public static class CompareCells
	{
		public static void CompareCellsFromDifferentPresentations()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareCellsFromDifferentPresentations.old.pptx";
			string targetPath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareCellsFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first cell from source presentation
			ComparisonCellBase sourceCell = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase)[0, 0];
			// Getting first cell from target presentation
			ComparisonCellBase targetCell = (targetPresentation.Slides[0].Shapes[0] as ComparisonTableBase)[0, 0];

			// Creating settings for comparison of cells
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing cells
			ISlidesCompareResult compareResult = sourceCell.CompareWith(targetCell, SlidesComparisonSettings);
			Console.WriteLine("Cells was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareCellsFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareCellsFromOnePresentations()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.Slides.Components.data.CompareCellsFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first cell from source presentation
			ComparisonCellBase sourceCell = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase)[0, 0];
			// Getting last cell from source presentation
			ComparisonCellBase targetCell = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase)[2, 1];

			// Creating settings for comparison of cells
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing cells
			ISlidesCompareResult compareResult = sourceCell.CompareWith(targetCell, SlidesComparisonSettings);
			Console.WriteLine("Cells was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareCellsFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareCellFromPresentationsWithCreatingCell()
		{
			string sourcePath =
				@"GroupDocs.Comparison.Samples.Slides.Components.data.CompareCellFromPresentationsWithCreatingCell.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first cell from source presentation
			ComparisonCellBase sourceCell = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase)[0, 0];
			// Creating Cell
			ComparisonCellBase targetCell = new ComparisonCell(200, 50, 0, 0);
			ComparisonParagraphBase paragraph = new ComparisonParagraph();
			paragraph.Text = "This cell was created.";
			targetCell.TextFrame.Paragraphs.Remove(targetCell.TextFrame.Paragraphs[0]);
			targetCell.TextFrame.Paragraphs.Add(paragraph);
			Console.WriteLine("New cell was created.");

			// Creating settings for comparison of cells
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing cells
			ISlidesCompareResult compareResult = sourceCell.CompareWith(targetCell, SlidesComparisonSettings);
			Console.WriteLine("Cells was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareCellFromPresentationsWithCreatingCell/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingCells()
		{
			// Creating cells
			ComparisonCellBase sourceCell = new ComparisonCell(200, 50, 0, 0);
			ComparisonParagraphBase sourceParagraph = new ComparisonParagraph();
			sourceCell.TextFrame.Paragraphs.Remove(sourceCell.TextFrame.Paragraphs[0]);
			sourceParagraph.Text = "This is source cell.";
			sourceCell.TextFrame.Paragraphs.Add(sourceParagraph);
			Console.WriteLine("New cell was created.");

			ComparisonCellBase targetCell = new ComparisonCell(200, 50, 0, 0);
			ComparisonParagraphBase targetParagraph = new ComparisonParagraph();
			targetCell.TextFrame.Paragraphs.Remove(targetCell.TextFrame.Paragraphs[0]);
			targetParagraph.Text = "This is target cell.";
			targetCell.TextFrame.Paragraphs.Add(targetParagraph);
			Console.WriteLine("New cell was created.");

			// Creating settings for comparison of cells
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing cells
			ISlidesCompareResult compareResult = sourceCell.CompareWith(targetCell, SlidesComparisonSettings);
			Console.WriteLine("Cells was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingCells/result.pptx";
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