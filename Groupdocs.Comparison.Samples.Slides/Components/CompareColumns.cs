using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Slides;
using GroupDocs.Comparison.Slides.Contracts;
using GroupDocs.Comparison.Slides.Contracts.Comparison;
using GroupDocs.Comparison.Slides.Contracts.Enums;

namespace Groupdocs.Comparison.Samples.Slides.Components
{
	public static class CompareColumns
	{
		public static void CompareColumnsFromDifferentPresentations()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareColumnsFromDifferentPresentations.old.pptx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareColumnsFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first Column from source presentation
			ComparisonColumnBase sourceColumn = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Columns[0];
			// Getting first Column from target presentation
			ComparisonColumnBase targetColumn = (targetPresentation.Slides[0].Shapes[0] as ComparisonTableBase).Columns[0];

			// Creating settings for comparison of Columns
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Columns
			ISlidesCompareResult compareResult = sourceColumn.CompareWith(targetColumn, SlidesComparisonSettings);
			Console.WriteLine("Columns was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareColumnsFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareColumnsFromOnePresentations()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareColumnsFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Column from source presentation
			ComparisonColumnBase sourceColumn = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Columns[0];
			// Getting second Column from source presentation
			ComparisonColumnBase targetColumn = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Columns[1];

			// Creating settings for comparison of Columns
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Columns
			ISlidesCompareResult compareResult = sourceColumn.CompareWith(targetColumn, SlidesComparisonSettings);
			Console.WriteLine("Columns was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareColumnsFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareColumnFromPresentationsWithCreatingColumn()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareColumnFromPresentationsWithCreatingColumn.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Column from source presentation
			ComparisonColumnBase sourceColumn = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase).Columns[0];
			// Creating Column
			ComparisonColumnBase targetColumn = new ComparisonColumn(new double[] {50, 50}, 200);
			targetColumn[0].TextFrame.Paragraphs[0].Text = "This is first cell in Column that was created.";
			targetColumn[1].TextFrame.Paragraphs[0].Text = "This is second cell in Column that was created.";
			Console.WriteLine("New Column was created.");

			// Creating settings for comparison of Columns
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Columns
			ISlidesCompareResult compareResult = sourceColumn.CompareWith(targetColumn, SlidesComparisonSettings);
			Console.WriteLine("Columns was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareColumnFromPresentationsWithCreatingColumn/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingColumns()
		{
			// Creating Columns
			ComparisonColumnBase sourceColumn = new ComparisonColumn(new double[] {50, 50}, 200);
			ComparisonParagraphBase sourceParagraph0 = new ComparisonParagraph();
			sourceParagraph0.Text = "This is first cell in source Column.";
			sourceColumn[0].TextFrame.Paragraphs.Add(sourceParagraph0);
			ComparisonParagraphBase sourceParagraph1 = new ComparisonParagraph();
			sourceParagraph1.Text = "This is second cell in source Column.";
			sourceColumn[1].TextFrame.Paragraphs.Add(sourceParagraph1);
			Console.WriteLine("New Column was created.");

			ComparisonColumnBase targetColumn = new ComparisonColumn(new double[] {50, 50}, 200);
			ComparisonParagraphBase targetParagraph0 = new ComparisonParagraph();
			targetParagraph0.Text = "This is first cell in target Column.";
			targetColumn[0].TextFrame.Paragraphs.Add(targetParagraph0);
			ComparisonParagraphBase targetParagraph1 = new ComparisonParagraph();
			targetParagraph1.Text = "This is second cell in target Column.";
			targetColumn[1].TextFrame.Paragraphs.Add(targetParagraph1);
			Console.WriteLine("New Column was created.");

			// Creating settings for comparison of Columns
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Columns
			ISlidesCompareResult compareResult = sourceColumn.CompareWith(targetColumn, SlidesComparisonSettings);
			Console.WriteLine("Columns was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingColumns/result.pptx";
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