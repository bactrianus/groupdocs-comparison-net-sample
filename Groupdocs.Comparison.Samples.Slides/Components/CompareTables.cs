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
	public static class CompareTables
	{
		public static void CompareTablesFromDifferentPresentations()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareTablesFromDifferentPresentations.old.pptx";
			string targetPath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareTablesFromDifferentPresentations.new.pptx";

			// Create to streams of presentations
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			// Opening two presentations
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");
			ComparisonPresentationBase targetPresentation = new ComparisonPresentation(targetStream);
			Console.WriteLine("Presentation with source path: " + targetPath + " was loaded.");

			// Getting first Table from source presentation
			ComparisonTableBase sourceTable = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase);
			// Getting first Table from target presentation
			ComparisonTableBase targetTable = (targetPresentation.Slides[0].Shapes[0] as ComparisonTableBase);

			// Creating settings for comparison of Tables
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Tables
			ISlidesCompareResult compareResult = sourceTable.CompareWith(targetTable, SlidesComparisonSettings);
			Console.WriteLine("Tables was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTablesFromDifferentPresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTablesFromOnePresentations()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.Slides.Components.data.CompareTablesFromOnePresentations.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Table from source presentation
			ComparisonTableBase sourceTable = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase);
			// Getting second Table from source presentation
			ComparisonTableBase targetTable = (sourcePresentation.Slides[1].Shapes[0] as ComparisonTableBase);

			// Creating settings for comparison of Tables
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Tables
			ISlidesCompareResult compareResult = sourceTable.CompareWith(targetTable, SlidesComparisonSettings);
			Console.WriteLine("Tables was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTablesFromOnePresentations/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTableFromPresentationsWithCreatingTable()
		{
			string sourcePath =
				@"Groupdocs.Comparison.Samples.Slides.Components.data.CompareTableFromPresentationsWithCreatingTable.old.pptx";

			// Create to stream of presentation
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);

			// Opening source presentation
			ComparisonPresentationBase sourcePresentation = new ComparisonPresentation(sourceStream);
			Console.WriteLine("Presentation with source path: " + sourcePath + " was loaded.");

			// Getting first Table from source presentation
			ComparisonTableBase sourceTable = (sourcePresentation.Slides[0].Shapes[0] as ComparisonTableBase);
			// Creating Table
			ComparisonTableBase targetTable = new ComparisonTable(100, 100, new double[] {200, 200}, new double[] {50, 50});
			targetTable[0, 0].TextFrame.Paragraphs[0].Text = "This is first cell in created table.";
			targetTable[0, 1].TextFrame.Paragraphs[0].Text = "This is second cell in created table.";
			targetTable[1, 0].TextFrame.Paragraphs[0].Text = "This is third cell in created table.";
			targetTable[1, 1].TextFrame.Paragraphs[0].Text = "This is fourth cell in created table.";
			Console.WriteLine("New Table was created.");

			// Creating settings for comparison of Tables
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Tables
			ISlidesCompareResult compareResult = sourceTable.CompareWith(targetTable, SlidesComparisonSettings);
			Console.WriteLine("Tables was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTableFromPresentationsWithCreatingTable/result.pptx";
			ComparisonPresentationBase result = compareResult.GetPresentation();
			Stream resultStream = new FileStream(resultPath, FileMode.Create);
			result.Save(resultStream, ComparisonSaveFormat.Pptx);
			resultStream.Close();
			Console.WriteLine("Result of comparison was saved to presentation with the folloving source path" +
			                  resultPath + ".");
			Console.WriteLine("===============================================");
			Console.WriteLine("");
		}

		public static void CompareTwoCreatingTables()
		{
			// Creating Tables
			ComparisonTableBase sourceTable = new ComparisonTable(100, 100, new double[] {200, 200}, new double[] {50, 50});
			ComparisonParagraphBase sourceParagraph00 = new ComparisonParagraph();
			sourceParagraph00.Text = "This is first cell in source table.";
			sourceTable[0, 0].TextFrame.Paragraphs.Add(sourceParagraph00);
			ComparisonParagraphBase sourceParagraph01 = new ComparisonParagraph();
			sourceParagraph01.Text = "This is second cell in source table.";
			sourceTable[0, 1].TextFrame.Paragraphs.Add(sourceParagraph01);
			ComparisonParagraphBase sourceParagraph10 = new ComparisonParagraph();
			sourceParagraph10.Text = "This is third cell in source table.";
			sourceTable[1, 0].TextFrame.Paragraphs.Add(sourceParagraph10);
			ComparisonParagraphBase sourceParagraph11 = new ComparisonParagraph();
			sourceParagraph11.Text = "This is fourth cell in source table.";
			sourceTable[1, 1].TextFrame.Paragraphs.Add(sourceParagraph11);
			Console.WriteLine("New Table was created.");

			ComparisonTableBase targetTable = new ComparisonTable(100, 100, new double[] {200, 200}, new double[] {50, 50});
			ComparisonParagraphBase targetParagraph00 = new ComparisonParagraph();
			targetParagraph00.Text = "This is first cell in target table.";
			targetTable[0, 0].TextFrame.Paragraphs.Add(targetParagraph00);
			ComparisonParagraphBase targetParagraph01 = new ComparisonParagraph();
			targetParagraph01.Text = "This is second cell in target table.";
			targetTable[0, 1].TextFrame.Paragraphs.Add(targetParagraph01);
			ComparisonParagraphBase targetParagraph10 = new ComparisonParagraph();
			targetParagraph10.Text = "This is third cell in target table.";
			targetTable[1, 0].TextFrame.Paragraphs.Add(targetParagraph10);
			ComparisonParagraphBase targetParagraph11 = new ComparisonParagraph();
			targetParagraph11.Text = "This is fourth cell in target table.";
			targetTable[1, 1].TextFrame.Paragraphs.Add(targetParagraph11);
			Console.WriteLine("New Table was created.");

			// Creating settings for comparison of Tables
			SlidesComparisonSettings SlidesComparisonSettings = new SlidesComparisonSettings();
			// Comparing Tables
			ISlidesCompareResult compareResult = sourceTable.CompareWith(targetTable, SlidesComparisonSettings);
			Console.WriteLine("Tables was compared.");

			// Saving result of comparison to new presentation
			string resultPath = @"./../../Components/testresult/CompareTwoCreatingTables/result.pptx";
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