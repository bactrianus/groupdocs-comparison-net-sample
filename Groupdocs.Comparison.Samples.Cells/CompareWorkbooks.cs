using System;
using System.Diagnostics;
using GroupDocs.Comparison.Cells.Contracts.Nodes;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace Groupdocs.Comparison.Samples.Cells
{
	internal class CompareWorkbooks
	{
		public static void CompareTwoWorkbooks(ComparisonWorkbookBase source, ComparisonWorkbookBase target, string filePath,
			string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two workbooks with the same quantity of worksheets");
			var result = source.CompareWith(target, settings);
			Console.WriteLine("Workbooks were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareTwoWorkbooksWithDeletedWorksheets(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two workbooks where target document doesn't contain some worksheets from source workbook");
			var result = source.CompareWith(target, settings);
			Console.WriteLine("Workbooks were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareTwoWorkbooksWithAddedWorksheets(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two workbooks where target workbook contains some worksheets that are not presented in source workbook");
			var result = source.CompareWith(target, settings);
			Console.WriteLine("Workbooks were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareTwoWorkbooksWithDeletedAndAddedWorksheets(ComparisonWorkbookBase source,
			ComparisonWorkbookBase target, string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two workbooks where both of them contain worksheets that are not presented in the opponent document");
			var result = source.CompareWith(target, settings);
			Console.WriteLine("Workbooks were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareTwoBigWorkbooks(ComparisonWorkbookBase source, ComparisonWorkbookBase target, string filePath
			/*, string imagePath*/)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two workbooks with big amount of worksheet and cells in there");
			Stopwatch watch = new Stopwatch();
			watch.Start();
			var result = source.CompareWith(target, settings);
			Console.WriteLine("Workbooks were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to file");

			watch.Stop();
			Console.WriteLine("Compared for - " + watch.ElapsedMilliseconds + "milliseconds\n\n");
		}
	}
}