using System;
using GroupDocs.Comparison.Cells.Contracts.Nodes;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples.Cells
{
	internal class CompareWorksheets
	{
		public static void CompareWorksheetsFromOneDocuments(ComparisonWorkbookBase source, string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two worksheet from one workbook");
			var result = source.Worksheets[1].CompareWith(source.Worksheets[2], settings);
			Console.WriteLine("Worksheets were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to a file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareWorksheetsFromDifferentDocuments(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two worksheet from different workbooks");
			var result = source.Worksheets[1].CompareWith(target.Worksheets[1], settings);
			Console.WriteLine("Worksheets were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to a file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareEmptyWorksheetWithNotEmpty(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two worksheet first worksheet is empty");
			var result = source.Worksheets[3].CompareWith(target.Worksheets[1], settings);
			Console.WriteLine("Worksheets were successfully compared");
		    ComparisonWorkbookBase workbook = result.GetWorkbook();
			workbook.Save(filePath);
			Console.WriteLine("Results were successfully saved to a file");
			//result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareWorksheetsWithSimilarContent(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two worksheet with similar content");
			var result = source.Worksheets[1].CompareWith(target.Worksheets[1], settings);
			Console.WriteLine("Worksheets were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to a file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}

		public static void CompareWorksheetsWithDifferentContent(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filePath, string imagePath)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two worksheet with different content");
			var result = source.Worksheets[3].CompareWith(target.Worksheets[3], settings);
			Console.WriteLine("Worksheets were successfully compared");
			result.GetWorkbook().Save(filePath);
			Console.WriteLine("Results were successfully saved to a file");
			result.GetWorkbook().SaveAsImages(imagePath);
			Console.WriteLine("Results were successfully saved to a bunch of images\n\n");
		}
	}
}