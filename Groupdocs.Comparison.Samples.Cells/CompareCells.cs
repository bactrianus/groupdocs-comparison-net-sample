using System;
using GroupDocs.Comparison.Cells.Contracts.Nodes;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace Groupdocs.Comparison.Samples.Cells
{
	public class CompareCells
	{
		public static void CompareCellsOnTheSameWorkbook(ComparisonWorkbookBase source, string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two cells on the same workbook and the same worksheet");
			var result = source.Worksheets[0].CellRange["A1"].CompareWith(source.Worksheets[0].CellRange["B1"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareCellsOnDifferentWorkbooks(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two cells on different workbook and the same worksheet");
			var result = source.Worksheets[0].CellRange["A1"].CompareWith(target.Worksheets[0].CellRange["B1"], settings);
            Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareFormulaCellWithTextCell(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two cells on different workbook where source cell contains a formula and target cell contains plain text");
			var result = source.Worksheets[0].CellRange["A2"].CompareWith(target.Worksheets[0].CellRange["A2"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareFormulaCellWithNumberedCell(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two cells on different workbook where source cell contains a formula and target cell contains number");
			var result = source.Worksheets[0].CellRange["A3"].CompareWith(target.Worksheets[0].CellRange["A3"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareTextCellWithAnotherTextCell(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two cells on different workbook where source and target cells contain plain text");
			var result = source.Worksheets[0].CellRange["A4"].CompareWith(target.Worksheets[0].CellRange["A4"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareStyledCellWithAnotherStyledCell(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine(
				"Comparing two cells on different workbook where source and target cells contain plain text with different style of its characters");
			var result = source.Worksheets[0].CellRange["A5"].CompareWith(target.Worksheets[0].CellRange["A5"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}

		public static void CompareErrorCellWithNonError(ComparisonWorkbookBase source, ComparisonWorkbookBase target,
			string filename)
		{
			var settings = new CellsComparisonSettings();
			Console.WriteLine("Comparing two cells on different workbook where source and target cells contain plain text");
			var result = source.Worksheets[0].CellRange["A6"].CompareWith(target.Worksheets[0].CellRange["A6"], settings);
			Console.WriteLine("Comparison Done. Creating a file with result changes");
			result.GetWorkbook().Save(filename);
			Console.WriteLine("File with result changes was successfully created\n\n");
		}
	}
}