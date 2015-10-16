using System.Reflection;
using GroupDocs.Comparison.Cells.Nodes;

namespace Groupdocs.Comparison.Samples.Cells
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			string target_file_name = "Groupdocs.Comparison.Samples.Cells.data.cwt_target.xlsx";

			string source_file_name = "Groupdocs.Comparison.Samples.Cells.data.cwt_source.xlsx";

			string target_wb_file_name = "Groupdocs.Comparison.Samples.Cells.data.wbt_target.xlsx";

			string source_wb_file_name = "Groupdocs.Comparison.Samples.Cells.data.wbt_source.xlsx";

			string source_big_file_name = "Groupdocs.Comparison.Samples.Cells.data.source.xls";

			string target_big_file_name = "Groupdocs.Comparison.Samples.Cells.data.target.xls";

			Assembly assembly = Assembly.GetExecutingAssembly();

			var target_file = assembly.GetManifestResourceStream(target_file_name);

			var source_file = assembly.GetManifestResourceStream(source_file_name);

			var target_wb_file = assembly.GetManifestResourceStream(target_wb_file_name);

			var source_wb_file = assembly.GetManifestResourceStream(source_wb_file_name);

			var source_big_file = assembly.GetManifestResourceStream(source_big_file_name);

			var target_big_file = assembly.GetManifestResourceStream(target_big_file_name);

			var source_workbook = new ComparisonWorkbook(source_file);

			var target_workbook = new ComparisonWorkbook(target_file);

			var source_wbt_workbook = new ComparisonWorkbook(source_wb_file);

			var target_wbt_workbook = new ComparisonWorkbook(target_wb_file);

			var target_big_workbook = new ComparisonWorkbook(target_big_file);

			var source_big_workbook = new ComparisonWorkbook(source_big_file);

			// Comparing two cells in one document and one worksheet
			CompareCells.CompareCellsOnTheSameWorkbook(source_workbook,
				"../../testresult/cells/1. Compare Cells On The Same Workbook/save.xlsx");

			// Comparing two cells on different documents
			CompareCells.CompareCellsOnDifferentWorkbooks(source_workbook, target_workbook,
				"../../testresult/cells/2. Compare Cells On Different Workbooks/save.xlsx");

			// Comparing cell with formula with cell with text
			CompareCells.CompareFormulaCellWithTextCell(source_workbook, target_workbook,
				"../../testresult/cells/3. Compare Formula Cell With Text Cell/save.xlsx");

			// Comparing formula cell with numbeerd cell
			CompareCells.CompareFormulaCellWithNumberedCell(source_workbook, target_workbook,
				"../../testresult/cells/4. Compare Formula Cell With Numbered Cell/save.xlsx");

			// Comparing text cell with text cell
			CompareCells.CompareTextCellWithAnotherTextCell(source_workbook, target_workbook,
				"../../testresult/cells/5. Compare Text Cell With Another Text Cell/save.xlsx");

			// Comparing Styled cell with another styled cell
			CompareCells.CompareStyledCellWithAnotherStyledCell(source_workbook, target_workbook,
				"../../testresult/cells/6. Compare Styled Cell With Another Styled Cell/save.xlsx");

			// Comparing error cell with non error
			CompareCells.CompareErrorCellWithNonError(source_workbook, target_workbook,
				"../../testresult/cells/7. Compare Error Cell With NonError/save.xlsx");

			CompareWorksheets.CompareWorksheetsFromOneDocuments(
				source_workbook,
				"../../testresult/worksheets/1. Compare Worksheets From One Documents/save.xlsx",
				"../../testresult/worksheets/1. Compare Worksheets From One Documents/");
			CompareWorksheets.CompareWorksheetsFromDifferentDocuments(
				source_workbook,
				target_workbook,
				"../../testresult/worksheets/2. Compare Worksheets From Different Documents/save.xlsx",
				"../../testresult/worksheets/2. Compare Worksheets From Different Documents/");
			CompareWorksheets.CompareEmptyWorksheetWithNotEmpty(
				source_workbook,
				target_workbook,
				"../../testresult/worksheets/3. Compare Empty Worksheet With Not Empty/save.xlsx",
				"../../testresult/worksheets/3. Compare Empty Worksheet With Not Empty/");

			//Compare two workbooks
			CompareWorkbooks.CompareTwoWorkbooksWithDeletedAndAddedWorksheets(source_wbt_workbook, target_wbt_workbook,
				@"../../testresult/workbooks/two random workbooks/save.xlsx", @"../../testresult/workbooks/two random workbooks/");

			// Comparing workbooks with default settings
			CompareWithSettings.CompareWithDefaultSettings();

			// Disabling the generation of summary page
			CompareWithSettings.DisableSummaryPageGeneration();

			// Hiding deleted content
			CompareWithSettings.HideDeletedContent();

			// Setting style properties for inserted components
			CompareWithSettings.SetPropertiesForInsertedComponents();

			// Setting separation chars for text
			CompareWithSettings.SetSeparationCharsForText();

			// Setting separators for deleted and added text
			CompareWithSettings.SetSeparatorsForDeletedAndAddedComponents();

			// Setting style properties for deleted components
			CompareWithSettings.SetStylePropertiesForDeletedComponents();

			//Console.ReadKey();
		}
	}
}