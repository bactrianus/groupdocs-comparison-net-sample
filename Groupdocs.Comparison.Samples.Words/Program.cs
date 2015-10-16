using Groupdocs.Comparison.Samples.Words.Components;
using Groupdocs.Comparison.Samples.Words.Documents;
using Groupdocs.Comparison.Samples.Words.Settings;

namespace Groupdocs.Comparison.Samples.Words
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			//Documents
			CompareDocumentsWithTables.WithTablesOnAppropriatePages();
			CompareDocumentsWithTables.WithTablesOnDefferentPages();
			CompareDocumentsWithTables.WithTablesWitchContainsDifferentCountOfColumns();
			CompareDocumentsWithTables.WithTablesWitchContainsDifferentCountOfColumnsAndDifferentCountOfRows();
			CompareDocumentsWithTables.WithTablesWitchContainsDifferentCountOfColumns();
			CompareDocumentsWithTables.WithTablesWitchContainsDifferentCountOfRows();

			CompareDocumentsWithText.WithTextOnAppropriatePages();
			CompareDocumentsWithText.WithTextOnDifferentPages();

			CompareDocumentsWithHyperlinks.WithIgnoreLinkSetting();
			CompareDocumentsWithHyperlinks.WithoutIgnoreLinkSetting();

			//Settings
			ComparisonWithDifferentSettings.CompareDocumentsWithComparisonDepthSetChars();
			ComparisonWithDifferentSettings.CompareDocumentsWithComparisonDepthSetWords();
			ComparisonWithDifferentSettings.CompareDocumentsWithGenerationSummaryPage();
			ComparisonWithDifferentSettings.CompareDocumentsWithOutGenerationSummaryPage();
			ComparisonWithDifferentSettings.CompareDocumentsWithOutShowDeletedContent();
			ComparisonWithDifferentSettings.CompareDocumentsWithShowDeletedContent();
			ComparisonWithDifferentSettings.CompareDocumentsWithWordsSepCharsSetSpace();
			ComparisonWithDifferentSettings.CompareDocumentsWithSettingStylesOnDelInsComponents();
            
            //Components
            CompareParagraphs.CompareParagraphsFromDifferentDocuments();
            CompareParagraphs.CompareParagraphFromDocumentWithCreatingParagraph();
            CompareParagraphs.CompareTwoCreatingParagraphs();

            CompareRows.CompareRowsFromDifferentDocuments();
            CompareRows.CompareRowFromDocumentWithCreatingRow();
            CompareRows.CompareTwoCreatingRows();

            CompareColumns.CompareColumnFromDocumentWithCreatingColumn();
            CompareColumns.CompareColumnsFromDifferentDocuments();
            CompareColumns.CompareTwoCreatingColumns();

            CompareCells.CompareCellFromDocumentWithCreatingCell();
		    CompareCells.CompareCellsFromDifferentDocuments();
            CompareCells.CompareTwoCreatingCells();

            CompareTables.CompareTableFromDocumentWithCreatingTable();
		    CompareTables.CompareTablesFromDifferentDocuments();
            CompareTables.CompareTwoCreatingTables();
            
		    //Console.ReadKey();
		}
	}
}