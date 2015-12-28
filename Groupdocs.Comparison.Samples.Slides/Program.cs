using GroupDocs.Comparison.Samples.Slides.Components;
using GroupDocs.Comparison.Samples.Slides.Presentations;
using GroupDocs.Comparison.Samples.Slides.Settings;

namespace GroupDocs.Comparison.Samples.Slides
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			//Components
			CompareSlides.CompareSlidesFromDifferentPresentations();
			CompareSlides.CompareSlidesFromOnePresentations();
			CompareSlides.CompareSlideFromPresentationsWithCreatingSlide();
			CompareSlides.CompareTwoCreatingSlides();

			CompareCells.CompareCellsFromDifferentPresentations();
			CompareCells.CompareCellsFromOnePresentations();
			CompareCells.CompareCellFromPresentationsWithCreatingCell();
			CompareCells.CompareTwoCreatingCells();

			CompareParagraphs.CompareParagraphFromPresentationsWithCreatingParagraph();
			CompareParagraphs.CompareParagraphsFromDifferentPresentations();
			CompareParagraphs.CompareParagraphsFromOnePresentations();
			CompareParagraphs.CompareTwoCreatingParagraphs();

			CompareAutoShapes.CompareAutoShapeFromPresentationsWithCreatingAutoShape();
			CompareAutoShapes.CompareAutoShapesFromDifferentPresentations();
			CompareAutoShapes.CompareAutoShapesFromOnePresentations();
			CompareAutoShapes.CompareTwoCreatingAutoShapes();

			CompareTables.CompareTableFromPresentationsWithCreatingTable();
			CompareTables.CompareTablesFromDifferentPresentations();
			CompareTables.CompareTablesFromOnePresentations();
			CompareTables.CompareTwoCreatingTables();

			CompareRows.CompareRowFromPresentationsWithCreatingRow();
			CompareRows.CompareRowsFromDifferentPresentations();
			CompareRows.CompareRowsFromOnePresentations();
			CompareRows.CompareTwoCreatingRows();

			CompareColumns.CompareColumnFromPresentationsWithCreatingColumn();
			CompareColumns.CompareColumnsFromDifferentPresentations();
			CompareColumns.CompareColumnsFromOnePresentations();
			CompareColumns.CompareTwoCreatingColumns();

			// Presentations
			ComparePresentationsWithAutoShapes.ComparePresentationsWithAutoShapesOnAppropriateSlides();
			ComparePresentationsWithAutoShapes.ComparePresentationsWithAutoShapesOnDifferentSlides();

			ComparePresentationsWithTables.ComparePresentationsWithTablesOnAppropriateSlides();
			ComparePresentationsWithTables.ComparePresentationsWithTablesOnDifferentSlides();
			ComparePresentationsWithTables.ComparePresentationsWithTablesWitchConyainsDifferentCountOfRows();
			ComparePresentationsWithTables.ComparePresentationsWithTablesWitchConyainsDifferentCountOfColumns();
			ComparePresentationsWithTables
				.ComparePresentationsWithTablesWitchConyainsDifferentCountOfColumnsAndDifferentCountOfRows();

			//Settings
			ComparisonWithDifferentSettings.ComparePresentationsWithComparisonDepthSetChars();
			ComparisonWithDifferentSettings.ComparePresentationsWithComparisonDepthSetWords();
			ComparisonWithDifferentSettings.ComparePresentationsWithGenerationSummaryPage();
			ComparisonWithDifferentSettings.ComparePresentationsWithOutGenerationSummaryPage();
			ComparisonWithDifferentSettings.ComparePresentationsWithOutShowDeletedContent();
			ComparisonWithDifferentSettings.ComparePresentationsWithShowDeletedContent();
			ComparisonWithDifferentSettings.ComparePresentationsWithWordsSepCharsSetSpace();
			ComparisonWithDifferentSettings.ComparePresentationsWithSettingStylesOnDelInsComponents();

			//Console.ReadKey();
		}
	}
}