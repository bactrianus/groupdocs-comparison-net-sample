using System;
using System.IO;

namespace Groupdocs.Comparison.Samples.Pdf
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			String dir = @"./Result";
			Directory.CreateDirectory(dir);

			ComparisonOfSimpleText.Case1(dir);
			ComparisonOfSimpleText.Case1_Reverse(dir);
			ComparisonOfSimpleText.Case2(dir);
			ComparisonOfSimpleText.Case2_Reverse(dir);
			ComparisonOfSimpleText.Case3(dir);
			ComparisonOfSimpleText.Case3_Reverse(dir);

			ComparisonUsingSettings.DefaultSettings(dir);
			ComparisonUsingSettings.GenerationSummaryPageIsFalse(dir);
			ComparisonUsingSettings.InsertedItemsStyleIsDefined(dir);
			ComparisonUsingSettings.DeletedItemsStyleIsDefined(dir);
			ComparisonUsingSettings.ShowDeletedContentIsFalse(dir);
			ComparisonUsingSettings.ComparisonDepthIsChars(dir);
			ComparisonUsingSettings.ComparisonDepthIsWords(dir);

			ResultComparingConsistsInsertedAndDeletedPages.Case1(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case1_Reverse(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case2(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case2_Reverse(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case3(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case3_Reverse(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case4(dir);
			ResultComparingConsistsInsertedAndDeletedPages.Case4_Reverse(dir);

			ComparisonOfDocumentWithTables.Case1(dir);
			ComparisonOfDocumentWithTables.Case1_Reverse(dir);
			ComparisonOfDocumentWithTables.Case2(dir);
			ComparisonOfDocumentWithTables.Case2_Reverse(dir);

			ComparisonOfDocumentWithXForms.Case1(dir);
			ComparisonOfDocumentWithXForms.Case1_Reverse(dir);
			ComparisonOfDocumentWithXForms.Case2(dir);
			ComparisonOfDocumentWithXForms.Case2_Reverse(dir);

            ComparisonOfDocumentWithXForms.DocumentsWithXFormsWhenDeletionAndInsertionRows(dir);
            ComparisonOfDocumentWithXForms.DocumentsWithXFormsWhenDeletionRowsWereMovedToNextNewPage(dir);

            //Console.ReadKey();
        }
	}
}