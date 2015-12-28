using System;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using GroupDocs.Comparison.Cells.Contracts.Nodes;
using GroupDocs.Comparison.Cells.Nodes;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples.Cells
{
	public static class CompareWithSettings
	{
		public static void CompareWithDefaultSettings()
		{
			var settings = new CellsComparisonSettings();

			Console.WriteLine("Comparison of documents with default settings has been started");
			Compare("../../testresult/settings/Default settings/save.xlsx", settings);
		}

		public static void HideDeletedContent()
		{
			var settings = new CellsComparisonSettings();
			settings.ShowDeletedContent = false;

			Console.WriteLine("Comparison of documents without showing deleted content has been started");
			Compare("../../testresult/settings/Hide deleted content/save.xlsx", settings);
		}

		public static void DisableSummaryPageGeneration()
		{
			var settings = new CellsComparisonSettings();
			settings.GenerateSummaryPage = false;

			Console.WriteLine("Comparison of documents without summary page generation has been started");
			Compare("../../testresult/settings/Disable summary page/save.xlsx", settings);
		}

		public static void SetStylePropertiesForDeletedComponents()
		{
			var settings = new CellsComparisonSettings
			{
				DeletedItemsStyle =
				{
					Color = Color.DarkGreen,
					IsBold = true,
					IsItalic = true,
					IsStrikeout = true,
					IsUnderlined = true
				}
			};

			Console.WriteLine("Comparison of documents with settings for deleted content has been started");
			Compare("../../testresult/settings/Deleted properties/save.xlsx", settings);
		}

		public static void SetPropertiesForInsertedComponents()
		{
			var settings = new CellsComparisonSettings();

			settings.InsertedItemsStyle.Color = Color.Cyan;
			settings.InsertedItemsStyle.IsBold = true;
			settings.InsertedItemsStyle.IsItalic = true;
			settings.InsertedItemsStyle.IsStrikeout = true;
			settings.InsertedItemsStyle.IsUnderlined = true;

			Console.WriteLine("Comparison of documents with settings for inserted content has been started");
			Compare("../../testresult/settings/Inserted properties/save.xlsx", settings);
		}

		public static void SetSeparatorsForDeletedAndAddedComponents()
		{
			var settings = new CellsComparisonSettings();
			settings.DeletedItemsStyle.BeginSeparatorString = "<";
			settings.DeletedItemsStyle.EndSeparatorString = ">";

			settings.InsertedItemsStyle.BeginSeparatorString = "|";
			settings.InsertedItemsStyle.EndSeparatorString = "|";

			Console.WriteLine("Comparison of documents with separators for deleted and added content has been started");
			Compare("../../testresult/settings/Separators/save.xlsx", settings);
		}

		public static void SetSeparationCharsForText()
		{
			var settings = new CellsComparisonSettings();
			settings.WordsSeparatorChars = new[] {'/', '|', ' '};

			Console.WriteLine("Comparison of documents with text separation chars settings has been started");
			Compare("../../testresult/settings/Separation chars/save.xlsx", settings);
		}

		private static void Compare(string fileName, CellsComparisonSettings settings)
		{
			string target_settings_file_name = "GroupDocs.Comparison.Samples.Cells.data.settings_target.xlsx";

			string source_settings_file_name = "GroupDocs.Comparison.Samples.Cells.data.settings_source.xlsx";

			var assembly = Assembly.GetExecutingAssembly();

			var source_settings_file = assembly.GetManifestResourceStream(source_settings_file_name);

			var target_settings_file = assembly.GetManifestResourceStream(target_settings_file_name);

			ComparisonWorkbookBase source = new ComparisonWorkbook(source_settings_file);
			ComparisonWorkbookBase target = new ComparisonWorkbook(target_settings_file);
			var watch = new Stopwatch();
			watch.Start();

			var result = source.CompareWith(target, settings);

			Console.WriteLine("Saving result to file");

			result.GetWorkbook().Save(fileName);
			watch.Stop();

			Console.WriteLine("Compared for " + watch.ElapsedMilliseconds/1000 + " seconds\n\n");
		}
	}
}