using System;
using System.Drawing;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples.Pdf
{
	public static class ComparisonUsingSettings
	{
		public static void DefaultSettings(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"DefaultSettings - filePDFSourceWithCountPages5_pdf was compared with filePDFTargetWithCountPages5_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFSourceWithCountPages5.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFTargetWithCountPages5.pdf";

			Comparing.ProcessComparing("Comparison pdf using DefaultSettings: ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void GenerationSummaryPageIsFalse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"GenerateSummaryPage = false - filePDFSourceWithCountPages5_pdf was compared with filePDFTargetWithCountPages5_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFSourceWithCountPages5.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFTargetWithCountPages5.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				GenerateSummaryPage = false
			};

			Comparing.ProcessComparing("Comparison pdf using setting of GenerateSummaryPage equal false: ", SourceFileName,
				TargetFileName, ResultPath, settings);
		}

		public static void InsertedItemsStyleIsDefined(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"InsertedItemsStyleIsDefined - filePDFSourceWithCountPages5_pdf was compared with filePDFTargetWithCountPages5_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFSourceWithCountPages5.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFTargetWithCountPages5.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				InsertedItemsStyle = new StyleSettings()
				{
					Color = Color.Green,
					Size = 20,
					BeginSeparatorString = "[",
					EndSeparatorString = "]"
				}
			};

			Comparing.ProcessComparing("Comparison pdf using setting of InsertedItemsStyle: ", SourceFileName, TargetFileName,
				ResultPath, settings);
		}

		public static void DeletedItemsStyleIsDefined(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"DeletedItemsStyleIsDefined - filePDFSourceWithCountPages5_pdf was compared with filePDFTargetWithCountPages5_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFSourceWithCountPages5.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFTargetWithCountPages5.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				DeletedItemsStyle = new StyleSettings()
				{
					Color = Color.Tomato,
					BeginSeparatorString = "{",
					EndSeparatorString = "}"
				}
			};

			Comparing.ProcessComparing("Comparison pdf using setting of InsertedItemsStyle: ", SourceFileName, TargetFileName,
				ResultPath, settings);
		}

		public static void ShowDeletedContentIsFalse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"ShowDeletedContentIsFalse - filePDFSourceWithCountPages5_pdf was compared with filePDFTargetWithCountPages5_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFSourceWithCountPages5.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFTargetWithCountPages5.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				ShowDeletedContent = false
			};

			Comparing.ProcessComparing("Comparison pdf using setting of ShowDeletedContent equal false: ", SourceFileName,
				TargetFileName, ResultPath, settings);
		}

		public static void ComparisonDepthIsChars(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"ComparisonDepthIsChars - Source1_pdf was compared with Target1_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.Source1.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.Target1.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				ComparisonDepth = ComparisonDepth.Chars
			};

			Comparing.ProcessComparing("Comparison pdf using setting of ComparisonDepth equal Chars: ", SourceFileName,
				TargetFileName, ResultPath, settings);
		}

		public static void ComparisonDepthIsWords(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"ComparisonDepthIsWords - Source1_pdf was compared with Target1_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.Source1.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.Target1.pdf";

			PdfComparisonSettings settings = new PdfComparisonSettings()
			{
				ComparisonDepth = ComparisonDepth.Words
			};

			Comparing.ProcessComparing("Comparison pdf using setting of ComparisonDepth equal Words: ", SourceFileName,
				TargetFileName, ResultPath, settings);
		}
	}
}