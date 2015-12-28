using System;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples.Pdf
{
	public static class ComparisonOfDocumentWithXForms
	{
		public static void Case1(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "xForm1_pdf was compared with xForm2_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm1.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm2.pdf";

			Comparing.ProcessComparing("Comparison of a document with xForms (case 1): ", SourceFileName, TargetFileName,
				ResultPath, new PdfComparisonSettings {ComparisonDepth = ComparisonDepth.Words});
		}

		public static void Case1_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "xForm2_pdf was compared with xForm1_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm2.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm1.pdf";

			Comparing.ProcessComparing("Comparison of a document with xForms (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath, new PdfComparisonSettings {ComparisonDepth = ComparisonDepth.Words});
		}

		public static void Case2(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "xForm3_pdf was compared with xForm4_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm3.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm4.pdf";

			Comparing.ProcessComparing("Comparison of a document with xForms (case 2): ", SourceFileName, TargetFileName,
				ResultPath, new PdfComparisonSettings {ComparisonDepth = ComparisonDepth.Words});
		}

		public static void Case2_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "xForm4_pdf was compared with xForm3_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm4.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.xForm3.pdf";

			Comparing.ProcessComparing("Comparison of a document with xForms (case 2, reverse): ", SourceFileName, TargetFileName,
				ResultPath, new PdfComparisonSettings {ComparisonDepth = ComparisonDepth.Words});
		}

        public static void DocumentsWithXFormsWhenDeletionAndInsertionRows(string dir)
        {
            string ResultPath = String.Format(@"{0}/{1}", dir, "TargetWith2PagesFirstIs3_pdf was compared with SourceWith2PagesFirstIs3_pdf.pdf");
            string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.TargetWith2PagesFirstIs3.pdf";
            string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.SourceWith2PagesFirstIs3.pdf";

            PdfComparisonSettings comparisonSettings = new PdfComparisonSettings();

            Comparing.ProcessComparing("Comparison of a document with xForms contains deletion and insertion rows: ", SourceFileName, TargetFileName,
                ResultPath, new PdfComparisonSettings { ComparisonDepth = ComparisonDepth.Chars });
        }
        
        public static void DocumentsWithXFormsWhenDeletionRowsWereMovedToNextNewPage(string dir)
        {
            string ResultPath = String.Format(@"{0}/{1}", dir, "SourceWith2PagesFirstIs3_pdf was compared with TargetWith2PagesFirstIs3_pdf.pdf");
            string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.SourceWith2PagesFirstIs3.pdf";
            string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.TargetWith2PagesFirstIs3.pdf";

            PdfComparisonSettings comparisonSettings = new PdfComparisonSettings();

            Comparing.ProcessComparing("Comparison of a document with xForms when deletion rows were moved to next new page: ", SourceFileName, TargetFileName,
                ResultPath, new PdfComparisonSettings { ComparisonDepth = ComparisonDepth.Chars });
        }
    }
}