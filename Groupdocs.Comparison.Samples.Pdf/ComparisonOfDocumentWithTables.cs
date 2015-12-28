using System;

namespace GroupDocs.Comparison.Samples.Pdf
{
	public static class ComparisonOfDocumentWithTables
	{
		public static void Case1(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"filePDFWithTable1_pdf was compared with filePDFWithTable2_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable1.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable2.pdf";

			Comparing.ProcessComparing("Comparison of a document with tables (case 1): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case1_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"filePDFWithTable2_pdf was compared with filePDFWithTable1_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable2.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable1.pdf";

			Comparing.ProcessComparing("Comparison of a document with tables (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case2(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"filePDFWithTable1_pdf was compared with filePDFWithTable3_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable1.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable3.pdf";

			Comparing.ProcessComparing("Comparison of a document with tables (case 2): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case2_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir,
				"filePDFWithTable3_pdf was compared with filePDFWithTable1_pdf.pdf");
			string SourceFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable3.pdf";
			string TargetFileName = "GroupDocs.Comparison.Samples.Pdf.data.filePDFWithTable1.pdf";

			Comparing.ProcessComparing("Comparison of a document with tables (case 2, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}
	}
}