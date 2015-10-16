using System;

namespace Groupdocs.Comparison.Samples.Pdf
{
	public static class ResultComparingConsistsInsertedAndDeletedPages
	{
		public static void Case1(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF1_pdf was compared with filePDF3_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF3.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 1): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case1_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF3_pdf was compared with filePDF1_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF3.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case2(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF2_pdf was compared with filePDF3_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF3.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 2): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case2_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF3_pdf was compared with filePDF2_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF3.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 2, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case3(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF1_pdf was compared with filePDF5_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF5.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 3): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case3_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF5_pdf was compared with filePDF1_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF5.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 3, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case4(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF2_pdf was compared with filePDF5_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF5.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 4): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case4_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF5_pdf was compared with filePDF2_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF5.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";

			Comparing.ProcessComparing("Comparison with the odd pages (case 4, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}
	}
}