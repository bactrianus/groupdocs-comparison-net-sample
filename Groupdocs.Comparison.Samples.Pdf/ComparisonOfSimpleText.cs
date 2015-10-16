using System;

namespace Groupdocs.Comparison.Samples.Pdf
{
	public static class ComparisonOfSimpleText
	{
		public static void Case1(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "Source1_pdf was compared with Target1_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.Source1.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.Target1.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case1_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "Target1_pdf was compared with Source1_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.Target1.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.Source1.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case2(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "Source2_pdf was compared with Target2_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.Source2.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.Target2.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case2_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "Target2_pdf was compared with Source2_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.Target2.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.Source2.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}

		public static void Case3(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF1_pdf was compared with filePDF2_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1): ", SourceFileName, TargetFileName, ResultPath);
		}

		public static void Case3_Reverse(string dir)
		{
			string ResultPath = String.Format(@"{0}/{1}", dir, "filePDF2_pdf was compared with filePDF1_pdf.pdf");
			string SourceFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF2.pdf";
			string TargetFileName = "Groupdocs.Comparison.Samples.Pdf.data.filePDF1.pdf";

			Comparing.ProcessComparing("Comparison of a simple text (case 1, reverse): ", SourceFileName, TargetFileName,
				ResultPath);
		}
	}
}