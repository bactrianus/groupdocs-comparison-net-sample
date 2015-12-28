using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Pdf;
using GroupDocs.Comparison.Pdf.Contracts.ComparedResult;

namespace GroupDocs.Comparison.Samples.Pdf
{
	public static class Comparing
	{
		public static void ProcessComparing(string title, string sourceFileName, string targetFileName, string resultPath,
			PdfComparisonSettings comparisonSettings = null)
		{
			Console.WriteLine(title);

			Stream resultStream = new MemoryStream();

			Assembly assembly = Assembly.GetExecutingAssembly();

			Stream sourceStream = assembly.GetManifestResourceStream(sourceFileName);
			Stream targetStream = assembly.GetManifestResourceStream(targetFileName);

			ComparisonPdfDocument sourcePdfDocument = new ComparisonPdfDocument(sourceStream);
			ComparisonPdfDocument targetPdfDocument = new ComparisonPdfDocument(targetStream);

			if (comparisonSettings == null) comparisonSettings = new PdfComparisonSettings();
			Console.WriteLine("");
			Console.WriteLine("Comparing...");

			IPdfComparedResult comparedPdfResult = sourcePdfDocument.CompareWith(targetPdfDocument, comparisonSettings);

			Console.WriteLine("End comparing!");
			Console.WriteLine("Saving Result...");
			comparedPdfResult.SaveTo(resultStream);

			Console.WriteLine("Result was saved to file: " + resultPath + ".");
			Console.WriteLine("===============================================================");
			Console.WriteLine("");

			using (var fsource = new FileStream(resultPath, FileMode.Create, FileAccess.Write))
			{
				resultStream.CopyTo(fsource);
			}
		}
	}
}