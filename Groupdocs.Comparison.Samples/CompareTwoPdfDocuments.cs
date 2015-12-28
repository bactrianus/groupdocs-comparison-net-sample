using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples
{
	internal class CompareTwoPdfDocuments
	{
		public static void CompareTwoPdfDocumentsFromStreamsWithSavingToFileAndSettings()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.data.Pdf.source.pdf";
			string targetPath = @"GroupDocs.Comparison.Samples.data.Pdf.target.pdf";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFileAndSettings/result.pdf";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of pdfDocuments has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Pdf,
				new PdfComparisonSettings());
			Console.WriteLine("PdfDocuments has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}
	}
}