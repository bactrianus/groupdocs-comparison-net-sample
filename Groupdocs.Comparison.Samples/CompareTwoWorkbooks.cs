using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples
{
	public static class CompareTwoWorkbooks
	{
		public static void CompareTwoWorkbooksFromStreamsWithSavingToFileAndSettings()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.data.Cells.source.xlsx";
			string targetPath = @"GroupDocs.Comparison.Samples.data.Cells.target.xlsx";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFileAndSettings/result.xlsx";

			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of workbooks has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Cells,
				new WordsComparisonSettings());
			Console.WriteLine("Documents have been compared...");
			Console.WriteLine("Result has been saved to file\n\n");
		}

/*        public static void CompareTwoDocumentsFromStreams()
        {
            string sourcePath = @"GroupDocs.Comparison.Samples.data.Words.source.docx";
            string targetPath = @"GroupDocs.Comparison.Samples.data.Words.target.docx";

            // Create two streams of documents
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
            Stream targetStream = assembly.GetManifestResourceStream(targetPath);

            Console.WriteLine("Comparison of documents has been started..");
            Stream result = Comparison.Compare(sourceStream, targetStream, ComparisonType.Words);
            Console.WriteLine("Documents has been compared...");

            string resultPath = @"./../../testresult/FromStreams/result.docx";
            IComparisonWorkbook doc = new ComparisonWorkbook(result);
            doc.Save(resultPath, ComparisonSaveFormat.Docx);
            Console.WriteLine("Result has been saved to file " + resultPath + ".");
        }*/
	}
}