using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace GroupDocs.Comparison.Samples
{
	internal class CompareTwoTextFiles
	{
		public static void CompareTwoTextFilesFromStreamsWithSavingToFileAndSettings()
		{
			string sourcePath = @"GroupDocs.Comparison.Samples.data.Text.source.txt";
			string targetPath = @"GroupDocs.Comparison.Samples.data.Text.target.txt";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFileAndSettings/result.txt";

			// Create two streams of TextFiles
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of textFiles has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Text,
				new TextComparisonSettings());
			Console.WriteLine("TextFiles has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}
	}
}