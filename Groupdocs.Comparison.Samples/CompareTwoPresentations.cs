using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common;
using GroupDocs.Comparison.Common.ComparisonSettings;

namespace Groupdocs.Comparison.Samples
{
	internal class CompareTwoPresentations
	{
		public static void CompareTwoPresentationsFromStreamsWithSavingToFileAndSettings()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.data.Slides.source.pptx";
			string targetPath = @"Groupdocs.Comparison.Samples.data.Slides.target.pptx";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFileAndSettings/result.pptx";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of presentations has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Slides,
				new SlidesComparisonSettings());
			Console.WriteLine("Presentations has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoPresentationsFromFilesWithSavingToFileAndSettings()
		{
			string sourcePath = @"./../../data/Slides/source.pptx";
			string targetPath = @"./../../data/Slides/target.pptx";
			string resultPath = @"./../../testresult/FromFilesWithSavingToFileAndSettings/result.pptx";

			Console.WriteLine("Comparison of presentations has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourcePath, targetPath, resultPath, ComparisonType.Slides,
				new SlidesComparisonSettings());
			Console.WriteLine("Presentations has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}
	}
}