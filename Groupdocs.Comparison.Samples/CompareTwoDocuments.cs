using System;
using System.IO;
using System.Reflection;
using GroupDocs.Comparison.Common;
using GroupDocs.Comparison.Common.ComparisonSettings;
using GroupDocs.Comparison.Words.Contracts;
using GroupDocs.Comparison.Words.Contracts.Enums;
using GroupDocs.Comparison.Words.Contracts.Nodes;
using GroupDocs.Comparison.Words.Nodes;

namespace Groupdocs.Comparison.Samples
{
	internal class CompareTwoDocuments
	{
		public static void CompareTwoDocumentsFromStreamsWithSavingToFileAndSettings()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.data.Words.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.data.Words.target.docx";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFileAndSettings/result.docx";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Words,
				new WordsComparisonSettings());
			Console.WriteLine("Documents has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoDocumentsFromStreamsWithSettings()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.data.Words.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.data.Words.target.docx";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, ComparisonType.Words,
				new WordsComparisonSettings());
			Console.WriteLine("Documents has been compared...");

			string resultPath = @"./../../testresult/FromStreamsWithSettings/result.docx";

			IComparisonDocument doc = new ComparisonDocument(result);
			doc.Save(resultPath, ComparisonSaveFormat.Docx);
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoDocumentsFromStreamsWithSavingToFile()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.data.Words.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.data.Words.target.docx";
			string resultPath = @"./../../testresult/FromStreamsWithSavingToFile/result.docx";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, resultPath, ComparisonType.Words);
			Console.WriteLine("Documents has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoDocumentsFromStreams()
		{
			string sourcePath = @"Groupdocs.Comparison.Samples.data.Words.source.docx";
			string targetPath = @"Groupdocs.Comparison.Samples.data.Words.target.docx";

			// Create two streams of documents
			Assembly assembly = Assembly.GetExecutingAssembly();
			Stream sourceStream = assembly.GetManifestResourceStream(sourcePath);
			Stream targetStream = assembly.GetManifestResourceStream(targetPath);

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourceStream, targetStream, ComparisonType.Words);
			Console.WriteLine("Documents has been compared...");

			string resultPath = @"./../../testresult/FromStreams/result.docx";
			IComparisonDocument doc = new ComparisonDocument(result);
			doc.Save(resultPath, ComparisonSaveFormat.Docx);
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

	    public static void CompareTwoDocumentsFromFilesWithSavingToFileAndSettings()
	    {
	        string sourcePath = @"./../../data/Words/source.docx";
	        string targetPath = @"./../../data/Words/target.docx";
	        string resultPath = @"./../../testresult/FromFilesWithSavingToFileAndSettings/result.docx";

	        Console.WriteLine("Comparison of documents has been started..");
	        GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
	        Stream result = comparison.Compare(sourcePath, targetPath, resultPath, ComparisonType.Words,
	            new WordsComparisonSettings());
	        Console.WriteLine("Documents has been compared...");
	        Console.WriteLine("Result has been saved to file " + resultPath + ".");
	    }

	    public static void CompareTwoDocumentsFromFilesWithSettings()
		{
			string sourcePath = @"./../../data/Words/source.docx";
			string targetPath = @"./../../data/Words/target.docx";

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourcePath, targetPath, ComparisonType.Words,
				new WordsComparisonSettings());
			Console.WriteLine("Documents has been compared...");

			string resultPath = @"./../../testresult/FromFilesWithSettings/result.docx";
			IComparisonDocument doc = new ComparisonDocument(result);
			doc.Save(resultPath, ComparisonSaveFormat.Docx);
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoDocumentsFromFilesWithSavingToFile()
		{
			string sourcePath = @"./../../data/Words/source.docx";
			string targetPath = @"./../../data/Words/target.docx";
			string resultPath = @"./../../testresult/FromFilesWithSettings/result.docx";

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourcePath, targetPath, resultPath, ComparisonType.Words);
			Console.WriteLine("Documents has been compared...");
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}

		public static void CompareTwoDocumentsFromFiles()
		{
			string sourcePath = @"./../../data/Words/source.docx";
			string targetPath = @"./../../data/Words/target.docx";

			Console.WriteLine("Comparison of documents has been started..");
            GroupDocs.Comparison.Comparison comparison = new GroupDocs.Comparison.Comparison();
            Stream result = comparison.Compare(sourcePath, targetPath, ComparisonType.Words);
			Console.WriteLine("Documents has been compared...");

			string resultPath = @"./../../testresult/FromFiles/result.docx";
			IComparisonDocument doc = new ComparisonDocument(result);
			doc.Save(resultPath, ComparisonSaveFormat.Docx);
			Console.WriteLine("Result has been saved to file " + resultPath + ".");
		}
	}
}