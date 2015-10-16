namespace Groupdocs.Comparison.Samples
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			//Word
			CompareTwoDocuments.CompareTwoDocumentsFromFiles();
			CompareTwoDocuments.CompareTwoDocumentsFromFilesWithSavingToFile();
			CompareTwoDocuments.CompareTwoDocumentsFromFilesWithSavingToFileAndSettings();
			CompareTwoDocuments.CompareTwoDocumentsFromFilesWithSettings();

			CompareTwoDocuments.CompareTwoDocumentsFromStreams();
			CompareTwoDocuments.CompareTwoDocumentsFromStreamsWithSavingToFile();
			CompareTwoDocuments.CompareTwoDocumentsFromStreamsWithSavingToFileAndSettings();
			CompareTwoDocuments.CompareTwoDocumentsFromStreamsWithSettings();

			//Slides
			CompareTwoPresentations.CompareTwoPresentationsFromStreamsWithSavingToFileAndSettings();
			CompareTwoPresentations.CompareTwoPresentationsFromFilesWithSavingToFileAndSettings();

			//Text
			CompareTwoTextFiles.CompareTwoTextFilesFromStreamsWithSavingToFileAndSettings();

			//Pdf
			CompareTwoPdfDocuments.CompareTwoPdfDocumentsFromStreamsWithSavingToFileAndSettings();

			//Cells
			CompareTwoWorkbooks.CompareTwoWorkbooksFromStreamsWithSavingToFileAndSettings();

			//Console.ReadKey();
		}
	}
}