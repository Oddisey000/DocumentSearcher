/*
 * Created by SharpDevelop.
 * User: pevi5001
 * Date: 1/31/2023
 * Time: 8:41 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;

namespace DocumentSearcher.Classes
{
	/// <summary>
	/// Description of DocumentExtraction.
	/// </summary>
	public class DocumentExtraction
	{
		private WordSearch document = new WordSearch();
		public DocumentExtraction()
		{
		}
		
		public void ExtractDocumentWord(string folder, string keyword)
		{
			DocumentLocation documentList = new DocumentLocation(folder);
			ArrayExtraction(0, documentList.Word.Count - 1, documentList.Word, keyword);
			
			Console.ForegroundColor = ConsoleColor.Blue;
			document.ShowContent();
			
			Console.Write(System.Environment.NewLine);
			Console.ForegroundColor = ConsoleColor.Red;
			document.ShowErrors();
		}
		
		private void ArrayExtraction(int start, int end, ArrayList array, string keyword)
		{
			int counter = 0;
			int resultSize = 0;
			
			for (int i = start; i <= end; i++)
            {
            	document.OpenDocument(array[i].ToString(), keyword);
            	counter++;
            	Console.ForegroundColor = ConsoleColor.Green;
            	Console.WriteLine("Processing file " + counter + " of " + array.Count);
            	
            	if (document.Word.Count >= resultSize)
            	{
            		Console.ForegroundColor = ConsoleColor.Blue;
					document.ShowContent();
					resultSize++;
            	}
            }
		}
	}
}
