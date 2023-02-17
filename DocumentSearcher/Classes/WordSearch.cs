/*
 * Created by SharpDevelop.
 * User: pevi5001
 * Date: 1/30/2023
 * Time: 8:00 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;

namespace DocumentSearcher.Classes
{
	/// <summary>
	/// Description of WordSearch.
	/// </summary>
	class WordSearch
	{
		private string document;
		private Application application = new Application();
		public string Document
		{
			get { return document; }
			set { Document = document; }
		}
		private ArrayList word = new ArrayList();
		public ArrayList Word
		{
			get { return word; }
			set { Word = word; }
		}
		private ArrayList error = new ArrayList();
		public ArrayList Error
		{
			get { return error; }
			set { Error = error; }
		}
		
		public WordSearch() {}
		
		public void OpenDocument(string file, string keyword)
		{
			//string keyword = "для усіх матеріалів виробництва";
			keyword.ToLower();
			
			try {
				WordprocessingDocument document = WordprocessingDocument.Open(file,false);
				Body body = document.MainDocumentPart.Document.Body;
				string content = body.InnerText;
				content.ToLower();
				if (content.Contains(keyword))
				{
					this.word.Add(file);
				}
				document.Close();

			} catch (Exception) {
				
				SearchInCorruptedDocuments(file, keyword);
			}
		}
		
		public void SearchInCorruptedDocuments(string file, string keyword)
		{
			string fileName = file;
			
			try {
				Microsoft.Office.Interop.Word.Document document = application.Documents.Open(file,false,true);
			
				this.document = document.Content.Text.Trim();
				if (this.document.Contains(keyword))
				{
					this.word.Add(document.FullName);
				}
				((_Document)document).Close();
				
			} catch (Exception) {

				this.error.Add(fileName);
			}
		}
		
		public void CleanUP()
		{
			((_Application)application).Quit();
		}
		
		public void ShowContent()
		{
			foreach(var item in this.word )
			{
			  Console.WriteLine(item);
			}
		}
		
		public void ShowErrors()
		{
			foreach(var item in this.error )
			{
			  Console.WriteLine(item);
			}
		}
	}
}
