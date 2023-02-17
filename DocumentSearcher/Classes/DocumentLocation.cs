/*
 * Created by SharpDevelop.
 * User: pevi5001
 * Date: 1/30/2023
 * Time: 7:46 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Collections;

namespace DocumentSearcher.Classes
{
	/// <summary>
	/// Description of Documents.
	/// </summary>
	class DocumentLocation
	{
		private ArrayList word = new ArrayList();
		private ArrayList excel = new ArrayList();
		private ArrayList powerPoint = new ArrayList();
		
		public ArrayList Word
		{
			get { return word; }
			set { Word = word; }
		}
		
		public ArrayList Excel
		{
			get { return excel; }
			set { Excel = excel; }
		}
		
		public ArrayList PowerPoint
		{
			get { return powerPoint; }
			set { PowerPoint = powerPoint; }
		}
		
		private void TraverseDirectory(DirectoryInfo directoryInfo, ArrayList arr)
		{
			var subdirectories = directoryInfo.EnumerateDirectories();
			
			foreach (var subdirectory in subdirectories)
			{
			    TraverseDirectory(subdirectory, arr);
			}
			
			var files = directoryInfo.EnumerateFiles();
			
			foreach (var file in files)
			{
			    HandleFile(file, arr);
			}
		}
		
		private void HandleFile(FileInfo file, ArrayList arr)
        {
			string wordExt = ".doc";
			if (file.Name.Contains(wordExt))
			{
				string el = file.Directory + "\\" + file.Name;
            	arr.Add(el);
			}
    	}
	
		
		public DocumentLocation(string path)
		{
			DirectoryInfo directory = new DirectoryInfo(path);
			TraverseDirectory(directory, this.word);
		}
	}
}
