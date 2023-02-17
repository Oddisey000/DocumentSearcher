/*
 * Created by SharpDevelop.
 * User: pevi5001
 * Date: 1/30/2023
 * Time: 7:44 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using DocumentSearcher.Classes;

namespace DocumentSearcher
{
	class Program
	{
		public static void Main(string[] args)
		{	
			var watch = System.Diagnostics.Stopwatch.StartNew();
			DocumentExtraction document = new DocumentExtraction();
			WordSearch document_test = new WordSearch();
			document.ExtractDocumentWord(@"I:\Documentation-Dokumente\02_Leoni AG & WSD Documentation\04_Work instructions (AA)", "Leoni");
			document_test.CleanUP();
			
			Console.Write(System.Environment.NewLine);
			Console.ForegroundColor = ConsoleColor.White;
			Console.WriteLine("Завершено");
			
			watch.Stop();
			var elapsedMs = watch.ElapsedMilliseconds;
			Console.WriteLine(elapsedMs);
			Console.ReadKey(true);
		}
	}
}