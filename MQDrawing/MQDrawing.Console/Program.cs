/*
 * Created by SharpDevelop.
 * User: 53785
 * Date: 2018/1/11
 * Time: 上午 10:14
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using CsvHelper;

namespace MQDrawing
{
	class Program
	{
		class DocumentFileInfo
		{
			public string DocumentNumber;
			public string Filename;
			public DateTime CreationTime;
		}
	
		
		const string MASTER_DOCUMENT_REGISTER_FILENAME = @".\MasterDocumentRegister.csv";
		const int MASTER_DOCUMENT_REGISTER_DOCUMENT_NUMBER_COLUMN_INDEX = 28;
		const int MASTER_DOCUMENT_REGISTER_DOCUMENT_TITLE_COLUMN_INDEX = 29;
		const int EXPORTED_DOCUMENT_NUMBER_INDEX = 6;
		const int EXPORTED_DOCUMENT_REVISION = 7;

		const string USER_INPUT_DIRECTORY = @".\Input";
		const string SYSTEM_OUTPUT_DIRECTORY = @".\Output";
		const string SYSTEM_TEMPLATE_DIRECTORY = @".\Template";
		
		static Dictionary<string, string> MasterDocumentRegister;
		static Dictionary<string, string> InputFileList;
		static SortedDictionary<string, DocumentFileInfo> InputFileInfo;
		
		public static void Main(string[] args)
		{
			Console.Write("== Mobile Query Drawing Brewer ==");
			Console.Write("\n== Liang, 2018/Jan/11          ==");
			MasterDocumentRegister = getMasterDocumentRegister(MASTER_DOCUMENT_REGISTER_FILENAME);
			InputFileList = getInputFileList(USER_INPUT_DIRECTORY);
			InputFileInfo = getInputFileInfo(USER_INPUT_DIRECTORY);
			copyInputFileToOutput(InputFileInfo);
			string htmlListContent = renderHtmlDataList(MasterDocumentRegister, InputFileList, InputFileInfo);
			
			StreamReader indexPageContent = File.OpenText(SYSTEM_TEMPLATE_DIRECTORY + @"\index.html");
			string htmlBody = indexPageContent.ReadToEnd();
			string outputIndexFilename = SYSTEM_OUTPUT_DIRECTORY + @"\index.html";
			File.WriteAllText(outputIndexFilename, string.Format(htmlBody, "iDocs+", htmlListContent));
			
			if (!Directory.Exists(SYSTEM_OUTPUT_DIRECTORY + @"\resource"))
				Directory.CreateDirectory(SYSTEM_OUTPUT_DIRECTORY + @"\resource");
			CloneDirectory(SYSTEM_TEMPLATE_DIRECTORY + @"\resource", SYSTEM_OUTPUT_DIRECTORY + @"\resource");
			
			Console.Write("\n== Done                        ==");
			Console.ReadKey(true);
		}
		
		static void copyInputFileToOutput(SortedDictionary<string, DocumentFileInfo> InputFiles)
		{
			Console.Write("\nCopy input files to output folder ");
			int index = 0;
			string documentPath = SYSTEM_OUTPUT_DIRECTORY + @"\documents\";
			if (!Directory.Exists(documentPath))
				Directory.CreateDirectory(documentPath);
			foreach (KeyValuePair<string, DocumentFileInfo> file in InputFiles) {
				if (File.Exists(documentPath + Path.GetFileName(file.Value.Filename)))
				    continue;
				File.Copy(file.Value.Filename, documentPath + Path.GetFileName(file.Value.Filename));
				if (index % 10 == 0)
					Console.Write(".");
				index++;
			}
		}
		static Dictionary<string, string> getInputFileList(string InputFileDirectory)
		{
			Console.Write("\nInitializing raw files list ...");
			Dictionary<string, string> result = new Dictionary<string, string>();
			
			string[] filenames = Directory.GetFiles(InputFileDirectory, "*export_report.csv", SearchOption.AllDirectories);
			foreach (string filename in filenames) {
				StreamReader txt = File.OpenText(filename);
				var csv = new CsvReader(txt);
				while (csv.Read()) {
					if (
						!result.ContainsKey(
							csv.GetField(
								EXPORTED_DOCUMENT_NUMBER_INDEX
							))) {
						string docNo = csv.GetField(EXPORTED_DOCUMENT_NUMBER_INDEX);
						docNo = docNo.Replace("192937-", "");
						string revision = csv.GetField(EXPORTED_DOCUMENT_REVISION);
						
						result.Add(docNo, revision);
					}
				}
			}
			return result;
		}
		static SortedDictionary<string, DocumentFileInfo> getInputFileInfo(string InputFileDirectory)
		{
			Console.Write("\nInitializing raw files ...");
			SortedDictionary<string, DocumentFileInfo> result = new SortedDictionary<string, DocumentFileInfo>();
			
			string[] filenames = Directory.GetFiles(InputFileDirectory, "*.pdf", SearchOption.AllDirectories);
			foreach (string filename in filenames) {
				
				if (filename.Contains("192937-")) {
					FileInfo file = new FileInfo(filename);
					string rawFilename = file.Name.ToUpper();
					string documentNo = rawFilename.Replace(".PDF", "").Replace("192937-", "");
					if (result.ContainsKey(documentNo)) {
						if (file.CreationTime > result[documentNo].CreationTime) {
							result[documentNo].Filename = file.FullName;
							result[documentNo].CreationTime = file.CreationTime;
						}
					} else {
						DocumentFileInfo doc = new DocumentFileInfo();
						doc.Filename = file.FullName;
						doc.CreationTime = file.CreationTime;
						doc.DocumentNumber = Path.GetFileNameWithoutExtension(file.Name).Replace("192937-", "");
						result.Add(documentNo, doc);						
					}

				}
				
			}
			return result;
		}
		static Dictionary<string,string> getMasterDocumentRegister(string Filename)
		{
			Console.Write("\nInitializing master document register ");
			
			Dictionary<string, string> result = new Dictionary<string, string>();
			StreamReader txt = File.OpenText(Filename);
			int index = 0;
			var csv = new CsvReader(txt);
			while (csv.Read()) {
				if (
					!result.ContainsKey(
						csv.GetField(
							MASTER_DOCUMENT_REGISTER_DOCUMENT_NUMBER_COLUMN_INDEX
						)))
					result.Add(
						csv.GetField(
							MASTER_DOCUMENT_REGISTER_DOCUMENT_NUMBER_COLUMN_INDEX
						),
						csv.GetField(
							MASTER_DOCUMENT_REGISTER_DOCUMENT_TITLE_COLUMN_INDEX
						));
				if (index % 2000 == 0)
					Console.Write(".");
				index++;
			}
			return result;
		}
		static string renderHtmlDataList(
			Dictionary<string, string> MasterDocument,
			Dictionary<string, string> InputFileList,
			SortedDictionary<string, DocumentFileInfo> InputFileInfo
		)
		{
			StringBuilder stb = new StringBuilder();
			string template = @"
				<li>
					<a rel=""external"" href=""./documents/{0}"">
						<h4>{1}</h4>
						<p>Title : {2}</p>
						<p>Revision : {3}</p>
					</a>
				</li>
			";
			foreach (KeyValuePair<string, DocumentFileInfo> item in InputFileInfo) {
				string row;
				if (MasterDocument.ContainsKey(item.Value.DocumentNumber) &&
				    InputFileList.ContainsKey(item.Value.DocumentNumber)) {
					row = string.Format(
						template,
						Path.GetFileName(item.Value.Filename),
						item.Value.DocumentNumber,
						MasterDocument[item.Key],
						InputFileList[item.Key]);
					stb.Append(row);
				}
				
			}
			return stb.ToString();
		}
		static void CloneDirectory(string root, string dest)
		{
			foreach (var directory in Directory.GetDirectories(root)) {
				string dirName = Path.GetFileName(directory);
				if (!Directory.Exists(Path.Combine(dest, dirName))) {
					Directory.CreateDirectory(Path.Combine(dest, dirName));
				}
				CloneDirectory(directory, Path.Combine(dest, dirName));
			}

			foreach (var file in Directory.GetFiles(root)) {
				if (File.Exists(Path.Combine(dest, Path.GetFileName(file))))
				    continue;
				File.Copy(file, Path.Combine(dest, Path.GetFileName(file)));
			}
		}
	}
}