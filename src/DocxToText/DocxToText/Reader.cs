#region License

// Copyright (c) McCreary, Veselka, Bragg & Allen, P.C.
// 
// This source code is subject to terms and conditions of the MIT License. 
// A copy of the license can be found in the License.txt file at the root 
// of this distribution. By using this source code in any fashion, you are 
// agreeing to be bound by the terms of the MIT License. 
// 
// You must not remove this notice from this software.

#endregion

using System;
using System.IO;
using System.Linq;
using System.Text;

using ICSharpCode.SharpZipLib.Zip;

using LinqToHtml;

namespace DocxToText
{
	public class Reader
	{
		private static string ExtractTextFromDocument(string text)
		{
			int tagStart = text.IndexOf('<');
			var document = HTMLParser.Parse(text.Substring(tagStart));
			var content = new StringBuilder();
			document
				.DescendantTags
				.Where(x => x.Type == "w:t")
				.Select(x => x.Content)
				.ToList()
				.ForEach(x => content.Append(x));
			return content.ToString();
		}

		private static ZipEntry FindTheDocumentXmlEntry(ZipFile zip, string documentXmlPath)
		{
			var entry = zip.Cast<ZipEntry>().FirstOrDefault(x => x.Name == documentXmlPath);
			if (entry != null)
			{
				return entry;
			}
			throw new ArgumentException("unable to find " + documentXmlPath + " in file " + zip.Name);
		}

		public string GetTextFromFile(string docxFileName)
		{
			if (!File.Exists(docxFileName))
			{
				throw new ArgumentException("file " + docxFileName + " does not exist.");
			}
			const string documentXml = "word/document.xml";
			string xml = LoadDocumentXml(docxFileName, documentXml);
			string content = ExtractTextFromDocument(xml);
			return content;
		}

		private static string LoadDocumentXml(string docxFileName, string documentXmlPath)
		{
			using (var zip = new ZipFile(docxFileName))
			{
				var entry = FindTheDocumentXmlEntry(zip, documentXmlPath);
				using (var inputStream = zip.GetInputStream(entry))
				{
					var bytes = new byte[entry.Size];
					inputStream.Read(bytes, 0, bytes.Length);
					string xml = Encoding.ASCII.GetString(bytes);
					return xml;
				}
			}
		}
	}
}