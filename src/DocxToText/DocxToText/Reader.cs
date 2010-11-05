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

using ICSharpCode.SharpZipLib.Zip;

namespace DocxToText
{
	public class Reader
	{
		private static ZipEntry FindTheDocumentXml(ZipFile zip, string documentXmlPath)
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
			LoadDocumentXml(docxFileName, documentXml);
			throw new NotImplementedException();
		}

		private static void LoadDocumentXml(string docxFileName, string documentXmlPath)
		{
			using (var zip = new ZipFile(docxFileName))
			{
				var document = FindTheDocumentXml(zip, documentXmlPath);
			}
		}
	}
}