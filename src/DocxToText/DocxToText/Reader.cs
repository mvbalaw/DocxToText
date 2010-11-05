#region License

// Copyright (c) McCreary, Veselka, Bragg, & Allen, P.C.
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

using ICSharpCode.SharpZipLib.Zip;

namespace DocxToText
{
	public class Reader
	{
		public string GetTextFromFile(string docxFileName)
		{
			if (!File.Exists(docxFileName))
			{
				throw new ArgumentException("file " + docxFileName + " does not exist.");
			}
			LoadDocumentXml(docxFileName);
			throw new NotImplementedException();
		}

		private void LoadDocumentXml(string docxFileName)
		{
			using (var zip = new ZipFile(docxFileName))
			{
			}
		}
	}
}