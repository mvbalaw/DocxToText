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
using System.Reflection;

using DocxToText;

using FluentAssert;

using ICSharpCode.SharpZipLib.Zip;

using NUnit.Framework;

namespace Tests
{
	public class ReaderTests
	{
		[TestFixture]
		public class When_asked_to_find_the_document_in_a_docx_file
		{
			private const string FilePath = @".\temp.docx";

			private Exception _exception;
			private string _fileName;
			private Reader _reader;

			[SetUp]
			public void BeforeEachTest()
			{
				_reader = new Reader();
			}

			[TearDown]
			public void AfterEachTest()
			{
				if (File.Exists(FilePath))
				{
					File.Delete(FilePath);
				}
			}

			[Test]
			public void Given_the_Zip_archive_does_not_contain_a_document_xml_file()
			{
				Test.Verify(
					with_the_name_of_a_Zip_archive_that_does_not_contain_a_document_xml_file,
					when_asked_to_find_the_document_in_docx_file,
					should_throw_an_ArgumentException
					);
			}

			[Test]
			public void Given_the_file_does_not_exist()
			{
				Test.Verify(
					with_the_name_of_a_file_that_does_not_exist,
					when_asked_to_find_the_document_in_docx_file,
					should_throw_an_ArgumentException
					);
			}

			[Test]
			public void Given_the_file_is_not_Zip_compressed()
			{
				Test.Verify(
					with_the_name_of_a_file_that_is_not_Zip_compressed,
					when_asked_to_find_the_document_in_docx_file,
					should_throw_a_ZipException
					);
			}

			private static void CreateFileFromEmbeddedResource(string resourceName, string targetPath)
			{
				var asm = Assembly.GetExecutingAssembly();
				string name = asm.GetManifestResourceNames().First(x => x.EndsWith(resourceName));
				using (var stream = asm.GetManifestResourceStream(name))
				{
// ReSharper disable PossibleNullReferenceException
					var bytes = new byte[stream.Length];
// ReSharper restore PossibleNullReferenceException
					stream.Read(bytes, 0, bytes.Length);
					File.WriteAllBytes(targetPath, bytes);
				}
			}

			private void should_throw_a_ZipException()
			{
				_exception.ShouldNotBeNull("should have thrown an exception");
				_exception.GetType().ShouldBeEqualTo(typeof(ZipException));
			}

			private void should_throw_an_ArgumentException()
			{
				_exception.ShouldNotBeNull("should have thrown an exception");
				_exception.GetType().ShouldBeEqualTo(typeof(ArgumentException));
			}

			private void when_asked_to_find_the_document_in_docx_file()
			{
				try
				{
					_reader.GetTextFromFile(_fileName);
				}
				catch (Exception e)
				{
					_exception = e;
				}
			}

			private void with_the_name_of_a_Zip_archive_that_does_not_contain_a_document_xml_file()
			{
				CreateFileFromEmbeddedResource("empty.zip", FilePath);
				_fileName = FilePath;
			}

			private void with_the_name_of_a_file_that_does_not_exist()
			{
				_fileName = "invalid.docx";
			}

			private void with_the_name_of_a_file_that_is_not_Zip_compressed()
			{
				_fileName = FilePath;
				File.WriteAllText(FilePath, "this is a test");
			}
		}
	}
}