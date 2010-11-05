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

using DocxToText;

using FluentAssert;

using NUnit.Framework;

namespace Tests
{
	public class ReaderTests
	{
		[TestFixture]
		public class When_asked_to_find_the_document_in_a_docx_file
		{
			private Exception _exception;
			private string _fileName;
			private Reader _reader;

			[SetUp]
			public void BeforeEachTest()
			{
				_reader = new Reader();
			}

			[Test]
			public void Given_the_file_does_not_exist()
			{
				Test.Verify(
					with_a_docx_file_that_does_not_exist,
					when_asked_to_find_the_document_in_docx_file,
					should_throw_an_Argument_exception
					);
			}

			private void should_throw_an_Argument_exception()
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

			private void with_a_docx_file_that_does_not_exist()
			{
				_fileName = "invalid.docx";
			}
		}
	}
}