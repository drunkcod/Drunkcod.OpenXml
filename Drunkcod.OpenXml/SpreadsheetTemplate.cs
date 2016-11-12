using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace Drunkcod.OpenXml
{
	public class SpreadsheetTemplate : IDisposable
	{
		readonly MemoryStream bytes;
		public readonly SpreadsheetDocument Document;

		public void Dispose() => Document.Dispose();

		SpreadsheetTemplate(MemoryStream bytes, SpreadsheetDocument document) {
			this.bytes = bytes;
			this.Document = document;
		}

		public static SpreadsheetTemplate Load(string path) {
			var result = new MemoryStream();
			using(var tmp = File.OpenRead(path))
				tmp.CopyTo(result);
			result.Position = 0;
			var doc = SpreadsheetDocument.Open(result, true);
			return new SpreadsheetTemplate(result, doc);
		}

		public void SaveAs(string path) {
			bytes.Position = 0;
			using(var output = File.Create(path))
				bytes.CopyTo(output);
		}
	}
}
