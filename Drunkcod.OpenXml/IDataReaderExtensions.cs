using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace Drunkcod.OpenXml
{
	public static class IDataReaderExtensions
	{
		public static IEnumerable<RowFragment> ToCells(this IDataReader reader, CellReference? anchor = null) {
			var headers = new Cell[reader.FieldCount];
			for(var i = 0; i != headers.Length; ++i) {
				headers[i] = InlineString(reader.GetName(i));
				if(i == 0 && anchor.HasValue)
					headers[i].CellReference = anchor.Value.ToString();

			}
			yield return new RowFragment(headers, anchor.HasValue ? (uint?)anchor.Value.RowIndex : null);
			var values = new object[reader.FieldCount];
			var converters = new Converter<object, Cell>[reader.FieldCount];
			for(var i = 0; i != converters.Length; ++i)
				converters[i] = GetConverter(reader.GetFieldType(i));
			var rowOffset = 0;
			while(reader.Read()) {
				++rowOffset;
				var row = new Cell[reader.FieldCount];
				for(var i = 0; i != reader.FieldCount; ++i) {
					row[i] = converters[i](reader.GetValue(i));
					if(i == 0 && anchor.HasValue)
						row[i].CellReference = anchor.Value.AddRows(rowOffset).ToString();
				}
				yield return new RowFragment(row, null);
			}
		}

		static Converter<object, Cell> GetConverter(Type cellType) {
			switch(cellType.FullName) {
				default: throw new NotSupportedException($"Unsupported column type {cellType}");
				case "System.String": return InlineString;
				case "System.Int32": return IntegerCell;
				case "System.Double": return DoubleCell;
			}
		}

		static Cell InlineString(object value) => new Cell { DataType = CellValues.InlineString, InlineString = new InlineString(new Text(value.ToString())) };
		static Cell IntegerCell(object value) => new Cell { CellValue = new CellValue(value.ToString()) };
		static Cell DoubleCell(object value) => new Cell { CellValue = new CellValue(((double)value).ToString(CultureInfo.InvariantCulture)) };
	}
}
