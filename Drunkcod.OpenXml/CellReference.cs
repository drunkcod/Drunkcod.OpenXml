using System;

namespace Drunkcod.OpenXml
{
	public struct CellReference
    {
		const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		readonly int row, column;

		public CellReference(int row, int column) {
			if(row <= 0 || column <= 0)
				throw new InvalidOperationException($"Invalid row or column ({row},{column})");
			this.row = row;
			this.column = column;
		}

		public bool IsEmpty => row == 0 && column == 0;
		public uint RowIndex => (uint)row;

		public CellReference AddRows(int rows) => new CellReference(row + rows, column);

		public override string ToString() {
			if(IsEmpty)
				return string.Empty;
			var r = new char[4];
			var n = 4;
			int part;
			var c = column - 1;
			do {
				c = Math.DivRem(c, Alphabet.Length, out part);
				r[--n] = Alphabet[part];
			} while(c > 0);
			return new string(r, n, r.Length - n) + row;
		}
	}
}
