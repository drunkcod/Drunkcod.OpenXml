using DocumentFormat.OpenXml.Spreadsheet;

namespace Drunkcod.OpenXml
{
	public struct RowFragment
	{
		public readonly uint? RowIndex;
		public readonly Cell[] Cells;

		public RowFragment(Cell[] cells, uint? rowIndex) {
			this.Cells = cells;
			this.RowIndex = rowIndex;
		}

		public Row ToRow() {
			var r = new Row(Cells);
			if(RowIndex.HasValue)
				r.RowIndex = RowIndex;
			return r;
		}
	}
}
