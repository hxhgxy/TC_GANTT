using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worksheet
{
	internal class Range
	{
	}

	[Serializable]
	public struct RangePosition
	{
		public int Row { get; set; }
		public int Col { get; set; }
		public int Rows { get; set; }
		public int Cols { get; set; }

		public int EndRow
		{
			get => Row + Rows - 1;
			set => Rows = value - Row + 1;
		}

		public int EndCol
		{
			get => Col + Cols - 1;
			set => Cols = value - Col + 1;
		}

		public CellPosition StartPos
		{
			get => new CellPosition(Row, Col);
			set
			{
				Row = value.Row;
				Col = value.Col;
			}
		}

		public CellPosition EndPos
		{
			get => new CellPosition(EndRow, EndCol);
			set
			{
				EndRow = value.Row;
				EndCol = value.Col;
			}
		}

		public RangePosition(CellPosition startPos, CellPosition endPos)
		{
			Row = Math.Min(startPos.Row, endPos.Row);
			Col = Math.Min(startPos.Col, endPos.Col);
			Rows = Math.Max(startPos.Row, endPos.Row) - Row + 1;
			Cols = Math.Max(startPos.Col, endPos.Col) - Col + 1;
		}

		public RangePosition(int row, int col, int rows, int cols)
		{
			Row = row;
			Col = col;
			Rows = rows;
			Cols = cols;
		}

		public bool Contains(CellPosition pos) => Contains(pos.Row, pos.Col);

		public bool Contains(int row, int col) =>
		    row >= Row && col >= Col && row <= EndRow && col <= EndCol;

		public override bool Equals(object obj) =>
		    obj is RangePosition range && Row == range.Row && Col == range.Col &&
		    Rows == range.Rows && Cols == range.Cols;

		//public override int GetHashCode() => HashCode.Combine(Row, Col, Rows, Cols);

		public static bool operator ==(RangePosition left, RangePosition right) => left.Equals(right);

		public static bool operator !=(RangePosition left, RangePosition right) => !(left == right);

		public override string ToString() => $"{StartPos}:{EndPos}";
	}

}
