using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Worksheet
{
	internal struct GridRegion
	{
		internal int StartRow;
		internal int EndRow;
		internal int StartCol;
		internal int EndCol;

		internal static readonly GridRegion Empty = new GridRegion()
		{
			StartRow = 0,
			StartCol = 0,
			EndRow = 0,
			EndCol = 0
		};

		public GridRegion(int startRow, int startCol, int endRow, int endCol)
		{
			this.StartRow = startRow;
			this.StartCol = startCol;
			this.EndRow = endRow;
			this.EndCol = endCol;
		}

		//public bool Contains(CellPosition pos)
		//{
		//	return Contains(pos.Row, pos.Col);
		//}

		public bool Contains(int row, int col)
		{
			return this.StartRow <= row && this.EndRow >= row && this.StartCol <= col && this.EndCol >= col;
		}

		//public bool Contains(RangePosition range)
		//{
		//	return range.Row >= this.StartRow && range.Col >= this.StartCol
		//	    && range.EndRow <= this.EndRow && range.EndCol <= this.EndCol;
		//}

		//public bool Intersect(RangePosition range)
		//{
		//	return (range.Row < this.StartRow && range.EndRow > this.StartRow)
		//	    || (range.Row < this.EndRow && range.EndRow > this.EndRow)
		//	    || (range.Col < this.StartCol && range.EndCol > this.StartCol)
		//	    || (range.Col < this.EndCol && range.EndCol > this.EndCol);
		//}

		//public bool IsOverlay(RangePosition range)
		//{
		//	return Contains(range) || Intersect(range);
		//}

		//public override bool Equals(object obj)
		//{
		//	if ((obj as VisibleRegion?) == null) return false;
		//	VisibleRegion vr2 = (VisibleRegion)obj;
		//	return StartRow == vr2.StartRow && StartCol == vr2.StartCol
		//	    && EndRow == vr2.EndRow && EndCol == vr2.EndCol;
		//}

		public override int GetHashCode()
		{
			return StartRow ^ StartCol ^ EndRow ^ EndCol;
		}

		//public static bool operator ==(VisibleRegion vr1, VisibleRegion vr2)
		//{
		//	return vr1.Equals(vr2);
		//}

		//public static bool operator !=(VisibleRegion vr1, VisibleRegion vr2)
		//{
		//	return !vr1.Equals(vr2);
		//}

		public bool IsEmpty { get { return this.Equals(Empty); } }
		public int Rows { get { return EndRow - StartRow + 1; } set { EndRow = StartRow + value - 1; } }
		public int Cols { get { return EndCol - StartCol + 1; } set { EndCol = StartCol + value - 1; } }

		public override string ToString()
		{
			return string.Format("GridRegion[{0},{1}-{2},{3}]", StartRow, StartCol, EndRow, EndCol);
		}

		//public RangePosition ToRange()
		//{
		//	return new RangePosition(StartRow, StartCol, EndRow - StartRow + 1, EndCol - StartCol + 1);
		//}
	}
}
