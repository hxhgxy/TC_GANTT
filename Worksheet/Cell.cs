using System;
using System.Collections.Generic;
using System.Windows.Controls;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using Worksheet.Utility;

namespace Worksheet
{
	using System.Globalization;
	using System.Windows;
	using System.Windows.Controls.Primitives;
	using System.Windows.Controls;
	using System.Windows.Media;
	using System.Windows.Input;

	public class Cell : FrameworkElement
	{
		#region Hierarchy
		private Cell _parent;

		public new Cell Parent
		{
			get { return _parent; }
			set
			{
				if (_parent != value)
				{
					if (_parent != null)
					{
						_parent.Children.Remove(this);
						worksheet.Rows[Row].Children.Remove(RowHeader);
					}
					_parent = value;

					if (_parent != null)
					{
						_parent.Children.Add(this);
						if (!_parent.RowHeader.Children.Contains(this.RowHeader))
						{
							_parent.RowHeader.Children.Add(this.RowHeader);
							this.RowHeader.Parent = _parent.RowHeader;
						}
					}
				}
			}
		}

		private List<Cell> _children = new List<Cell>();
		public List<Cell> Children
		{
			get { return _children; }
			set
			{
				_children = value;
				_parent.RowHeader.Children.Clear();

				foreach (Cell child in _children)
				{
					child.Parent = this; // 设置每个子单元格的父单元格

					_parent.RowHeader.Children.Add(child.RowHeader);
					child.RowHeader.Parent = this.RowHeader;
				}
			}
		}

		public RowHeader RowHeader { get { return worksheet.Rows[Row]; } }

		private bool _isExpanded = true;

		public bool IsExpanded
		{
			get { return _isExpanded; }
			set
			{
				_isExpanded = value;

				this.RowHeader.IsExpanded = value;
			}
		}

		public double toggleX { get; private set; }
		public double toggleY { get; private set; }
		public double toggleButtonSize { get; private set; } = 10;

		#endregion // Hierarchy

		#region Constructor
		[NonSerialized]
		private Worksheet worksheet;
		public Worksheet Worksheet { get { return this.worksheet; } }
		//public Cell()
		//{
		//	//if (this.Children.Count > 0)
		//	//{
		//		toggleButton = new Button();
		//		toggleButton.Width = 10;
		//		toggleButton.Height = 10;
		//		toggleButton.Margin = new Thickness(2);
		//		toggleButton.Click += ToggleButton_Click;
		//		toggleButton.Content = "+";

		//		// 设置按钮的样式
		//		toggleButton.Style = CellStyles.ToggleButtonStyle;

		//		// 将按钮添加到 UI 树中
		//		AddVisualChild(toggleButton);
		//		AddLogicalChild(toggleButton);
		//	//}
		//}

		//private void ToggleButton_Click(object sender, RoutedEventArgs e)
		//{
		//	// 执行折叠/展开逻辑
		//	IsExpanded = !IsExpanded;
		//	toggleButton.Content = IsExpanded ? "-" : "+";
		//	InvalidateVisual(); // 重新绘制
		//}

		internal Cell(Worksheet worksheet)
		{
			this.worksheet = worksheet;
		}

		internal Cell(Worksheet worksheet, int row, int column)
		{
			this.worksheet = worksheet;
			this.InternalPos.Row = row;
			this.InternalPos.Col = column;
			//this.MouseDown += OnMouseDown;
			//this.MouseUp += OnMouseUp;
		}
		#endregion // Constructor

		#region Position
		internal CellPosition InternalPos;
		internal int InternalRow
		{
			get { return this.InternalPos.Row; }
			set { this.InternalPos.Row = value; }
		}
		internal int InternalCol
		{
			get { return this.InternalPos.Col; }
			set { this.InternalPos.Col = value; }
		}
		public int Row { get { return this.InternalPos.Row; } }
		public int Column { get { return this.InternalPos.Col; } }
		public CellPosition Position { get { return this.InternalPos; } }
		public string Address
		{
			get
			{
				return Utility.Utility.ToAddress(this.InternalPos.Row, this.InternalPos.Col);
			}
		}
		#endregion // Position

		#region Rowspan & Colspan
		private short colspan;
		internal short Colspan
		{
			get { return colspan; }
			set { colspan = value; }
		}
		private short rowspan;
		internal short Rowspan
		{
			get { return rowspan; }
			set { rowspan = value; }
		}
		public short GetColspan() { return colspan; }
		public short GetRowspan() { return rowspan; }
		#endregion // Rowspan & Colspan

		#region Location & Size
		[NonSerialized]
		private Rect bounds;
		internal Rect Bounds
		{
			get { return bounds; }
			set { bounds = value; }
		}
		internal double Width
		{
			get { return bounds.Width; }
			set { bounds.Width = value; }
		}
		internal double Height
		{
			get { return bounds.Height; }
			set { bounds.Height = value; }
		}
		internal double Top
		{
			get { return bounds.Y; }
			set { bounds.Y = value; }
		}
		internal double Left
		{
			get { return bounds.X; }
			set { bounds.X = value; }
		}
		internal double Right
		{
			get { return bounds.Right; }
			set { bounds.Width += bounds.Right - value; }
		}
		internal double Bottom
		{
			get { return bounds.Bottom; }
			set { bounds.Height += bounds.Bottom - value; }
		}
		#endregion // Location & Size

		#region Data, Display
		internal string InnerData { get; set; }
		public string Data
		{
			get { return InnerData; }
			set
			{
				if (this.worksheet != null)
				{
					this.worksheet.SetSingleCellData(this.Row, this.Column, value);
				}
				else
				{
					this.InnerData = value;
				}
			}
		}

		public string DisplayText { get { return InnerData?.ToString(); } }
		public bool IsReadOnly { get; set; }
		#endregion // Data and Display

		#region Style
		public Brush Background { get; set; } = Brushes.Transparent;
		public Brush Foreground { get; set; } = Brushes.Black;
		public Typeface Font { get; set; } = new Typeface(new FontFamily("Arial"), FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
		public double FontSize { get; set; } = 12;
		public double? BorderAllThickness { get; set; } = null;
		public Brush? BorderAllBrush { get; set; } = null;

		// 单独的边框属性
		public Brush BorderTopBrush { get; set; }
		public double BorderTopThickness { get; set; }

		public Brush BorderBottomBrush { get; set; }
		public double BorderBottomThickness { get; set; }

		public Brush BorderLeftBrush { get; set; }
		public double BorderLeftThickness { get; set; }

		public Brush BorderRightBrush { get; set; }
		public double BorderRightThickness { get; set; }
		public double Indent { get; set; } = 0;

		public TextAlignment TextAlignment { get; set; } = TextAlignment.Left;

		// 新增 SetStyle 方法
		public void SetStyle(string fontName, double fontSize, TextAlignment textAlignment, double indent, Brush foreground, FontWeight fontWeight)
		{
			//FontWeight fontweightTask = this.Children.Count > 0 ? FontWeights.Bold : fontWeight;

			this.Font = new Typeface(new FontFamily(fontName), FontStyles.Normal, fontWeight, FontStretches.Normal);
			this.FontSize = fontSize;
			this.TextAlignment = textAlignment;
			this.Indent = indent;
			this.Foreground = foreground;
		}
		#endregion // Style

		#region Border Wraps
		private CellBorderProperty borderProperty = null;
		public CellBorderProperty Border
		{
			get
			{
				if (borderProperty == null)
				{
					borderProperty = new CellBorderProperty(this);
				}
				return borderProperty;
			}
		}
		#endregion // Border Wraps

		#region Utility
		public override string ToString()
		{
			return "Cell[" + this.Address + "]";
		}
		#endregion // Utility

		#region Drawing
		public void Draw(DrawingContext dc, Rect cellRect, Visual visual)
		{

			// 绘制背景
			dc.DrawRectangle(Background, null, cellRect);

			if (BorderAllBrush != null && BorderAllThickness != null)
			{
				Pen borderPen = new Pen(BorderAllBrush, (double)BorderAllThickness);
				dc.DrawRectangle(null, borderPen, cellRect);
			}
			else
			{
				// 绘制边框
				if (BorderTopBrush != null && BorderTopThickness > 0)
				{
					Pen topPen = new Pen(BorderTopBrush, BorderTopThickness);
					dc.DrawLine(topPen, new Point(cellRect.Left, cellRect.Top), new Point(cellRect.Right, cellRect.Top));
				}

				if (BorderBottomBrush != null && BorderBottomThickness > 0)
				{
					Pen bottomPen = new Pen(BorderBottomBrush, BorderBottomThickness);
					dc.DrawLine(bottomPen, new Point(cellRect.Left, cellRect.Bottom), new Point(cellRect.Right, cellRect.Bottom));
				}

				if (BorderLeftBrush != null && BorderLeftThickness > 0)
				{
					Pen leftPen = new Pen(BorderLeftBrush, BorderLeftThickness);
					dc.DrawLine(leftPen, new Point(cellRect.Left, cellRect.Top), new Point(cellRect.Left, cellRect.Bottom));
				}

				if (BorderRightBrush != null && BorderRightThickness > 0)
				{
					Pen rightPen = new Pen(BorderRightBrush, BorderRightThickness);
					dc.DrawLine(rightPen, new Point(cellRect.Right, cellRect.Top), new Point(cellRect.Right, cellRect.Bottom));
				}
			}

			// 绘制折叠/展开按钮
			//double toggleButtonSize = 10; // 按钮的大小
			double toggleButtonMargin = 0; // 按钮和文本之间的间距
			double textX = cellRect.X + 2 + Indent; // 默认左对齐，加 2 个像素的内边距和缩进量

			toggleX = cellRect.X + 2 + Indent - 20;
			toggleY = cellRect.Y + (cellRect.Height - toggleButtonSize) / 2;

			if (Children.Count > 0)
			{
				// 计算按钮的位置
				Rect toggleButtonRect = new Rect(toggleX, cellRect.Y + (cellRect.Height - toggleButtonSize) / 2, toggleButtonSize, toggleButtonSize);

				// 绘制按钮背景
				dc.DrawRoundedRectangle(Brushes.LightGray, new Pen(Brushes.Gray, 1), toggleButtonRect, 2, 2);
				//dc.DrawRectangle(Brushes.LightGray, new Pen(Brushes.Gray, 1), toggleButtonRect);

				// 绘制按钮的加号或减号
				Pen togglePen = new Pen(Brushes.Black, 1);
				Point center = new Point(toggleButtonRect.Left + toggleButtonSize / 2, toggleButtonRect.Top + toggleButtonSize / 2);

				if (this.IsExpanded)
				{
					// 绘制减号
					dc.DrawLine(togglePen, new Point(center.X - 3, center.Y), new Point(center.X + 3, center.Y));
				}
				else
				{
					// 绘制加号
					dc.DrawLine(togglePen, new Point(center.X - 3, center.Y), new Point(center.X + 3, center.Y));
					dc.DrawLine(togglePen, new Point(center.X, center.Y - 3), new Point(center.X, center.Y + 3));
				}

				// 调整文本位置
				textX = toggleButtonRect.Right + toggleButtonMargin;
			}

			if (this.InternalCol == 1)
			{
				int tiger = 1;
			}

			// 绘制文本
			if (InnerData != null)
			{
				FormattedText formattedText = new FormattedText(
				    InnerData,
				    CultureInfo.CurrentCulture,
				    FlowDirection.LeftToRight,
				    Font,
				    FontSize,
				    Foreground,
				    VisualTreeHelper.GetDpi(visual).PixelsPerDip);

				// 根据对齐方式调整文本位置
				switch (TextAlignment)
				{
					case TextAlignment.Center:
						textX = cellRect.X + (cellRect.Width - formattedText.Width) / 2;
						break;
					case TextAlignment.Right:
						textX = cellRect.X + cellRect.Width - formattedText.Width - 2 - Indent;
						break;
					case TextAlignment.Left:
					default:
						textX = cellRect.X + 2 + Indent;
						break;
				}

				double textY = cellRect.Y + (cellRect.Height - formattedText.Height) / 2;
				dc.DrawText(formattedText, new Point(textX, textY));
			}
		}

		#endregion // Drawing
	}

	// 自定义事件参数类
	public class CellEventArgs : EventArgs
	{
		public Cell Cell { get; }

		public CellEventArgs(Cell cell)
		{
			Cell = cell;
		}
	}

	public struct CellPosition
	{
		public int Row { get; set; }
		public int Col { get; set; }

		public CellPosition(int row, int col)
		{
			Row = row;
			Col = col;
		}
		//public bool Equals(string address)
		//{
		//	return !string.IsNullOrEmpty(address) && IsValidAddress(address) && Equals(new CellPosition(address));
		//}
		public static bool operator ==(CellPosition r1, CellPosition r2)
		{
			return r1.Equals(r2);
		}
		public static bool operator !=(CellPosition r1, CellPosition r2)
		{
			return !r1.Equals(r2);
		}
		public static bool Equals(CellPosition pos1, CellPosition pos2)
		{
			return Equals(pos1, pos2.Row, pos2.Col);
		}
		public static bool Equals(CellPosition pos, int row, int col)
		{
			return pos.Row == row && pos.Col == col;
		}
		//public override int GetHashCode()
		//{
		//	return Row ^ Col ^ positionProperties;
		//}
	}

	public class CellBorderProperty
	{
		public Brush BorderBrush { get; set; }
		public Thickness BorderThickness { get; set; }

		public CellBorderProperty(Cell cell)
		{
			BorderBrush = Brushes.Black;
			BorderThickness = new Thickness(1);
		}
	}

	public static class CellStyles
	{
		public static Style ToggleButtonStyle;

		static CellStyles()
		{
			// 定义 ToggleButton 的样式
			ToggleButtonStyle = new Style(typeof(Button));
			ControlTemplate template = new ControlTemplate(typeof(Button));

			FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
			borderFactory.SetValue(Border.BackgroundProperty, Brushes.LightGray);
			borderFactory.SetValue(Border.BorderBrushProperty, Brushes.Gray);
			borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1));
			borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));

			FrameworkElementFactory contentPresenterFactory = new FrameworkElementFactory(typeof(ContentPresenter));
			contentPresenterFactory.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
			contentPresenterFactory.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);

			borderFactory.AppendChild(contentPresenterFactory);
			template.VisualTree = borderFactory;

			ToggleButtonStyle.Setters.Add(new Setter(Button.TemplateProperty, template));
		}
	}
}
