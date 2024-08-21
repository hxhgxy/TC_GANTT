using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows;
using System.Windows.Controls;
using static System.Net.Mime.MediaTypeNames;

namespace Worksheet
{
	public class ColumnHeader
	{
		public int Index { get; }
		public string Title { get; private set; }
		public double Width { get; set; }

		public double Left { get; internal set; }
				
		public TextAlignment TextAlignment { get; set; } = TextAlignment.Center;
		public FontWeight FontWeight { get; set; } = FontWeights.Normal;

		public bool IsSelected { get; set; }

		// 新增方法：设置列标题
		public void SetTitle(string title, TextAlignment textAlignment,FontWeight fontWeight)
		{
			Title = title;
			TextAlignment = textAlignment;
			FontWeight = fontWeight;
		}

		// 新增 Tooltip 属性
		public ToolTip Tooltip { get; }

		public ColumnHeader(int index, string title, double width, double left)
		{
			Index = index;
			Title = title;
			Width = width;
			Left = left;
			IsSelected = false;
		}

		public void Draw(DrawingContext dc, double x, double y, double width, double height, Typeface typeface, double fontSize, Brush textBrush)
		{
			Color backColor;
			Pen linePen;
			if (IsSelected)
			{
				backColor = Color.FromArgb(255, 210, 210, 210);
				linePen = new Pen(new SolidColorBrush(Colors.Green), 1.5);
			}
			else
			{
				backColor = Color.FromArgb(255, 240, 240, 240);
				linePen = new Pen(new SolidColorBrush(Color.FromArgb(255, 210, 210, 210)), 1);
			}
			dc.DrawRectangle(new SolidColorBrush(backColor), null, new Rect(x, y, width, height));
			dc.DrawLine(linePen, new Point(x, y + height), new Point(x + width, y + height));

			FormattedText formattedText = new FormattedText(
			    Title,
			    System.Globalization.CultureInfo.CurrentCulture,
			    FlowDirection.LeftToRight,
			    new Typeface(typeface.FontFamily, typeface.Style, FontWeight, typeface.Stretch),
			    fontSize,
			    textBrush,
			    1.25);
			//{
			//	TextAlignment = TextAlignment
			//};

			double textX = x + (width - formattedText.Width) / 2;
			double textY = y + (height - formattedText.Height) / 2;

			switch (TextAlignment)
			{
				case TextAlignment.Left:
					textX = x;
					break;
				case TextAlignment.Right:
					textX = x + (width - formattedText.Width);
					break;
				case TextAlignment.Center:
					// 已经在上面计算了 textX
					break;
				case TextAlignment.Justify:
					// Justify 对齐方式在这种情况下不常用，可以视需求处理
					break;
			}

			dc.DrawText(formattedText, new Point(textX, textY));
		}

		//public void Draw(DrawingContext dc, double x, double y, double width, double height, Typeface typeface, double fontSize, Brush textBrush)
		//{
		//	Color backColor;
		//	Pen linePen;
		//	if (IsSelected)
		//	{
		//		backColor = Color.FromArgb(255, 210, 210, 210);
		//		linePen = new Pen(new SolidColorBrush(Colors.Green), 1.5);
		//	}
		//	else
		//	{
		//		backColor = Color.FromArgb(255, 240, 240, 240);
		//		linePen = new Pen(new SolidColorBrush(Color.FromArgb(255, 210, 210, 210)), 1);
		//	}

		//	dc.DrawRectangle(new SolidColorBrush(backColor), null, new Rect(x, y, width, height));
		//	dc.DrawLine(linePen, new Point(x, y + height), new Point(x + width, y + height));

		//	FormattedText formattedText = new FormattedText(
		//	    Title,
		//	    System.Globalization.CultureInfo.CurrentCulture,
		//	    FlowDirection.LeftToRight,
		//	    typeface,
		//	    fontSize,
		//	    textBrush,
		//	    1.25);
		//	dc.DrawText(formattedText, new Point(x + (width - formattedText.Width) / 2, y + (height - formattedText.Height) / 2));
		//}
	}

	public class RowHeader
	{
		#region Hierarchy
		public RowHeader Parent { get; set; }
		public bool IsVisible { get; set; } = true;
		public List<RowHeader> Children { get; set; } = new List<RowHeader>();

		private bool _isExpanded;

		public bool IsExpanded
		{
			get { return _isExpanded; }
			set 
			{ 
				_isExpanded = value; 
				if(Children.Count > 0)
				{
					foreach (RowHeader child in Children)
					{						
						SetChildrenVisibility(child, value);
					}
					worksheet.UpdateRowPositions();
				}
			}
		}

		private void SetChildrenVisibility(RowHeader rowHeader,bool isVisible)
		{
			rowHeader.IsVisible = isVisible;

			if (!isVisible)
			{
				Cell cell = worksheet.GetOrCreateCell(rowHeader.Index, worksheet.HierarchyColIndex);
				cell.IsExpanded = isVisible;

				if (rowHeader.Children.Count > 0)
				{
					rowHeader.IsExpanded = isVisible;
				}

				foreach (RowHeader child in rowHeader.Children)
				{
					SetChildrenVisibility(child, isVisible);
				}
			}
		}

		#endregion // Hierarchy
		[NonSerialized]
		private Worksheet worksheet;
		public Worksheet Worksheet { get { return this.worksheet; } }
		public int Index { get; }
		public string Title { get; }
		public double Height { get; set; }
		public double Top { get; internal set; }
		public bool IsSelected { get; set; }


		public RowHeader(Worksheet worksheet, int index, string title, double height, double top)
		{
			this.worksheet = worksheet;
			Index = index;
			Title = title;
			Height = height;
			Top = top;
			IsSelected = false;
			IsVisible=true;
		}

		public void Draw(DrawingContext dc, double x, double y, double width, double height, Typeface typeface, double fontSize, Brush textBrush)
		{
			Color backColor;
			Pen linePen;
			if (IsSelected)
			{
				backColor = Color.FromArgb(255, 210, 210, 210);
				linePen = new Pen(new SolidColorBrush(Colors.Green),1.5);
			}
			else
			{
				backColor = Color.FromArgb(255, 240, 240, 240);
				linePen = new Pen(new SolidColorBrush(Color.FromArgb(255, 210, 210, 210)), 1);
			}

			dc.DrawRectangle(new SolidColorBrush(backColor), null, new Rect(x, y, width, height));
			dc.DrawLine(linePen, new Point(x + width, y), new Point(x + width, y + height));

			FormattedText formattedText = new FormattedText(
			    Title,
			    System.Globalization.CultureInfo.CurrentCulture,
			    FlowDirection.LeftToRight,
			    typeface,
			    fontSize,
			    textBrush,
			    1.25);
			dc.DrawText(formattedText, new Point(x + (width - formattedText.Width) / 2, y + (height - formattedText.Height) / 2));
		}
	}
}
