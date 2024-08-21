using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls.Primitives;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows;
using System.Windows.Input;
using System.Diagnostics;
using Worksheet.Utility;
using System.Windows.Media.Media3D;
using System.Windows.Shapes;
using System.ComponentModel;
using System.Windows.Threading;
using System.Globalization;
using System.Drawing.Drawing2D;
using Worksheet.Models;
using Microsoft.SqlServer.Server;
using System.Windows.Media.Effects;

namespace Worksheet
{
	public class Worksheet : Canvas, INotifyPropertyChanged
	{
		#region 通知接口
		public event PropertyChangedEventHandler PropertyChanged;

		protected void OnPropertyChanged(string propertyName)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
		#endregion

		#region 甘特图
		#region 属性
		private int taskPointed { get; set; } = -1;
		private double _initalScrollValue;

		public double InitialScrollValue
		{
			get { return _initalScrollValue; }
			set
			{
				_initalScrollValue = value;
				this.horScrollBarGantt.Value = _initalScrollValue;
				OnPropertyChanged(nameof(InitialScrollValue));
			}
		}


		//private bool initialSetting4ScrollBar { get; set; } = true;
		/// <summary>
		/// 用户设置控件最小日期
		/// </summary>
		private DateTime _startDate = new DateTime(2024, 8, 1);
		public DateTime StartDate
		{
			get
			{
				return _startDate;
			}
			set
			{
				_startDate = value;
				OnPropertyChanged(nameof(StartDate));
				if (ScaleMode == 0 || ScaleMode == 1)
				{
					RealStartDate = _startDate.AddDays(DateHelper.GetFirstDayOfWeek(_startDate));
				}
				else
				{
					RealStartDate = _startDate.AddDays(-DateHelper.GetFirstDayOfMonth(_startDate) + 1);
				}
			}

		}
		/// <summary>
		/// 用户设置控件最大日期
		/// </summary>
		private DateTime _endDate = new DateTime(2024, 12, 31);
		public DateTime EndDate
		{
			get { return _endDate; }
			set
			{
				_endDate = value;
				OnPropertyChanged(nameof(EndDate));
				if (ScaleMode == 0 || ScaleMode == 1)
				{
					RealEndDate = _endDate.AddDays(DateHelper.GetFirstDayOfWeek(_endDate) + 6);
				}
				else
				{
					RealEndDate = _endDate.AddDays(DateHelper.GetLastDayOfMonth(_endDate));
				}
			}
		}
		/// <summary>
		/// 根据日期刻度模式，计算得到的最小日期
		/// </summary>
		public DateTime RealStartDate { get; private set; }
		/// <summary>
		/// 根据日期刻度模式，计算得到的最大日期
		/// </summary>
		public DateTime RealEndDate { get; private set; }
		/// <summary>
		/// 刻度模式：<br/>
		/// 0 - 周 + 日期<br/>
		/// 1 - 月 + 周<br/>
		/// 2 - 季度 + 月
		/// </summary>
		private int _scaleMode = 0;
		public int ScaleMode
		{
			get
			{
				return _scaleMode;
			}
			set
			{
				_scaleMode = value;
				OnPropertyChanged(nameof(ScaleMode));
				OnPropertyChanged(nameof(StartDate));
				OnPropertyChanged(nameof(EndDate));
			}
		}
		/// <summary>
		/// 一个日期代表的绘制宽度
		/// </summary>
		private int _widthRatio = 15;
		public int WidthRatio
		{
			get { return _widthRatio; }
			set
			{
				_widthRatio = value;
				OnPropertyChanged(nameof(WidthRatio));
			}
		}
		private double ganttBarSize = 18;
		#endregion

		#region 方法
		/// <summary>
		/// 控件最小日期
		/// </summary>
		/// <param name="startDate"></param>
		//public void SetStartDate(DateTime startDate)
		//{
		//	this.startDate = startDate;
		//}
		/// <summary>
		/// 设置控件最大日期
		/// </summary>
		/// <param name="endDate"></param>
		//public void SetEndDate(DateTime endDate)
		//{
		//	this.endDate = endDate;
		//}
		/// <summary>
		/// 设置日期刻度模式
		/// </summary>
		/// <param name="scaleMode"></param>
		//public void SetScaleMode(int scaleMode)
		//{
		//	if (scaleMode < 0 || scaleMode > 2)
		//	{
		//		MessageBox.Show("0 - 周 + 日期\n1 - 月 + 周\n2 - 季度 + 月", "有效设置");
		//		return;
		//	}
		//	this.ScaleMode = scaleMode;
		//}

		//public void SetWidthRatio(int widthRatio)
		//{
		//	if (widthRatio <= 0) { return; }

		//	this.WidthRatio = widthRatio;
		//}
		/// <summary>
		/// 根据刻度模式计算真实最小日期
		/// </summary>
		/// <param name="startDate">用户设置的最小日期</param>
		/// <param name="scaleMode">用户设置刻度模式</param>
		/// <returns></returns>
		//private void CalRealDates()
		//{
		//	if (ScaleMode == 0 || ScaleMode == 1)
		//	{
		//		this.realStartDate = this.startDate.AddDays(DateHelper.GetFirstDayOfWeek(this.startDate));
		//		this.realEndDate = this.endDate.AddDays(DateHelper.GetFirstDayOfWeek(endDate) + 6);
		//	}
		//	else
		//	{
		//		this.realStartDate = this.startDate.AddDays(-DateHelper.GetFirstDayOfMonth(this.startDate) + 1);
		//		this.realEndDate = this.endDate.AddDays(DateHelper.GetLastDayOfMonth(endDate));
		//	}
		//}

		/// <summary>
		/// 绘制日期刻度，有三种模式
		/// </summary>
		/// <param name="dc"></param>
		private void DrawScale(DrawingContext dc)
		{
			//CalRealDates();

			// 分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);
			// 可供绘图的宽度
			double totalGanttWidth = this.ActualWidth - splitterPos - gridSplitter.Width - ScrollbarSize;
			// 总日期范围
			TimeSpan difference = RealEndDate - RealStartDate;
			int totalDays = difference.Days + 1;
			double totalDateWidth = totalDays * WidthRatio;

			// 甘特图水平滚动条数值设置
			horScrollBarGantt.Maximum = totalDateWidth;
			horScrollBarGantt.SmallChange = WidthRatio;
			switch (ScaleMode)
			{
				case 0:
					horScrollBarGantt.LargeChange = 7 * WidthRatio;
					break;
				default:
					horScrollBarGantt.LargeChange = 30 * WidthRatio;
					break;
			}

			// 滚动条当前值
			double scrollOffset = horScrollBarGantt.Value;

			// 可见区域的起始和结束位置
			double visibleStartX = scrollOffset;
			double visibleEndX = scrollOffset + totalGanttWidth;

			// 可见区域的起始和结束日期
			int visibleStartDayIndex = (int)(visibleStartX / WidthRatio);
			int visibleEndDayIndex = (int)(visibleEndX / WidthRatio);

			// 确保索引在有效范围内
			visibleStartDayIndex = Math.Max(0, visibleStartDayIndex);
			visibleEndDayIndex = Math.Min(totalDays - 1, visibleEndDayIndex);

			// 设置绘制矩形和文本的样式
			SolidColorBrush rectBrush = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240));
			SolidColorBrush weekendBrush = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)); // 浅灰色
			Pen rectPen = new Pen(new SolidColorBrush(Color.FromArgb(255, 210, 210, 210)), 1);
			Typeface typeface = new Typeface("Arial");
			double fontSize = 10;
			Brush textBrush = Brushes.Black;

			#region 周 + 日期
			if (ScaleMode == 0)
			{
				// 从可见区域的起始日期到结束日期绘制日期刻度
				// 绘制下部小刻度，日为单位
				for (int i = visibleStartDayIndex; i <= visibleEndDayIndex; i++)
				{
					DateTime currentDate = RealStartDate.AddDays(i);
					double xPos = i * WidthRatio - scrollOffset + splitterPos + gridSplitter.Width;

					// 判断是否为周末
					bool isWeekend = currentDate.DayOfWeek == DayOfWeek.Saturday || currentDate.DayOfWeek == DayOfWeek.Sunday;

					// 绘制矩形
					Rect rect = new Rect(xPos, ColumnHeaderHeight / 2, WidthRatio, ColumnHeaderHeight / 2); // 假设矩形高度为 30
					dc.DrawRectangle(isWeekend ? weekendBrush : rectBrush, rectPen, rect);
					if (isWeekend)
					{
						Rect rectWeekend = new Rect(xPos, ColumnHeaderHeight, WidthRatio, this.ActualHeight - ColumnHeaderHeight - ScrollbarSize);
						dc.DrawRectangle(weekendBrush, null, rectWeekend);
					}

					// 绘制日期文本
					FormattedText formattedText = new FormattedText(
					    currentDate.ToString("dd"), // 显示日期的日部分
					    System.Globalization.CultureInfo.InvariantCulture,
					    FlowDirection.LeftToRight,
					    typeface,
					    fontSize,
					    textBrush,
					    96 // dpi
					);

					// 计算文本的位置，使其居中
					double textX = xPos + (WidthRatio - formattedText.Width) / 2;
					double textY = (ColumnHeaderHeight / 2 - formattedText.Height) / 2 + ColumnHeaderHeight / 2; // 假设矩形高度为 30
					dc.DrawText(formattedText, new Point(textX, textY));
				}

				// 绘制上部大刻度，周为单位
				int visibleStartWeekIndex = visibleStartDayIndex / 7;
				int visibleEndWeekIndex = visibleEndDayIndex / 7;

				for (int i = visibleStartWeekIndex; i <= visibleEndWeekIndex; i++)
				{
					DateTime weekStartDate = RealStartDate.AddDays(i * 7);
					double xPos = i * 7 * WidthRatio - scrollOffset + splitterPos + gridSplitter.Width;

					// 绘制矩形
					Rect rect = new Rect(xPos, 0, 7 * WidthRatio, ColumnHeaderHeight / 2); // 假设矩形高度为 30
					dc.DrawRectangle(rectBrush, rectPen, rect);

					// 获取年份和周数
					int year = weekStartDate.Year;
					int weekNumber = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(weekStartDate, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

					// 格式化为 "yyyy week xx"
					string weekText = $"{year} week {weekNumber}";

					// 绘制周文本
					FormattedText formattedText = new FormattedText(
					    weekText, // 显示年份和周数
					    System.Globalization.CultureInfo.InvariantCulture,
					    FlowDirection.LeftToRight,
					    typeface,
					    fontSize,
					    textBrush,
					    96 // dpi
					);

					// 计算文本的位置，使其居中
					double textX = xPos + (7 * WidthRatio - formattedText.Width) / 2;
					double textY = (ColumnHeaderHeight / 2 - formattedText.Height) / 2; // 假设矩形高度为 30
					dc.DrawText(formattedText, new Point(textX, textY));
				}
			}
			#endregion
		}
		private List<GanttTask> _taskList;

		public List<GanttTask> TaskList
		{
			get { return _taskList; }
			set
			{
				_taskList = value;
				OnPropertyChanged(nameof(TaskList));
			}
		}

		public void FindCriticalPath()
		{
			foreach (var task in TaskList)
			{
				if (task.Predecessors == null) continue;

				if (!task.Predecessors.Any())
				{
					task.EarlyStart = task.StartDate;
				}
				else
				{
					var maxEF = DateTime.MinValue;
					foreach (var (predecessorIndex, dependencyType) in task.Predecessors)
					{
						var predecessor = TaskList[predecessorIndex];
						DateTime tempES;
						switch (dependencyType)
						{
							case "FS":
								tempES = predecessor.EarlyFinish;
								break;
							case "SS":
								tempES = predecessor.EarlyStart;
								break;
							case "FF":
								tempES = predecessor.EarlyFinish.AddDays(-task.Duration);
								break;
							case "SF":
								tempES = predecessor.EarlyStart.AddDays(-task.Duration);
								break;
							default:
								throw new InvalidOperationException($"Unknown dependency type: {dependencyType}");
						}
						if (tempES > maxEF)
						{
							maxEF = tempES;
						}
					}
					task.EarlyStart = maxEF;
				}
				task.EarlyFinish = task.EarlyStart.AddDays(task.Duration);
			}

			var projectEndDate = TaskList.Max(t => t.EarlyFinish);
			foreach (var task in TaskList.OrderByDescending(t => t.Index))
			{
				if (!TaskList.Any(t => t.Predecessors != null && t.Predecessors.Any(p => p.PredecessorIndex == task.Index)))
				{
					task.LateFinish = projectEndDate;
				}
				else
				{
					var minLS = DateTime.MaxValue;
					foreach (var successor in TaskList.Where(t => t.Predecessors != null && t.Predecessors.Any(p => p.PredecessorIndex == task.Index)))
					{
						foreach (var (predecessorIndex, dependencyType) in successor.Predecessors)
						{
							if (predecessorIndex == task.Index)
							{
								DateTime tempLS;
								switch (dependencyType)
								{
									case "FS":
										tempLS = successor.LateStart;
										break;
									case "SS":
										tempLS = successor.LateStart;
										break;
									case "FF":
										tempLS = successor.LateFinish.AddDays(-task.Duration);
										break;
									case "SF":
										tempLS = successor.LateFinish;
										break;
									default:
										throw new InvalidOperationException($"Unknown dependency type: {dependencyType}");
								}
								if (tempLS < minLS)
								{
									minLS = tempLS;
								}
							}
						}
					}
					task.LateFinish = minLS;
				}

				// 检查 LateFinish 是否被正确计算
				if (task.LateFinish == DateTime.MinValue)
				{
					task.LateFinish = projectEndDate;
				}

				task.LateStart = task.LateFinish.AddDays(-task.Duration);
			}

			foreach (var task in TaskList)
			{
				if (task.EarlyStart == task.LateStart)
				{
					task.IsCritical = true;
				}
				else
				{
					task.IsCritical = false;
				}
			}
		}

		/// <summary>
		/// 获取任务列表
		/// </summary>
		/// <returns></returns>
		//private void GetTasks()
		//{
		//	if (_taskList == null)
		//	{
		//		this.TaskList = new List<GanttTask>();
		//	}
		//	else
		//	{
		//		this.TaskList.Clear();
		//	}

		//	for (int i = 0; i < Rows.Count; i++)
		//	{
		//		Cell cell;
		//		cell = GetOrCreateCell(i, 1);
		//		if (cell.InnerData != null)
		//		{
		//			GanttTask task = new GanttTask();

		//			task.Index = i;
		//			task.TaskName = cell.InnerData;
		//			if (Rows[i].Children.Count > 0)
		//			{
		//				task.HasChildren = true;
		//			}
		//			cell = GetOrCreateCell(i, 2);
		//			task.StartDate = String2Date(cell.InnerData);
		//			cell = GetOrCreateCell(i, 3);
		//			task.EndDate = String2Date(cell.InnerData);
		//			cell = GetOrCreateCell(i, 4);
		//			task.Duration = Convert.ToDouble(cell.InnerData);
		//			cell = GetOrCreateCell(i, 5);
		//			task.Completion = String2Double(cell.InnerData);
		//			cell = GetOrCreateCell(i, 6);

		//			if (!string.IsNullOrEmpty(cell.InnerData))
		//			{
		//				string tempString = cell.InnerData.Replace(" ", "");
		//				task.Predecessors = tempString.Split(',').ToList();
		//			}
		//			//task.Predecessors = cell.InnerData;

		//			TaskList.Add(task);
		//		}
		//	}
		//	//控件初始化时，将甘特图滚动条滚动到项目最开始时间
		//	//只做一次设置，后面用户可以随意拖拽滚动条
		//	//if (TaskList.Count > 0 & initialSetting4ScrollBar)
		//	//{
		//	//	DateTime earliestStartDate = TaskList.Min(task => task.StartDate);

		//	//	// 计算滚动条的位置
		//	//	double offsetToEarliestDate = (earliestStartDate - realStartDate).TotalDays * widthRatio;

		//	//	// 设置滚动条的值
		//	//	horScrollBarGantt.Value = Math.Max(0, offsetToEarliestDate);
		//	//	initialSetting4ScrollBar = false;
		//	//}

		//	//return tasks;
		//}
		/// <summary>
		/// 甘特图阴影
		/// </summary>
		//private DropShadowEffect shadowEffect = new DropShadowEffect
		//{
		//	Color = Colors.Black,
		//	ShadowDepth = 5,
		//	BlurRadius = 10,
		//	Opacity = 0.5
		//};

		public double String2Double(string strComplet)
		{
			if (string.IsNullOrEmpty(strComplet)) return 0;
			// 移除百分号
			string numberString = strComplet.TrimEnd('%');
			return Convert.ToDouble(numberString) / 100d;
		}

		public DateTime String2Date(string strDate)
		{
			if (string.IsNullOrEmpty(strDate)) return DateTime.Today;
			string format = "yyyy-MM-dd";
			return DateTime.ParseExact(strDate, format, CultureInfo.InvariantCulture);
		}

		private void DrawGanttChart(DrawingContext dc)
		{
			// 分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);
			// 可供绘图的宽度
			double totalGanttWidth = this.ActualWidth - splitterPos - gridSplitter.Width - ScrollbarSize;
			// 总日期范围
			TimeSpan difference = RealEndDate - RealStartDate;
			int totalDays = difference.Days + 1;
			double totalDateWidth = totalDays * WidthRatio;

			// 获取滚动条的水平和垂直偏移量
			double horizontalOffset = horScrollBarGantt.Value;
			double verticalOffset = verScrollBar.Value;

			//GetTasks();
			if (TaskList == null) return;

			// 用于计算可见任务的实际 Y 值
			double currentY = ColumnHeaderHeight;

			for (int index = 0; index < TaskList.Count; index++)
			{
				if (!Rows[index].IsVisible)
				{
					continue;
				}

				GanttTask task = TaskList[index];

				// 计算任务条形的位置和大小
				double taskStartX = (task.StartDate - RealStartDate).TotalDays * WidthRatio + splitterPos + gridSplitter.Width - horizontalOffset;
				double taskWidth = task.Duration * WidthRatio;

				// 计算任务条形的 Y 值
				double taskY = currentY - verticalOffset + (ColumnHeaderHeight - ganttBarSize) / 4;

				// 更新 currentY 以便下一个可见任务使用
				currentY += RowHeaderHeight;

				// 检查任务条形是否在可见区域内，并裁剪任务条形
				double visibleStartX = Math.Max(taskStartX, splitterPos + gridSplitter.Width);
				double visibleEndX = Math.Min(taskStartX + taskWidth, splitterPos + gridSplitter.Width + totalGanttWidth);
				double visibleWidth = visibleEndX - visibleStartX;

				double visibleStartY = Math.Max(taskY, ColumnHeaderHeight); // 确保任务条形不在列标题之上
				double visibleEndY = Math.Min(taskY + ganttBarSize, this.ActualHeight);
				double visibleHeight = visibleEndY - visibleStartY;

				if (visibleWidth > 0 && visibleHeight > 0)
				{
					if (task.HasChildren)
					{
						double bracketHeight = ganttBarSize / 2;

						// 左边短竖线
						dc.DrawLine(new Pen(new SolidColorBrush(Color.FromArgb(255, 50, 50, 50)), 3),
						    new Point(visibleStartX, visibleStartY + bracketHeight),
						    new Point(visibleStartX, visibleStartY));

						// 上面长横线
						dc.DrawLine(new Pen(new SolidColorBrush(Color.FromArgb(255, 50, 50, 50)), 3),
						    new Point(visibleStartX, visibleStartY + 1.5),
						    new Point(visibleEndX, visibleStartY + 1.5));

						// 右边短竖线
						dc.DrawLine(new Pen(new SolidColorBrush(Color.FromArgb(255, 50, 50, 50)), 3),
						    new Point(visibleEndX, visibleStartY),
						    new Point(visibleEndX, visibleStartY + bracketHeight));
					}
					else
					{
						//可见部分
						Rect taskRect = new Rect(visibleStartX, visibleStartY, visibleWidth, visibleHeight);
						//原始，包括不可见部分
						Rect taskRectOriginal = new Rect(taskStartX, visibleStartY, visibleWidth, visibleHeight);

						// 绘制圆角矩形作为任务条形
						double cornerRadius = 3.0; // 圆角半径

						SolidColorBrush barBackColor;
						if (task.IsCritical)
						{
							barBackColor = new SolidColorBrush(Color.FromArgb(255, 255, 150, 170));
						}
						else
						{
							barBackColor = (SolidColorBrush)new BrushConverter().ConvertFromString("#b4e1fa");
						}

						dc.DrawRoundedRectangle(barBackColor, null, taskRect, cornerRadius, cornerRadius);

						// 计算完成部分的矩形
						double completionWidth = taskWidth * task.Completion;
						double barCompletStart = Math.Max(taskStartX, taskRect.X);

						completionWidth += taskStartX - taskRect.X;

						if (completionWidth < 0)
						{
							completionWidth = 0;
						}

						Rect completionRect = new Rect(barCompletStart, taskRect.Y, completionWidth, taskRect.Height);

						// 创建圆角矩形几何体
						var taskGeometry = new RectangleGeometry(taskRect, cornerRadius, cornerRadius);
						var completionGeometry = new RectangleGeometry(completionRect);

						// 排除未完成部分
						var completedGeometry = new CombinedGeometry(GeometryCombineMode.Intersect, taskGeometry, completionGeometry);

						// 绘制完成部分
						SolidColorBrush completionColor;
						if (task.IsCritical)
						{
							completionColor = new SolidColorBrush(Color.FromArgb(255, 255, 60, 80));
						}
						else
						{
							completionColor = (SolidColorBrush)new BrushConverter().ConvertFromString("#FF64C8FF");
						}
						dc.DrawGeometry(completionColor, null, completedGeometry);

						//如果鼠标在本甘特条范围之内，则绘制边框
						if (task.Index == taskPointed)
						{
							dc.DrawRoundedRectangle(null, new Pen(new SolidColorBrush(Colors.Gray), 1.5), taskRect, cornerRadius, cornerRadius);
						}
					}
				}
				// 绘制前置任务的连接线
				if (task.Predecessors != null)
				{
					foreach (var (predecessorIndex, dependencyType) in task.Predecessors)
					{
						GanttTask predecessorTask = TaskList[predecessorIndex];

						double preTaskStartX = (predecessorTask.StartDate - RealStartDate).TotalDays * WidthRatio + splitterPos + gridSplitter.Width - horizontalOffset;
						double preTaskEndX = ((predecessorTask.EndDate - RealStartDate).TotalDays + 1) * WidthRatio + splitterPos + gridSplitter.Width - horizontalOffset;
						double preTaskY = ColumnHeaderHeight + predecessorIndex * RowHeaderHeight - verticalOffset + (ColumnHeaderHeight - ganttBarSize) / 4;

						Point startPoint = new Point();
						Point endPoint = new Point();

						switch (dependencyType.ToUpper())
						{
							case "FS":
								startPoint = new Point(preTaskEndX, preTaskY + ganttBarSize / 2);
								endPoint = new Point(taskStartX + Math.Min(10, taskWidth * 0.2), taskY);
								break;
							case "SS":
								startPoint = new Point(preTaskStartX, preTaskY + ganttBarSize / 2);
								endPoint = new Point(taskStartX, taskY + ganttBarSize / 2);
								break;
							case "SF":
								startPoint = new Point(preTaskStartX, preTaskY + ganttBarSize / 2);
								endPoint = new Point(taskStartX + taskWidth, taskY + ganttBarSize / 2);
								break;
							case "FF":
								startPoint = new Point(preTaskEndX, preTaskY + ganttBarSize / 2);
								endPoint = new Point(taskStartX + taskWidth, taskY + ganttBarSize / 2);
								break;
							default:
								// 默认为 FS
								startPoint = new Point(preTaskEndX, preTaskY + ganttBarSize / 2);
								endPoint = new Point(taskStartX + Math.Min(10, taskWidth * 0.2), taskY);
								break;
						}

						// 调用 DrawElbowConnector 进行绘制
						DrawElbowConnector(dc, startPoint, endPoint, splitterPos, gridSplitter.Width, totalGanttWidth, dependencyType.ToUpper(), (task.IsCritical && predecessorTask.IsCritical));
					}
				}

			}
		}

		private void DrawElbowConnector(DrawingContext dc, Point startPoint, Point endPoint, double splitterPos, double splitterWidth, double totalGanttWidth, string connectionType, bool isCritical)
		{
			// 调整起始点和终点，使其在可见区域内
			double minX = splitterPos + splitterWidth;
			double maxX = minX + totalGanttWidth;
			double originalEndX = endPoint.X;

			bool startPointClipped = false;
			bool endPointClipped = false;

			if (startPoint.X < minX && endPoint.X < minX)
			{
				// 如果两个点都在gridsplitter左边，则不绘制连接符
				return;
			}
			else if (startPoint.X < minX || endPoint.X < minX)
			{
				// 如果只有一个点在gridsplitter左边，计算可见部分
				if (startPoint.X < minX)
				{
					startPoint = ClipPointToVisibleArea(startPoint, minX, maxX);
					startPointClipped = true;
				}
				if (endPoint.X < minX)
				{
					endPoint = ClipPointToVisibleArea(endPoint, minX, maxX);
					endPointClipped = true;
				}
			}

			Pen linePen;

			if (isCritical)
			{
				linePen = new Pen(new SolidColorBrush(Color.FromArgb(255, 255, 60, 80)), 1);
			}
			else
			{
				linePen = new Pen((SolidColorBrush)new BrushConverter().ConvertFromString("#FF64C8FF"), 1);
			}

			switch (connectionType)
			{
				case "FS":
					if (startPoint.X < endPoint.X)
					{
						// 从左到右的短横，从上到下的竖线，从右到左的横线
						dc.DrawLine(linePen, startPoint, new Point(endPoint.X, startPoint.Y));
						dc.DrawLine(linePen, new Point(endPoint.X, startPoint.Y), new Point(endPoint.X, endPoint.Y));
					}
					else
					{
						// 从左到右的短横，从上到下的竖线，从右到左的横线，从上到下的竖线
						double midY = (endPoint.Y - startPoint.Y) / 2 + startPoint.Y;
						dc.DrawLine(linePen, startPoint, new Point(startPoint.X + 10, startPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, startPoint.Y), new Point(startPoint.X + 10, midY));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, midY), new Point(endPoint.X, midY));
						dc.DrawLine(linePen, new Point(endPoint.X, midY), endPoint);
					}
					break;

				case "SS":
					// 中括号形式的连接符，从右向左的短横，从上到下的竖线，从左到右的短横
					double leftX = Math.Min(startPoint.X, endPoint.X) - 10;

					dc.DrawLine(linePen, startPoint, new Point(leftX, startPoint.Y));
					dc.DrawLine(linePen, new Point(leftX, startPoint.Y), new Point(leftX, endPoint.Y));
					dc.DrawLine(linePen, new Point(leftX, endPoint.Y), endPoint);
					break;

				case "FF":
					// 右中括号形式的连接符
					double rightX = Math.Max(startPoint.X, endPoint.X) + 10;

					if (startPoint.X < endPoint.X)
					{
						// 如果起点在终点左边
						dc.DrawLine(linePen, startPoint, new Point(rightX, startPoint.Y));
						dc.DrawLine(linePen, new Point(rightX, startPoint.Y), new Point(rightX, endPoint.Y));
						dc.DrawLine(linePen, new Point(rightX, endPoint.Y), endPoint);
					}
					else
					{
						// 如果起点在终点右边，绘制可见部分
						if (endPointClipped)
						{
							// 终点在gridsplitter左边
							dc.DrawLine(linePen, startPoint, new Point(rightX, startPoint.Y));
							dc.DrawLine(linePen, new Point(rightX, startPoint.Y), new Point(rightX, endPoint.Y));
							dc.DrawLine(linePen, new Point(rightX, endPoint.Y), endPoint);
						}
						else
						{
							// 终点在可见区域
							dc.DrawLine(linePen, startPoint, new Point(rightX, startPoint.Y));
							dc.DrawLine(linePen, new Point(rightX, startPoint.Y), new Point(rightX, endPoint.Y));
							dc.DrawLine(linePen, new Point(rightX, endPoint.Y), endPoint);
						}
					}
					break;

				case "SF":
					if (startPoint.X > endPoint.X)
					{
						// 从右到左的短横，从上到下的竖线，从右到左的短横
						dc.DrawLine(linePen, startPoint, new Point(startPoint.X - 10, startPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X - 10, startPoint.Y), new Point(startPoint.X - 10, endPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X - 10, endPoint.Y), endPoint);
					}
					else
					{
						// 从右到左的短横，从上到下的竖线，从左到右的短横，从上到下的竖线，从右到左的短横
						double midY = (endPoint.Y - startPoint.Y) / 2 + startPoint.Y;
						dc.DrawLine(linePen, startPoint, new Point(startPoint.X - 10, startPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X - 10, startPoint.Y), new Point(startPoint.X - 10, midY));
						dc.DrawLine(linePen, new Point(startPoint.X - 10, midY), new Point(endPoint.X + 10, midY));
						dc.DrawLine(linePen, new Point(endPoint.X + 10, midY), new Point(endPoint.X + 10, endPoint.Y));
						dc.DrawLine(linePen, new Point(endPoint.X + 10, endPoint.Y), endPoint);
					}
					break;

				default:
					// 默认为 FS
					if (startPoint.X < endPoint.X)
					{
						// 从左到右的短横，从上到下的竖线，从右到左的横线
						dc.DrawLine(linePen, startPoint, new Point(startPoint.X + 10, startPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, startPoint.Y), new Point(startPoint.X + 10, endPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, endPoint.Y), endPoint);
					}
					else
					{
						double midY = (endPoint.Y - startPoint.Y) / 2 + startPoint.Y;
						// 从左到右的短横，从上到下的竖线，从右到左的横线，从上到下的竖线
						dc.DrawLine(linePen, startPoint, new Point(startPoint.X + 10, startPoint.Y));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, startPoint.Y), new Point(startPoint.X + 10, midY));
						dc.DrawLine(linePen, new Point(startPoint.X + 10, midY), new Point(endPoint.X - 10, midY));
						dc.DrawLine(linePen, new Point(endPoint.X - 10, midY), endPoint);
					}
					break;
			}

			// 只有当终点在可见区域时才绘制终点的三角形
			if (originalEndX > minX)
			{
				Point trianglePoint1, trianglePoint2, trianglePoint3;

				if (connectionType.ToUpper() == "SS")
				{
					// SS 类型的连接符箭头向右
					trianglePoint1 = new Point(endPoint.X - 5, endPoint.Y - 4);
					trianglePoint2 = new Point(endPoint.X - 5, endPoint.Y + 4);
					trianglePoint3 = new Point(endPoint.X, endPoint.Y);
				}
				else if (connectionType.ToUpper() == "FF" || connectionType.ToUpper() == "SF")
				{
					// FF 和 SF 类型的连接符箭头向左
					trianglePoint1 = new Point(endPoint.X + 5, endPoint.Y - 4);
					trianglePoint2 = new Point(endPoint.X + 5, endPoint.Y + 4);
					trianglePoint3 = new Point(endPoint.X, endPoint.Y);
				}
				else if (startPoint.Y < endPoint.Y)
				{
					// 起点在终点上面，箭头向下
					trianglePoint1 = new Point(endPoint.X - 4, endPoint.Y - 5);
					trianglePoint2 = new Point(endPoint.X + 4, endPoint.Y - 5);
					trianglePoint3 = new Point(endPoint.X, endPoint.Y);
				}
				else
				{
					// 起点在终点下面，箭头向上
					trianglePoint1 = new Point(endPoint.X - 4, endPoint.Y + 5);
					trianglePoint2 = new Point(endPoint.X + 4, endPoint.Y + 5);
					trianglePoint3 = new Point(endPoint.X, endPoint.Y);
				}

				StreamGeometry triangle = new StreamGeometry();
				using (StreamGeometryContext ctx = triangle.Open())
				{
					ctx.BeginFigure(trianglePoint1, true, true);
					ctx.LineTo(trianglePoint2, true, false);
					ctx.LineTo(trianglePoint3, true, false);
				}

				SolidColorBrush angleBrush;

				if (isCritical)
				{
					angleBrush = new SolidColorBrush(Color.FromArgb(255, 255, 60, 80));
				}
				else
				{
					angleBrush = (SolidColorBrush)new BrushConverter().ConvertFromString("#FF64C8FF");
				}
				dc.DrawGeometry(angleBrush, null, triangle);
			}
		}

		private Point ClipPointToVisibleArea(Point point, double minX, double maxX)
		{
			if (point.X < minX)
			{
				point.X = minX;
			}
			if (point.X > maxX)
			{
				point.X = maxX;
			}
			return point;
		}


		#endregion

		#region 事件

		private void OnScrollBarGanttValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			if (isHandlingScrollEvent)
			{
				return;
			}

			isHandlingScrollEvent = true;

			try
			{
				InvalidateVisual();
			}
			finally
			{
				isHandlingScrollEvent = false;
			}
		}

		#endregion
		#endregion

		#region 事件
		// 定义 SelectionChanged 事件
		public event EventHandler<CellEventArgs> SelectionChanged;

		// 定义 Changed 事件
		public event EventHandler<CellEventArgs> Changed;

		// 当前选中的单元格
		public Cell SelectedCell
		{
			get
			{
				if (selectedRowIndex >= 0 && selectedRowIndex < Rows.Count &&
				    selectedColIndex >= 0 && selectedColIndex < Cols.Count)
				{
					return GetOrCreateCell(selectedRowIndex, selectedColIndex);
				}
				return null;
			}
		}

		// 触发 SelectionChanged 事件
		protected virtual void OnSelectionChanged(CellEventArgs e)
		{
			SelectionChanged?.Invoke(this, e);
		}

		// 触发 Changed 事件
		protected virtual void OnChanged(CellEventArgs e)
		{
			Changed?.Invoke(this, e);
		}
		#endregion

		#region 属性，变量
		public int HierarchyColIndex { get; set; } = 1;
		/// <summary>
		/// 是否高亮所选行
		/// </summary>
		private bool highlight = true;
		public bool Highlight
		{
			get => highlight;
			set
			{
				if (highlight != value)
				{
					highlight = value;
					OnPropertyChanged(nameof(Highlight));
					InvalidateVisual(); // 触发重绘
				}
			}
		}

		/// <summary>
		/// 单元格编辑器
		/// </summary>
		private TextBox editTextBox;
		// 当前选中的行和列索引
		private int selectedRowIndex = 0;
		private int selectedColIndex = 0;

		public int SelectedRowIndex
		{
			get { return selectedRowIndex; }
			set
			{
				if (selectedRowIndex != value)
				{
					selectedRowIndex = value;
					//OnSelectionChanged(new CellEventArgs(SelectedCell));
				}
			}
		}

		public int SelectedColIndex
		{
			get { return selectedColIndex; }
			set
			{
				if (selectedColIndex != value)
				{
					selectedColIndex = value;
					//OnSelectionChanged(new CellEventArgs(SelectedCell));
				}
			}
		}

		/// <summary>
		/// 临时存储单元格内容
		/// </summary>
		private string tempCellValue;
		private bool isTyping = false;
		/// <summary>
		/// 储存单元格数据的数组
		/// </summary>
		private string[,] cellValues;
		/// <summary>
		/// 滚动条宽度，渲染时，需要扣除
		/// </summary>
		private const int ScrollbarSize = 18;
		/// <summary>
		/// Excel和Gantt视图分割
		/// </summary>
		private GridSplitter gridSplitter;
		/// <summary>
		/// Excel水平滚动条
		/// </summary>
		private ScrollBar horScrollBarGrid;
		/// <summary>
		/// 甘特图水平滚动条
		/// </summary>
		private ScrollBar horScrollBarGantt;
		/// <summary>
		/// 垂直滚动条
		/// </summary>
		private ScrollBar verScrollBar;
		/// <summary>
		/// 全选按钮
		/// </summary>
		private Button selectAllButton; // 全选按钮
		/// <summary>
		/// 默认列数
		/// </summary>
		private static int DefaultCols = 1;
		/// <summary>
		/// 默认行数
		/// </summary>
		private static int DefaultRows = 3;
		/// <summary>
		/// 列集合
		/// </summary>
		private List<ColumnHeader> _cols = new List<ColumnHeader>(DefaultCols);

		public List<ColumnHeader> Cols
		{
			get { return _cols; }
			set { _cols = value; }
		}

		//public List<ColumnHeader> Cols { get; set; } = new List<ColumnHeader>(DefaultCols);
		/// <summary>
		/// 行集合
		/// </summary>
		private List<RowHeader> _rows = new List<RowHeader>(DefaultRows);

		public List<RowHeader> Rows
		{
			get { return _rows; }
			set { _rows = value; }
		}

		//public List<RowHeader> Rows { get; set; } = new List<RowHeader>(DefaultRows);
		/// <summary>
		/// 标志在调节列宽
		/// </summary>
		private bool isResizingColumn = false;
		/// <summary>
		/// 正在调节列宽的列索引
		/// </summary>
		private int resizingColumnIndex = -1;
		/// <summary>
		/// 初始鼠标坐标X
		/// </summary>
		private double initialMouseX;
		/// <summary>
		/// 初始列宽
		/// </summary>
		private double initialColumnWidth;
		/// <summary>
		/// 标志正在处理滚动条滚动事件
		/// </summary>
		private bool isHandlingScrollEvent = false;
		/// <summary>
		/// 标志在调节行高
		/// </summary>
		private bool isResizingRow = false;
		/// <summary>
		/// 正在调节行高的行索引
		/// </summary>
		private int resizingRowIndex = -1;
		/// <summary>
		/// 初始鼠标坐标Y
		/// </summary>
		private double initialMouseY;
		/// <summary>
		/// 初始行高
		/// </summary>
		private double initialRowHeight;
		/// <summary>
		/// 可见区域
		/// </summary>
		private GridRegion visibleRegion = new GridRegion();
		/// <summary>
		/// 列标题宽度，列宽
		/// </summary>
		public double ColumnHeaderWidth { get; set; } = 100; // 默认列标题宽度
		/// <summary>
		/// 列标题高度
		/// </summary>
		public double ColumnHeaderHeight { get; set; } = 40; // 默认列标题高度
		/// <summary>
		/// 行标题宽度
		/// </summary>
		public double RowHeaderWidth { get; set; } = 40; // 默认行标题宽度
		/// <summary>
		/// 行标题高度，行高
		/// </summary>
		public double RowHeaderHeight { get; set; } = 30; // 默认行标题高度
		/// <summary>
		/// 字典，工作表中单元格集合
		/// </summary>
		private Dictionary<string, Cell> cells;
		/// <summary>
		/// 标志正在初始化
		/// </summary>
		private bool isInitialized = false;

		private DispatcherTimer tooltipTimer;
		private ToolTip currentTooltip;
		private DateTime lastTooltipDate;

		/// <summary>
		/// 工作表背景
		/// </summary>
		public SolidColorBrush BackBrush { get; set; } = new SolidColorBrush(Colors.White);
		public SolidColorBrush BorderBrush { get; set; } = new SolidColorBrush(Color.FromArgb(255, 210, 210, 210));
		#endregion

		#region 构造函数
		public Worksheet()
		{
			this.SnapsToDevicePixels = true;
			this.Focusable = true;
			this.FocusVisualStyle = null;

			this.gridSplitter = new GridSplitter()
			{
				VerticalAlignment = VerticalAlignment.Stretch,
				HorizontalAlignment = HorizontalAlignment.Center,
				Background = new SolidColorBrush(Colors.LightGray),
				Visibility = Visibility.Visible,
				IsTabStop = false,
				Width = 4,
			};
			this.Children.Add(gridSplitter);
			this.gridSplitter.DragDelta += GridSplitter_DragDelta;

			this.horScrollBarGantt = new ScrollBar()
			{
				Orientation = Orientation.Horizontal,
				Height = ScrollbarSize,
				SmallChange = 1,
				IsTabStop = false,
				BorderBrush = new SolidColorBrush(Colors.LightGray),
			};
			this.Children.Add(horScrollBarGantt);

			this.horScrollBarGrid = new ScrollBar()
			{
				Orientation = Orientation.Horizontal,
				Height = ScrollbarSize,
				SmallChange = 1,
				IsTabStop = false,
				BorderBrush = new SolidColorBrush(Colors.LightGray),
			};

			this.verScrollBar = new ScrollBar()
			{
				Orientation = Orientation.Vertical,
				Width = ScrollbarSize,
				SmallChange = 1,
				IsTabStop = false,
				BorderBrush = new SolidColorBrush(Colors.LightGray),
			};

			this.selectAllButton = new Button()
			{
				Content = "",
				BorderThickness = new Thickness(0, 0, 0, 1),
				BorderBrush = new SolidColorBrush(Color.FromArgb(255, 210, 210, 210)),
				Background = new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)),
				IsTabStop = false,
				Width = RowHeaderWidth,
				Height = ColumnHeaderHeight
			};
			this.selectAllButton.Click += (s, e) => SelectAll();

			this.Children.Add(this.horScrollBarGrid);
			this.Children.Add(this.verScrollBar);
			this.Children.Add(this.selectAllButton);
			this.Background = new SolidColorBrush(Colors.White);
			this.Resize(DefaultRows, DefaultCols);

			this.horScrollBarGantt.ValueChanged += OnScrollBarGanttValueChanged;

			this.horScrollBarGrid.ValueChanged += OnScrollBarValueChanged;
			this.verScrollBar.ValueChanged += OnScrollBarValueChanged;

			this.LayoutUpdated += OnLayoutUpdated;

			this.MouseLeftButtonDown += OnMouseLeftButtonDown;
			this.MouseRightButtonDown += OnMouseRightButtonDown;
			this.MouseMove += OnMouseMove;
			this.MouseLeftButtonUp += OnMouseLeftButtonUp;
			//this.MouseEnter += OnMouseEnter;

			editTextBox = new TextBox()
			{
				Visibility = Visibility.Collapsed,
				BorderThickness = new Thickness(0),
				CaretBrush = new SolidColorBrush(Colors.Green),
				SelectionBrush = new SolidColorBrush(Colors.Green),
				VerticalContentAlignment = VerticalAlignment.Center,
				AutoWordSelection = true,
				Background = new SolidColorBrush(Colors.Transparent),
				IsTabStop = false,
			};
			editTextBox.LostFocus += EditTextBox_LostFocus;
			this.Children.Add(editTextBox);

			// 初始化完成
			isInitialized = true;

			// 初始化单元格存储
			cells = new Dictionary<string, Cell>();

			//this.SizeChanged += Worksheet_SizeChanged;

			//UpdateScrollBarMaxValue();
			UpdateSelectionRange();

			this.KeyDown += Worksheet_KeyDown;
			// 订阅 TextCompositionManager 事件
			TextCompositionManager.AddPreviewTextInputStartHandler(this, OnPreviewTextInputStart);
			TextCompositionManager.AddPreviewTextInputHandler(this, OnPreviewTextInput);
			TextCompositionManager.AddPreviewTextInputUpdateHandler(this, OnTextInputUpdate);

			// 初始化 Tooltip 计时器
			tooltipTimer = new DispatcherTimer
			{
				Interval = TimeSpan.FromSeconds(3)
			};
			tooltipTimer.Tick += TooltipTimer_Tick;
		}

		private void TooltipTimer_Tick(object sender, EventArgs e)
		{
			// 隐藏当前 Tooltip
			if (currentTooltip != null)
			{
				currentTooltip.IsOpen = false;
				currentTooltip = null;
			}

			// 停止计时器
			tooltipTimer.Stop();
		}

		private void GridSplitter_DragDelta(object sender, DragDeltaEventArgs e)
		{
			double currentLeft = Canvas.GetLeft(gridSplitter);
			double newLeft = currentLeft + e.HorizontalChange;

			if (newLeft <= RowHeaderWidth + ColumnHeaderWidth) newLeft = RowHeaderWidth + ColumnHeaderWidth;
			if (newLeft >= this.ActualWidth - ScrollbarSize - gridSplitter.Width)
				newLeft = this.ActualWidth - ScrollbarSize - gridSplitter.Width;

			Canvas.SetLeft(gridSplitter, newLeft);
			ArrangeScrollBars();
			//UpdateScrollBarMaxValue();
			this.InvalidateVisual();
		}

		private void Worksheet_SizeChanged(object sender, SizeChangedEventArgs e)
		{
			//UpdateScrollBarMaxValue();
		}
		#endregion

		#region 单元格编辑
		private void OnPreviewTextInputStart(object sender, TextCompositionEventArgs e)
		{
			if (editTextBox.Visibility != Visibility.Visible)
			{
				StartEditing();
			}
		}

		private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				// 处理最终文本
				editTextBox.Text += e.Text;
				editTextBox.CaretIndex = editTextBox.Text.Length;

				e.Handled = true;

				InvalidateVisual();
			}
		}

		private void OnTextInputUpdate(object sender, TextCompositionEventArgs e)
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				// 处理组合文本
				editTextBox.Text += e.TextComposition.CompositionText;
				editTextBox.CaretIndex = editTextBox.Text.Length;

				e.Handled = true;

				InvalidateVisual();
			}
		}

		private void Worksheet_KeyDown(object sender, KeyEventArgs e)
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				// 处理 TextBox 可见时的按键事件
				switch (e.Key)
				{
					case Key.Enter:
						SaveCurrentEdit(); // 保存当前编辑内容
						MoveSelection(1, 0); // 移动到下一行
						e.Handled = true;
						break;
					case Key.Tab:
						SaveCurrentEdit(); // 保存当前编辑内容
						MoveSelection(0, 1); // 移动到下一列
						e.Handled = true;
						break;
					case Key.Escape:
						CancelEdit(); // 取消编辑并恢复旧内容
						e.Handled = true;
						break;
					case Key.Left:
					case Key.Right:
					case Key.Up:
					case Key.Down:
						// 允许 TextBox 内部光标移动
						break;
					default:
						// 其他按键直接进入编辑模式
						editTextBox.Focus();
						break;
				}
			}
			else
			{
				// 处理 TextBox 不可见时的按键事件
				switch (e.Key)
				{
					case Key.Enter:
						MoveSelection(1, 0);
						e.Handled = true;
						break;
					case Key.Tab:
						MoveSelection(0, 1);
						e.Handled = true;
						break;
					case Key.Left:
						MoveSelection(0, -1);
						e.Handled = true;
						break;
					case Key.Right:
						MoveSelection(0, 1);
						e.Handled = true;
						break;
					case Key.Up:
						MoveSelection(-1, 0);
						e.Handled = true;
						break;
					case Key.Down:
						MoveSelection(1, 0);
						e.Handled = true;
						break;
					case Key.LeftShift:
						e.Handled = true;
						break;
					case Key.RightShift:
						e.Handled = true;
						break;
					default:
						// 其他按键直接进入编辑模式
						StartEditing();
						break;
				}
			}

			InvalidateVisual();
		}

		private void CancelEdit()
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				// 恢复旧内容
				SetCellValue(SelectedRowIndex, SelectedColIndex, tempCellValue);

				// 隐藏 TextBox
				editTextBox.Visibility = Visibility.Collapsed;

				// 更新视图
				InvalidateVisual();
			}
		}

		private void SaveCurrentEdit()
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				SetCellValue(SelectedRowIndex, SelectedColIndex, editTextBox.Text);

				// 隐藏 TextBox
				editTextBox.Clear();
				editTextBox.Visibility = Visibility.Collapsed;

				// 更新视图
				InvalidateVisual();
			}
		}

		private void EditTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			SaveCurrentEdit();
			isTyping = false;
		}

		private void MoveSelection(int rowOffset, int colOffset)
		{
			SelectedRowIndex = selectionRange.Row;
			SelectedColIndex = selectionRange.Col;

			if (SelectedRowIndex == 0 && SelectedColIndex == 0)
			{
				if (rowOffset < 0 || colOffset < 0) return;
			}

			if (SelectedRowIndex == DefaultRows - 1 && SelectedColIndex == DefaultCols - 1)
			{
				if (rowOffset > 0 || colOffset > 0)
				{
					SelectedColIndex = 0;
					SelectedRowIndex = 0;
					verScrollBar.Value = 0;
					horScrollBarGrid.Value = 0;
				}
			}
			else
			{
				SelectedRowIndex += rowOffset;
				SelectedColIndex += colOffset;

				if (SelectedRowIndex >= DefaultRows)
				{
					SelectedRowIndex = 0;
					SelectedColIndex++;
					verScrollBar.Value = 0;
				}
				else if (SelectedRowIndex < 0)
				{
					SelectedRowIndex = DefaultRows - 1;
					SelectedColIndex--;
				}

				if (SelectedColIndex >= DefaultCols)
				{
					SelectedColIndex = 0;
					SelectedRowIndex++;
					horScrollBarGrid.Value = 0;
				}
				else if (SelectedColIndex < 0)
				{
					SelectedColIndex = DefaultCols - 1;
					SelectedRowIndex--;
				}
			}

			Cols.ToList().ForEach(col => col.IsSelected = false);
			Rows.ToList().ForEach(row => row.IsSelected = false);

			selectionRange.Row = SelectedRowIndex;
			selectionRange.Col = SelectedColIndex;

			selectionRange.Rows = 1;
			selectionRange.Cols = 1;

			Cols[SelectedColIndex].IsSelected = true;
			Rows[SelectedRowIndex].IsSelected = true;

			InvalidateVisual();
		}

		private void StartEditing()
		{
			var col = Cols[SelectedColIndex];
			var row = Rows[SelectedRowIndex];

			editTextBox.Visibility = Visibility.Visible;
			Canvas.SetLeft(editTextBox, col.Left + RowHeaderWidth + 1);
			Canvas.SetTop(editTextBox, row.Top + ColumnHeaderHeight + 1);
			editTextBox.Width = col.Width - 2;
			editTextBox.Height = row.Height - 2;
			editTextBox.Focus();

			isTyping = true;
			//InvalidateVisual();
		}
		#endregion

		#region 视图渲染
		/// <summary>
		/// 绘制工作表Worksheet
		/// </summary>
		/// <param name="dc"></param>
		private void DrawView(DrawingContext dc)
		{
			UpdateScrollBarMaxValue();
			//先绘制甘特图刻度，确保图层顺序正确
			DrawScale(dc);

			#region 表格部分

			//分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);
			//绘制表格区域
			Rect rectGrid = new Rect(0, 0, splitterPos, this.ActualHeight);
			dc.DrawRectangle(new SolidColorBrush(Colors.Gray), null, rectGrid);

			var splitterLinePen = new Pen(new SolidColorBrush(Color.FromArgb(255, 200, 200, 200)), 0.5);
			var textBrush = Brushes.Black;
			Typeface typeface = new Typeface("Arial");
			double fontSize = 12;

			UpdateVisibleRegion(); // 更新可见区域

			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			double visibleWidth = splitterPos/* - SystemParameters.VerticalScrollBarWidth*/; // 减去垂直滚动条的宽度
			double visibleHeight = this.ActualHeight - SystemParameters.HorizontalScrollBarHeight; // 减去水平滚动条的高度
			double visibleColWidth = Cols[visibleRegion.EndCol].Width + Cols[visibleRegion.EndCol].Left - offsetX + RowHeaderWidth;
			double visibleRowHeight = Rows[visibleRegion.EndRow].Height + Rows[visibleRegion.EndRow].Top - offsetY + ColumnHeaderHeight;

			// 绘制可见区域背景
			Rect backRect = new Rect(0, 0, Math.Min(visibleColWidth, visibleWidth), Math.Min(visibleRowHeight, visibleHeight));
			dc.DrawRectangle(BackBrush, new Pen(BorderBrush, 1), backRect);

			// 绘制滚动条下面的灰色
			Rect scrollRect = new Rect(0, visibleHeight, splitterPos, ScrollbarSize);
			dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(255, 240, 240, 240)), null, scrollRect);

			// 绘制行标题
			double currentY = ColumnHeaderHeight/* - offsetY*/;
			for (int i = visibleRegion.StartRow; i <= visibleRegion.EndRow; i++)
			{
				double height = Rows[i].Height;

				if (height > 0 && Rows[i].IsVisible)
				{
					if (currentY + height > visibleHeight)
					{
						height = visibleHeight - currentY;
					}

					Rows[i].Draw(dc, 0, currentY, RowHeaderWidth, height, typeface, fontSize, textBrush);

					// 绘制行网格线
					dc.DrawLine(splitterLinePen, new Point(0, currentY), new Point(Math.Min(visibleColWidth, visibleWidth), currentY));


					// 如果Highlight为真
					if (Highlight && i == SelectedRowIndex)
					{
						var highlightPen = new Pen(Brushes.Green, 1.5);

						Rect rectHighlight = new Rect(new Point(0, currentY), new Point(this.ActualWidth/*Math.Min(visibleColWidth, visibleWidth)*/, currentY + height));
						dc.DrawLine(new Pen(new SolidColorBrush(Color.FromArgb(255, 180, 180, 180)), 1.5), new Point(0, currentY), new Point(this.ActualWidth, currentY));
						dc.DrawLine(new Pen(new SolidColorBrush(Color.FromArgb(255, 180, 180, 180)), 1.5), new Point(0, currentY + height), new Point(this.ActualWidth, currentY + height));

						//dc.DrawRectangle(new SolidColorBrush(Color.FromArgb(100, 200, 200, 200)), null, rectHighlight);
					}
				}

				currentY += height;
			}

			// 绘制列标题和内容
			double currentX = RowHeaderWidth/* - offsetX*/;
			for (int i = visibleRegion.StartCol; i <= visibleRegion.EndCol; i++)
			{
				double width = Cols[i].Width;

				if (width > 0)
				{
					// 确保列宽不会超出可见区域
					if (currentX + width > visibleWidth)
					{
						width = visibleWidth - currentX;
					}

					Cols[i].Draw(dc, currentX, 0, width, ColumnHeaderHeight, typeface, fontSize, textBrush);

					// 绘制网格线
					dc.DrawLine(splitterLinePen, new Point(currentX, 0), new Point(currentX, Math.Min(visibleRowHeight, visibleHeight)));

				}

				currentX += width;
			}

			// 绘制最后一列的右边界网格线
			dc.DrawLine(splitterLinePen, new Point(currentX, 0), new Point(currentX, Math.Min(visibleRowHeight, visibleHeight)));

			#endregion

			#region 甘特图部分
			DrawGanttChart(dc);
			#endregion
		}

		/// <summary>
		/// 绘制单元格
		/// </summary>
		/// <param name="dc"></param>
		private void DrawCells(DrawingContext dc)
		{
			//分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);

			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			foreach (var cellPair in cells)
			{
				string address = cellPair.Key;
				Cell cell = cellPair.Value;

				// 解析地址
				var position = Utility.Utility.FromAddress(address);
				int column = position.Col;
				int row = position.Row;
				if (Rows[row].IsVisible)
				{

					// 确保单元格在可见区域内
					if (column < visibleRegion.StartCol || column > visibleRegion.EndCol ||
					    row < visibleRegion.StartRow || row > visibleRegion.EndRow)
					{
						continue;
					}

					// 计算单元格的矩形区域
					double cellX = Cols.Take(column).Sum(c => c.Width) + RowHeaderWidth - offsetX;
					double cellY = Rows.Take(row).Sum(r => r.Height) + ColumnHeaderHeight - offsetY;
					double cellWidth = Cols[column].Width;
					double cellHeight = Rows[row].Height;

					Rect cellRect = new Rect(cellX, cellY, cellWidth, cellHeight);

					if (cellRect.IntersectsWith(new Rect(0, 0, splitterPos, this.ActualHeight))) // 仅绘制可见区域内的单元格
					{
						cell.Draw(dc, cellRect, this); // 传递当前的 Worksheet 实例作为 Visual 参数
					}
				}
			}
		}
		#endregion

		#region 行列
		// 设置列标题的方法
		public void SetColumnTitle(int columnIndex, string title, TextAlignment textAlignment, FontWeight fontWeight)
		{
			if (columnIndex < 0 || columnIndex >= Cols.Count)
			{
				throw new ArgumentOutOfRangeException(nameof(columnIndex), "Invalid column index.");
			}

			Cols[columnIndex].SetTitle(title, textAlignment, fontWeight);
			InvalidateVisual(); // 重新绘制控件以更新显示
		}

		// 设置列宽度
		public void SetColumnWidth(int columnIndex, double width)
		{
			if (width < 0) return;

			if (columnIndex >= 0 && columnIndex < Cols.Count)
			{
				Cols[columnIndex].Width = width;
				UpdateColumnPositions(); // 更新列的 Left 和 Right 属性
				UpdateScrollBarMaxValue(); // 更新滚动条的最大值
				this.InvalidateVisual();
			}
		}

		// 更新每列的 Left 和 Right 属性
		private void UpdateColumnPositions()
		{
			double currentLeft = 0; // 从行标题宽度之后开始

			for (int i = 0; i < Cols.Count; i++)
			{
				Cols[i].Left = currentLeft;
				currentLeft += Cols[i].Width;
			}
		}

		// 设置行高度
		public void SetRowHeight(int index, double height)
		{
			if (height < 0) return;
			if (index >= 0 && index < Rows.Count)
			{
				Rows[index].Height = height;
				UpdateRowPositions(); // 更新行的 Top 属性
				UpdateScrollBarMaxValue(); // 更新滚动条的最大值
				this.InvalidateVisual();
			}
		}

		public void SetColumnHeaderHeight(double height)
		{
			ColumnHeaderHeight = height;
			selectAllButton.Height = height;
			UpdateRowPositions();
			UpdateScrollBarMaxValue();
			this.InvalidateVisual();
		}

		// 更新每行的 Top 属性
		public void UpdateRowPositions()
		{
			double currentTop = 0; // 从列标题高度之后开始

			for (int i = 0; i < Rows.Count; i++)
			{
				if (Rows[i].IsVisible)
				{
					Rows[i].Height = RowHeaderHeight;
				}
				else
				{
					Rows[i].Height = 0;
				}

				Rows[i].Top = currentTop;
				currentTop += Rows[i].Height;
			}
		}

		public void SetCols(int colCount)
		{
			DefaultCols = colCount;
			Resize(-1, colCount);
		}

		public void SetRows(int rowCount)
		{
			DefaultRows = rowCount;
			Resize(rowCount, -1);
		}

		// 添加列
		public void AppendColumns(int count)
		{
			if (count < 0)
			{
				MessageBox.Show("添加的列数必须大于0。");
				return;
			}

			double x = Cols.Count == 0 ? 0 : Cols[Cols.Count - 1].Left + Cols[Cols.Count - 1].Width;
			int total = Cols.Count + count;
			for (int i = Cols.Count; i < total; i++)
			{
				Cols.Add(new ColumnHeader(i,
				    Utility.Utility.GetAlphaChar(i),
				    ColumnHeaderWidth, x));
				x += ColumnHeaderWidth;
			}
		}

		// 添加行
		public void AppendRows(int count)
		{
			if (count < 0)
			{
				MessageBox.Show("添加的行数必须大于0。");
				return;
			}

			double y = Rows.Count == 0 ? 0 : Rows[Rows.Count - 1].Top + Rows[Rows.Count - 1].Height;
			int total = Rows.Count + count;
			for (int i = Rows.Count; i < total; i++)
			{
				Rows.Add(new RowHeader(this, i, (i + 1).ToString(), RowHeaderHeight, y));
				y += RowHeaderHeight;
			}
			InvalidateVisual(); // 强制重绘以显示新行
		}
		#endregion

		#region 滚动条
		// 布局滚动条和全选按钮
		private void ArrangeScrollBars()
		{
			double availableHeight = this.ActualHeight - ScrollbarSize; // 实际高度减去水平滚动条的高度
			double availableWidth = this.ActualWidth - ScrollbarSize - RowHeaderWidth; // 实际宽度减去垂直滚动条的宽度和行标题宽度

			double splitterPos = Canvas.GetLeft(gridSplitter);

			if (double.IsNaN(splitterPos))
			{
				Canvas.SetLeft(gridSplitter, availableWidth / 2 + RowHeaderWidth);
			}
			Canvas.SetTop(gridSplitter, 0);
			gridSplitter.Height = this.ActualHeight;

			double availableGridWidth = splitterPos - RowHeaderWidth;
			double availableGanttWidth = availableWidth - splitterPos - gridSplitter.Width + RowHeaderWidth;

			// 设置垂直滚动条的高度
			verScrollBar.Height = Math.Max(0, availableHeight);
			// 设置水平滚动条的宽度
			horScrollBarGrid.Width = Math.Max(0, availableGridWidth);

			horScrollBarGantt.Width = Math.Max(0, availableGanttWidth);

			// 设置垂直滚动条的位置
			Canvas.SetLeft(verScrollBar, this.ActualWidth - ScrollbarSize);
			Canvas.SetTop(verScrollBar, 0);
			// 设置水平滚动条的位置
			Canvas.SetLeft(horScrollBarGrid, RowHeaderWidth);
			Canvas.SetTop(horScrollBarGrid, this.ActualHeight - ScrollbarSize);

			Canvas.SetLeft(horScrollBarGantt, gridSplitter.Width + splitterPos);
			Canvas.SetTop(horScrollBarGantt, this.ActualHeight - ScrollbarSize);

			// 设置全选按钮的位置
			Canvas.SetLeft(selectAllButton, 0);
			Canvas.SetTop(selectAllButton, 0);

			double totalWidth = Cols.Sum(col => col.Width);
			double totalHeight = Rows.Sum(row => row.Height);


			horScrollBarGrid.LargeChange = availableGridWidth > 0 ? 3 : 0;
			verScrollBar.LargeChange = availableHeight > 0 ? 3 : 0;
		}

		// 更新滚动条的最大值
		private void UpdateScrollBarMaxValue()
		{
			//分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);

			// 更新水平滚动条的最大值
			double totalWidth = 0 /*RowHeaderWidth*/;

			foreach (var col in Cols)
			{
				totalWidth += col.Width;
			}

			// 假设 CanvasWidth 是可见区域的宽度
			double canvasWidth = splitterPos;

			// 计算水平滚动条的最大值
			if (totalWidth > canvasWidth)
			{
				// 确保能够滚动到最右边列的右边缘
				double maxWidth = 0;
				double gap = totalWidth - canvasWidth + RowHeaderWidth;
				for (int i = 0; i < Cols.Count - 1; i++)
				{
					maxWidth += Cols[i].Width;
					if (gap - maxWidth < Cols[i + 1].Width)
					{
						maxWidth += Cols[i + 1].Width;
						break;
					}
				}

				this.horScrollBarGrid.Maximum = maxWidth;
			}
			else
			{
				this.horScrollBarGrid.Maximum = 0;
			}

			// 更新垂直滚动条的最大值
			double totalHeight = 0 /*ColumnHeaderHeight*/;

			foreach (var row in Rows)
			{
				totalHeight += row.Height;
			}

			// 假设 CanvasHeight 是可见区域的高度
			double canvasHeight = this.ActualHeight;

			// 计算垂直滚动条的最大值
			if (totalHeight > canvasHeight)
			{
				// 确保能够滚动到最底部行的底部边缘
				double maxHeight = 0;
				double gapHeight = totalHeight - canvasHeight + ColumnHeaderHeight;
				for (int i = 0; i < Rows.Count - 1; i++)
				{
					maxHeight += Rows[i].Height;
					if (gapHeight - maxHeight < Rows[i + 1].Height)
					{
						maxHeight += Rows[i + 1].Height;
						break;
					}
				}

				this.verScrollBar.Maximum = maxHeight;

				//double lastRowHeight = Rows.Count > 0 ? Rows[Rows.Count - 1].Height : 0;
				//this.verScrollBar.Maximum = totalHeight - canvasHeight + lastRowHeight;
			}
			else
			{
				this.verScrollBar.Maximum = 0;
			}
		}

		// 滚动条值变化事件处理
		private void OnScrollBarValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
		{
			if (isHandlingScrollEvent)
			{
				return;
			}

			isHandlingScrollEvent = true;

			try
			{
				InvalidateVisual();

				ScrollBar scrollBar = sender as ScrollBar;

				if (scrollBar.Orientation == Orientation.Horizontal)
				{
					HandleHorizontalScrollBar(e);
				}
				else if (scrollBar.Orientation == Orientation.Vertical)
				{
					HandleVerticalScrollBar(e);
				}

				scrollBar.ToolTip = scrollBar.Value.ToString();
			}
			finally
			{
				isHandlingScrollEvent = false;
			}
		}

		private void HandleHorizontalScrollBar(RoutedPropertyChangedEventArgs<double> e)
		{
			if (Cols.Count > 0)
			{
				//分隔条位置
				double splitterPos = Canvas.GetLeft(gridSplitter);

				double changeAmount = e.NewValue - e.OldValue;
				double widthShift = 0;

				if (changeAmount > 0)
				{
					// 向右滚动
					if (changeAmount == 1)
					{
						// 移动一列
						widthShift = this.Cols[this.visibleRegion.StartCol].Width;
						this.horScrollBarGrid.Value += -changeAmount;

						this.horScrollBarGrid.Value += widthShift;
					}
					else
					{
						widthShift += this.Cols[this.visibleRegion.EndCol].Left;

						this.horScrollBarGrid.Value = widthShift;
					}
				}
				else
				{
					// 向左滚动
					if (changeAmount == -1)
					{
						// 移动一列
						widthShift = -this.Cols[this.visibleRegion.StartCol - 1].Width;
						this.horScrollBarGrid.Value += -changeAmount;

						this.horScrollBarGrid.Value += widthShift;
					}
					else
					{
						// 大量移动
						double visibleWidth = splitterPos /*- SystemParameters.VerticalScrollBarWidth*/;
						double totalWidth = 0;
						int i = this.visibleRegion.StartCol;

						// 计算新的可见区域最左边是哪一列
						while (i > 0 && totalWidth < visibleWidth)
						{
							i--;
							totalWidth += this.Cols[i].Width;
						}

						// 确保前一屏的最后一列完整显示在新屏幕上
						if (this.visibleRegion.StartCol > 0)
						{
							totalWidth = 0;
							i = this.visibleRegion.StartCol - 1;
							while (i >= 0 && totalWidth + this.Cols[i].Width <= visibleWidth)
							{
								totalWidth += this.Cols[i].Width;
								i--;
							}
						}

						// 更新滚动条的值
						this.horScrollBarGrid.Value = this.Cols[Math.Max(0, i + 1)].Left;
					}
				}

				// 确保滚动条的值不会小于0
				if (this.horScrollBarGrid.Value < 0)
				{
					this.horScrollBarGrid.Value = 0;
				}
			}
		}

		private void HandleVerticalScrollBar(RoutedPropertyChangedEventArgs<double> e)
		{
			if (Rows.Count > 0)
			{
				double changeAmount = e.NewValue - e.OldValue;
				double heightShift = 0;

				if (changeAmount > 0)
				{
					// 向下滚动
					if (changeAmount == 1)
					{
						// 移动一行
						heightShift = this.Rows[this.visibleRegion.StartRow].Height;
						this.verScrollBar.Value += -changeAmount;

						this.verScrollBar.Value += heightShift;
					}
					else
					{
						heightShift += this.Rows[this.visibleRegion.EndRow].Top;

						this.verScrollBar.Value = heightShift;
					}
				}
				else
				{
					// 向上滚动
					if (changeAmount == -1)
					{
						// 移动一行
						heightShift = -this.Rows[this.visibleRegion.StartRow - 1].Height;
						this.verScrollBar.Value += -changeAmount;

						this.verScrollBar.Value += heightShift;
					}
					else
					{
						// 大量移动
						double visibleHeight = this.ActualHeight - SystemParameters.HorizontalScrollBarHeight;
						double totalHeight = 0;
						int i = this.visibleRegion.StartRow;

						// 计算新的可见区域最上边是哪一行
						while (i > 0 && totalHeight < visibleHeight)
						{
							i--;
							totalHeight += this.Rows[i].Height;
						}

						// 确保前一屏的最后一行完整显示在新屏幕上
						if (this.visibleRegion.StartRow > 0)
						{
							totalHeight = 0;
							i = this.visibleRegion.StartRow - 1;
							while (i >= 0 && totalHeight + this.Rows[i].Height <= visibleHeight)
							{
								totalHeight += this.Rows[i].Height;
								i--;
							}
						}

						// 更新滚动条的值
						this.verScrollBar.Value = this.Rows[Math.Max(0, i + 1)].Top;
					}
				}

				// 确保滚动条的值不会小于0
				if (this.verScrollBar.Value < 0)
				{
					this.verScrollBar.Value = 0;
				}
			}
		}

		#endregion

		#region 鼠标
		private void OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
			Point mousePosition = e.GetPosition(this);

			var cellPosX = GetColumnIndexAtPosition(mousePosition.X);
			var cellPosY = GetRowIndexAtPosition(mousePosition.Y);
			var cellClicked = GetOrCreateCell(cellPosY, cellPosX);

			if (mousePosition.X >= cellClicked.toggleX
				&& mousePosition.X <= cellClicked.toggleX + cellClicked.toggleButtonSize
				&& mousePosition.Y >= cellClicked.toggleY
				&& mousePosition.Y <= cellClicked.toggleY + cellClicked.toggleButtonSize)
			{
				return;
			}

			// 计算鼠标点击位置对应的单元格
			int colIndex = GetColumnIndexAtPosition(mousePosition.X);
			int rowIndex = GetRowIndexAtPosition(mousePosition.Y);
			Cell cell = GetOrCreateCell(rowIndex, colIndex);


			string value = cell.InnerData;
			SetCellValue(rowIndex, colIndex, null);

			//这里强制更新视图，是要隐藏单元格内容，避免重叠
			InvalidateVisual();

			if (colIndex >= 0 && colIndex < Cols.Count && rowIndex >= 0 && rowIndex < Rows.Count)
			{
				var col = Cols[colIndex];
				var row = Rows[rowIndex];



				// 显示 TextBox 并设置其位置和大小
				editTextBox.Text = value; // 获取单元格的当前值


				editTextBox.Visibility = Visibility.Visible;
				Canvas.SetLeft(editTextBox, col.Left + RowHeaderWidth + 1);
				Canvas.SetTop(editTextBox, row.Top + ColumnHeaderHeight + 1);
				editTextBox.Width = col.Width - 2;
				editTextBox.Height = row.Height - 2;
				editTextBox.SelectionStart = 0;
				editTextBox.SelectionLength = editTextBox.Text.Length;
				//editTextBox.SelectedText = editTextBox.Text;
				editTextBox.Focus();
			}

			isSelecting = false;
		}

		private void OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
		{
			Point clickPosition = e.GetPosition(this);
			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			var cellPosX = GetColumnIndexAtPosition(clickPosition.X);
			var cellPosY = GetRowIndexAtPosition(clickPosition.Y);
			var cellClicked = GetOrCreateCell(cellPosY, cellPosX);

			MessageBox.Show(cellClicked.IsExpanded.ToString());

		}

		// 鼠标左键点击事件处理
		private void OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
		{
			if (editTextBox.Visibility == Visibility.Visible)
			{
				EditTextBox_LostFocus(sender, e);
			}

			Point clickPosition = e.GetPosition(this);
			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			if (e.ClickCount == 2)
			{
				OnMouseDoubleClick(sender, e);
				return;
			}

			var cellPosX = GetColumnIndexAtPosition(clickPosition.X);
			var cellPosY = GetRowIndexAtPosition(clickPosition.Y);
			var cellClicked = GetOrCreateCell(cellPosY, cellPosX);

			if (clickPosition.X >= cellClicked.toggleX
				&& clickPosition.X <= cellClicked.toggleX + cellClicked.toggleButtonSize
				&& clickPosition.Y >= cellClicked.toggleY
				&& clickPosition.Y <= cellClicked.toggleY + cellClicked.toggleButtonSize)
			{
				cellClicked.IsExpanded = !cellClicked.IsExpanded;
			}

			// 检查列宽调整
			if (clickPosition.Y <= ColumnHeaderHeight)
			{
				int index = visibleRegion.StartCol;
				double currentX = RowHeaderWidth;

				for (int i = 0; i < visibleRegion.Cols; i++)
				{
					currentX += this.Cols[index].Width;
					if (Math.Abs(clickPosition.X - currentX) < 5)
					{
						isResizingColumn = true;
						resizingColumnIndex = index;
						initialMouseX = clickPosition.X;
						initialColumnWidth = Cols[index].Width;
						Mouse.Capture(this);
						return;
					}
					index++;
				}

				// 点击在列标题区域，不改变选区
				return;
			}

			// 检查行高调整
			if (clickPosition.X <= RowHeaderWidth)
			{
				double currentY = ColumnHeaderHeight;
				for (int i = 0; i < Rows.Count; i++)
				{
					currentY += Rows[i].Height;
					if (Math.Abs(clickPosition.Y - currentY) < 5)
					{
						isResizingRow = true;
						resizingRowIndex = i;
						initialMouseY = clickPosition.Y;
						initialRowHeight = Rows[i].Height;
						Mouse.Capture(this);
						return;
					}
				}

				// 点击在行标题区域，不改变选区
				return;
			}

			// 检查是否点击在可见区域之外
			if (clickPosition.X > RowHeaderWidth + GetVisibleWidth() || clickPosition.Y > ColumnHeaderHeight + GetVisibleHeight())
			{
				// 点击在可见区域之外，不改变选区
				return;
			}

			// 开始选择区域
			isSelecting = true;
			selectionStart = new Point(clickPosition.X + offsetX /*- RowHeaderWidth*/, clickPosition.Y + offsetY /*- ColumnHeaderHeight*/);
			selectionEnd = selectionStart;
			UpdateSelectionRange();  // 立即更新选区范围
			this.InvalidateVisual();
			Mouse.Capture(this);
		}

		// 鼠标移动事件处理
		private void OnMouseMove(object sender, MouseEventArgs e)
		{
			Point currentPosition = e.GetPosition(this);
			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			#region 表格部分
			bool isNearColumnEdge = false;
			bool isNearRowEdge = false;

			var cellPosX = GetColumnIndexAtPosition(currentPosition.X);
			var cellPosY = GetRowIndexAtPosition(currentPosition.Y);
			var cellClicked = GetOrCreateCell(cellPosY, cellPosX);
			if (currentPosition.X >= cellClicked.toggleX
			    && currentPosition.X <= cellClicked.toggleX + cellClicked.toggleButtonSize
			    && currentPosition.Y >= cellClicked.toggleY
			    && currentPosition.Y <= cellClicked.toggleY + cellClicked.toggleButtonSize)
			{
				this.Cursor = Cursors.Hand;
				return;
			}
			else
			{
				this.Cursor = Cursors.Arrow;
			}

			// 只有当光标在列标题区域时，才检查列宽调整
			if (currentPosition.Y <= ColumnHeaderHeight)
			{
				int index = visibleRegion.StartCol;
				double currentX = RowHeaderWidth;

				for (int i = 0; i < visibleRegion.Cols; i++)
				{
					currentX += this.Cols[index].Width;
					if (Math.Abs(currentPosition.X - currentX) < 5)
					{
						isNearColumnEdge = true;
						break;
					}
					index++;
				}
			}

			// 只有当光标在行标题区域时，才检查行高调整
			if (currentPosition.X <= RowHeaderWidth)
			{
				double currentY = ColumnHeaderHeight;
				for (int i = 0; i < Rows.Count; i++)
				{
					currentY += Rows[i].Height;
					if (Math.Abs(currentPosition.Y - currentY) < 5)
					{
						isNearRowEdge = true;
						break;
					}
				}
			}

			if (isNearColumnEdge)
			{
				this.Cursor = Cursors.SizeWE;
			}
			else if (isNearRowEdge)
			{
				this.Cursor = Cursors.SizeNS;
			}
			else
			{
				this.Cursor = Cursors.Arrow;
			}

			// 列宽调整
			if (isResizingColumn && resizingColumnIndex >= 0)
			{
				double deltaX = currentPosition.X - initialMouseX;
				double newWidth = initialColumnWidth + deltaX;
				if (newWidth > 10)
				{
					SetColumnWidth(resizingColumnIndex, newWidth);
				}
			}

			// 行高调整
			if (isResizingRow && resizingRowIndex >= 0)
			{
				double deltaY = currentPosition.Y - initialMouseY;
				double newHeight = initialRowHeight + deltaY;
				if (newHeight > 10)
				{
					SetRowHeight(resizingRowIndex, newHeight);
				}
			}

			// 更新选区
			if (isSelecting)
			{
				selectionEnd = new Point(currentPosition.X + offsetX /*- RowHeaderWidth*/, currentPosition.Y + offsetY /*- ColumnHeaderHeight*/);
				UpdateSelectionRange();
				this.InvalidateVisual();
			}
			#endregion

			#region 甘特图部分
			double splitterPos = Canvas.GetLeft(gridSplitter);
			double scrollOffset = horScrollBarGantt.Value;
			double relativeX = currentPosition.X - splitterPos - gridSplitter.Width + scrollOffset;

			// 计算甘特图的总宽度
			double totalGanttWidth = (RealEndDate - RealStartDate).TotalDays * WidthRatio;

			if (relativeX >= 0 && relativeX < totalGanttWidth)
			{
				int dayIndex = (int)(relativeX / WidthRatio);
				DateTime currentDate = RealStartDate.AddDays(dayIndex);

				// 检查鼠标是否在下部小刻度区域
				if (currentPosition.Y >= ColumnHeaderHeight / 2 && currentPosition.Y <= ColumnHeaderHeight)
				{
					// 仅在日期变化时更新 Tooltip
					if (currentDate != lastTooltipDate)
					{
						lastTooltipDate = currentDate;

						// 关闭当前 Tooltip
						if (currentTooltip != null)
						{
							currentTooltip.IsOpen = false;
						}

						string tooltipText = currentDate.ToString("yyyy-MM-dd");

						currentTooltip = new ToolTip
						{
							Content = tooltipText,
							IsOpen = true,
							Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse
						};

						// 重置并启动计时器
						tooltipTimer.Stop();
						tooltipTimer.Start();
					}
				}
				// 检查鼠标是否在上部大刻度区域
				else if (currentPosition.Y >= 0 && currentPosition.Y < ColumnHeaderHeight / 2)
				{
					int weekIndex = dayIndex / 7;
					DateTime weekStartDate = RealStartDate.AddDays(weekIndex * 7);
					DateTime weekEndDate = weekStartDate.AddDays(6);

					// 仅在周变化时更新 Tooltip
					if (weekStartDate != lastTooltipDate)
					{
						lastTooltipDate = weekStartDate;

						// 关闭当前 Tooltip
						if (currentTooltip != null)
						{
							currentTooltip.IsOpen = false;
						}

						string tooltipText = $"{weekStartDate:yyyy/MM/dd} - {weekEndDate:yyyy/MM/dd}";

						currentTooltip = new ToolTip
						{
							Content = tooltipText,
							IsOpen = true,
							Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse
						};

						// 重置并启动计时器
						tooltipTimer.Stop();
						tooltipTimer.Start();
					}
				}
				else
				{
					// 检查鼠标是否悬停在甘特图任务条形上
					//GetTasks();
					if (TaskList == null) return;
					double currentY = ColumnHeaderHeight;
					bool tooltipShown = false;

					for (int index = 0; index < TaskList.Count; index++)
					{
						if (!Rows[index].IsVisible)
						{
							continue;
						}

						GanttTask task = TaskList[index];

						// 计算任务条形的位置和大小
						double taskStartX = (task.StartDate - RealStartDate).TotalDays * WidthRatio + splitterPos + gridSplitter.Width - scrollOffset;
						double taskWidth = task.Duration * WidthRatio;

						// 计算任务条形的 Y 值
						double taskY = currentY - verScrollBar.Value + (ColumnHeaderHeight - ganttBarSize) / 4;

						// 更新 currentY 以便下一个可见任务使用
						currentY += RowHeaderHeight;

						// 检查任务条形是否在可见区域内，并裁剪任务条形
						double visibleStartX = Math.Max(taskStartX, splitterPos + gridSplitter.Width);
						double visibleEndX = Math.Min(taskStartX + taskWidth, splitterPos + gridSplitter.Width + totalGanttWidth);
						double visibleWidth = visibleEndX - visibleStartX;

						double visibleStartY = Math.Max(taskY, ColumnHeaderHeight); // 确保任务条形不在列标题之上
						double visibleEndY = Math.Min(taskY + ganttBarSize, this.ActualHeight);
						double visibleHeight = visibleEndY - visibleStartY;

						if (visibleWidth > 0 && visibleHeight > 0)
						{
							Rect taskRect = new Rect(visibleStartX, visibleStartY, visibleWidth, visibleHeight);

							// 检查鼠标是否在任务条形上
							if (taskRect.Contains(currentPosition))
							{
								if (currentTooltip != null)
								{
									currentTooltip.IsOpen = false;
								}

								string tooltipText = $"{task.TaskName}  [{task.Completion:P}]";

								taskPointed = task.Index;

								currentTooltip = new ToolTip
								{
									Content = tooltipText,
									IsOpen = true,
									Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse
								};

								tooltipShown = true;
								break;
							}
							else
							{
								taskPointed = -1;
							}
						}
					}

					if (!tooltipShown)
					{
						if (currentTooltip != null)
						{
							currentTooltip.IsOpen = false;
							currentTooltip = null;
							lastTooltipDate = DateTime.MinValue;
						}
					}
					InvalidateVisual();
				}
			}
			else
			{
				// 如果鼠标不在有效范围内，关闭 Tooltip
				if (currentTooltip != null)
				{
					currentTooltip.IsOpen = false;
					currentTooltip = null;
					lastTooltipDate = DateTime.MinValue;
				}
			}

			#endregion
		}

		// 鼠标左键释放事件处理
		private void OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
		{
			isResizingColumn = false;
			resizingColumnIndex = -1;
			isResizingRow = false;
			resizingRowIndex = -1;
			isSelecting = false;
			Mouse.Capture(null);

			// 触发 SelectionChanged 事件
			OnSelectionChanged(new CellEventArgs(SelectedCell));
		}

		// 更新选区范围
		private void UpdateSelectionRange()
		{
			int startCol = GetColumnIndexAtPosition(selectionStart.X);
			int endCol = GetColumnIndexAtPosition(selectionEnd.X);
			int startRow = GetRowIndexAtPosition(selectionStart.Y);
			int endRow = GetRowIndexAtPosition(selectionEnd.Y);

			SelectedColIndex = startCol;
			SelectedRowIndex = startRow;

			selectionRange = new RangePosition(
			    Math.Min(startRow, endRow),
			    Math.Min(startCol, endCol),
			    Math.Abs(endRow - startRow) + 1,
			    Math.Abs(endCol - startCol) + 1
			);

			for (int i = 0; i < Cols.Count; i++)
			{
				Cols[i].IsSelected = i >= startCol && i <= endCol;
			}

			for (int i = 0; i < Rows.Count; i++)
			{
				Rows[i].IsSelected = i >= startRow && i <= endRow;
			}
		}

		// 获取指定位置的列索引
		private int GetColumnIndexAtPosition(double position)
		{
			double currentX = 0;
			for (int i = 0; i < Cols.Count; i++)
			{
				currentX += Cols[i].Width;
				if (currentX > position - RowHeaderWidth)
				{
					return i;
				}
			}
			return -1;
		}

		// 获取指定位置的行索引
		private int GetRowIndexAtPosition(double position)
		{
			double currentY = 0;
			for (int i = 0; i < Rows.Count; i++)
			{
				currentY += Rows[i].Height;
				if (currentY > position - ColumnHeaderHeight)
				{
					return i;
				}
			}
			return -1;
		}
		#endregion

		private void OnLayoutUpdated(object sender, EventArgs e)
		{
			if (isInitialized)
			{
				ArrangeScrollBars();
				this.LayoutUpdated -= OnLayoutUpdated; // 只需处理一次
			}
		}

		// 重写窗口大小变化事件
		protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
		{
			base.OnRenderSizeChanged(sizeInfo);
			ArrangeScrollBars();
			InvalidateVisual();
		}

		// 绘制视图
		// 渲染
		protected override void OnRender(DrawingContext dc)
		{
			base.OnRender(dc);
			DrawView(dc);
			DrawCells(dc);
			DrawSelection(dc);
		}
		/// <summary>
		/// 设置哪一列是层级列，比如项目管理中的任务名称
		/// </summary>
		/// <param name="col">列索引</param>
		public void SetHierarchyCol(int col)
		{
			this.HierarchyColIndex = col;
		}

		private void ShowTooltip(Point currentPosition)
		{
			// 只有当光标在列标题区域时，才显示 Tooltip
			if (currentPosition.Y <= ColumnHeaderHeight)
			{
				double currentX = RowHeaderWidth - horScrollBarGrid.Value;
				for (int i = 0; i < Cols.Count; i++)
				{
					currentX += Cols[i].Width;
					if (currentPosition.X >= currentX - Cols[i].Width && currentPosition.X <= currentX)
					{
						// 显示 Tooltip
						if (currentTooltip != Cols[i].Tooltip)
						{
							if (currentTooltip != null)
							{
								currentTooltip.IsOpen = false;
							}
							currentTooltip = Cols[i].Tooltip;
							currentTooltip.IsOpen = true;
							currentTooltip.PlacementTarget = this;
							currentTooltip.Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse;

							// 设置3秒后自动关闭
							var timer = new System.Windows.Threading.DispatcherTimer
							{
								Interval = TimeSpan.FromSeconds(3)
							};
							timer.Tick += (s, args) =>
							{
								currentTooltip.IsOpen = false;
								timer.Stop();
							};
							timer.Start();
						}
						return;
					}
				}
			}

			// 如果光标不在任何列标题上，关闭当前 Tooltip
			if (currentTooltip != null)
			{
				currentTooltip.IsOpen = false;
				currentTooltip = null;
			}
		}

		private double GetVisibleWidth()
		{
			double visibleWidth = 0;
			for (int i = visibleRegion.StartCol; i < visibleRegion.StartCol + visibleRegion.Cols; i++)
			{
				visibleWidth += Cols[i].Width;
			}
			return visibleWidth;
		}

		private double GetVisibleHeight()
		{
			double visibleHeight = 0;
			for (int i = visibleRegion.StartRow; i < visibleRegion.StartRow + visibleRegion.Rows; i++)
			{
				visibleHeight += Rows[i].Height;
			}
			return visibleHeight;
		}
		private void UpdateVisibleRegion()
		{
			//分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);

			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			double visibleWidth = splitterPos /*- SystemParameters.VerticalScrollBarWidth*/; // 减去垂直滚动条的宽度
			double visibleHeight = this.ActualHeight - SystemParameters.HorizontalScrollBarHeight; // 减去水平滚动条的高度

			// 更新可见区域信息
			double rowOffsetY = 0;
			visibleRegion.StartRow = 0;

			// 计算可见区域的起始行
			for (int i = 0; i < Rows.Count; i++)
			{
				if (rowOffsetY + Rows[i].Height > offsetY)
				{
					visibleRegion.StartRow = i;
					break;
				}
				rowOffsetY += Rows[i].Height;
			}

			// 计算可见区域的结束行
			double remainingHeight = visibleHeight - ColumnHeaderHeight;
			visibleRegion.EndRow = visibleRegion.StartRow;

			for (int i = visibleRegion.StartRow; i < Rows.Count; i++)
			{
				if (remainingHeight <= 0)
					break;

				remainingHeight -= Rows[i].Height;
				visibleRegion.EndRow = i;
			}

			// 找到第一个可见的列
			visibleRegion.StartCol = 0;
			double colOffsetX = 0;
			for (int i = 0; i < Cols.Count; i++)
			{
				if (colOffsetX + Cols[i].Width > offsetX)
				{
					visibleRegion.StartCol = i;
					break;
				}
				colOffsetX += Cols[i].Width;
			}

			// 计算可见的列数
			double remainingWidth = visibleWidth - RowHeaderWidth;
			visibleRegion.EndCol = visibleRegion.StartCol;

			for (int i = visibleRegion.StartCol; i < Cols.Count; i++)
			{
				if (remainingWidth <= 0)
					break;

				remainingWidth -= Cols[i].Width;
				visibleRegion.EndCol = i;
			}
		}
		// 全选功能
		private void SelectAll()
		{
			foreach (var col in Cols)
			{
				col.IsSelected = true;
			}
			foreach (var row in Rows)
			{
				row.IsSelected = true;
			}
			InvalidateVisual();
		}

		// 调整列数和行数
		public void Resize(int Rows, int colsCount)
		{
			if (colsCount > 0)
			{
				if (colsCount > this.Cols.Count)
				{
					AppendColumns(colsCount - this.Cols.Count);
				}
				else if (colsCount < this.Cols.Count)
				{
					// DeleteColumns(colsCount, this.Cols.Count - colsCount);
				}
			}
			if (Rows > 0)
			{
				if (Rows > this.Rows.Count)
				{
					AppendRows(Rows - this.Rows.Count);
				}
			}

			//UpdateVisibleRegion();
		}


		#region 单元格赋值和设置格式
		public void SetCellStyle(string address,
			string fontName,
			double fontSize,
			TextAlignment textAlignment,
			double indent,
			SolidColorBrush foreground)
		{
			SetCellStyle(address, fontName, fontSize, textAlignment, indent, foreground, FontWeights.Normal);
		}
		public void SetCellStyle(string address,
			string fontName,
			double fontSize,
			TextAlignment textAlignment,
			double indent,
			SolidColorBrush foreground,
			FontWeight fontWeight)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.SetStyle(
				fontName: fontName,
				fontSize: fontSize,
				textAlignment: textAlignment,
				indent: indent,
				foreground: foreground,
				fontWeight: fontWeight);
		}

		private string GetCellValue(int rowIndex, int colIndex)
		{
			// 获取单元格的当前值（根据您的数据结构实现）
			// 这里假设有一个二维数组 `cellValues` 存储单元格数据
			Cell cell = GetOrCreateCell(rowIndex, colIndex);
			return cell.InnerData;
		}

		private void SetCellValue(int rowIndex, int colIndex, string value)
		{
			// 设置单元格的值（根据您的数据结构实现）
			// 这里假设有一个二维数组 `cellValues` 存储单元格数据
			Cell cell = GetOrCreateCell(rowIndex, colIndex);
			if (cell.InnerData != value)
			{
				cell.InnerData = value;
				OnChanged(new CellEventArgs(cell));
			}
		}

		public void SetValue(string address, string value)
		{
			SetValue(address, value, null);
		}
		public void SetValue(string address, string value, string parent)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			SetSingleCellData(cell.Row, cell.Column, value);

			if (parent != null)
			{
				var posParent = Utility.Utility.FromAddress(parent);
				var cellParent = GetOrCreateCell(posParent.Row, posParent.Col);
				cell.Parent = cellParent;
			}
		}

		public void SetBackground(string address, Brush background)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.Background = background;
			InvalidateVisual();
		}

		public void SetForeground(string address, Brush foreground)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.Foreground = foreground;
			InvalidateVisual();
		}

		public void SetFont(string address, Typeface font, double fontSize)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.Font = font;
			cell.FontSize = fontSize;
			InvalidateVisual();
		}

		public void SetBorder(string address, double thickness, Brush brush)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.BorderAllThickness = thickness;
			cell.BorderAllBrush = brush;
			InvalidateVisual();
		}

		public Cell GetOrCreateCell(int row, int col)
		{
			string address = Utility.Utility.ToAddress(row, col);
			if (!cells.ContainsKey(address))
			{
				cells[address] = new Cell(this, row, col);
			}
			return cells[address];
		}

		// 修改 SetSingleCellData 方法
		public void SetSingleCellData(int row, int col, string data)
		{
			var cell = GetOrCreateCell(row, col);
			cell.InnerData = data; // 设置 InnerData 而不是 Data
			OnChanged(new CellEventArgs(cell));

			InvalidateVisual(); // 触发重绘
		}

		public void SetCellBorder(string address, Brush topBrush, double topThickness, Brush bottomBrush, double bottomThickness, Brush leftBrush, double leftThickness, Brush rightBrush, double rightThickness)
		{
			var position = Utility.Utility.FromAddress(address);
			var cell = GetOrCreateCell(position.Row, position.Col);
			cell.BorderTopBrush = topBrush;
			cell.BorderTopThickness = topThickness;
			cell.BorderBottomBrush = bottomBrush;
			cell.BorderBottomThickness = bottomThickness;
			cell.BorderLeftBrush = leftBrush;
			cell.BorderLeftThickness = leftThickness;
			cell.BorderRightBrush = rightBrush;
			cell.BorderRightThickness = rightThickness;

		}
		#endregion

		#region 单元格区域
		private bool isSelecting = false;
		private Point selectionStart;
		private Point selectionEnd;

		private CellPosition selStart = new CellPosition(0, 0);
		private CellPosition selEnd = new CellPosition(0, 0);
		private CellPosition focusPos = new CellPosition(0, 0);
		private RangePosition selectionRange = new RangePosition(0, 0, 1, 1);

		public CellPosition FocusPos
		{
			get => focusPos;
			set
			{
				if (focusPos != value)
				{
					focusPos = value;
					//FocusPosChanged?.Invoke(this, new CellPosEventArgs(value));
				}
			}
		}

		public RangePosition SelectionRange
		{
			get => selectionRange;
			set
			{
				if (selectionRange != value)
				{
					selectionRange = value;
					//SelectionRangeChanged?.Invoke(this, new RangeEventArgs(value));
				}
			}
		}

		//public event EventHandler<CellPosEventArgs> FocusPosChanged;
		//public event EventHandler<RangeEventArgs> SelectionRangeChanged;

		public void SelectRange(CellPosition start, CellPosition end)
		{
			selStart = start;
			selEnd = end;
			SelectionRange = new RangePosition(start, end);
			FocusPos = start;
		}

		private void DrawSelection(DrawingContext dc)
		{
			//分隔条位置
			double splitterPos = Canvas.GetLeft(gridSplitter);

			double offsetX = horScrollBarGrid.Value;
			double offsetY = verScrollBar.Value;

			// 获取选区的起点和终点
			var startCol = selectionRange.StartPos.Col;
			var endCol = selectionRange.EndPos.Col;
			var startRow = selectionRange.StartPos.Row;
			var endRow = selectionRange.EndPos.Row;

			// 检查选区是否超出有效行列范围
			if (startCol < 0 || endCol >= Cols.Count || startRow < 0 || endRow >= Rows.Count)
			{
				// 如果选区超出范围，直接返回，不绘制选区
				return;
			}

			// 计算选区矩形，考虑滚动条的偏移量
			double selectionLeft = Cols[startCol].Left + RowHeaderWidth - offsetX;
			double selectionTop = Rows[startRow].Top + ColumnHeaderHeight - offsetY;
			double selectionRight = Cols[endCol].Left + Cols[endCol].Width + RowHeaderWidth - offsetX;
			double selectionBottom = Rows[endRow].Top + Rows[endRow].Height + ColumnHeaderHeight - offsetY;

			// 检查选区是否在可见区域内
			if (selectionRight <= RowHeaderWidth || selectionBottom <= ColumnHeaderHeight ||
			    selectionLeft >= splitterPos || selectionTop >= this.ActualHeight)
			{
				// 如果选区不在可见区域内，直接返回，不绘制选区
				return;
			}

			// 限制选区在可见区域内
			selectionLeft = Math.Max(selectionLeft, RowHeaderWidth);
			selectionTop = Math.Max(selectionTop, ColumnHeaderHeight);
			selectionRight = Math.Min(selectionRight, splitterPos);
			selectionBottom = Math.Min(selectionBottom, this.ActualHeight);

			// 计算选区矩形
			Rect selectionRect = new Rect(selectionLeft, selectionTop, selectionRight - selectionLeft, selectionBottom - selectionTop);

			// 计算活动单元格矩形，考虑滚动条的偏移量
			var activeCellCol = Cols[startCol];
			var activeCellRow = Rows[startRow];
			Rect activeCellRect = new Rect(activeCellCol.Left + RowHeaderWidth - offsetX, activeCellRow.Top + ColumnHeaderHeight - offsetY, activeCellCol.Width, activeCellRow.Height);

			// 检查活动单元格是否在可见区域内
			if (activeCellRect.Right <= RowHeaderWidth || activeCellRect.Bottom <= ColumnHeaderHeight ||
			    activeCellRect.Left >= splitterPos || activeCellRect.Top >= this.ActualHeight)
			{
				// 如果活动单元格不在可见区域内，直接返回，不绘制活动单元格
				return;
			}

			// 限制活动单元格在可见区域内
			activeCellRect = new Rect(
			    Math.Max(activeCellRect.Left, RowHeaderWidth),
			    Math.Max(activeCellRect.Top, ColumnHeaderHeight),
			    Math.Min(activeCellRect.Right, splitterPos) - Math.Max(activeCellRect.Left, RowHeaderWidth),
			    Math.Min(activeCellRect.Bottom, this.ActualHeight) - Math.Max(activeCellRect.Top, ColumnHeaderHeight)
			);

			// 绘制半透明灰色背景
			SolidColorBrush selectionFillColor = new SolidColorBrush(Color.FromArgb(50, 0, 0, 0));

			// 创建选区几何体
			var selectionGeometry = new RectangleGeometry(selectionRect);

			// 创建活动单元格几何体
			var activeCellGeometry = new RectangleGeometry(activeCellRect);

			// 排除活动单元格区域
			var combinedGeometry = new CombinedGeometry(GeometryCombineMode.Exclude, selectionGeometry, activeCellGeometry);

			// 绘制选区（排除活动单元格区域）
			dc.DrawGeometry(selectionFillColor, null, combinedGeometry);

			// 绘制选区边框
			Pen selectionBorderPen = new Pen(new SolidColorBrush(Colors.Green), 1);
			dc.DrawRectangle(null, selectionBorderPen, selectionRect);

			// 绘制右下角的拖拽手柄
			Rect thumbRect = new Rect(selectionRect.Right - 4, selectionRect.Bottom - 4, 5, 5);
			dc.DrawRectangle(new SolidColorBrush(Colors.Green), new Pen(new SolidColorBrush(Color.FromArgb(255, 255, 255, 255)), 1), thumbRect);
		}



		#endregion
	}
}
