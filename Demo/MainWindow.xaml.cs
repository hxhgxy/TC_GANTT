using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Worksheet;
using Worksheet.Models;

namespace Demo
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();

			this.DataContext = this;

			tcWorksheet.SetRows(100);
			tcWorksheet.SetCols(10);
			tcWorksheet.Focus();

			#region 表格部分

			#region 任务

			tcWorksheet.SetHierarchyCol(1);
			tcWorksheet.SetColumnTitle(1, "任务", TextAlignment.Left, FontWeights.ExtraBold);
			tcWorksheet.SetValue("B1", "1、任务1");
			tcWorksheet.SetCellStyle("B1", "微软雅黑", 13, TextAlignment.Left, 20, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetValue("B2", "1.1、任务1.1", "B1");
			tcWorksheet.SetCellStyle("B2", "微软雅黑", 13, TextAlignment.Left, 40, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetValue("B3", "1.1.1、任务1.1.1", "B2");
			tcWorksheet.SetCellStyle("B3", "微软雅黑", 13, TextAlignment.Left, 60, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B4", "1.1.2、任务1.1.2", "B2");
			tcWorksheet.SetCellStyle("B4", "微软雅黑", 13, TextAlignment.Left, 60, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B5", "1.2、任务1.2", "B1");
			tcWorksheet.SetCellStyle("B5", "微软雅黑", 13, TextAlignment.Left, 40, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetValue("B6", "1.2.1、任务1.2.1", "B5");
			tcWorksheet.SetCellStyle("B6", "微软雅黑", 13, TextAlignment.Left, 60, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B7", "1.2.2、任务1.2.2", "B5");
			tcWorksheet.SetCellStyle("B7", "微软雅黑", 13, TextAlignment.Left, 60, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B8", "1.2.3、任务1.2.3", "B5");
			tcWorksheet.SetCellStyle("B8", "微软雅黑", 13, TextAlignment.Left, 60, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B9", "1.3、任务1.3", "B1");
			tcWorksheet.SetCellStyle("B9", "微软雅黑", 13, TextAlignment.Left, 40, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B10", "2、任务2");
			tcWorksheet.SetCellStyle("B10", "微软雅黑", 13, TextAlignment.Left, 20, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetValue("B11", "2.1、任务2.1", "B10");
			tcWorksheet.SetCellStyle("B11", "微软雅黑", 13, TextAlignment.Left, 40, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			tcWorksheet.SetValue("B12", "2.2、任务2.2", "B10");
			tcWorksheet.SetCellStyle("B12", "微软雅黑", 13, TextAlignment.Left, 40, new SolidColorBrush(Colors.Black)/*, FontWeights.Normal*/);
			#endregion

			#region 开始日期
			tcWorksheet.SetColumnTitle(2, "开始日期", TextAlignment.Right, FontWeights.Bold);
			tcWorksheet.SetValue("C1", DateString(2024, 8, 20));
			tcWorksheet.SetValue("C2", DateString(2024, 8, 20));
			tcWorksheet.SetValue("C3", DateString(2024, 8, 20));
			tcWorksheet.SetValue("C4", DateString(2024, 8, 24));
			tcWorksheet.SetValue("C5", DateString(2024, 8, 31));
			tcWorksheet.SetValue("C6", DateString(2024, 8, 31));
			tcWorksheet.SetValue("C7", DateString(2024, 9, 3));
			tcWorksheet.SetValue("C8", DateString(2024, 9, 11));
			tcWorksheet.SetValue("C9", DateString(2024, 9, 16));
			tcWorksheet.SetValue("C10", DateString(2024, 9, 22));
			tcWorksheet.SetValue("C11", DateString(2024, 9, 22));
			tcWorksheet.SetValue("C12", DateString(2024, 10, 4));
			tcWorksheet.SetCellStyle("C1", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("C2", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("C3", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C4", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C5", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("C6", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C7", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C8", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C9", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C10", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("C11", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("C12", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			#endregion

			#region 结束日期
			tcWorksheet.SetColumnTitle(3, "结束日期", TextAlignment.Right, FontWeights.Bold);
			tcWorksheet.SetValue("D1", DateString(2024, 9, 21));
			tcWorksheet.SetValue("D2", DateString(2024, 8, 30));
			tcWorksheet.SetValue("D3", DateString(2024, 8, 23));
			tcWorksheet.SetValue("D4", DateString(2024, 8, 30));
			tcWorksheet.SetValue("D5", DateString(2024, 9, 15));
			tcWorksheet.SetValue("D6", DateString(2024, 9, 2));
			tcWorksheet.SetValue("D7", DateString(2024, 9, 10));
			tcWorksheet.SetValue("D8", DateString(2024, 9, 15));
			tcWorksheet.SetValue("D9", DateString(2024, 9, 21));
			tcWorksheet.SetValue("D10", DateString(2024, 10, 3));
			tcWorksheet.SetValue("D11", DateString(2024, 10, 8));
			tcWorksheet.SetValue("D12", DateString(2024, 10, 6));
			tcWorksheet.SetCellStyle("D1", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("D2", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("D3", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D4", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D5", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("D6", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D7", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D8", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D9", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D10", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("D11", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("D12", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			#endregion

			#region 时长
			tcWorksheet.SetColumnTitle(4, "时长", TextAlignment.Right, FontWeights.Bold);
			tcWorksheet.SetValue("E1", "33");
			tcWorksheet.SetValue("E2", "11");
			tcWorksheet.SetValue("E3", "4");
			tcWorksheet.SetValue("E4", "7");
			tcWorksheet.SetValue("E5", "16");
			tcWorksheet.SetValue("E6", "3");
			tcWorksheet.SetValue("E7", "8");
			tcWorksheet.SetValue("E8", "5");
			tcWorksheet.SetValue("E9", "6");
			tcWorksheet.SetValue("E10", "12");
			tcWorksheet.SetValue("E11", "17");
			tcWorksheet.SetValue("E12", "3");
			tcWorksheet.SetCellStyle("E1", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("E2", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("E3", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E4", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E5", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("E6", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E7", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E8", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E9", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E10", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black), FontWeights.Bold);
			tcWorksheet.SetCellStyle("E11", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("E12", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			#endregion

			#region 完成率
			tcWorksheet.SetColumnTitle(5, "完成率", TextAlignment.Right, FontWeights.Bold);
			tcWorksheet.SetValue("F3", "100%");
			tcWorksheet.SetValue("F4", "80%");
			tcWorksheet.SetValue("F6", "60%");
			tcWorksheet.SetValue("F7", "50%");
			tcWorksheet.SetValue("F8", "25%");
			tcWorksheet.SetValue("F9", "10%");
			tcWorksheet.SetValue("F11", "0%");
			tcWorksheet.SetValue("F12", "0%");

			tcWorksheet.SetCellStyle("F3", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("F4", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));

			tcWorksheet.SetCellStyle("F6", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("F7", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("F8", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("F9", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));

			tcWorksheet.SetCellStyle("F11", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("F12", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			#endregion

			#region 前置任务
			tcWorksheet.SetColumnTitle(6, "前置任务", TextAlignment.Right, FontWeights.Bold);
			tcWorksheet.SetValue("G4", "3FS");
			tcWorksheet.SetValue("G6", "4FS");
			tcWorksheet.SetValue("G7", "6FS");
			tcWorksheet.SetValue("G8", "7FS");
			tcWorksheet.SetValue("G9", "8FS");
			tcWorksheet.SetValue("G11", "9FS");
			tcWorksheet.SetValue("G12", "11SF");

			tcWorksheet.SetCellStyle("G4", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));

			tcWorksheet.SetCellStyle("G6", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("G7", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("G8", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("G9", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));

			tcWorksheet.SetCellStyle("G11", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));
			tcWorksheet.SetCellStyle("G12", "微软雅黑", 13, TextAlignment.Right, 0, new SolidColorBrush(Colors.Black));

			#endregion

			#region 负责人
			tcWorksheet.SetColumnTitle(7, "负责人", TextAlignment.Center, FontWeights.Bold);
			#endregion

			tcWorksheet.SetColumnWidth(0, 30);
			tcWorksheet.SetColumnWidth(1, 200);
			tcWorksheet.SetColumnWidth(4, 60);
			tcWorksheet.SetColumnWidth(5, 60);
			tcWorksheet.SetColumnWidth(6, 80);

			tcWorksheet.SetColumnHeaderHeight(40);

			tcWorksheet.SelectionChanged += TcWorksheet_SelectionChanged;
			tcWorksheet.Changed += TcWorksheet_Changed;
			#endregion

			#region 甘特图部分
			//设置刻度模式
			tcWorksheet.ScaleMode = 0;
			//设置日期范围
			tcWorksheet.StartDate = new DateTime(2024, 1, 1);
			tcWorksheet.EndDate = new DateTime(2024, 12, 31);

			//给任务集合赋值
			if (tcWorksheet.TaskList == null)
			{
				tcWorksheet.TaskList = new List<GanttTask>();
			}
			else
			{
				tcWorksheet.TaskList.Clear();
			}

			for (int i = 0; i < tcWorksheet.Rows.Count; i++)
			{
				Cell cell;
				cell = tcWorksheet.GetOrCreateCell(i, 1);
				if (cell.Data != null)
				{
					GanttTask task = new GanttTask();

					task.Index = i;
					task.TaskName = cell.Data;
					if (tcWorksheet.Rows[i].Children.Count > 0)
					{
						task.HasChildren = true;
					}
					cell = tcWorksheet.GetOrCreateCell(i, 2);
					task.StartDate = tcWorksheet.String2Date(cell.Data);
					cell = tcWorksheet.GetOrCreateCell(i, 3);
					task.EndDate = tcWorksheet.String2Date(cell.Data);
					cell = tcWorksheet.GetOrCreateCell(i, 4);
					task.Duration = Convert.ToDouble(cell.Data);
					cell = tcWorksheet.GetOrCreateCell(i, 5);
					task.Completion = tcWorksheet.String2Double(cell.Data);
					cell = tcWorksheet.GetOrCreateCell(i, 6);

					if (!string.IsNullOrEmpty(cell.Data))
					{
						string tempString = cell.Data.Replace(" ", "");
						var predecessors = tempString.Split(',')
										     .Select(p =>
										     {
											     string dependencyType;
											     int predecessorIndex;

											     if (p.Length > 2 && char.IsLetter(p[p.Length - 2]) && char.IsLetter(p[p.Length - 1]))
											     {
												     // 格式为 3FS, 4FF 等
												     dependencyType = p.Substring(p.Length - 2); // 获取最后两个字符作为依赖关系类型
												     predecessorIndex = int.Parse(p.Substring(0, p.Length - 2)) - 1; // 获取前面的部分作为前置任务索引
											     }
											     else
											     {
												     // 格式为 3 等，默认为 FS
												     dependencyType = "FS";
												     predecessorIndex = int.Parse(p);
											     }

											     return (PredecessorIndex: predecessorIndex, DependencyType: dependencyType);
										     })
										     .ToList();
						task.Predecessors = predecessors;
					}
					else
					{
						task.Predecessors = null;
					}

					tcWorksheet.TaskList.Add(task);
				}
			}

			if (tcWorksheet.TaskList.Count > 0) 
				tcWorksheet.FindCriticalPath();

			//控件初始化时，将甘特图滚动条滚动到项目最开始时间
			//只做一次设置，后面用户可以随意拖拽滚动条
			if (tcWorksheet.TaskList.Count > 0 /*& tcWorksheet.initialSetting4ScrollBar*/)
			{
				DateTime earliestStartDate = tcWorksheet.TaskList.Min(task => task.StartDate);

				// 计算滚动条的位置
				double offsetToEarliestDate = (earliestStartDate - tcWorksheet.RealStartDate).TotalDays * tcWorksheet.WidthRatio;

				// 设置滚动条的值
				tcWorksheet.InitialScrollValue = Math.Max(0, offsetToEarliestDate);
				//initialSetting4ScrollBar = false;
			}

			#endregion
		}

		private void TcWorksheet_Changed(object sender, CellEventArgs e)
		{
			int rowIndex = e.Cell.Row;
			int columnIndex = e.Cell.Column;

			if (rowIndex <= tcWorksheet.TaskList.Count - 1)
			{
				//修改现有任务数据
				switch (columnIndex)
				{
					case 0:
						break;
					case 1:
						tcWorksheet.TaskList[rowIndex].TaskName = e.Cell.Data;
						break;
					case 2:
						tcWorksheet.TaskList[rowIndex].StartDate = tcWorksheet.String2Date(e.Cell.Data);
						break;
					case 3:
						tcWorksheet.TaskList[rowIndex].EndDate = tcWorksheet.String2Date(e.Cell.Data);
						break;
					case 4:
						if (e.Cell.Data == "")
						{
							tcWorksheet.TaskList[rowIndex].Duration = 0;
						}
						else
						{
							tcWorksheet.TaskList[rowIndex].Duration = Convert.ToInt32(e.Cell.Data);
						}

						break;
					case 5:
						tcWorksheet.TaskList[rowIndex].Completion = tcWorksheet.String2Double(e.Cell.Data);
						break;
					case 6:
						if (!string.IsNullOrEmpty(e.Cell.Data))
						{
							string tempString = e.Cell.Data.Replace(" ", "");
							var predecessors = tempString.Split(',')
											     .Select(p =>
											     {
												     string dependencyType;
												     int predecessorIndex;

												     if (p.Length > 2 && char.IsLetter(p[p.Length - 2]) && char.IsLetter(p[p.Length - 1]))
												     {
													     // 格式为 3FS, 4FF 等
													     dependencyType = p.Substring(p.Length - 2); // 获取最后两个字符作为依赖关系类型
													     predecessorIndex = int.Parse(p.Substring(0, p.Length - 2)) - 1; // 获取前面的部分作为前置任务索引
												     }
												     else
												     {
													     // 格式为 3 等，默认为 FS
													     dependencyType = "FS";
													     predecessorIndex = int.Parse(p);
												     }

												     return (PredecessorIndex: predecessorIndex, DependencyType: dependencyType);
											     })
											     .ToList();
							tcWorksheet.TaskList[rowIndex].Predecessors = predecessors;
						}
						else
						{
							tcWorksheet.TaskList[rowIndex].Predecessors = null;
						}

						break;
					default:
						break;
				}

				tcWorksheet.FindCriticalPath();
			}
		}

		private void TcWorksheet_SelectionChanged(object sender, CellEventArgs e)
		{
			//MessageBox.Show(e.Cell.Address + " is selected.");
		}

		private string DateString(int year, int month, int day)
		{
			return new DateTime(year, month, day).ToString("yyyy-MM-dd");
		}
	}
}
