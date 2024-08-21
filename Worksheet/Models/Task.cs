using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace Worksheet.Models
{
	public class GanttTask : INotifyPropertyChanged
	{
		#region 通知接口
		public event PropertyChangedEventHandler PropertyChanged;

		protected void OnPropertyChanged(string propertyName)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
		#endregion
		/// <summary>
		/// 对应的行号
		/// </summary>
		public int Index { get; set; }
		public bool HasChildren { get; set; } = false;
		public string TaskName { get; set; }
		public DateTime StartDate { get; set; }
		public DateTime EndDate { get; set; }
		public double Duration { get; set; }
		public double Completion { get; set; }
		public List<(int PredecessorIndex, string DependencyType)> Predecessors { get; set; } // 前置任务的索引
		public string Owner { get; set; }
		/// <summary>
		/// 是否在关键路径上
		/// </summary>
		public bool IsCritical { get; set; }

		public DateTime EarlyStart { get; set; }
		public DateTime EarlyFinish { get; set; }
		public DateTime LateStart { get; set; }
		public DateTime LateFinish { get; set; }
	}
}
