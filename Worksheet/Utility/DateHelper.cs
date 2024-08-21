using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Worksheet.Models;

namespace Worksheet.Utility
{
	public class DateHelper
	{
		/// <summary>
		/// 获取指定日期所在星期的第一天
		/// 这里返回的是第一天与参数日期的差异
		/// </summary>
		/// <param name="date">指定的日期，整型</param>
		/// <returns>本星期第一天与指定日期的差异</returns>
		public static int GetFirstDayOfWeek(DateTime date)
		{
			//DateTime dt = DateTime.FromOADate(date);
			// 获取当前日期是一周中的第几天
			DayOfWeek dayOfWeek = date.DayOfWeek;
			return -(int)dayOfWeek;
		}
		/// <summary>
		/// 获取Excel日期序列在该月的天
		/// </summary>
		/// <param name="date"></param>
		/// <returns></returns>
		public static int GetFirstDayOfMonth(DateTime date)
		{
			//DateTime dateTime = DateTime.FromOADate(date);
			return (int)date.Day;
		}
		/// <summary>
		/// 获取日期所在月的最后一天
		/// </summary>
		/// <param name="date"></param>
		/// <returns></returns>
		public static int GetLastDayOfMonth(DateTime date)
		{
			//DateTime dateTime = DateTime.FromOADate(date);
			DateTime lastDayOfMonth = new DateTime(date.Year, date.Month, 1).AddMonths(1).AddDays(-1);

			TimeSpan timeSpan = lastDayOfMonth - date;
			return (int)timeSpan.TotalDays;
		}
		/// <summary>
		/// 获取日期所在季度的最后一天
		/// </summary>
		/// <param name="date"></param>
		/// <returns></returns>
		public static int GetLastDayOfQuarter(int date)
		{
			DateTime currentDate = DateTime.FromOADate(date);

			// 计算所在季度的最后月份（3、6、9、12月）
			int endOfMonth = (currentDate.Month - 1) / 3 * 3 + 3;
			if (endOfMonth > 12)
			{
				endOfMonth -= 12;
			}

			// 获取所在季度的最后一天
			DateTime lastDayOfQuarter = new DateTime(currentDate.Year, endOfMonth, DateTime.DaysInMonth(currentDate.Year, endOfMonth));

			TimeSpan timeSpan = lastDayOfQuarter - currentDate;
			return (int)timeSpan.TotalDays;
		}
		/// <summary>
		/// DateTime转换为Excel日期序列
		/// </summary>
		/// <param name="date">日期</param>
		/// <returns>整型日期序列</returns>
		/// <exception cref="ArgumentException"></exception>
		public static int DateTimeToExcelSerial(DateTime date)
		{
			// Excel日期从1900年1月1日开始计数，但需要注意1900年在Excel中是特殊处理的（它认为1900年是闰年，实际上不是）
			if (date < new DateTime(1900, 1, 1))
			{
				throw new ArgumentException("日期不能早于1900年1月1日");
			}

			// 对于1900年1月1日至1900年3月1日之间的日期需要特殊处理，因为Excel中这部分日期计算有误
			if (date >= new DateTime(1900, 2, 29) && date < new DateTime(1900, 3, 1))
			{
				return date.Subtract(new DateTime(1899, 12, 30)).Days - 1;
			}
			else
			{
				return date.Subtract(new DateTime(1899, 12, 30)).Days;
			}
		}
		/// <summary>
		/// 返回集合
		/// 日期，小组天数
		/// </summary>
		/// <param name="startDate">开始日期</param>
		/// <param name="endDate">结束日期</param>
		/// <param name="scaleMode">模式</param>
		/// <returns></returns>
		public static ObservableCollection<BigScaleModel> GetDates(DateTime startDate, DateTime endDate, string scaleMode)
		{
			ObservableCollection<BigScaleModel> dates = new ObservableCollection<BigScaleModel>();
			//月份
			if (scaleMode == "M")
			{
				DateTime currentDate = startDate;
				while (currentDate < endDate)
				{
					BigScaleModel model = new BigScaleModel();
					model.ScaleDate = new DateTime(currentDate.Year, currentDate.Month, 1); // 每个月的第一天

					// 如果已到达或超过endDate，则跳出循环
					if (model.ScaleDate > endDate)
					{
						break;
					}

					if (currentDate == startDate) // 对于非月初的情况
					{
						model.ScaleDate = currentDate;
						// 计算从当前日期到该月月底的天数
						int remainingDays = DateTime.DaysInMonth(currentDate.Year, currentDate.Month) - (currentDate.Day - 1);
						model.DaysInScale = remainingDays;
					}
					else // 对于月初的情况，即每个月的第一天
					{
						model.DaysInScale = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);

						if (model.ScaleDate.AddMonths(1).AddDays(-1) > endDate) // 检查是否超过endDate
						{
							model.DaysInScale = endDate.Subtract(model.ScaleDate).Days; // 调整为只包含endDate之前的天数
						}
					}

					dates.Add(model);

					// 移动到下一个月的第一天
					currentDate = currentDate.AddMonths(1);

				}
				return dates;
			}
			//季度
			else
			{
				DateTime currentDate = startDate;

				while (currentDate < endDate)
				{
					// 确定季度开始和结束日期
					DateTime quarterStartDateOld = new DateTime(currentDate.Year, ((currentDate.Month - 1) / 3) * 3 + 1, 1);
					DateTime quarterEndDate = GetEndOfQuarter(currentDate);
					DateTime quarterStartDate = quarterEndDate.AddDays(1);

					if (quarterEndDate > endDate)
					{
						quarterEndDate = endDate;
					}

					// 创建新的模型并设置季度开始日期
					BigScaleModel model = new BigScaleModel();
					if (currentDate == startDate)
					{
						model.ScaleDate = currentDate;
					}
					else
					{
						model.ScaleDate = quarterStartDateOld;
					}

					// 计算剩余天数（考虑非季度第一天的情况）
					if (currentDate != quarterStartDate)
					{
						model.DaysInScale = quarterEndDate.Subtract(currentDate).Days + 1;
					}
					else
					{
						model.DaysInScale = quarterEndDate.Subtract(model.ScaleDate).Days + 1;
					}

					dates.Add(model);

					if (currentDate > endDate)
					{
						break;
					}

					currentDate = quarterEndDate.AddDays(1);
				}

				return dates;
			}
		}
		public static DateTime GetEndOfQuarter(DateTime date)
		{
			int quarterEndMonth = ((date.Month - 1) / 3 + 1) * 3;
			if (quarterEndMonth > 12)
			{
				quarterEndMonth -= 12;
				date = new DateTime(date.Year + 1, quarterEndMonth, 1);
			}
			else
			{
				date = new DateTime(date.Year, quarterEndMonth, 1);
			}
			return date.AddMonths(1).AddDays(-1);
		}
	}
}
