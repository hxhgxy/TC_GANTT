using System;
using System.ComponentModel;

namespace Worksheet.Models
{
	public class BigScaleModel : INotifyPropertyChanged
	{
		#region INotifyPropertyChanged接口
		public event PropertyChangedEventHandler PropertyChanged;
		private void OnPropertyChanged(string propertyName)
		{
			var handler = PropertyChanged;
			if (handler != null)
			{
				handler(this, new PropertyChangedEventArgs(propertyName));

			}
		}
		#endregion

		private DateTime _scaleDate;

		public DateTime ScaleDate
		{
			get { return _scaleDate; }
			set
			{
				_scaleDate = value;
				OnPropertyChanged(nameof(ScaleDate));
			}
		}

		private int _daysInScale = 1;

		public int DaysInScale
		{
			get { return _daysInScale; }
			set
			{
				_daysInScale = value;
				OnPropertyChanged(nameof(DaysInScale));
			}
		}

	}
}
