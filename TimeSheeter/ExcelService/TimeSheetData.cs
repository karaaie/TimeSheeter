using System;

namespace ExcelService
{
	public class TimeSheetData
	{
		public WorkDayData Monday { get; set; }
		public WorkDayData Tuesday { get; set; }
		public WorkDayData Wednesday { get; set; }
		public WorkDayData Thursday { get; set; }
		public WorkDayData Friday { get; set; }
		public WorkDayData Saturday { get; set; }
		public WorkDayData Sunday { get; set; }
	}
}