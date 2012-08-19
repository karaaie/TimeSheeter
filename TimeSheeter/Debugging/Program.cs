using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelService;

namespace Debugging
{
	class Program
	{
		static void Main(string[] args)
		{
			var timeSheetData = new TimeSheetData
			                    	{
			                    		Monday = new WorkDayData
			                    		         	{
			                    		         		DayStarted = new DateTime(1, 1, 1, 7, 50, 0),
			                    		         		DayEnded = new DateTime(1, 1, 1, 17, 5, 0),
			                    		         		BreakStarted = new DateTime(1, 1, 1, 11, 30, 0),
			                    		         		BreakEnded = new DateTime(1, 1, 1, 12, 30, 0)
			                    		         	},
			                    		Tuesday = new WorkDayData
			                    		          	{
			                    		          		DayStarted = new DateTime(1, 1, 1, 7, 51, 0),
			                    		          		DayEnded = new DateTime(1, 1, 1, 17, 6, 0),
			                    		          		BreakStarted = new DateTime(1, 1, 1, 12, 30, 0),
			                    		          		BreakEnded = new DateTime(1, 1, 1, 13, 30, 0)
			                    		          	},
										Wednesday = new WorkDayData
										{
											DayStarted = new DateTime(1, 1, 1, 7, 51, 0),
											DayEnded = new DateTime(1, 1, 1, 17, 6, 0),
											BreakStarted = new DateTime(1, 1, 1, 12, 30, 0),
											BreakEnded = new DateTime(1, 1, 1, 13, 30, 0)
										},
										Thursday = new WorkDayData
										{
											DayStarted = new DateTime(1, 1, 1, 7, 51, 0),
											DayEnded = new DateTime(1, 1, 1, 17, 6, 0),
											BreakStarted = new DateTime(1, 1, 1, 12, 30, 0),
											BreakEnded = new DateTime(1, 1, 1, 13, 30, 0)
										},
										Friday = new WorkDayData
										{
											DayStarted = new DateTime(1, 1, 1, 7, 51, 0),
											DayEnded = new DateTime(1, 1, 1, 17, 6, 0),
											BreakStarted = new DateTime(1, 1, 1, 12, 30, 0),
											BreakEnded = new DateTime(1, 1, 1, 13, 30, 0)
										}
			                    	};


			ExcelTimeSheetCreator creator = new ExcelTimeSheetCreator();
			creator.LoadDataIntoExcelTemplateAndSaveToLocation(timeSheetData, @"D:\Users\Kamil\Mina dokument\Visual Studio 2010\Projects\TimeSheeter\resources\test1.xlsx");
		}
	}
}
