using System;
using System.Reflection;
using System.Resources;
using Microsoft.Office.Interop.Excel;

namespace ExcelService
{
	public class ExcelTimeSheetCreator
	{
		private Application _excelApplication;

		public ExcelTimeSheetCreator()
		{
			_excelApplication = new Application();
		}

		public void LoadDataIntoExcelTemplateAndSaveToLocation(TimeSheetData timeSheetData, string location)
		{
			var emptyWorksheet = GetTheMainWorkSheet();
			var loadedWorksheet = PutTimesheetDataIntoExcelSheet(emptyWorksheet, timeSheetData);
			loadedWorksheet.SaveAs(location);
		}

		private Worksheet PutTimesheetDataIntoExcelSheet(Worksheet worksheet, TimeSheetData timeSheetData)
		{
			worksheet = PutDayDataInToRow("2" , timeSheetData.Monday, worksheet);
			worksheet = PutDayDataInToRow("3", timeSheetData.Tuesday, worksheet);
			worksheet = PutDayDataInToRow("4", timeSheetData.Wednesday, worksheet);
			worksheet = PutDayDataInToRow("5", timeSheetData.Thursday, worksheet);
			worksheet = PutDayDataInToRow("6", timeSheetData.Friday, worksheet);
			return worksheet;
		}

		private Worksheet PutDayDataInToRow(string rowNumber, WorkDayData workDayData, Worksheet worksheet)
		{
			worksheet.Range["B"+rowNumber].Value = workDayData.DayStarted.TimeOfDay.ToString();
			worksheet.Range["C"+rowNumber].Value = workDayData.DayEnded.TimeOfDay.ToString();
			worksheet.Range["D"+rowNumber].Value = workDayData.BreakStarted.TimeOfDay.ToString();
			worksheet.Range["E"+rowNumber].Value = workDayData.BreakEnded.TimeOfDay.ToString();
			worksheet.Range["F"+rowNumber].Value = workDayData.BreakEnded.Subtract(workDayData.BreakStarted).Minutes.ToString();
			return worksheet;
		}

		private Worksheet GetTheMainWorkSheet()
		{
			var excelWorkbook = GetExcelTemplate();
			return excelWorkbook.Worksheets.Item[1] as Worksheet;
		}

		private Workbook GetExcelTemplate()
		{
			var excelWorkBook =
				_excelApplication.Workbooks.Open(
					@"D:\Users\Kamil\Mina dokument\Visual Studio 2010\Projects\TimeSheeter\resources\Tidrapport_120827_120902_KamilHakim.xlsx");
			return excelWorkBook;
		}
	}
}