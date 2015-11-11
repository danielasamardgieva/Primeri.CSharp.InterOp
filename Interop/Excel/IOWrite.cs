using System;
using InteropExcel=Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
	public class IOWrite
	{
		private Datastruct _data;
		private InteropExcel.Application excel;

		public IOWrite (Datastruct data)
		{
		}
		public bool exportTable()
		{
			try
			{
				//Подготовка
				excel= new InteropExcel.ApplicationClass ();
				if (excel==null) return false;
				excel.Visible=false;

				InteropExcel.Workbook workbook=excel.Workbooks.Add();
				if(workbook==null) return false;

				InteropExcel.Worksheet sheet=(InteropExcel.worksheet) workbook.worksheets[1]
				sheet.Name="Таблица1"


				//Попълване на таблицата

				//Запаметяване и затваряне
				workbook.SaveCopyAs(getPath());
				excel.DisplayAlerts=false; //изключване на всички съобщения на Excel



				workbook.Close();
				excel.Quit ();
					return true;
			}catch{
			}
				return false;
		}
		public void addRow(DataRow _row)
		{
			try {


			} catch {
			}
		}	
		public void runFile()
		{
			try {
				System.Diagnostics.Process.Start (getPath ());
			} catch {
			}
		}
		private string getPath ()
		{
			return System.IO.Path.Combine (AppDomain.CurrentDomain.BaseDirectory, "Table.xls");
		}
	}
}

