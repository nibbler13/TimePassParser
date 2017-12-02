using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimePassParser {
	class Excel {
		public static string WriteToExcel(Dictionary<string, Employee> employees, BackgroundWorker backgroundWorker,
			double progressFrom, double progressTo) {
			double progressCurrent = progressFrom;

			string templateFile = Environment.CurrentDirectory + "\\Template.xlsx";
			string resultFilePrefix = "Result_";

			backgroundWorker.ReportProgress((int)progressCurrent, "Выгрузка данных в Excel");

			if (!File.Exists(templateFile)) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Не удалось найти файл шаблона: " + templateFile);
				return string.Empty;
			}

			string resultFile = Environment.CurrentDirectory + "\\" + resultFilePrefix + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
			try {
				File.Copy(templateFile, resultFile);
			} catch (Exception e) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Не удалось скопировать файл шаблона в новый файл: " + resultFile + ", " + e.Message);
				return string.Empty;
			}

			int totalRows = 0;
			foreach (KeyValuePair<string, Employee> employee in employees)
				totalRows += employee.Value.Days.Count;

			double progressStep = (progressTo - progressFrom) / totalRows;

			IWorkbook workbook;
			using (FileStream stream = new FileStream(resultFile, FileMode.Open, FileAccess.Read)) {
				workbook = new XSSFWorkbook(stream);
				stream.Close();
			}

			int rowNumber = 1;
			int columnNumber = 0;

			ISheet sheet = workbook.GetSheet("Data");
			ICreationHelper creationHelper = workbook.GetCreationHelper();

			List<string> notFoundMkbCodes = new List<string>();

			foreach (KeyValuePair<string, Employee> employee in employees) {
				foreach (KeyValuePair<string, Employee.DayInfo> dayInfoPair in employee.Value.Days) {
					backgroundWorker.ReportProgress((int)progressCurrent, "");
					progressCurrent += progressStep;

					IRow row = sheet.CreateRow(rowNumber);

					string[] data = new string[] {
						employee.Key,
						dayInfoPair.Key,
						dayInfoPair.Value.Passes.ToString(),
						dayInfoPair.Value.WorkingTime.ToString(),
						dayInfoPair.Value.FirstEnterTime,
						dayInfoPair.Value.LastExitTime
					};

					foreach (string value in data) {
						ICell cell = row.CreateCell(columnNumber);
						cell.SetCellValue(creationHelper.CreateRichTextString(value));
						columnNumber++;
					}

					columnNumber = 0;
					rowNumber++;
				}
			}
			
			using (FileStream stream = new FileStream(resultFile, FileMode.Open, FileAccess.Write)) {
				workbook.Write(stream);
				stream.Close();
			}

			workbook.Close();
			return resultFile;
		}


		public static DataTable ReadExcelFile(string fileFullPath, string sheetName, 
			BackgroundWorker backgroundWorker, int currentProgress) {
			DataTable dataTable = new DataTable();
			backgroundWorker.ReportProgress((int)currentProgress, "Считывание файла: " + fileFullPath);

			if (!File.Exists(fileFullPath)) {
				backgroundWorker.ReportProgress(currentProgress, "Не удается найти файл: " + fileFullPath);
				return dataTable;
			}

			try {
				using (OleDbConnection conn = new OleDbConnection()) {
					conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileFullPath + ";" +
						"Extended Properties='Excel 12.0 Xml;HDR=NO;'";

					using (OleDbCommand comm = new OleDbCommand()) {
						if (string.IsNullOrEmpty(sheetName)) {
							conn.Open();
							DataTable dtSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
							sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
							conn.Close();
						} else
							sheetName += "$";

						comm.CommandText = "Select * from [" + sheetName + "]";
						comm.Connection = conn;

						using (OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter()) {
							oleDbDataAdapter.SelectCommand = comm;
							oleDbDataAdapter.Fill(dataTable);
						}
					}
				}
			} catch (Exception e) {
				backgroundWorker.ReportProgress(0, e.Message);
			}

			return dataTable;
		}
	}
}
