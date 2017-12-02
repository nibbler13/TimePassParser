using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TimePassParser {
	class Parser {
		public static void Run(string fileName, BackgroundWorker backgroundWorker) {
			double progressCurrent = 0;
			DataTable dataTable = Excel.ReadExcelFile(fileName, String.Empty, backgroundWorker, (int)progressCurrent);

			int totalRows = dataTable.Rows.Count;
			if (totalRows == 0) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Не удалось считать книгу Excel, либо файл пуст");
				return;
			}

			progressCurrent += 10;
			backgroundWorker.ReportProgress((int)progressCurrent, "Считано строк: " + dataTable.Rows.Count);

			double progressStep = 80.0d / (double)totalRows;
			Dictionary<string, Employee> employees = new Dictionary<string, Employee>();
			for (int rowEmployee = 1; rowEmployee < totalRows; rowEmployee++) {
				//Console.WriteLine("rowEmployee: " + rowEmployee);
				try {
					string nameEmployee = dataTable.Rows[rowEmployee][0].ToString();
					Employee employee = new Employee(nameEmployee);

					for (int rowDate = rowEmployee; rowDate < totalRows; rowDate++) {
						//Console.WriteLine("rowDate: " + rowDate);
						string nameDate = dataTable.Rows[rowDate][0].ToString();

						if (!nameDate.Equals(nameEmployee)) {
							rowEmployee = rowDate - 1;
							break;
						}

						string dateRow = dataTable.Rows[rowDate][1].ToString();
						Employee.DayInfo dayInfo = new Employee.DayInfo(dateRow);

						string previousTimePass = string.Empty;
						string previousMarkEnter = string.Empty;
						
						for (int rowPass = rowDate; rowPass < totalRows; rowPass++) {
							//Console.WriteLine("rowPass: " + rowPass);
							string namePass = dataTable.Rows[rowPass][0].ToString();
							string datePass = dataTable.Rows[rowPass][1].ToString();

							if (!datePass.Equals(dateRow) ||
								!namePass.Equals(nameEmployee)) {

								if (!string.IsNullOrEmpty(previousTimePass) &&
									!string.IsNullOrEmpty(previousMarkEnter)) {
									///
									/// todo!!!!
									///
								}

								rowDate = rowPass - 1;
								break;
							}

							string timePass = dataTable.Rows[rowPass][2].ToString().ToLower().Trim(' ');
							string markEnter = dataTable.Rows[rowPass][3].ToString().ToLower().Trim(' ');

							dayInfo.Passes++;

							if (string.IsNullOrEmpty(dayInfo.FirstEnterTime) &&
								markEnter.Equals("да"))
								dayInfo.FirstEnterTime = timePass;
							if (markEnter.Equals("нет"))
								dayInfo.LastExitTime = timePass;

							if (string.IsNullOrEmpty(previousTimePass) &&
								string.IsNullOrEmpty(previousMarkEnter)) {
								previousTimePass = timePass;
								previousMarkEnter = markEnter;
								continue;
							}
							
							if (previousMarkEnter.Equals("да")) {
								if (markEnter.Equals("нет")) {
									TimeSpan timeEnter;
									TimeSpan.TryParse(previousTimePass, null, out timeEnter);
									TimeSpan timeExit;
									TimeSpan.TryParse(timePass, null, out timeExit);
									dayInfo.WorkingTime += timeExit - timeEnter;
									//Console.WriteLine("rowPass: " + rowPass);
								} else {

								}
							} else {
								if (markEnter.Equals("да")) {

								} else {
								
								}
							}

							previousTimePass = string.Empty;
							previousMarkEnter = string.Empty;

							if (rowPass == totalRows - 1) {
								rowDate = rowPass;
								break;
							}
						}

						//backgroundWorker.ReportProgress((int)(progressCurrent + progressStep * rowDate), "Строка: " + rowDate);
						if (!employee.Days.ContainsKey(dateRow))
							employee.Days.Add(dateRow, dayInfo);

						if (rowDate == totalRows - 1) {
							rowEmployee = rowDate;
							break;
						}
					}

					backgroundWorker.ReportProgress((int)(progressCurrent + progressStep * rowEmployee), "Сотрудник: " + nameEmployee);
					if (!employees.ContainsKey(nameEmployee))
						employees.Add(nameEmployee, employee);
				} catch (Exception e) {
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				}
			}

			progressCurrent += 80.0d;
			backgroundWorker.ReportProgress((int)progressCurrent, "Сотрудников в списке: " + employees.Count);

			string resultFile = Excel.WriteToExcel(employees, backgroundWorker, progressCurrent, 100);
			
			if (string.IsNullOrEmpty(resultFile)) {
				backgroundWorker.ReportProgress((int)progressCurrent, "Не удалось записать данные в файл");
			} else {
				Process.Start(resultFile);
				backgroundWorker.ReportProgress((int)progressCurrent, "Результат анализа сохранен в файл: " + resultFile);
			}
		}
	}
}
