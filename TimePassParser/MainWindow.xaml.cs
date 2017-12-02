using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace TimePassParser {
	/// <summary>
	/// Логика взаимодействия для MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window {
		public MainWindow() {
			InitializeComponent();
		}

		private void ButtonSelect_Click(object sender, RoutedEventArgs e) {
			OpenFileDialog openFileDialog = new OpenFileDialog();
			openFileDialog.Filter = "Книга Excel (*.xls*)|*.xls*";
			openFileDialog.CheckFileExists = true;
			openFileDialog.CheckPathExists = true;
			openFileDialog.Multiselect = false;
			openFileDialog.RestoreDirectory = true;

			if (openFileDialog.ShowDialog() == true)
				textBoxSelected.Text = openFileDialog.FileName;
		}

		private void ButtonParse_Click(object sender, RoutedEventArgs e) {
			if (string.IsNullOrEmpty(textBoxSelected.Text)) {
				MessageBox.Show(this, "Для анализа необходимо выбрать файл", "Не выбран файл", 
					MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}


			BackgroundWorker backgroundWorker = new BackgroundWorker();
			backgroundWorker.WorkerReportsProgress = true;
			backgroundWorker.ProgressChanged += Worker_ProgressChanged;
			backgroundWorker.DoWork += Worker_DoWork;
			backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
			backgroundWorker.RunWorkerAsync(textBoxSelected.Text);
		}

		private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			if (e.Error != null) {
				MessageBox.Show(this, e.Error.Message + Environment.NewLine + e.Error.StackTrace, "Завершено с ошибкой",
					MessageBoxButton.OK, MessageBoxImage.Error);
				return;
			}

			MessageBox.Show(this, "Завершено успешно", "", MessageBoxButton.OK, MessageBoxImage.Information);
		}

		private void Worker_DoWork(object sender, DoWorkEventArgs e) {
			Parser.Run(e.Argument as string, sender as BackgroundWorker);
		}

		private void Worker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			progressBarResult.Value = e.ProgressPercentage;

			if (!string.IsNullOrEmpty(e.UserState.ToString()))
				textBoxResult.Text = DateTime.Now.ToLongTimeString() + ": " +
					e.UserState.ToString() + Environment.NewLine + textBoxResult.Text;
		}
	}
}
