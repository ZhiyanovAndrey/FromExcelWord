using FromExcelWord.Models;
//using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Linq;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace FromExcelWord
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
      
        private string _path = string.Empty;

        private WordExporter _wordExporter;


        public MainWindow()
        {
            _wordExporter = new WordExporter();
            InitializeComponent();
        }


     
        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                var query = from p in OpenExcelFile.GetPerson(_path)
                            join d in OpenExcelFile.GetDepartment(_path) on p.Department equals d.DepartmentId
                            join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                            select new { Отдел = d.Name, ФИО = $"{p.SurName} {p.FirstName}", TaskName = t.TaskId };

                // количество задач у сотрудников по отделам
                var query1 = query.GroupBy(q => new { q.Отдел, q.ФИО }).Select(g => new { g.Key.Отдел, g.Key.ФИО, Count = g.Count() });
                var query2 = query1.GroupBy(q => new { q.Отдел }).Select(g => new { g.Key.Отдел, Count = g.Sum(x => x.Count) }).OrderByDescending(d => d.Count);

                datagrid1.ItemsSource = query2;


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }



            try
            {
                var query = from p in OpenExcelFile.GetPerson(_path)
                            join t in OpenExcelFile.GetTask(_path) on p.PersonNumber equals t.PersonNumber
                            select new
                            {
                                Name = $"{p.SurName.Trim()} {p.FirstName.Trim().First()}. {p.MiddleName.FirstOrDefault()}.",
                                TaskName = t.TaskId
                            };

                // количество задач у сотрудников
                datagrid2.ItemsSource = query.GroupBy(p => p.Name).Select(g => new { Name = g.Key, Count = g.Count() }).OrderByDescending(d => d.Count);

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }



        private void Button_Click_4(object sender, RoutedEventArgs e)
        {

            _wordExporter.WordExport(datagrid1, datagrid2);

        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog()
                {
                    CheckFileExists = false,
                    CheckPathExists = true,
                    Multiselect = false,
                    Title = "Выберите файл"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    _path = openFileDialog.FileName;

                };
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }
    }
}
