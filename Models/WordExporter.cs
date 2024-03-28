using System;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace FromExcelWord.Models
{
    public class WordExporter
    {

        public void WordExport(DataGrid grd1, DataGrid grd2)
        {
            if (grd1 == null || grd1.Items.Count <= 0)
            {
                MessageBox.Show("Данные для экспорта не обнаружены.");
                return;
            }
            var app = new Word.Application();
            app.Visible = true;
            Word.Document doc = app.Documents.Add();


            try
            {
                InsertDataWord(doc, grd1, grd2);

                // Сохраняем документ Word и закрываем приложение
                doc.Save();
                doc.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте: {ex.Message}");
            }
            finally
            {


            }
        }

        private void InsertDataWord(Word.Document doc, DataGrid grd1, DataGrid grd2)
        {

            // Создаем абзац заголовка
            Word.Paragraph paragraph = doc.Paragraphs.Add();
            paragraph.Range.Text = "Отчет по загрузке";
            object style_name = "Заголовок 1";
            paragraph.Range.set_Style(ref style_name);
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.InsertParagraphAfter();
            paragraph.Range.InsertParagraphAfter();


            // создаем таблицу в новый параграф
            Word.Paragraph paragraph1 = doc.Paragraphs.Add();
            var table = paragraph1.Range.Tables.Add(doc.Range(), grd1.Items.Count + grd2.Items.Count + 1, grd1.Columns.Count);




            table.Cell(1, 1).Range.Text = "Отдел";
            table.Cell(1, 2).Range.Text = "Количество задач";

            table.Application.Selection.Tables[1].Borders.Enable = 1; // включаем все границы
            //Стиль заголовка таблицы
            table.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
            table.Application.Selection.Tables[1].Rows[1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray375;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Color = Word.WdColor.wdColorWhite;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Calibri";
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 11;
            table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            // строки с количеством задач по отделам
            for (int i = 0; i < grd1.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)grd1.ItemContainerGenerator.ContainerFromIndex(i);

                for (int j = 0; j < grd1.Columns.Count; j++)
                {
                    if (grd1.Columns[j] != null)
                    {

                        TextBlock cellContent = grd1.Columns[j].GetCellContent(row) as TextBlock;
                        string cellValue = cellContent == null ? "" : cellContent.Text;
                        table.Cell(i + 2, j + 1).Range.Text = cellValue;


                    }

                }

            }

            // строки с количеством задач у сотрудников
            for (int i = 0; i < grd2.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)grd2.ItemContainerGenerator.ContainerFromIndex(i);
                for (int j = 0; j < grd2.Columns.Count; j++)
                {
                    if (grd2.Columns[j] != null)
                    {

                        TextBlock cellContent = grd1.Columns[j].GetCellContent(row) as TextBlock;
                        string cellValue = cellContent == null ? "" : cellContent.Text;
                        table.Cell(i + 2 + grd1.Items.Count, j + 1).Range.Text = cellValue;
                    }

                }

            }



            for (int j = 1; j < grd1.Items.Count + 1; j++)
            {
                table.Application.Selection.Tables[1].Rows[j + 1].Cells[2]
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                table.Application.Selection.Tables[1].Rows[j + 1].Range.Bold = 1;
                table.Application.Selection.Tables[1].Rows[j + 1].Shading.BackgroundPatternColor = Word.WdColor.wdColorGray10;

            }

            for (int j = grd1.Items.Count; j <= grd2.Items.Count + grd1.Items.Count; j++)
            {
                table.Application.Selection.Tables[1].Rows[j + 1].Cells[2]
                    .Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            }


            //table.Application.Selection.Tables[1].Rows[2].Cells[2].Range.Bold = 1; // ячейка 4 жирным


            table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            table.Application.Selection.Tables[1].Borders.Enable = 1; // включаем все границы

            //Стиль заголовка таблицы
            table.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Calibri";
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 11;
            table.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

        }
    }
}

