﻿using Microsoft.Office.Interop.Word;
using System;
using System.Windows;
using System.Windows.Controls;

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
            var app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            var doc = app.Documents.Add();



            try
            {
                InsertDataWord(doc, grd1, grd2);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте: {ex.Message}");
            }
            finally
            {

                // Сохраняем документ Word и закрываем приложение
                doc.Save();
                doc.Close();
                app.Quit();
            }
        }

        private void InsertDataWord(Document doc, DataGrid grd1, DataGrid grd2)
        {
            // создаем таблицу
            var table = doc.Tables.Add(doc.Range(), grd1.Items.Count + grd2.Items.Count + 1, grd1.Columns.Count);
            //Формат таблицы

            //table.Application.Selection.Tables[1].Rows[1].Select();
            //table.Application.Selection.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            //table.Application.Selection.Tables[1].Select();
            //table.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
            //table.Application.Selection.Tables[1].Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            //table.Application.Selection.Tables[1].Rows[1].Select();
            //table.Application.Selection.InsertRowsAbove(1);
            //table.Application.Selection.Tables[1].Rows[1].Select();



            // таблица с количеством задач по отделам
            for (int j = 0; j < grd1.Columns.Count; j++)
            {

                table.Cell(1, 1).Range.Text = "Отдел";
                table.Cell(1, 2).Range.Text = "Количество задач";
            }
            table.Application.Selection.Tables[1].Borders.Enable = 1; // включаем все границы
            //Стиль заголовка таблицы
            table.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Calibri";
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 11;
            table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;


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

            // таблица с количеством задач у сотрудников


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

           

            for (int j = 0; j < grd1.Columns.Count; j++)
            {
                table.Application.Selection.Tables[1].Rows[j+1].Cells[2].Range.Bold = 1;
            }




            //table.Application.Selection.Tables[1].Rows[2].Cells[2].Range.Bold = 1; // ячейка 4 жирным


            table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            table.Application.Selection.Tables[1].Borders.Enable = 1; // включаем все границы

            //Стиль заголовка таблицы
            table.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Calibri";
            table.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 11;
            table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //table.Range.ParagraphFormat.SpaceAfter = 6; // отступ

            //table.Application.Selection.Tables[1].Rows[1].Alignment= WdRowAlignment.wdAlignRowCenter;

            //table.Rows[1].Range.Font.Bold = 1;
            //table.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //table.Borders.Enable = 1;
        }


        //Workbook wb = new Workbook(path);

        //    foreach (Worksheet worksheet in wb.Worksheets)
        //    {
        //        MessageBox.Show(worksheet.Name);
        //    }


    }
}

