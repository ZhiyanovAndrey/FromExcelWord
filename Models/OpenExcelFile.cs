using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FromExcelWord.Models
{
    public class OpenExcelFile
    {
      
   

        public static IEnumerable<Person> GetPerson(string path)
        {
            Workbook wb = new Workbook(path);
            // Получить рабочий лист 1
            using (Worksheet worksheet = wb.Worksheets[1])
            {
                // Получить количество строк и столбцов
                int rows = worksheet.Cells.MaxDataRow;

                // Цикл по строкам
                for (int i = 1; i <= rows; i++)
                {
                    var person = new Person
                    {
                        PersonNumber = worksheet.Cells[i, 0].StringValue,
                        SurName = worksheet.Cells[i, 1].StringValue,
                        FirstName = worksheet.Cells[i, 2].StringValue,
                        MiddleName = worksheet.Cells[i, 3].StringValue,
                        Birthday = worksheet.Cells[i, 4].DateTimeValue,
                        Department = worksheet.Cells[i, 5].IntValue,

                    };
                    // И возвращаем его
                    yield return person;
                }
            };
        }


        public static IEnumerable<Department> GetDepartment(string path)
        {
            Workbook wb = new Workbook(path);
            // Получить рабочий лист 2
            using (Worksheet worksheet = wb.Worksheets[2])
            {
                int rows = worksheet.Cells.MaxDataRow;
                for (int i = 1; i <= rows; i++)
                {
                    var department = new Department
                    {
                        DepartmentId = worksheet.Cells[i, 0].IntValue,
                        Name = worksheet.Cells[i, 1].StringValue,
                    };

                    yield return department;
                }

            };
        }

        public static IEnumerable<PersonTask> GetTask(string path)
        {
            Workbook wb = new Workbook(path);
            // Получить рабочий лист 3
            using (Worksheet worksheet = wb.Worksheets[7])
            {
                int rows = worksheet.Cells.MaxDataRow;
                for (int i = 1; i <= rows; i++)
                {
                    var tasks = new PersonTask
                    {
                        TaskId = worksheet.Cells[i, 0].StringValue,
                        PersonNumber = worksheet.Cells[i, 1].StringValue,
                    };

                    yield return tasks;
                }

            };
        }
    }
}
