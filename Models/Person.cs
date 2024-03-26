using System;

namespace FromExcelWord.Models
{
    public class Person
    {
        // Табельный номер	Фамилия	Имя 	Отчество	Дата рождения	Отдел

        public string PersonNumber { get; set; }
        public string SurName { get; set; }
        public string FirstName { get; set; }
        public string MiddleName { get; set; }

        public DateTime Birthday { get; set; }
        public int Department { get; set; }

        public Person() { }


    }
}
