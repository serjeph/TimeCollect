﻿using System.ComponentModel.DataAnnotations;

namespace TimeCollect.Models
{
    public class Employee
    {
        public int EmployeeId { get; set; }
        public string Name { get; set; }
        public string Nickname { get; set; }
        public string Team { get; set; }

        [Required(ErrorMessage = "Spreadsheet ID is required.")]
        public string SpreadsheetId { get; set; }

    }
}
