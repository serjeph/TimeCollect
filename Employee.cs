using System.ComponentModel.DataAnnotations;

namespace TimeCollect
{
    public class Employee
    {
        public string Name { get; set; }
        public string Nickname { get; set; }

        [Required(ErrorMessage = "Spreadsheet ID is required.")]
        public string SpreadsheetId { get; set; }
    }
}
