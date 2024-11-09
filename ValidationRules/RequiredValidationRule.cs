using System.Globalization;
using System.Windows.Controls;

namespace TimeCollect.ValidationRules
{
    public class RequiredValidationRule : ValidationRule
    {
        public string ErrorMessage { get; set; }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            if (string.IsNullOrWhiteSpace(value?.ToString()))
            {
                return new ValidationResult(false, ErrorMessage);
            }
            return ValidationResult.ValidResult;
        }
    }
}
