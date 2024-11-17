using System.Globalization;
using System.Windows.Controls;

namespace TimeCollect.ValidationRules
{
    /// <summary>
    /// A validation rule that checks if a value is provided (not null or whitespace)
    /// </summary>
    public class RequiredValidationRule : ValidationRule
    {
        /// <summary>
        /// Gets or sets the error message to display if the validation fails.
        /// </summary>
        public string ErrorMessage { get; set; }


        /// <summary>
        /// Validates the given value.
        /// </summary>
        /// <param name="value">The value to validate.</param>
        /// <param name="cultureInfo">The culture to use for validation.</param>
        /// <returns>A ValidationResult indicationg whether the value is valid.</returns>
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
