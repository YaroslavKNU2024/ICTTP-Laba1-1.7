using System.ComponentModel.DataAnnotations;

namespace LabaOne.ViewModel
{
    public class RegisterViewModel
    {
        [Required(ErrorMessage = "Дане поле є обов'язковим")]
        [Display(Name = "Email")]
        [DataType(DataType.EmailAddress, ErrorMessage = "E-mail невірний.")]
        [RegularExpression("^[a-zA-Z0-9_\\.-]+@([a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,6}$", ErrorMessage = "E-mail невірний.")]
        public string Email { get; set; }

        [Required(ErrorMessage = "Дане поле є обов'язковим")]
        [Range(1900, 2022, ErrorMessage = "Введіть дійсний рік народження.")]
        [Display(Name = "Рік народження")]
        public int Year { get; set; }

        [Required(ErrorMessage = "Дане поле є обов'язковим")]
        [Display(Name = "Пароль")]
        [StringLength(100, ErrorMessage = "Поле {0} повинне мати мінімум {2} і максимум {1} символів.", MinimumLength = 8)]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        [Required(ErrorMessage = "Дане поле є обов'язковим")]
        [Compare("Password", ErrorMessage = "Паролі не співпадають")]
        [Display(Name = "Підтвердження паролю")]
        [DataType(DataType.Password)]
        public string PasswordConfirm { get; set; }

    }
}
