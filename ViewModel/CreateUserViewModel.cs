using System.ComponentModel.DataAnnotations;

namespace LabaOne.ViewModel
{
    public class CreateUserViewModel
    {
        [Required]
        [Display(Name = "Email")]
        [DataType(DataType.EmailAddress)]
        public string UserEmail { get; set; }


        [Required]
        [Display(Name = "Рік народження")]
        [Range(1900, 2022)]
        public int UserYear { get; set; }


        [Required]
        [Display(Name = "Пароль")]
        [DataType(DataType.Password)]
        public string UserPassword { get; set; }
    }
}
