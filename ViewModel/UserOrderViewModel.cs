using System.ComponentModel.DataAnnotations;

namespace LabaOne.ViewModel
{
    public class UserOrderViewModel
    {
        [Required(ErrorMessage = "Поле не повинне бути порожнім")]
        [Display(Name = "ПІБ")]
        [DataType(DataType.Text)]
        public string ClientName { get; set; }


        [Required(ErrorMessage = "Поле не повинне бути порожнім")]
        [Display(Name = "Номер телефону")]
        [DataType(DataType.PhoneNumber)]
        public string ClientPhoneNumber { get; set; }


        [Required(ErrorMessage = "Поле не повинне бути порожнім")]
        [Display(Name = "Адресa")]
        [DataType(DataType.Text)]
        public string ClientAddress { get; set; }



        [Display(Name = "Магазин")]
        public int IdStore { get; set; }
        [Display(Name = "Менеджер")]
        public string EmployeesName { get; set; }
        [Display(Name = "Доставка")]
        public int DeliveryName { get; set; }

        [Display(Name = "Книга")]
        public int MyBook { get; set; }
        [Display(Name = "Книги")]
        public IList<CountriesVariant> CountriesVariantList { get; set; }
        public IList<SymptomsVariant> SymptomsVariantList { get; set; }

        public int UserId { get; set; }
        public int ClientId { get; set; }
        public VirusGroup VirusGroup { get; set; }
        public Virus Virus{ get; set; }
        public Variant Variant { get; set; }
        public Country Country { get; set; }
        public Symptom Symptom { get; set; }
        
    }
}
