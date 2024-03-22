using System.ComponentModel.DataAnnotations;

namespace QRVCard.Models
{
    public class Contact
    {
        [Required(ErrorMessage = "First name is required")]
        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Required(ErrorMessage = "Mobile no is required")]
        [StringLength(15)]
        public string Mobile { get; set; }
        
        //[Required(ErrorMessage = "Phone no is required")]
        //[StringLength(15)]
        [Display(Name = "Phone Number")]
        //[RegularExpression("[a-z A-Z]", ErrorMessage = "Please enter valid phone no")]
        public string PhoneNumber { get; set; }
        
        public string Email { get; set; }

        public string Company { get; set; }

        public string Website { get; set; }

    }
}
