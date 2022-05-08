using Microsoft.AspNetCore.Identity;

namespace LabaOne.Models
{
    public class User: IdentityUser
    {
        public int Year { get; set; }
        //public int client_Id { get; set; }
    }
}
