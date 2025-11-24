namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;

    public partial class Users
    {
        [Key]
        public int UserId { get; set; }

        [Required]
        public string Login { get; set; }

        [Required]
        public string UserPassword { get; set; }

        public int RoleId { get; set; }

        public int? EstActif { get; set; }
    }
}
