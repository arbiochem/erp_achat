namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("RoleAutorisation")]
    public partial class RoleAutorisation
    {
        public int Id { get; set; }

        public int EstAutorise { get; set; }

        public int? IdRole { get; set; }

        public int IdRubrique { get; set; }
    }
}
