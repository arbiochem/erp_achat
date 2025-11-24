namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("Rubrique")]
    public partial class Rubrique
    {
        public int Id { get; set; }

        [Required]
        [StringLength(100)]
        public string Nom { get; set; }

        public int Niveau { get; set; }

        public int? IdParent { get; set; }
    }
}
