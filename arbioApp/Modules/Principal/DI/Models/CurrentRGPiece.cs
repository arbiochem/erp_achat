namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    [Table("CurrentRGPiece")]
    public partial class CurrentRGPiece
    {
        [Key]
        [Column("CurrentRGPiece")]
        [StringLength(50)]
        public string CurrentRGPiece1 { get; set; }
    }
}
