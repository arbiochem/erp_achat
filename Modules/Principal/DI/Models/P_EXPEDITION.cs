namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class P_EXPEDITION
    {
        [StringLength(35)]
        public string E_Intitule { get; set; }

        [StringLength(3)]
        public string E_Mode { get; set; }

        [StringLength(19)]
        public string AR_Ref { get; set; }

        public short? E_TypeFrais { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? E_ValFrais { get; set; }

        public short? E_TypeLigneFrais { get; set; }

        public short? E_TypeFranco { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? E_ValFranco { get; set; }

        public short? E_TypeLigneFranco { get; set; }

        public short? E_TypeCalcul { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
