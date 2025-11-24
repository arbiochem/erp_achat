namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;

    public partial class P_CONDLIVR
    {
        [StringLength(35)]
        public string C_Intitule { get; set; }

        [StringLength(5)]
        public string C_Mode { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
