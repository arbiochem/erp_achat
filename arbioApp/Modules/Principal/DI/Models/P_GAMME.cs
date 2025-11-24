namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;

    public partial class P_GAMME
    {
        [StringLength(35)]
        public string G_Intitule { get; set; }

        public short? G_Type { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
