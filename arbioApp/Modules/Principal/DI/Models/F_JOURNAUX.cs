namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_JOURNAUX
    {
        [Required]
        [StringLength(7)]
        public string JO_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(8)]
        public byte[] cbJO_Num { get; set; }

        [StringLength(35)]
        public string JO_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbJO_Intitule { get; set; }

        [StringLength(13)]
        public string CG_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_Num { get; set; }

        public short? JO_Type { get; set; }

        public short? JO_NumPiece { get; set; }

        public short? JO_Contrepartie { get; set; }

        public short? JO_SaisAnal { get; set; }

        public short? JO_NotCalcTot { get; set; }

        public short? JO_Rappro { get; set; }

        public short? JO_Sommeil { get; set; }

        public short? JO_IFRS { get; set; }

        public short? JO_Reglement { get; set; }

        public short? JO_SuiviTreso { get; set; }

        public short? JO_LettrageSaisie { get; set; }

        public short? cbProt { get; set; }

        [Key]
        public int cbMarq { get; set; }

        [StringLength(4)]
        public string cbCreateur { get; set; }

        public DateTime? cbModification { get; set; }

        public int? cbReplication { get; set; }

        public short? cbFlag { get; set; }

        public DateTime? cbCreation { get; set; }

        public Guid? cbCreationUser { get; set; }

        //public short? JO_Protec { get; set; }
    }
}
