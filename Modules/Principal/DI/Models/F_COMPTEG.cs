namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_COMPTEG
    {
        [Required]
        [StringLength(13)]
        public string CG_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(15)]
        public byte[] cbCG_Num { get; set; }

        public short? CG_Type { get; set; }

        [StringLength(35)]
        public string CG_Intitule { get; set; }

        [StringLength(17)]
        public string CG_Classement { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCG_Classement { get; set; }

        public short? N_Nature { get; set; }

        public short? CG_Report { get; set; }

        [StringLength(13)]
        public string CR_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCR_Num { get; set; }

        [StringLength(7)]
        public string CG_Raccourci { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(8)]
        public byte[] cbCG_Raccourci { get; set; }

        public short? CG_Saut { get; set; }

        public short? CG_Regroup { get; set; }

        public short? CG_Analytique { get; set; }

        public short? CG_Echeance { get; set; }

        public short? CG_Quantite { get; set; }

        public short? CG_Lettrage { get; set; }

        public short? CG_Tiers { get; set; }

        public short? CG_Devise { get; set; }

        public short? N_Devise { get; set; }

        [StringLength(5)]
        public string TA_Code { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(6)]
        public byte[] cbTA_Code { get; set; }

        public short? CG_Sommeil { get; set; }

        public short? CG_ReportAnal { get; set; }

        public short? CG_LettrageSaisie { get; set; }

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

        //[Column("Derniere ecriture", TypeName = "smalldatetime")]
        //public DateTime? Derniere_ecriture { get; set; }

        //[Column("Nb ecritures", TypeName = "numeric")]
        //public decimal? Nb_ecritures { get; set; }
    }
}
