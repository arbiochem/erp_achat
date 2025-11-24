namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_ARTCOMPTA
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        public short? ACP_Type { get; set; }

        public short? ACP_Champ { get; set; }

        [StringLength(13)]
        public string ACP_ComptaCPT_CompteG { get; set; }

        [StringLength(13)]
        public string ACP_ComptaCPT_CompteA { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_Taxe1 { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_Taxe2 { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_Taxe3 { get; set; }

        public DateTime? ACP_ComptaCPT_Date1 { get; set; }

        public DateTime? ACP_ComptaCPT_Date2 { get; set; }

        public DateTime? ACP_ComptaCPT_Date3 { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_TaxeAnc1 { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_TaxeAnc2 { get; set; }

        [StringLength(5)]
        public string ACP_ComptaCPT_TaxeAnc3 { get; set; }

        public short? ACP_TypeFacture { get; set; }

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
    }
}
