namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_CAISSE
    {
        public int? CA_No { get; set; }

        [StringLength(35)]
        public string CA_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbCA_Intitule { get; set; }

        public int DE_No { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        public int? CO_NoCaissier { get; set; }

        public int? cbCO_NoCaissier { get; set; }

        [Required]
        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [Required]
        [StringLength(7)]
        public string JO_Num { get; set; }

        public short? CA_IdentifCaissier { get; set; }

        public short? N_Comptoir { get; set; }

        public short? N_Clavier { get; set; }

        public short? CA_LignesAfficheur { get; set; }

        public short? CA_ColonnesAfficheur { get; set; }

        public short? CA_ImpTicket { get; set; }

        public short? CA_SaisieVendeur { get; set; }

        public short? CA_Souche { get; set; }

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
