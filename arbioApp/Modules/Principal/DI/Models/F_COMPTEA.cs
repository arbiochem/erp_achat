namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_COMPTEA
    {
        public short N_Analytique { get; set; }

        [Required]
        [StringLength(13)]
        public string CA_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(15)]
        public byte[] cbCA_Num { get; set; }

        [StringLength(35)]
        public string CA_Intitule { get; set; }

        public short? CA_Type { get; set; }

        [StringLength(17)]
        public string CA_Classement { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCA_Classement { get; set; }

        [StringLength(7)]
        public string CA_Raccourci { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(8)]
        public byte[] cbCA_Raccourci { get; set; }

        public short? CA_Report { get; set; }

        public short? N_Analyse { get; set; }

        public short? CA_Saut { get; set; }

        public short? CA_Sommeil { get; set; }

        public short? CA_Domaine { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? CA_Achat { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? CA_Vente { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        public short? CA_Statut { get; set; }

        public DateTime? CA_DateCreationAffaire { get; set; }

        public DateTime? CA_DateAcceptAffaire { get; set; }

        public DateTime? CA_DateDebutAffaire { get; set; }

        public DateTime? CA_DateFinAffaire { get; set; }

        public short? CA_ModeFacturation { get; set; }

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

        [Column("Nb de ventilations", TypeName = "numeric")]
        public decimal? Nb_de_ventilations { get; set; }
    }
}
