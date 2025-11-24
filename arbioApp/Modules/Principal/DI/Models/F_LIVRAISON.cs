namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_LIVRAISON
    {
        public int LI_No { get; set; }

        [Required]
        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [StringLength(69)]
        public string LI_Intitule { get; set; }

        [StringLength(35)]
        public string LI_Adresse { get; set; }

        [StringLength(35)]
        public string LI_Complement { get; set; }

        [StringLength(9)]
        public string LI_CodePostal { get; set; }

        [StringLength(35)]
        public string LI_Ville { get; set; }

        [StringLength(25)]
        public string LI_CodeRegion { get; set; }

        [StringLength(35)]
        public string LI_Pays { get; set; }

        [StringLength(35)]
        public string LI_Contact { get; set; }

        public short? N_Expedition { get; set; }

        public short? N_Condition { get; set; }

        public short? LI_Principal { get; set; }

        [StringLength(21)]
        public string LI_Telephone { get; set; }

        [StringLength(21)]
        public string LI_Telecopie { get; set; }

        [StringLength(69)]
        public string LI_EMail { get; set; }

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

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(70)]
        public byte[] cbLI_Intitule { get; set; }

        [StringLength(69)]
        public string LI_Commentaire { get; set; }

        public short? LI_DelaiTransport { get; set; }
    }
}
