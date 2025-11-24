namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_DEPOT
    {
        public int? DE_No { get; set; }

        [Required]
        [StringLength(35)]
        public string DE_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbDE_Intitule { get; set; }

        [StringLength(35)]
        public string DE_Adresse { get; set; }

        [StringLength(35)]
        public string DE_Complement { get; set; }

        [StringLength(9)]
        public string DE_CodePostal { get; set; }

        [StringLength(35)]
        public string DE_Ville { get; set; }

        [StringLength(35)]
        public string DE_Contact { get; set; }

        public short? DE_Principal { get; set; }

        public short? DE_CatCompta { get; set; }

        [StringLength(25)]
        public string DE_Region { get; set; }

        [StringLength(35)]
        public string DE_Pays { get; set; }

        [StringLength(69)]
        public string DE_EMail { get; set; }

        [StringLength(9)]
        public string DE_Code { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDE_Code { get; set; }

        [StringLength(21)]
        public string DE_Telephone { get; set; }

        [StringLength(21)]
        public string DE_Telecopie { get; set; }

        public int? DE_Replication { get; set; }

        public int? DP_NoDefaut { get; set; }

        public int? cbDP_NoDefaut { get; set; }

        public short? DE_Exclure { get; set; }

        public short? DE_Souche01 { get; set; }

        public short? DE_Souche02 { get; set; }

        public short? DE_Souche03 { get; set; }

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
