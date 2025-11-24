namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_COLLABORATEUR
    {
        public int CO_No { get; set; }

        [StringLength(35)]
        public string CO_Nom { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbCO_Nom { get; set; }

        [StringLength(35)]
        public string CO_Prenom { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbCO_Prenom { get; set; }

        [StringLength(35)]
        public string CO_Fonction { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbCO_Fonction { get; set; }

        [StringLength(35)]
        public string CO_Adresse { get; set; }

        [StringLength(35)]
        public string CO_Complement { get; set; }

        [StringLength(9)]
        public string CO_CodePostal { get; set; }

        [StringLength(35)]
        public string CO_Ville { get; set; }

        [StringLength(25)]
        public string CO_CodeRegion { get; set; }

        [StringLength(35)]
        public string CO_Pays { get; set; }

        [StringLength(35)]
        public string CO_Service { get; set; }

        public short? CO_Vendeur { get; set; }

        public short? CO_Caissier { get; set; }

        public short? CO_Acheteur { get; set; }

        [StringLength(21)]
        public string CO_Telephone { get; set; }

        [StringLength(21)]
        public string CO_Telecopie { get; set; }

        [StringLength(69)]
        public string CO_EMail { get; set; }

        public short? CO_Receptionnaire { get; set; }

        public int? PROT_No { get; set; }

        public int? cbPROT_No { get; set; }

        [StringLength(21)]
        public string CO_TelPortable { get; set; }

        public short? CO_ChargeRecouvr { get; set; }

        [StringLength(11)]
        public string CO_Matricule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(12)]
        public byte[] cbCO_Matricule { get; set; }

        public short? CO_Financier { get; set; }

        public short? CO_Transmission { get; set; }

        [StringLength(35)]
        public string CO_Facebook { get; set; }

        [StringLength(35)]
        public string CO_LinkedIn { get; set; }

        [StringLength(35)]
        public string CO_Skype { get; set; }

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

        //public short? CO_Sommeil { get; set; }
    }
}
