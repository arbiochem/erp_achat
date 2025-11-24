namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;

    public partial class P_REGLEMENT
    {
        [StringLength(35)]
        public string R_Intitule { get; set; }

        [StringLength(3)]
        public string R_Code { get; set; }

        public short? R_ModePaieDebit { get; set; }

        public short? R_ModePaieCredit { get; set; }

        [StringLength(3)]
        public string IB_AFBDecaissPrinc { get; set; }

        [StringLength(3)]
        public string IB_AFBEncaissPrinc { get; set; }

        public int? EB_NoDecaiss { get; set; }

        public int? EB_NoEncaiss { get; set; }

        [StringLength(3)]
        public string R_EdiCode { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }

        public short? R_PaiementLigne { get; set; }
    }
}
