namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_PREFERENCES
    {
        [StringLength(19)]
        public string PR_RefEsc { get; set; }

        public short? PR_AlerteAg { get; set; }

        public short? PR_DelaiPreAlert { get; set; }

        public short? PR_Confirmation { get; set; }

        public short? PR_ChoixRemise { get; set; }

        public short? PR_UnitePoids { get; set; }

        public short? PR_Marge { get; set; }

        public short? PR_CodeEAN { get; set; }

        public short? PR_Prefixe20 { get; set; }

        public short? PR_UnitePrix { get; set; }

        public short? PR_Poids { get; set; }

        public short? PR_ModifImport { get; set; }

        public short? PR_StockNeg { get; set; }

        public short? PR_DelaiLivr { get; set; }

        [StringLength(19)]
        public string PR_RefTaxeNP { get; set; }

        [StringLength(13)]
        public string CG_NumCli { get; set; }

        [StringLength(13)]
        public string CG_NumFrs { get; set; }

        [StringLength(17)]
        public string CT_Num { get; set; }

        public short? PR_LignesAfficheur { get; set; }

        public short? PR_ColonnesAfficheur { get; set; }

        public short? PR_IdentifCaissier { get; set; }

        public short? PR_PrixTTC { get; set; }

        [StringLength(13)]
        public string CG_NumVirement { get; set; }

        public short? PR_Souche { get; set; }

        public short? PR_RegroupTicket { get; set; }

        public short? PR_RegroupRglt { get; set; }

        [StringLength(69)]
        public string PR_EMail { get; set; }

        public int? DE_No { get; set; }

        public short? PR_ArtNonLivre { get; set; }

        public short? PR_RecalculModele { get; set; }

        public short? PR_MajPAFrs { get; set; }

        public short? PR_CreeFrs { get; set; }

        public short? N_Reglement { get; set; }

        public short? PR_DateFact { get; set; }

        public short? PR_CliCaisse { get; set; }

        [StringLength(13)]
        public string CG_NumComptoirDebit { get; set; }

        [StringLength(13)]
        public string CG_NumComptoirCredit { get; set; }

        public short? PR_FondCaisse { get; set; }

        public short? PR_Integration { get; set; }

        public short? PR_ComptaBonAchat { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? PR_MontantMaxTicket { get; set; }

        public int? CD_No { get; set; }

        public short? PR_FormatFacture { get; set; }

        [StringLength(69)]
        public string PR_Certificat { get; set; }

        public short? PR_CreeArticle { get; set; }

        public short? PR_CreeTiers { get; set; }

        public short? PR_CreeAffaire { get; set; }

        public short? PR_EFacture { get; set; }

        public short? PR_ClotureMajCpta { get; set; }

        public short? PR_ExportStructCpta { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
