namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_PARAMETRECIAL
    {
        public short? P_ChpAnal { get; set; }

        public short? P_ChpAnalDoc { get; set; }

        public short? P_Conversion { get; set; }

        public short? P_Indispo { get; set; }

        public short? P_Echeances { get; set; }

        public short? P_StockSaisie { get; set; }

        public short? P_PeriodEncours { get; set; }

        public short? P_BaseEncours { get; set; }

        public short? P_Contremarque { get; set; }

        public short? P_ReportAchat { get; set; }

        public short? P_Analytique { get; set; }

        public short? P_Devise { get; set; }

        public short? P_TransfertDevise { get; set; }

        [StringLength(7)]
        public string P_Journal01 { get; set; }

        [StringLength(7)]
        public string P_Journal02 { get; set; }

        [StringLength(7)]
        public string P_Journal03 { get; set; }

        [StringLength(7)]
        public string P_Journal04 { get; set; }

        [StringLength(7)]
        public string P_Journal05 { get; set; }

        [StringLength(7)]
        public string P_Journal06 { get; set; }

        public short? P_Piece01 { get; set; }

        public short? P_Piece02 { get; set; }

        public short? P_Piece03 { get; set; }

        public short? P_Piece04 { get; set; }

        [StringLength(35)]
        public string P_Libelle01 { get; set; }

        [StringLength(35)]
        public string P_Libelle02 { get; set; }

        [StringLength(35)]
        public string P_Libelle03 { get; set; }

        [StringLength(35)]
        public string P_Libelle04 { get; set; }

        public short? P_LigneNeg { get; set; }

        public short? P_Quantite { get; set; }

        public short? P_InfoLibre { get; set; }

        public short? P_Seuil { get; set; }

        public short? P_Appel { get; set; }

        public short? P_Arrondi { get; set; }

        public short? P_TransfertCaisse { get; set; }

        [StringLength(7)]
        public string P_JournalCaisse { get; set; }

        public short? P_PieceEcrReg { get; set; }

        public short? P_RefEcrReg { get; set; }

        public short? P_MvtCaisse { get; set; }

        [StringLength(7)]
        public string P_JournEcartRegl { get; set; }

        [StringLength(13)]
        public string P_DebitRegl { get; set; }

        [StringLength(13)]
        public string P_CreditRegl { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? P_MaxEcart { get; set; }

        public short? P_CptaCaisse { get; set; }

        [StringLength(13)]
        public string P_DebitCaisse { get; set; }

        [StringLength(13)]
        public string P_CreditCaisse { get; set; }

        [StringLength(7)]
        public string P_JournEcartConv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? P_MaxEcartConv { get; set; }

        public short? P_UniciteLot { get; set; }

        public short? P_UniteTemps { get; set; }

        public short? P_TransfertIFRS { get; set; }

        public short? P_ConversionEng { get; set; }

        public short? P_AnalytiqueEng { get; set; }

        public short? P_DeviseEng { get; set; }

        public short? P_TransfertDeviseEng { get; set; }

        public short? P_LigneNegEng { get; set; }

        public short? P_QuantiteEng { get; set; }

        public short? P_InfoLibreEng { get; set; }

        public short? P_TransfertIFRSEng { get; set; }

        public short? P_ValSerie01 { get; set; }

        public short? P_ValSerie02 { get; set; }

        public DateTime? P_ExercicePrevision { get; set; }

        public short? P_PeriodicitePrevision { get; set; }

        public short? P_InterditSommeil { get; set; }

        public short? P_GestionMulti { get; set; }

        public short? P_Priorite { get; set; }

        public short? P_GestionControle { get; set; }

        public short? P_ReportInfos { get; set; }

        public short? P_ReportPrixRev { get; set; }

        public short? P_ComptaRegl { get; set; }

        public short? P_BaseCalculMarge { get; set; }

        public short? P_GestionPlanning { get; set; }

        public short? P_EvtPlanning { get; set; }

        public short? P_RefStructuree { get; set; }

        [StringLength(21)]
        public string P_LibellePoidsNet { get; set; }

        [StringLength(21)]
        public string P_LibellePoidsBrut { get; set; }

        [StringLength(21)]
        public string P_LibelleCondi { get; set; }

        [StringLength(21)]
        public string P_LibelleEntete1 { get; set; }

        [StringLength(21)]
        public string P_LibelleEntete2 { get; set; }

        [StringLength(21)]
        public string P_LibelleEntete3 { get; set; }

        [StringLength(21)]
        public string P_LibelleEntete4 { get; set; }

        [StringLength(21)]
        public string P_LibelleReference { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? P_TauxEscompte { get; set; }

        [StringLength(61)]
        public string P_ExonerationTVA { get; set; }

        [StringLength(19)]
        public string AR_Ref { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
