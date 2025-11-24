using Objets100cLib;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.CodeParser;
using arbioApp.Modules.Principal.DI._2_Documents;
using System.Data.SqlClient;

namespace arbioApp.Services
{
    internal class F_DOCENTETEService
    {
        // =====================================================================================================
        // DEBUT DECLARATION DES VARIABLES =====================================================================
        // =====================================================================================================
        private readonly AppDbContext _context;
        private readonly F_DOCENTETERepository _f_DOCENTETERepository;
        private readonly F_COMPTETRepository _f_COMPTETRepository;
        private readonly P_EXPEDITIONRepository _p_EXPEDITIONRepository;
        private readonly F_COLLABORATEURRepository _f_COLLABORATEURRepository;
        private readonly F_LIVRAISONRepository _f_LIVRAISONRepository;
        // =====================================================================================================
        // FIN DECLARATION DES VARIABLES =======================================================================
        // =====================================================================================================


        // ===================================================================================================
        // DEBUT CONSTRUCTEUR ================================================================================
        // ===================================================================================================
        public F_DOCENTETEService(F_DOCENTETERepository f_DOCENTETERepository)
        {
            _context = new AppDbContext();

            _f_DOCENTETERepository = f_DOCENTETERepository;
            _f_COMPTETRepository = new F_COMPTETRepository(_context);
            _p_EXPEDITIONRepository = new P_EXPEDITIONRepository(_context);
            _f_COLLABORATEURRepository = new F_COLLABORATEURRepository(_context);
            _f_LIVRAISONRepository = new F_LIVRAISONRepository(_context);
        }
        // ===================================================================================================
        // FIN CONSTRUCTEUR ==================================================================================
        // ===================================================================================================

        // ===================================================================================================
        // DEBUT FONCTIONS NE NECESSITANT LE REPOSITORY ======================================================
        // ===================================================================================================

        // FORMATTER LA PRESENTATION DU NUMERO DE PIECE ======================================================
        public string FormatPieceNo(int maxNo, string prefixe)
        {
            string doPieceActu;
            if (maxNo <= 9)
            {
                doPieceActu = prefixe + "0000" + maxNo.ToString();
                return doPieceActu;
            }
            else if (maxNo <= 99 && maxNo > 9)
            {
                doPieceActu = prefixe + "000" + maxNo.ToString();
                return doPieceActu;
            }
            else if (maxNo <= 999 && maxNo > 99)
            {
                doPieceActu = prefixe + "00" + maxNo.ToString();
                return doPieceActu;
            }
            else if (maxNo <= 9999 && maxNo > 999)
            {
                doPieceActu = prefixe + "0" + maxNo.ToString();
                return doPieceActu;
            }
            else if (maxNo <= 99999 && maxNo > 9999)
            {
                doPieceActu = prefixe + maxNo.ToString();
                return doPieceActu;
            }
            else
            {
                return null;
            }
        }





        // TYPE DU DOCUMENT ==============================================================================
        public DocumentType GetDocumentType(string docType)
        {
            if (docType == "Devis")
            {
                return DocumentType.DocumentTypeVenteDevis;
            }
            else if (docType == "Bon de commande")
            {
                return DocumentType.DocumentTypeVenteCommande;
            }
            else if (docType == "Préparation de livraison")
            {
                return DocumentType.DocumentTypeVentePrepaLivraison;
            }
            else if (docType == "Bon de livraison")
            {
                return DocumentType.DocumentTypeVenteLivraison;
            }
            else if (docType == "Bon de retour")
            {
                return DocumentType.DocumentTypeVenteReprise;
            }
            else if (docType == "Bon d'avoir finanicier")
            {
                return DocumentType.DocumentTypeVenteAvoir;
            }
            else if (docType == "Facture")
            {
                return DocumentType.DocumentTypeAchatFacture;
            }
            else if (docType == "Facture de retour")
            {
                return DocumentType.DocumentTypeAchatFacture;
            }
            else if (docType == "Facture d'avoir")
            {
                return DocumentType.DocumentTypeAchatAvoir;
            }
            else
            {
                return DocumentType.DocumentTypeAchatFacture;
            }
        }





        // NUMERO DU TYPE DE DOCUMENT =======================================================================
        public int? GetDocTypeNo(string docType)
        {
            if (docType == "Projet d'achat" || docType == "APA")
            {
                return 10;
            }
            else if (docType == "Préparation de commande" || docType == "APC")
            {
                return 11;
            }
            else if (docType == "Bon de commande" || docType == "ABC")
            {
                return 12;
            }
            else if (docType == "Bon de livraison" || docType == "ABL")
            {
                return 13;
            }
            else if (docType == "Bon de retour" || docType == "ABR")
            {
                return 14;
            }
            else if (docType == "Bon d'avoir" || docType == "ABA")
            {
                return 15;
            }
            else if (docType == "Facture" || docType == "Facture de retour" || docType == "Facture d'avoir" || docType == "AFA"
                 || docType == "AFR" || docType == "AFV")
            {
                return 16;
            }
            else if (docType == "Facture comptabilisée")
            {
                return 17;
            }
            else
            {
                return 200;
            }
        }





        // NOM DU TYPE DE DOCUMENT =======================================================================
        public string GetDocTypeName(int docType, string DO_Piece)
        {
            if (docType == 10)
            {
                return "Projet d'achat";
            }
            else if (docType == 11)
            {
                return "Préparation de commande";
            }
            else if (docType == 12)
            {
                return "Bon de commande";
            }
            else if (docType == 13)
            {
                return "Bon de livraison";
            }
            else if (docType == 14)
            {
                return "Bon de retour";
            }
            else if (docType == 15)
            {
                return "Bon d'avoir";
            }
            else if (docType == 16)
            {
                if (DO_Piece.StartsWith("AFA"))
                    return "Facture";
                else if (DO_Piece.StartsWith("AFR"))
                    return "Facture de retour";
                else
                    return "Facture d'avoir";
            }
            else if (docType == 17)
            {
                return "Facture comptabilisée";
            }
            else
            {
                return null;
            }
        }



        // RECHERCHE DES VALEURS DE TRANSACTION ET REGIME ==============================================================
        public (int transaction, int regime) GetTransacEtRegime(string typeDoc)
        {
            if (typeDoc == "Devis" || typeDoc == "Bon de commande" || typeDoc == "Préparation de livraison" || typeDoc == "Bon de livraison" || typeDoc == "Facture")
            {
                return (11, 21);
            }
            else if (typeDoc == "Bon de retour" || typeDoc == "Bon d'avoir finanicier" || typeDoc == "Facture de retour" || typeDoc == "Facture d'avoir")
            {
                return (21, 25);
            }
            else
            {
                return (0, 0);
            }
        }
        // ===========================================================================================================
        // FIN FONCTIONS NE NECESSITANT LE REPOSITORY ================================================================
        // ===========================================================================================================

        public string GetDepotNameByNo(int? deNo)
        {
            using (var ctx = new AppDbContext())
            {
                var depot = ctx.F_DEPOT.FirstOrDefault(d => d.DE_No == deNo);
                return depot?.DE_Intitule ?? "Inconnu";
            }
        }




        // ===========================================================================================================
        // DEBUT FONCTIONS NECESSITANT LE REPOSITORY =================================================================
        // ===========================================================================================================
        public void InsertNewF_DOCENTETE(string typeDoc, string noPiece, F_COMPTET frns,
            short? numExpedit, string caNum, int? numCaisse,
            int? numCaissier, string expeditInt, DateTime dateLivrPrev,
            DateTime dateLivrRealise, string refer, string DO_Coord01,
            int representant, int DE_No, int N_Devise,
            decimal DO_TxEscompte, string TA_Code, string _prefix,
            int number, decimal Do_Taxe1)
        {
            try
            {
                int? DO_Type = GetDocTypeNo(typeDoc);
                decimal? DO_Cours = _f_COMPTETRepository.GetF_COMPTET_Cours_N_Devise(frns.CT_Num);
                string CG_NumPrinc = _f_COMPTETRepository.GetF_COMPTET_CG_NumPrinc(frns.CT_Num).ToString();
                (int, int) tupleTransactionRegime = GetTransacEtRegime(typeDoc);
                DateTime now = DateTime.Now;
                string heureNow = "000" + now.Hour + now.Minute + now.Second;
                caNum = "0";

                //F_COLLABORATEUR representantDOCENTETE = _f_COLLABORATEURRepository.GetBy_CO_Nom_And_CO_Prenom(representant.CO_Nom);
                //int cono = (int)representantDOCENTETE.CO_No;

                P_EXPEDITION expedit = _p_EXPEDITIONRepository.Get_P_EXPEDITIONBy_E_Intitule(expeditInt);

                F_DOCENTETE newDocEnTete = new F_DOCENTETE
                {
                    DO_Type = (short)DO_Type,
                    DO_Piece = noPiece,
                    DO_Ref = refer,
                    DO_Tiers = frns.CT_Num,
                    CO_No = representant,
                    cbCO_No = representant,
                    DO_Period = 0,
                    DO_Devise = (short?)N_Devise,
                    DO_Cours = DO_Cours,
                    DE_No = DE_No,
                    cbDE_No = DE_No,
                    CT_NumPayeur = frns.CT_Num,
                    DO_Expedit = numExpedit,
                    DO_NbFacture = 1,
                    DO_BLFact = 0,
                    DO_TxEscompte = DO_TxEscompte,
                    CA_Num = caNum,
                    DO_Coord01 = DO_Coord01,
                    DO_DateLivr = dateLivrPrev,
                    DO_Condition = frns.N_Condition ?? 0,
                    DO_Tarif = frns.N_CatTarif ?? 0,
                    DO_Transaction = (short?)tupleTransactionRegime.Item1 ?? 0,
                    DO_Langue = frns.CT_Langue ?? 0,
                    DO_Regime = (short?)tupleTransactionRegime.Item2 ?? 0,
                    N_CatCompta = frns.N_CatCompta ?? 0,
                    CG_Num = CG_NumPrinc,
                    DO_Heure = heureNow,
                    CA_No = numCaisse ?? 0,
                    cbCA_No = numCaisse == 0 ? null : numCaisse,
                    CO_NoCaissier = numCaissier ?? 0,
                    cbCO_NoCaissier = numCaissier == 0 ? null : numCaissier,
                    CA_NumIFRS = frns.CA_NumIFRS == null ? "" : frns.CA_NumIFRS,
                    DO_TypeFrais = expedit.E_TypeFrais ?? 0,
                    DO_ValFrais = expedit.E_ValFrais ?? 0,
                    DO_TypeLigneFrais = expedit.E_TypeLigneFrais ?? 0,
                    DO_TypeFranco = expedit.E_TypeFranco ?? 0,
                    DO_ValFranco = expedit.E_ValFranco ?? 0,
                    DO_TypeLigneFranco = expedit.E_TypeLigneFranco ?? 0,
                    DO_Taxe1 = Do_Taxe1,
                    DO_DateLivrRealisee = dateLivrRealise,
                    DO_Statut = 0, 
                    DO_CodeTaxe1 = TA_Code,
                    cbCreationUser = FrmMdiParent._id_user
                    //Commentaires = commentaires,
                    //Divers = divers,

                    //----------------------------------------AJOUTÉ PAR DEFAUT DANS F_DOCENTETERepository
                    //DO_Domaine = 0,
                    //DO_Date = DateTime.Now,
                    //DO_Reliquat = 0,
                    //DO_Imprim = 0,
                    //DO_Coord02 = "",
                    //DO_Coord03 = "",
                    //DO_Coord04 = "",
                    //DO_Souche = 0,
                    //DO_Colisage = 1,
                    //DO_TypeColis = 1,
                    //DO_Ecart = 0,
                    //DO_Ventile = 0,
                    //AB_No = 0,
                    //DO_DebutAbo = new DateTime(1753, 01, 01, 00, 00, 00),
                    //DO_FinAbo = new DateTime(1753, 01, 01, 00, 00, 00),
                    //DO_DebutPeriod = new DateTime(1753, 01, 01, 00, 00, 00),
                    //DO_FinPeriod = new DateTime(1753, 01, 01, 00, 00, 00),
                    //DO_Statut = 2,
                    //DO_Transfere = 0,
                    //DO_Cloture = 0,
                    //DO_NoWeb = "",
                    //DO_Attente = 0,
                    //DO_Provenance = 0,
                    //MR_No = 0,
                    //DO_TypeTaux1 = 0,
                    //DO_TypeTaxe1 = 0,
                    //DO_Taxe2 = 0,
                    //DO_TypeTaux2 = 0,
                    //DO_TypeTaxe2 = 0,
                    //DO_Taxe3 = 0,
                    //DO_TypeTaux3 = 0,
                    //DO_TypeTaxe3 = 0,
                    //DO_MajCpta = 0,
                    //DO_Motif = "",
                    //CT_NumCentrale = null,
                    //DO_Contact = "",
                    //DO_FactureElec = 0,
                    //DO_TypeTransac = 0,
                    //DO_DateExpedition = new DateTime(1753, 01, 01, 00, 00, 00),
                    //DO_FactureFrs = "",
                    //DO_PieceOrig = "",
                    //DO_GUID = null,
                    //DO_EStatut = 0,
                    //DO_DemandeRegul = 0,
                    //ET_No = 0,
                    //cbET_No = null,
                    //DO_Valide = 0,
                    //DO_Coffre = 0,
                    //DO_CodeTaxe2 = null,
                    //DO_CodeTaxe3 = null,
                    //DO_TotalHT = 0,
                    //DO_StatutBAP = 0,
                    //DO_Escompte = 1,
                    //DO_DocType = 6,
                    //DO_TypeCalcul = 0,
                    //DO_FactureFile = null,
                    //DO_TotalHTNet = 0,
                    //DO_TotalTTC = 0,
                    //DO_NetAPayer = 0,
                    //DO_MontantRegle = 0,
                    //DO_RefPaiement = null,
                    //DO_AdressePaiement = "",
                    //DO_PaiementLigne = 0,
                    //DO_MotifDevis = 0,
                    //DO_Conversion = 0
                    //cbCreateur = "COLS",
                };
                //MessageBox.Show(@"typeDoc = " + DO_Type.ToString() + System.Environment.NewLine +
                //        "noPiece = " + noPiece.ToString() + System.Environment.NewLine +
                //        "frns = " + frns.CT_Num.ToString() + System.Environment.NewLine +
                //        "numExpedit = " + numExpedit.ToString() + System.Environment.NewLine +
                //        "caNum = " + caNum.ToString() + System.Environment.NewLine +
                //        "numCaisse = " + numCaisse.ToString() + System.Environment.NewLine +
                //        "numCaissier = " + numCaissier.ToString() + System.Environment.NewLine +
                //        "expeditInt = " + expeditInt.ToString() + System.Environment.NewLine +
                //        "dateLivrPrev = " + dateLivrPrev.ToString() + System.Environment.NewLine +
                //        "dateLivrRealise = " + dateLivrRealise.ToString() + System.Environment.NewLine +
                //        "refer = " + refer.ToString() + System.Environment.NewLine +
                //       "DO_Coord01 " + DO_Coord01.ToString() + System.Environment.NewLine +
                //        "representant = " + representant.ToString() + System.Environment.NewLine +
                //        "DE_No = " + DE_No.ToString() + System.Environment.NewLine +
                //        "N_Devise " + N_Devise.ToString() + System.Environment.NewLine +
                //       "DO_TxEscompte = " + DO_TxEscompte.ToString() + System.Environment.NewLine +
                //        "TA_Code = " + TA_Code.ToString(), "Test", MessageBoxButtons.OK, MessageBoxIcon.Information);


                _f_DOCENTETERepository.Add(newDocEnTete);
                frmEditDocument.UpdateSequence(_prefix, number);
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        //public string GetCurrentDocNumber(string docType, List<F_DOCENTETE> listeDocs)
        //{
        //    int maxNumber;
        //    if (docType == "Projet d'achat")
        //    {
        //        List<string> listeDOPiece = listeDocs.Where(doc => doc.DO_Piece.StartsWith("APA")).Select(doc => doc.DO_Piece).ToList();
        //        List<int> listeNumDOPiece = new List<int>();
        //        if (listeDOPiece.Count() == 0)
        //        {
        //            maxNumber = 1;
        //        }
        //        else
        //        {
        //            foreach (var piece in listeDOPiece)
        //            {
        //                listeNumDOPiece.Add(int.Parse(piece.Substring(2)));
        //            }
        //            maxNumber = listeNumDOPiece.Max() + 1;
        //        }
        //        return FormatPieceNo(maxNumber, "APA");
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}
        public F_DOCENTETE GetDocByPiece(string doPiece, List<F_DOCENTETE> listeDocs)
        {
            return listeDocs
                .FirstOrDefault(doc => doc.DO_Piece == doPiece);
        }


        public void UpdateProprietesF_DOCENTETE(string currentDocPieceNo, DateTime dateLivrPrevu, DateTime dateLivrReal,
    string reference, string caNum, int representant, short? numExpedit,
    string expeditInt, string entete, int DE_No, decimal Do_Taxe1, short Do_Statut, string DoTiers)
        {
            F_DOCENTETE f_DOCENTETE = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(currentDocPieceNo);
            if (f_DOCENTETE == null)
                throw new Exception("Document non trouvé pour mise à jour.");

            F_COLLABORATEUR representantDOCENTETE = _f_COLLABORATEURRepository.GetBy_CO_No(representant);
            P_EXPEDITION expedit = _p_EXPEDITIONRepository.Get_P_EXPEDITIONBy_E_Intitule(expeditInt);
            //f_DOCENTETE.DO_Intitule = "";
            f_DOCENTETE.DO_DateLivr = dateLivrPrevu;
            f_DOCENTETE.DO_DateLivrRealisee = dateLivrReal;
            f_DOCENTETE.DO_Ref = reference;
            f_DOCENTETE.CA_Num = caNum;
            f_DOCENTETE.CO_No = representant;
            f_DOCENTETE.DO_Tiers = DoTiers;
            f_DOCENTETE.CT_NumPayeur = DoTiers;
            //f_DOCENTETE.cbCO_No = 2;
            f_DOCENTETE.DO_Expedit = numExpedit;
            f_DOCENTETE.DO_ValFranco = expedit.E_ValFranco;
            f_DOCENTETE.DO_ValFrais = expedit.E_ValFrais;
            f_DOCENTETE.DO_Coord01 = entete;
            f_DOCENTETE.DE_No = DE_No;
            f_DOCENTETE.DO_Taxe1 = Do_Taxe1;
            f_DOCENTETE.DO_Statut = Do_Statut;
            _f_DOCENTETERepository.UpdateProprietesF_DOCENTETE(f_DOCENTETE);
        }
        public void TransformF_DOCENTETE(int cbMarqSource, short newDoType, string currentDocPieceNo, string newDocPieceNo, DateTime dateLivrPrevu, DateTime dateLivrReal,
string reference, string caNum, int representant, short? numExpedit,
string expeditInt, string entete, int DE_No, decimal Do_Taxe1, short Do_Statut, DateTime newDoDate)
        {
            //F_DOCENTETE f_DOCENTETE = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(currentDocPieceNo);
            //if (f_DOCENTETE == null)
            //    throw new Exception("Document non trouvé pour mise à jour.");
            F_DOCENTETE f_DOCENTETE = _f_DOCENTETERepository.GetBy_CbMarq(cbMarqSource);
            if (f_DOCENTETE == null)
                throw new Exception("Document non trouvé pour mise à jour.");

            F_COLLABORATEUR representantDOCENTETE = _f_COLLABORATEURRepository.GetBy_CO_No(representant);
            P_EXPEDITION expedit = _p_EXPEDITIONRepository.Get_P_EXPEDITIONBy_E_Intitule(expeditInt);
            f_DOCENTETE.DO_Type = newDoType;
            f_DOCENTETE.DO_Piece = newDocPieceNo;
            f_DOCENTETE.DO_DateLivr = dateLivrPrevu;
            f_DOCENTETE.DO_DateLivrRealisee = dateLivrReal;
            f_DOCENTETE.DO_Ref = reference;
            f_DOCENTETE.CA_Num = caNum;
            f_DOCENTETE.CO_No = representant;
            //f_DOCENTETE.cbCO_No = 2;
            f_DOCENTETE.DO_Expedit = numExpedit;
            f_DOCENTETE.DO_ValFranco = expedit.E_ValFranco;
            f_DOCENTETE.DO_ValFrais = expedit.E_ValFrais;
            f_DOCENTETE.DO_Coord01 = entete;
            f_DOCENTETE.DE_No = DE_No;
            f_DOCENTETE.DO_Taxe1 = Do_Taxe1;
            f_DOCENTETE.DO_Statut = Do_Statut;
            f_DOCENTETE.DO_Date = newDoDate;
            _f_DOCENTETERepository.UpdateProprietesF_DOCENTETE(f_DOCENTETE);
        }


        public void UpdateDO_Totaux_HT_Net_TTC(string DO_Piece, decimal? currentPUHT, int currentQuantite, decimal? previousDO_TotalHTLigne, decimal? previousDO_TotalTTCLigne)
        {
            F_DOCENTETE document = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(DO_Piece);

            //decimal? currentDO_TotalHTLigne = currentQuantite * currentPUHT;
            //decimal? newDO_TotalHT = document.DO_TotalHT - previousDO_TotalHTLigne + currentDO_TotalHTLigne;
            //decimal? DO_TotalHTNet = newDO_TotalHT + document.DO_ValFrais - (newDO_TotalHT * document.DO_TxEscompte) / 100;
            //decimal? DO_TotalTTC = DO_TotalHTNet + (DO_TotalHTNet * document.DO_Taxe1) / 100;
            decimal? currentDO_TotalHTLigne = previousDO_TotalHTLigne;
            decimal? currentDO_TotalTTCLigne = previousDO_TotalTTCLigne;

            _f_DOCENTETERepository.UpdateDO_Totaux_HT_Net_TTC_Repo(DO_Piece, currentDO_TotalHTLigne, currentDO_TotalHTLigne, currentDO_TotalTTCLigne);

        }




        //public void UpdateDO_TotalHTAfterDelete(string DO_Piece, decimal? previousDO_TotalHTLigne)
        //{
        //    F_DOCENTETE document = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(DO_Piece);
        //    decimal? newDO_TotalHT = document.DO_TotalHT - previousDO_TotalHTLigne;
        //    decimal? DO_TotalHTNet = newDO_TotalHT + document.DO_ValFrais - (newDO_TotalHT * document.DO_TxEscompte) / 100;
        //    decimal? DO_TotalTTC = DO_TotalHTNet + (DO_TotalHTNet * document.DO_Taxe1) / 100;

        //    _f_DOCENTETERepository.UpdateDO_Totaux_HT_Net_TTC_Repo(DO_Piece, newDO_TotalHT, DO_TotalHTNet, DO_TotalTTC);
        //}
        public void UpdateDO_TotalHTAfterDelete(string DO_Piece, decimal? previousDO_TotalHTLigne)
        {
            F_DOCENTETE document = _f_DOCENTETERepository.GetBy_DO_Piece_And_Type(DO_Piece);

            decimal oldTotalHT = document.DO_TotalHT ?? 0;
            decimal lineAmount = previousDO_TotalHTLigne ?? 0;

            decimal newDO_TotalHT = oldTotalHT - lineAmount;
            if (newDO_TotalHT < 0) newDO_TotalHT = 0;

            decimal valFrais = document.DO_ValFrais ?? 0;
            decimal txEscompte = document.DO_TxEscompte ?? 0;
            decimal taxe1 = document.DO_Taxe1 ?? 0;

            decimal DO_TotalHTNet = newDO_TotalHT + valFrais - (newDO_TotalHT * txEscompte / 100);
            decimal DO_TotalTTC = DO_TotalHTNet + (DO_TotalHTNet * taxe1 / 100);

            _f_DOCENTETERepository.UpdateDO_Totaux_HT_Net_TTC_Repo(
                DO_Piece,
                newDO_TotalHT,
                DO_TotalHTNet,
                DO_TotalTTC
            );
        }




        // =========================================================================================================
        // FIN FONCTIONS NECESSITANT LE REPOSITORY =================================================================
        // =========================================================================================================
    }
}
