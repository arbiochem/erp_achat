using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using DevExpress.DataProcessing.InMemoryDataProcessor.GraphGenerator;
using DevExpress.XtraSpreadsheet.DocumentFormats.Xlsb;
using DevExpress.XtraWaitForm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI.Services
{
    internal class F_DOCLIGNEService
    {
        // =============================================================================================================
        // DEBUT DECLARATION DES VARIABLES =============================================================================
        // =============================================================================================================
        private readonly AppDbContext _context;

        private readonly F_DOCLIGNERepository _f_DOCLIGNERepository;
        private readonly F_AGENDARepository _f_AGENDARepository;
        private readonly F_ARTGAMMERepository _f_ARTGAMMERepository;
        private readonly F_DOCENTETERepository _f_DOCENTETERepository;
        private readonly F_ARTFOURNISSRepository _f_ARTFOURNISSRepository;

        private readonly F_DOCENTETEService _f_DOCENTETEService;
        private readonly F_ARTFOURNISSService _f_ARTFOURNISSService;
        // =============================================================================================================
        // FIN DECLARATION DES VARIABLES ===============================================================================
        // =============================================================================================================









        // ==============================================================================================================
        // DEBUT CONSTRUCTEUR ===========================================================================================
        // ==============================================================================================================
        public F_DOCLIGNEService(AppDbContext context, F_DOCLIGNERepository fDOCLIGNERepository)
        {
            _context = new AppDbContext();
            _f_DOCLIGNERepository = fDOCLIGNERepository;
            _f_AGENDARepository = new F_AGENDARepository(_context);
            _f_ARTGAMMERepository = new F_ARTGAMMERepository(_context);
            _f_DOCENTETERepository = new F_DOCENTETERepository(_context);
            _f_ARTFOURNISSRepository = new F_ARTFOURNISSRepository(_context);

            _f_DOCENTETEService = new F_DOCENTETEService(_f_DOCENTETERepository);
            _f_ARTFOURNISSService = new F_ARTFOURNISSService(_f_ARTFOURNISSRepository);
        }
        // ===============================================================================================================
        // FIN CONSTRUCTEUR ==============================================================================================
        // ===============================================================================================================









        // ================================================================================================================
        // DEBUT FONCTIONS NECESSITANT LES REPOSITORIES ===================================================================
        // ================================================================================================================
        public string GetF_DOCLIGNES_Lies_Au_AG_No_ToDelete(string EG_Enumere)
        {
            F_ARTGAMME f_ARTGAMME = _f_ARTGAMMERepository.GetByEG_Enumere(EG_Enumere);

            if (f_ARTGAMME != null)
            {
                List<F_DOCLIGNE> f_DOCLIGNEs;
                using (AppDbContext context = new AppDbContext())
                {
                    f_DOCLIGNEs = _f_DOCLIGNERepository.Get_F_DOCLIGNE_HavingAG_No(f_ARTGAMME.AG_No);
                }
                if (f_DOCLIGNEs.Count > 1)
                {
                    string DO_Piece = f_DOCLIGNEs[0].DO_Piece;
                    return DO_Piece;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }





        // ================================================================================================================
        // DEBUT INSERT ===================================================================================================
        public void AjouterF_DOCLIGNE(
            short typeDoc,
            string CT_NumClient,
            string _currentDocPieceNo,
            DateTime DO_Date,
            int? numeroLigneDL_Ligne,
            F_DOCENTETE docEnCours,
            string AR_Ref,
            string DL_Design,
            decimal DL_Taxe1,
            decimal DL_Qte,
            string typeDocument,
            F_ARTICLE articleChoisi,
            string txtBxQuantiteText,
            string txtBxRemiseText,
            string TextBoxPUNetText,
            F_COLLABORATEUR collab,
            short DL_NoRef,
            decimal DL_PUTTC,
            DateTime DO_DateLivr,
            string CA_NumText,
            string TextBoxMontantTTCText,
            string TextBoxMontantHTText,
            DateTime dateTimePicker3Value,
            int? DE_No,
            int dlno,
            int? retenu)
        {
            //F_ARTFOURNISS fournisseur = _f_ARTFOURNISSService.GetByARRefAndPrincipal(articleChoisi.AR_Ref);
            F_ARTSTOCK depotArtStock = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == articleChoisi.AR_Ref && artStck.AS_Principal == 1).FirstOrDefault();
            F_DOCLIGNE docligne = new F_DOCLIGNE();



            // Valeurs qui ne sont pas par défauts =========================================================
            docligne.DO_Type = typeDoc;
            docligne.DO_Domaine = GetDocLigneDomaine(docligne.DO_Type);
            docligne.CT_Num = CT_NumClient;
            docligne.DO_Piece = _currentDocPieceNo;
            docligne.DO_Date = DO_Date;
            docligne.DL_DateBC = DO_Date;
            docligne.DL_DateBL = DO_Date;
            //docligne.DO_Ref = docEnCours.DO_Ref == null ? "" : "";
            docligne.AR_Ref = AR_Ref;
            docligne.DL_Design = DL_Design;
            docligne.DL_Qte = DL_Qte;
            docligne.DL_QteBC = DL_Qte;
            //docligne.AF_RefFourniss = fournisseur == null ? "" : fournisseur.AF_RefFourniss;
            docligne.DL_NoRef = DL_NoRef;
            docligne.DL_PUTTC = DL_PUTTC;
            docligne.DO_DateLivr = DO_DateLivr;
            docligne.CA_Num = CA_NumText;
            docligne.DL_DatePL = dateTimePicker3Value;
            //docligne.DL_DateDE = docEnCours.DO_Date;
            //docligne.CA_No = frmEditDocument.CaisseNo;

            docligne.DL_Ligne = numeroLigneDL_Ligne;

            docligne.CO_No = collab != null ? collab.CO_No : 0;
            docligne.DE_No = DE_No;
            docligne.DL_No = dlno;

            docligne.EU_Qte = Convert.ToDecimal(txtBxQuantiteText);

            docligne.DL_PoidsNet = Convert.ToDecimal(txtBxQuantiteText) * articleChoisi.AR_PoidsNet;
            docligne.DL_PoidsBrut = Convert.ToDecimal(txtBxQuantiteText) * articleChoisi.AR_PoidsBrut;

            docligne.DL_PrixUnitaire = Convert.ToDecimal(TextBoxPUNetText);
            docligne.DL_PrixRU = Convert.ToDecimal(TextBoxPUNetText);
            docligne.DL_CMUP = Convert.ToDecimal(TextBoxPUNetText);
            docligne.DL_MontantHT = Convert.ToDecimal(TextBoxMontantHTText);
            docligne.DL_MontantTTC = Convert.ToDecimal(TextBoxMontantTTCText);

            docligne.DL_QteDE = typeDoc == 0 ? Convert.ToDecimal(txtBxQuantiteText) : 0;
            docligne.DL_Remise01REM_Valeur = (txtBxRemiseText == "") ? 0 : Convert.ToDecimal(txtBxRemiseText);



            // Valeurs par défauts ==========================================================================
            docligne.DL_DateAvancement = new DateTime(1753, 01, 01, 00, 00, 00);

            docligne.DL_Remise02REM_Valeur = 0;
            docligne.DL_Remise02REM_Type = 0;
            docligne.DL_Remise03REM_Valeur = 0;
            docligne.DL_Remise03REM_Type = 0;
            docligne.DL_TNomencl = 0;
            docligne.DL_TRemPied = 0;
            docligne.DL_TRemExep = 0;
            docligne.DL_PUBC = 0;
            docligne.DL_Taxe1 = DL_Taxe1;
            docligne.DL_TypeTaux1 = 0;
            docligne.DL_TypeTaxe1 = 0;
            docligne.DL_Taxe2 = 0;
            docligne.DL_TypeTaux2 = 0;
            docligne.DL_TypeTaxe2 = 0;
            docligne.DL_Taxe3 = 0;
            docligne.DL_TypeTaux3 = 0;
            docligne.DL_TypeTaxe3 = 0;
            docligne.DT_No = 0;
            docligne.DL_TTC = 0;
            docligne.DL_TypePL = 0;
            docligne.DL_PUDevise = 0;
            docligne.DL_Frais = 0;
            docligne.DL_NonLivre = 0;
            docligne.DL_FactPoids = 0;
            docligne.DL_Escompte = 0;
            docligne.DL_NoLink = 0;
            docligne.DL_QteRessource = 0;
            docligne.DL_PieceOFProd = 0;
            docligne.DL_NoSousTotal = 0;

            docligne.DL_Remise01REM_Type = 0; // (Remise en pourcent) // MBOLA TSY MAINTSY MIOVA REHEFA PRISE EN COMPTE NY REMISES HAFA
            docligne.AG_No1 = 0; // TODO: Gammes mbola tsy prise en compte
            docligne.AG_No2 = 0; // TODO: Gammes mbola tsy prise en compte

            docligne.DL_Valorise = 1;

            docligne.AC_RefClient = "";
            docligne.DL_PiecePL = "";
            docligne.DL_NoColis = "";
            docligne.DL_PieceBL = "";
            docligne.DL_PieceDE = "";
            docligne.DL_Operation = "";

            docligne.AR_RefCompose = null;
            docligne.RP_Code = null;
            docligne.DL_CodeTaxe1 = null;
            docligne.DL_CodeTaxe2 = null;
            docligne.DL_CodeTaxe3 = null;
            //docligne.Colisage = null;
            //docligne.Unité_de_colisage = null;
            //docligne.Commentaires = null;
            docligne.cbCreationUser = FrmMdiParent._id_user;
            docligne.Retenu = retenu == 1 ? true : false; // Retenu à la source, 0 = Non, 1 = Oui


            // Valeurs conditionnelles ==========================================================================
            if ((short)_f_DOCENTETEService.GetDocTypeNo(typeDocument) == 0 || (short)_f_DOCENTETEService.GetDocTypeNo(typeDocument) == 1)
                docligne.DL_QteBL = 0;
            else
                docligne.DL_QteBL = 1;
            if (articleChoisi.AR_SuiviStock == 0)
            {
                docligne.DL_MvtStock = 0;
            }
            else
            {
                if (typeDoc == 0 || typeDoc == 1 || typeDoc == 2 || typeDoc == 5)
                {
                    docligne.DL_MvtStock = 0;
                }
                else if (typeDoc == 3 || typeDoc == 6 || typeDoc == 7)
                {
                    docligne.DL_MvtStock = 3;
                }
                else if (typeDoc == 4 || typeDoc == 16 || typeDoc == 17)
                {
                    docligne.DL_MvtStock = 1;
                }
                else
                {
                    docligne.DL_MvtStock = 0;
                }
            }
            if (articleChoisi.AR_Condition != 0)
            {
                // TODO: Mbola tsy choix avy amin'ny utilisateur ireto lignes ireot fa valeur par défaut.
                F_CONDITION condition = _context.F_CONDITION.Where(cond => cond.AR_Ref == articleChoisi.AR_Ref).FirstOrDefault();
                docligne.EU_Enumere = condition.EC_Enumere;
            }
            else
            {
                docligne.EU_Enumere = _context.P_UNITE.Where(unit => unit.cbMarq == articleChoisi.AR_UniteVen).Select(unit => unit.U_Intitule).FirstOrDefault();
            }
            if (CA_NumText != "")
            {
                int indexEspace = CA_NumText.IndexOf(' ');
                docligne.CA_Num = CA_NumText.Substring(0, indexEspace);
            }
            if (typeDoc == 1 || typeDoc == 2 || typeDoc == 3 || typeDoc == 4 || typeDoc == 5 || typeDoc == 6 || typeDoc == 7)
                docligne.DL_QtePL = Convert.ToInt32(txtBxQuantiteText);
            else
                docligne.DL_QtePL = 0;



            // Insertion dans la base ==================================================================
            _f_DOCLIGNERepository.Add(docligne);
        }
        // FIN INSERT =======================================================================================
        // ==================================================================================================





        // =================================================================================================
        // DEBUT UPDATE ====================================================================================
        public void UpdateF_DOCLIGNE(string DO_Piece, string CT_Num, string AR_Ref, string DL_Designe, decimal puBrut,
            int DL_No, int quantite, string typeDocument, decimal DL_Taxe1, decimal DL_MontantHT, decimal DL_MontantTTC,
            int? retenu, decimal remise, string DL_PieceFourniss, DateTime DL_DatePieceFourniss, decimal DL_MontantRegle)
        {
            try
            {
                // F_DOCLIGNE f_DOCLIGNEToUpdate = _context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.AR_Ref == AR_Ref && dl.DL_No == (DL_No ?? dl.DL_No) && dl.DE_No == De_No).FirstOrDefault();
                F_DOCLIGNE f_DOCLIGNEToUpdate = _context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.AR_Ref == AR_Ref).FirstOrDefault();

                //using (_context)
                //{
                //    var conn = _context.Database.Connection; // EF6
                //    MessageBox.Show($"Base = {conn.Database}, Serveur = {conn.DataSource}");
                //}


                var f_ARTICLE = _context.F_ARTICLE.Where(art => art.AR_Ref == AR_Ref).Select(art => new { art.AR_PoidsBrut, art.AR_PoidsNet }).FirstOrDefault();



                if (f_DOCLIGNEToUpdate == null)
                {
                    MessageBox.Show("Ligne de document introuvable !");
                }







                var QteEtMontantArticle = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && artStck.DE_No == f_DOCLIGNEToUpdate.DE_No)
                                                    .Select(artStck => new { artStck.AS_QteSto, artStck.AS_MontSto }).FirstOrDefault();

                if (typeDocument == "Devis")
                {
                    f_DOCLIGNEToUpdate.DL_Qte = quantite;
                    f_DOCLIGNEToUpdate.EU_Qte = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBC = quantite;
                    f_DOCLIGNEToUpdate.DL_QteDE = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBL = 0;
                    f_DOCLIGNEToUpdate.DL_QtePL = 0;
                }
                else if (typeDocument == "Bon de commande")
                {
                    f_DOCLIGNEToUpdate.DL_Qte = quantite;
                    f_DOCLIGNEToUpdate.EU_Qte = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBC = quantite;
                    f_DOCLIGNEToUpdate.DL_QteDE = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBL = quantite;
                    f_DOCLIGNEToUpdate.DL_QtePL = 0;
                }
                else if (typeDocument == "Facture de retour" || typeDocument == "Facture d'avoir")
                {
                    f_DOCLIGNEToUpdate.DL_Qte = -quantite;
                    f_DOCLIGNEToUpdate.EU_Qte = -quantite;
                    f_DOCLIGNEToUpdate.DL_QteBC = -quantite;
                    f_DOCLIGNEToUpdate.DL_QteDE = -quantite;
                    f_DOCLIGNEToUpdate.DL_QteBL = -quantite;
                    f_DOCLIGNEToUpdate.DL_QtePL = -quantite;
                }
                else
                {
                    f_DOCLIGNEToUpdate.DL_Qte = quantite;
                    f_DOCLIGNEToUpdate.EU_Qte = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBC = quantite;
                    f_DOCLIGNEToUpdate.DL_QteDE = quantite;
                    f_DOCLIGNEToUpdate.DL_QteBL = quantite;
                    f_DOCLIGNEToUpdate.DL_QtePL = quantite;

                }


                f_DOCLIGNEToUpdate.CT_Num = CT_Num;
                f_DOCLIGNEToUpdate.AR_Ref = AR_Ref;
                f_DOCLIGNEToUpdate.DL_Design = f_DOCLIGNEToUpdate.DL_Design;
                f_DOCLIGNEToUpdate.DL_PrixUnitaire = puBrut;
                f_DOCLIGNEToUpdate.DL_Taxe1 = DL_Taxe1;
                f_DOCLIGNEToUpdate.DL_MontantHT = DL_MontantHT;
                f_DOCLIGNEToUpdate.DL_MontantTTC = DL_MontantTTC;
                f_DOCLIGNEToUpdate.DL_MontantRegle = DL_MontantRegle;
                f_DOCLIGNEToUpdate.DL_Remise01REM_Valeur = remise;

                // Mise à jour des propriétés concernant le poids
                f_DOCLIGNEToUpdate.DL_PoidsBrut = quantite * f_ARTICLE.AR_PoidsBrut;
                f_DOCLIGNEToUpdate.DL_PoidsNet = quantite * f_ARTICLE.AR_PoidsNet;

                // Mise à jour DL_CMUP et DL_PrixRU
                //decimal? CMUP = QteEtMontantArticle.AS_MontSto / (QteEtMontantArticle.AS_QteSto == 0 ? 1 : QteEtMontantArticle.AS_QteSto);
                //f_DOCLIGNEToUpdate.DL_CMUP = CMUP;
                //f_DOCLIGNEToUpdate.DL_PrixRU = CMUP;
                f_DOCLIGNEToUpdate.Retenu = retenu == 1 ? true : false;
                f_DOCLIGNEToUpdate.DL_PieceFourniss = DL_PieceFourniss;
                f_DOCLIGNEToUpdate.DL_DatePieceFourniss = DL_DatePieceFourniss;

                // Enregistrement des modifications via le Repository
                _f_DOCLIGNERepository.Update(f_DOCLIGNEToUpdate);

                _context.Entry(f_DOCLIGNEToUpdate).Reload();
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        public void TransformF_DOCLIGNE(short DO_Type, string DO_Piece, int DL_No)
        {
            try
            {
                DO_Piece = DO_Piece.Substring(DO_Piece.Length - 8);
                F_DOCLIGNE f_DOCLIGNEToUpdate = _context.F_DOCLIGNE.Where(dl => dl.DL_No == DL_No).FirstOrDefault();
                f_DOCLIGNEToUpdate.DO_Type = DO_Type;
                switch (DO_Type)
                {
                    case 12:
                        f_DOCLIGNEToUpdate.DO_Piece = "ABC" + DO_Piece; //BON DE COMMANDE
                        break;
                    case 13:
                        f_DOCLIGNEToUpdate.DO_Piece = "ABL" + DO_Piece; //BON DE LIVRAISON
                        break;
                    case 14:
                        f_DOCLIGNEToUpdate.DO_Piece = "ABR" + DO_Piece; //BON DE RETOUR
                        break;
                    case 16:
                        f_DOCLIGNEToUpdate.DO_Piece = "AFA" + DO_Piece; //FACTURE D'ACHAT
                        break;
                    default:
                        throw new ArgumentException("Invalid document type", nameof(DO_Type));
                        break;
                }


                _f_DOCLIGNERepository.Update(f_DOCLIGNEToUpdate);
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public void DeleteF_DOCLIGNE(string DO_Piece, int? DL_Ligne)
        {
            F_DOCLIGNE f_DOCLIGNEToDelete = _context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.DL_Ligne == DL_Ligne).FirstOrDefault();
            F_AGENDA f_AGENDA = _context.F_AGENDA.Where(a => a.DL_No == f_DOCLIGNEToDelete.DL_No).FirstOrDefault();

            if (f_AGENDA != null)
            {
                _f_AGENDARepository.DeleteF_AGENDA(f_AGENDA);
            }

            _f_DOCLIGNERepository.Delete(f_DOCLIGNEToDelete);
        }




        // FIN UPDATE ====================================================================================================
        // ===============================================================================================================

        // ===============================================================================================================
        // FIN FONCTIONS NECESSITANT LES REPOSITORIES ====================================================================
        // ===============================================================================================================









        // ===================================================================================================================
        // DEBUT FONCTIONS NE NECESSITANT PAS LES REPOSITORIES ===============================================================
        // ===================================================================================================================
        public short? GetDocLigneDomaine(short docType)
        {
            List<int> typeZero = new List<int> { 0, 1, 2, 3, 4, 5, 6, 7 };
            if (typeZero.Contains(docType))
            {
                return 0;       //VENTE
            }
            else
            {
                return 1;       //ACHAT
            }
        }
        // ==================================================================================================================
        // FIN FONCTIONS NE NECESSITANT PAS LES REPOSITORIES ================================================================
        // ==================================================================================================================





    }
}
