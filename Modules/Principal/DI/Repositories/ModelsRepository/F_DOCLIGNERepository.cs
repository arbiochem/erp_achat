using arbioApp.Models;
using arbioApp.Modules.Principal.DI._2_Documents;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    internal class F_DOCLIGNERepository
    {
        // ================================================================================
        // DEBUT ECLARATION DES VARIABLES =================================================
        // ================================================================================
        private readonly AppDbContext _context;
        // ================================================================================
        // FIN DECLARATION DES VARIABLES ==================================================
        // ================================================================================





        // ================================================================================
        // DEBUT CONSTRUCTEUR =============================================================
        // ================================================================================
        public F_DOCLIGNERepository(AppDbContext context)
        {
            _context = context;
        }
        // ================================================================================
        // FIN CONSTRUCTEUR ===============================================================
        // ================================================================================


        public string GetF_DOCLIGNES_CTNum(int DL_No)
        {
            // On recherche la ligne correspondante
            var docLigne = _context.F_DOCLIGNE.FirstOrDefault(d => d.DL_No == DL_No);

            // Si non trouvé → null ou string.Empty
            if (docLigne == null)
                return string.Empty;

            // On retourne le champ CT_Num
            return docLigne.CT_Num;
        }







        // ================================================================================
        // DEBUT GET ======================================================================
        // ================================================================================
        public List<F_DOCLIGNE> Get_F_DOCLIGNE_HavingAG_No(int? AG_No)
        {
            using (AppDbContext context = new AppDbContext())
            {
                List<F_DOCLIGNE> f_DOCLIGNEs = context.F_DOCLIGNE.Where(dl => dl.AG_No1 == AG_No || dl.AG_No2 == AG_No).ToList();
                return f_DOCLIGNEs;
            }
        }

        public F_DOCLIGNE GetF_DOCLIGNE_By_DO_Piece_AR_Ref_DL_Ligne(string DO_Piece, string AR_Ref, int? DL_Ligne)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.AR_Ref == AR_Ref && dl.DL_Ligne == DL_Ligne).FirstOrDefault();
            }
        }

        public List<F_DOCLIGNE> GetAll_F_DOCLIGNE_Of_DOCENTETE(string DO_Piece)
        {
            DO_Piece = DO_Piece.Substring(DO_Piece.Length - 8);
            using (AppDbContext context = new AppDbContext())
            {
                return _context.F_DOCLIGNE.Where(dl => dl.DO_Piece.EndsWith(DO_Piece)).ToList();
            }
        }
        // ================================================================================
        // FIN GET ========================================================================
        // ================================================================================





        public void Add(F_DOCLIGNE f_DOCLIGNE)
        {
            try
            {
                string query = @"
                   DISABLE TRIGGER [dbo].[TG_CBINS_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                   DISABLE TRIGGER [dbo].[TG_INS_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                   DISABLE TRIGGER [dbo].[TG_INS_CPTAF_DOCLIGNE] ON [dbo].[F_DOCLIGNE];

                   Insert INTO [dbo].[F_DOCLIGNE] (
                       DO_Domaine,
                       DO_Type,
                       CT_Num,
                       DO_Piece,
                       DL_PieceBC,
                       DL_PieceBL,
                       DO_Date,
                       DL_DateBC,
                       DL_DateBL,
                       DL_Ligne,
                       DO_Ref,
                       DL_TNomencl,
                       DL_TRemPied,
                       DL_TRemExep,
                       AR_Ref,
                       DL_Design,
                       DL_Qte,
                       DL_QteBC,
                       DL_QteBL,
                       DL_PoidsNet,
                       DL_PoidsBrut,
                       DL_Remise01REM_Valeur,
                       DL_Remise01REM_Type,
                       DL_Remise02REM_Valeur,
                       DL_Remise02REM_Type,
                       DL_Remise03REM_Valeur,
                       DL_Remise03REM_Type,
                       DL_PrixUnitaire,
                       DL_PUBC,
                       DL_Taxe1,
                       DL_TypeTaux1,
                       DL_TypeTaxe1,
                       DL_Taxe2,
                       DL_TypeTaux2,
                       DL_TypeTaxe2,
                       CO_No,
                       cbCO_No,
                       AG_No1,
                       AG_No2,
                       DL_PrixRU,
                       DL_CMUP,
                       DL_MvtStock,
                       DT_No,
                       cbDT_No,
                       AF_RefFourniss,
                       EU_Enumere,
                       EU_Qte,
                       DL_TTC,
                       DE_No,
                       cbDE_No,
                       DL_NoRef,
                       DL_TypePL,
                       DL_PUDevise,
                       DL_PUTTC,
                       DL_No,
                       DO_DateLivr,
                       CA_Num,
                       DL_Taxe3,
                       DL_TypeTaux3,
                       DL_TypeTaxe3,
                       DL_Frais,
                       DL_Valorise,
                       AR_RefCompose,
                       DL_NonLivre,
                       AC_RefClient,
                       DL_MontantHT,
                       DL_MontantTTC,
                       DL_FactPoids,
                       DL_Escompte,
                       DL_PiecePL,
                       DL_DatePL,
                       DL_QtePL,
                       DL_NoColis,
                       DL_NoLink,
                       cbDL_NoLink,
                       RP_Code,
                       DL_QteRessource,
                       DL_DateAvancement,
                       PF_Num,
                       DL_CodeTaxe1,
                       DL_CodeTaxe2,
                       DL_CodeTaxe3,
                       DL_PieceOFProd,
                       DL_PieceDE,
                       DL_DateDE,
                       DL_QteDE,
                       DL_Operation,
                       DL_NoSousTotal,
                       CA_No,
                       cbCA_No,
                       cbCreationUser,
                       Retenu
                   ) 
                   values(
           {0},
                        {1},
                        {2},
                        {3},
                        {4},
                        {5},
                        {6},
                        {7},
                        {8},
                        {9},
                        {10},
                        {11},
                        {12},
                        {13},
                        {14},
                        {15},
                        {16},
                        {17},
                        {18},
                        {19},
                        {20},
                        {21},
                        {22},
                        {23},
                        {24},
                        {25},
                        {26},
                        {27},
                        {28},
                        {29},
                        {30},
                        {31},
                        {32},
                        {33},
                        {34},
                        {35},
                        {36},
                        {37},
                        {38},
                        {39},
                        {40},
                        {41},
                        {42},
                        {43},
                        {44},
                        {45},
                        {46},
                        {47},
                        {48},
                        {49},
                        {50},
                        {51},
                        {52},
                        {53},
                        {54},
                        {55},
                        {56},
                        {57},
                        {58},
                        {59},
                        {60},
                        {61},
                        {62},
                        {63},
                        {64},
                        {65},
                        {66},
                        {67},
                        {68},
                        {69},
                        {70},
                        {71},
                        {72},
                        {73},
                        {74},
                        {75},
                        {76},
                        {77},
                        {78},
                        {79},
                        {80},
                        {81},
                        {82},
                        {83},
                        {84},
                        {85},
                        {86},
                        {87},
                        {88},
                        {89},
                        {90},  
                        {91}
                   );

                   ENABLE TRIGGER [dbo].[TG_CBINS_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                   ENABLE TRIGGER [dbo].[TG_INS_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                   ENABLE TRIGGER [dbo].[TG_INS_CPTAF_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
               ";

                //using (frmQuery debugForm = new frmQuery(query))
                //{
                //    debugForm.ShowDialog();

                //     DialogResult result = MessageBox.Show(
                //        "Voulez-vous exécuter cette requête ?",
                //        "Confirmation",
                //        MessageBoxButtons.YesNo,
                //        MessageBoxIcon.Question
                //    );

                //    if (result == DialogResult.No)
                //    {
                //        return; // Arrêter si l'utilisateur annule
                //    }
                //}

                using (AppDbContext context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        query,
                        f_DOCLIGNE.DO_Domaine,
                        f_DOCLIGNE.DO_Type,
                        f_DOCLIGNE.CT_Num,
                        f_DOCLIGNE.DO_Piece,
                        f_DOCLIGNE.DL_PieceBC,
                        f_DOCLIGNE.DL_PieceBL,
                        f_DOCLIGNE.DO_Date,
                        f_DOCLIGNE.DL_DateBC,
                        f_DOCLIGNE.DL_DateBL,
                        f_DOCLIGNE.DL_Ligne,        //
                        f_DOCLIGNE.DO_Ref,
                        f_DOCLIGNE.DL_TNomencl,
                        f_DOCLIGNE.DL_TRemPied,
                        f_DOCLIGNE.DL_TRemExep,
                        f_DOCLIGNE.AR_Ref,
                        f_DOCLIGNE.DL_Design,
                        f_DOCLIGNE.DL_Qte,
                        f_DOCLIGNE.DL_QteBC,
                        f_DOCLIGNE.DL_QteBL,
                        f_DOCLIGNE.DL_PoidsNet,
                        f_DOCLIGNE.DL_PoidsBrut,
                        f_DOCLIGNE.DL_Remise01REM_Valeur,
                        f_DOCLIGNE.DL_Remise01REM_Type,
                        f_DOCLIGNE.DL_Remise02REM_Valeur,
                        f_DOCLIGNE.DL_Remise02REM_Type,
                        f_DOCLIGNE.DL_Remise03REM_Valeur,
                        f_DOCLIGNE.DL_Remise03REM_Type,
                        f_DOCLIGNE.DL_PrixUnitaire,
                        f_DOCLIGNE.DL_PUBC,
                        f_DOCLIGNE.DL_Taxe1,
                        f_DOCLIGNE.DL_TypeTaux1,
                        f_DOCLIGNE.DL_TypeTaxe1,
                        f_DOCLIGNE.DL_Taxe2,
                        f_DOCLIGNE.DL_TypeTaux2,
                        f_DOCLIGNE.DL_TypeTaxe2,
                        f_DOCLIGNE.CO_No,
                        f_DOCLIGNE.CO_No,
                        f_DOCLIGNE.AG_No1,
                        f_DOCLIGNE.AG_No2,
                        f_DOCLIGNE.DL_PrixRU,
                        f_DOCLIGNE.DL_CMUP,
                        f_DOCLIGNE.DL_MvtStock,
                        f_DOCLIGNE.DT_No,
                        f_DOCLIGNE.DT_No,
                        f_DOCLIGNE.AF_RefFourniss,
                        f_DOCLIGNE.EU_Enumere,
                        f_DOCLIGNE.EU_Qte,
                        f_DOCLIGNE.DL_TTC,
                        f_DOCLIGNE.DE_No,           //
                        f_DOCLIGNE.DE_No,         //
                        f_DOCLIGNE.DL_NoRef,
                        f_DOCLIGNE.DL_TypePL,
                        f_DOCLIGNE.DL_PUDevise,
                        f_DOCLIGNE.DL_PUTTC,
                        f_DOCLIGNE.DL_No,               //
                        f_DOCLIGNE.DO_DateLivr,
                        f_DOCLIGNE.CA_Num,
                        f_DOCLIGNE.DL_Taxe3,
                        f_DOCLIGNE.DL_TypeTaux3,
                        f_DOCLIGNE.DL_TypeTaxe3,
                        f_DOCLIGNE.DL_Frais,
                        f_DOCLIGNE.DL_Valorise,
                        f_DOCLIGNE.AR_RefCompose,
                        f_DOCLIGNE.DL_NonLivre,
                        f_DOCLIGNE.AC_RefClient,
                        f_DOCLIGNE.DL_MontantHT,
                        f_DOCLIGNE.DL_MontantTTC,
                        f_DOCLIGNE.DL_FactPoids,
                        f_DOCLIGNE.DL_Escompte,
                        f_DOCLIGNE.DL_PiecePL,
                        f_DOCLIGNE.DL_DatePL,
                        f_DOCLIGNE.DL_QtePL,
                        f_DOCLIGNE.DL_NoColis,
                        f_DOCLIGNE.DL_NoLink,
                        f_DOCLIGNE.cbDL_NoLink,
                        f_DOCLIGNE.RP_Code,
                        f_DOCLIGNE.DL_QteRessource,
                        f_DOCLIGNE.DL_DateAvancement,
                        f_DOCLIGNE.PF_Num ?? "",
                        f_DOCLIGNE.DL_CodeTaxe1,
                        f_DOCLIGNE.DL_CodeTaxe2,
                        f_DOCLIGNE.DL_CodeTaxe3,
                        f_DOCLIGNE.DL_PieceOFProd,
                        f_DOCLIGNE.DL_PieceDE,
                        f_DOCLIGNE.DL_DateDE,
                        f_DOCLIGNE.DL_QteDE,
                        f_DOCLIGNE.DL_Operation,
                        f_DOCLIGNE.DL_NoSousTotal,
                        f_DOCLIGNE.CA_No,
                        f_DOCLIGNE.cbCA_No,
                        f_DOCLIGNE.cbCreationUser,
                        f_DOCLIGNE.Retenu
                    );
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }




        public void Update(F_DOCLIGNE f_DOCLIGNE)
        {
            try
            {
                F_DOCLIGNE _f_DOCLIGNEToUpdate = _context.F_DOCLIGNE.Where(dl => dl.AR_Ref == f_DOCLIGNE.AR_Ref && dl.DL_No == f_DOCLIGNE.DL_No).FirstOrDefault();

                string queryUpdateF_DOCLIGNE = @"
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                DISABLE TRIGGER [dbo].[TG_UPD_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                DISABLE TRIGGER [dbo].[TG_UPD_CPTAF_DOCLIGNE] ON [dbo].[F_DOCLIGNE];

                UPDATE F_DOCLIGNE
                SET
                    DO_Type = @DO_Type,
                    DO_Piece = @DO_Piece,
                    AR_Ref = @AR_Ref, 
                    DL_Design = @DL_Design,
                    CT_Num  = @CT_Num,
                    DL_PrixUnitaire = @DL_PrixUnitaire,
                	DL_QteBL = @DL_QteBL,
                	DL_QteBC = @DL_QteBC,
                	DL_QtePL = @DL_QtePL,
                	DL_QteDE = @DL_QteDE,
                    DL_Taxe1 = @DL_Taxe1,
                	DL_Qte = @DL_Qte,
                	EU_Qte = @EU_Qte,
                    DL_Remise01REM_Valeur = @DL_Remise01REM_Valeur,
                	DL_PoidsNet = @DL_PoidsNet,
                	DL_PoidsBrut = @DL_PoidsBrut,
                	DL_PrixRU = @DL_PrixRU,
                	DL_MontantHT = @DL_MontantHT,
                	DL_MontantTTC = @DL_MontantTTC,
                    Retenu = @Retenu,
                    DL_MontantRegle = @DL_MontantRegle,
                    DL_PieceFourniss = @DL_PieceFourniss,
                    DL_DatePieceFourniss = @DL_DatePieceFourniss,
                    cbCreationUser = @cbCreationUser
                   
                WHERE DL_No = @DL_No;

                ENABLE TRIGGER [dbo].[TG_CBUPD_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                ENABLE TRIGGER [dbo].[TG_UPD_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                ENABLE TRIGGER [dbo].[TG_UPD_CPTAF_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
            ";

                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        queryUpdateF_DOCLIGNE,
                        new SqlParameter("@DO_Type", f_DOCLIGNE.DO_Type),
                        new SqlParameter("@DO_Piece", f_DOCLIGNE.DO_Piece),
                        new SqlParameter("@AR_Ref", f_DOCLIGNE.AR_Ref),
                        new SqlParameter("@DL_Design", f_DOCLIGNE.DL_Design),
                        new SqlParameter("@CT_Num", f_DOCLIGNE.CT_Num),
                        new SqlParameter("@DL_PrixUnitaire", f_DOCLIGNE.DL_PrixUnitaire),
                        new SqlParameter("@DL_QteBL", f_DOCLIGNE.DL_QteBL),
                        new SqlParameter("@DL_QteBC", f_DOCLIGNE.DL_QteBC),
                        new SqlParameter("@DL_QtePL", f_DOCLIGNE.DL_QtePL),
                        new SqlParameter("@DL_QteDE", f_DOCLIGNE.DL_QteDE),
                        new SqlParameter("@DL_Taxe1", f_DOCLIGNE.DL_Taxe1),
                        new SqlParameter("@DL_Qte", f_DOCLIGNE.DL_Qte),
                        new SqlParameter("@EU_Qte", f_DOCLIGNE.EU_Qte),
                        new SqlParameter("@DL_Remise01REM_Valeur", f_DOCLIGNE.DL_Remise01REM_Valeur),
                        new SqlParameter("@DL_PoidsNet", f_DOCLIGNE.DL_PoidsNet),
                        new SqlParameter("@DL_PoidsBrut", f_DOCLIGNE.DL_PoidsBrut),
                        new SqlParameter("@DL_PrixRU", f_DOCLIGNE.DL_PrixRU),
                        new SqlParameter("@DL_MontantHT", f_DOCLIGNE.DL_MontantHT),
                        new SqlParameter("@DL_MontantTTC", f_DOCLIGNE.DL_MontantTTC),                        
                        new SqlParameter("@Retenu", f_DOCLIGNE.Retenu),
                        new SqlParameter("@DL_MontantRegle", f_DOCLIGNE.DL_MontantRegle ?? 0),
                        new SqlParameter("@DL_PieceFourniss", f_DOCLIGNE.DL_PieceFourniss ?? ""),
                        new SqlParameter("@DL_DatePieceFourniss", f_DOCLIGNE.DL_DatePieceFourniss ?? DateTime.Now),
                        new SqlParameter("@cbCreationUser", FrmMdiParent._id_user),
                        new SqlParameter("@DL_No", f_DOCLIGNE.DL_No)
                    );
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }





        public void Delete(F_DOCLIGNE f_DOCLIGNEToDelete)
        {
            string queryDeleteF_DOCLIGNE = @"
                DISABLE TRIGGER [dbo].[TG_CBDEL_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                DISABLE TRIGGER [dbo].[TG_DEL_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];

                DELETE FROM F_DOCLIGNE WHERE DL_No = @DL_No;

                ENABLE TRIGGER [dbo].[TG_CBDEL_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
                ENABLE TRIGGER [dbo].[TG_DEL_F_DOCLIGNE] ON [dbo].[F_DOCLIGNE];
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryDeleteF_DOCLIGNE,
                    new SqlParameter("@DL_No", f_DOCLIGNEToDelete.DL_No)
                );
            }
        }

        public decimal GetMontantRegleByPieceArRef(string dlpiece, string arref, string ctnum)
        {
            dlpiece = dlpiece.Substring(dlpiece.Length - 8);
            string rglibelle = dlpiece + "_" + arref + "_" + ctnum;
            
            return _context.F_CREGLEMENT
                   .Where(r => r.RG_Libelle.Contains(rglibelle))
                   .Sum(r => (decimal?)r.RG_Montant) ?? 0;
        }
    }

}
