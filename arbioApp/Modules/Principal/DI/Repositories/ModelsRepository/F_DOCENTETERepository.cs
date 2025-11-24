using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace arbioApp.Repositories.ModelsRepository
{
    internal class F_DOCENTETERepository
    {
		// ==============================================================================
		// DEBUT DECLARATION DES VARIABLES ==============================================
		// ==============================================================================
        private readonly AppDbContext _context;
        // ==============================================================================
        // FIN DECLARATION DES VARIABLES ================================================
        // ==============================================================================










        // ==============================================================================
        // DEBUT CONSTRUCTEUR ===========================================================
        // ==============================================================================
        public F_DOCENTETERepository(AppDbContext context)
        {
            _context = context;
        }
		// ==============================================================================
		// FIN CONSTRUCTEUR =============================================================
		// ==============================================================================










		// ==============================================================================
		// DEBUT GET ====================================================================
		// ==============================================================================
		public List<F_DOCENTETE> GetAll()
		{
			using (AppDbContext context = new AppDbContext())
			{
				return context.F_DOCENTETE.ToList();
			}
		}



        //      public F_DOCENTETE GetBy_DO_Piece(string DO_Piece)
        //{
        //	using (AppDbContext context = new AppDbContext())
        //	{
        //		return context.F_DOCENTETE.Where(doc => doc.DO_Piece == DO_Piece).FirstOrDefault();
        //          }
        //}

        public F_DOCENTETE GetBy_DO_Piece_And_Type(string DO_Piece)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_DOCENTETE
                    .FirstOrDefault(doc => doc.DO_Piece == DO_Piece);
            }
        }
        public F_DOCENTETE GetBy_CbMarq(int cbMarqSource)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_DOCENTETE
                    .FirstOrDefault(doc => doc.cbMarq == cbMarqSource);
            }
        }

        // ==============================================================================
        // FIN GET ======================================================================
        // ==============================================================================










        // ==============================================================================
        // DEBUT INSERT =================================================================
        // ==============================================================================
        public void Add(F_DOCENTETE f_DOCENTETE)
        {
			try
			{
				//------------------------------------------------------------TEST TYPE
                var parameters = new List<SqlParameter>
					{
						
						new SqlParameter("@DO_Type", f_DOCENTETE.DO_Type),
                        new SqlParameter("@DO_Piece", f_DOCENTETE.DO_Piece),
                        new SqlParameter("@DO_Ref", f_DOCENTETE.DO_Ref),
                        new SqlParameter("@DO_Tiers", f_DOCENTETE.DO_Tiers),
                        new SqlParameter("@CO_No", f_DOCENTETE.CO_No),
                        new SqlParameter("@cbCO_No ", f_DOCENTETE.cbCO_No ?? (object)DBNull.Value),
                        new SqlParameter("@DO_Period", 1),
                        new SqlParameter("@DO_Devise", f_DOCENTETE.DO_Devise),
                        new SqlParameter("@DO_Cours", f_DOCENTETE.DO_Cours),
                        new SqlParameter("@DE_No", f_DOCENTETE.DE_No),
                        new SqlParameter("@cbDE_No", f_DOCENTETE.DE_No),
                        new SqlParameter("@CT_NumPayeur", f_DOCENTETE.CT_NumPayeur),
                        new SqlParameter("@DO_Expedit", f_DOCENTETE.DO_Expedit),
                        new SqlParameter("@DO_NbFacture", f_DOCENTETE.DO_NbFacture),
                        new SqlParameter("@DO_BLFact", f_DOCENTETE.DO_BLFact),
                        new SqlParameter("@DO_TxEscompte", f_DOCENTETE.DO_TxEscompte),
                        new SqlParameter("@CA_Num", f_DOCENTETE.CA_Num),
                        new SqlParameter("@DO_Coord01", f_DOCENTETE.DO_Coord01),
                        new SqlParameter("@DO_DateLivr", SqlDbType.DateTime) { Value = f_DOCENTETE.DO_DateLivr },
                        new SqlParameter("@DO_Condition", 1),
                        new SqlParameter("@DO_Tarif", 1),
                        new SqlParameter("@DO_Transaction", 11),
                        new SqlParameter("@DO_Langue", f_DOCENTETE.DO_Langue),
                        new SqlParameter("@DO_Regime", 11),
                        new SqlParameter("@N_CatCompta", 1),
                        new SqlParameter("@CG_Num", f_DOCENTETE.CG_Num),
                        new SqlParameter("@DO_Heure", f_DOCENTETE.DO_Heure),
                        new SqlParameter("@CA_No", f_DOCENTETE.CA_No),
                        new SqlParameter("@cbCA_No ", f_DOCENTETE.cbCA_No ?? (object)DBNull.Value),
                        new SqlParameter("@CO_NoCaissier", f_DOCENTETE.CO_NoCaissier),
                        new SqlParameter("@cbCO_NoCaissier ", f_DOCENTETE.cbCO_NoCaissier ?? (object)DBNull.Value),
                        new SqlParameter("@CA_NumIFRS", f_DOCENTETE.CA_NumIFRS),
                        new SqlParameter("@DO_TypeFrais", f_DOCENTETE.DO_TypeFrais),
                        new SqlParameter("@DO_ValFrais", f_DOCENTETE.DO_ValFrais),
                        new SqlParameter("@DO_TypeLigneFrais", f_DOCENTETE.DO_TypeLigneFrais),
                        new SqlParameter("@DO_TypeFranco", f_DOCENTETE.DO_TypeFranco),
                        new SqlParameter("@DO_ValFranco", f_DOCENTETE.DO_ValFranco),
                        new SqlParameter("@DO_TypeLigneFranco", f_DOCENTETE.DO_TypeLigneFranco),
                        new SqlParameter("@DO_Taxe1", f_DOCENTETE.DO_Taxe1),
                        new SqlParameter("@DO_DateLivrRealisee", SqlDbType.DateTime) { Value = f_DOCENTETE.DO_DateLivrRealisee },
                        new SqlParameter("@cbCreationUser", f_DOCENTETE.cbCreationUser)
                    };


                foreach (var p in parameters)
                {
                    if (p.SqlDbType == SqlDbType.DateTime)
                    {
                        if (p.Value == null || p.Value == DBNull.Value) continue;

                        if (p.Value is DateTime dt)
                        {
                            if (dt < (DateTime)System.Data.SqlTypes.SqlDateTime.MinValue || dt > (DateTime)System.Data.SqlTypes.SqlDateTime.MaxValue)
                            {
                                MessageBox.Show($"Paramètre {p.ParameterName} a une date hors limites : {dt}");
                                return; // stop avant insertion
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Paramètre {p.ParameterName} n'est pas un DateTime valide !");
                            return;
                        }
                    }
                }

				//--------------------------------------------------------------------------------------------------------------------------

                string query = @"
				DISABLE TRIGGER [dbo].[TG_CBINS_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				DISABLE TRIGGER [dbo].[TG_INS_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				DISABLE TRIGGER [dbo].[TG_INS_CPTAF_DOCENTETE] ON [dbo].[F_DOCENTETE];

                Insert INTO [dbo].[F_DOCENTETE] (
							DO_Type,
							DO_Piece,
							DO_Ref,
							DO_Tiers,
							CO_No,
							cbCO_No,
							DO_Period,
							DO_Devise,
							DO_Cours,
							DE_No,
							cbDE_No,
							CT_NumPayeur,
							DO_Expedit,
							DO_NbFacture,
							DO_BLFact,
							DO_TxEscompte,
							CA_Num,
							DO_Coord01,
							DO_DateLivr,
							DO_Condition,
							DO_Tarif,
							DO_Transaction,
							DO_Langue,
							DO_Regime,
							N_CatCompta,
							CG_Num,
							DO_Heure,
							CA_No,
							cbCA_No,
							CO_NoCaissier,
							cbCO_NoCaissier,
							CA_NumIFRS,
							DO_TypeFrais,
							DO_ValFrais,
							DO_TypeLigneFrais,
							DO_TypeFranco,
							DO_ValFranco,
							DO_TypeLigneFranco,
							DO_Taxe1,
							DO_DateLivrRealisee,
							--DO_CodeTaxe1,
							DO_Domaine,
							DO_Date,
							DO_Reliquat,
							DO_Imprim,
							DO_Coord02,
							DO_Coord03,
							DO_Coord04,
							DO_Souche,
							DO_Colisage,
							DO_TypeColis,
							DO_Ecart,
							DO_Ventile,
							AB_No,
							DO_DebutAbo,
							DO_FinAbo,
							DO_DebutPeriod,
							DO_FinPeriod,
							DO_Statut,
							DO_Transfere,
							DO_Cloture,
							DO_NoWeb,
							DO_Attente,
							DO_Provenance,
							MR_No,
							DO_TypeTaux1,
							DO_TypeTaxe1,
							DO_Taxe2,
							DO_TypeTaux2,
							DO_TypeTaxe2,
							DO_Taxe3,
							DO_TypeTaux3,
							DO_TypeTaxe3,
							DO_MajCpta,
							DO_Motif,
							CT_NumCentrale,
							DO_Contact,
							DO_FactureElec,
							DO_TypeTransac,
							DO_DateExpedition,
							DO_FactureFrs,
							DO_PieceOrig,
							DO_GUID,
							DO_EStatut,
							DO_DemandeRegul,
							ET_No,
							cbET_No,
							DO_Valide,
							DO_Coffre,
							DO_CodeTaxe2,
							DO_CodeTaxe3,
							DO_TotalHT,
							DO_StatutBAP,
							DO_Escompte,
							DO_DocType,
							DO_TypeCalcul,
							DO_FactureFile,
							DO_TotalHTNet,
							DO_TotalTTC,
							DO_NetAPayer,
							DO_MontantRegle,
							DO_RefPaiement,
							DO_AdressePaiement,
							DO_PaiementLigne,
							DO_MotifDevis,
							cbCreateur,
							cbCreationUser
				)
				values(
					@DO_Type,
					@DO_Piece,
					@DO_Ref,
					@DO_Tiers,
					@CO_No,
					@cbCO_No,
					@DO_Period,
					@DO_Devise,
					@DO_Cours,
					@DE_No,
					@cbDE_No,
					@CT_NumPayeur,
					@DO_Expedit,
					@DO_NbFacture,
					@DO_BLFact,
					@DO_TxEscompte,
					@CA_Num,
					@DO_Coord01,
					@DO_DateLivr,
					@DO_Condition,
					@DO_Tarif,
					@DO_Transaction,
					@DO_Langue,
					@DO_Regime,
					@N_CatCompta,
					@CG_Num,
					@DO_Heure,
					@CA_No,
					@cbCA_No,
					@CO_NoCaissier,
					@cbCO_NoCaissier,
					@CA_NumIFRS,
					@DO_TypeFrais,
					@DO_ValFrais,
					@DO_TypeLigneFrais,
					@DO_TypeFranco,
					@DO_ValFranco,
					@DO_TypeLigneFranco,
					@DO_Taxe1,
					@DO_DateLivrRealisee,
					--@DO_CodeTaxe1,
					1,
					CONVERT([datetime2](0),getdate()),
					0,
					0,
					'',
					'',
					'',
					0,
					1,
					1,
					0,
					0,
					0,
					'1753-01-01 00:00:00',
					'1753-01-01 00:00:00',
					'1753-01-01 00:00:00',
					'1753-01-01 00:00:00',
					0, --DO_Statut pour nouveau projet
					0,
					0,
					'',
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					0,
					'',
					NULL,
					0,
					0,
					0,
					'1753-01-01 00:00:00',
					'',
					'',
					NULL,
					0,
					0,
					0,
					NULL,
					0,
					0,
					NULL,
					NULL,
					0,
					0,
					1,
					10,
					0,
					NULL,
					0,
					0,
					0,
					0,
					NULL,
					'',
					0,
					0,
					'COLS',
					@cbCreationUser

				);

				ENABLE TRIGGER [dbo].[TG_CBINS_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				ENABLE TRIGGER [dbo].[TG_INS_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				ENABLE TRIGGER [dbo].[TG_INS_CPTAF_DOCENTETE] ON [dbo].[F_DOCENTETE];
			";

                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        query,
                        new SqlParameter("@DO_Type", f_DOCENTETE.DO_Type),
						new SqlParameter("@DO_Piece", f_DOCENTETE.DO_Piece),
						new SqlParameter("@DO_Ref", f_DOCENTETE.DO_Ref),
						new SqlParameter("@DO_Tiers", f_DOCENTETE.DO_Tiers),
						new SqlParameter("@CO_No", f_DOCENTETE.CO_No),
						new SqlParameter("@cbCO_No ", f_DOCENTETE.cbCO_No ?? (object)DBNull.Value),
						new SqlParameter("@DO_Period", f_DOCENTETE.DO_Period),
						new SqlParameter("@DO_Devise", f_DOCENTETE.DO_Devise),
						new SqlParameter("@DO_Cours", f_DOCENTETE.DO_Cours),
						new SqlParameter("@DE_No", f_DOCENTETE.DE_No),
						new SqlParameter("@cbDE_No", f_DOCENTETE.DE_No),
						new SqlParameter("@CT_NumPayeur", f_DOCENTETE.CT_NumPayeur),
						new SqlParameter("@DO_Expedit", f_DOCENTETE.DO_Expedit),
						new SqlParameter("@DO_NbFacture", f_DOCENTETE.DO_NbFacture),
						new SqlParameter("@DO_BLFact", f_DOCENTETE.DO_BLFact),
						new SqlParameter("@DO_TxEscompte", f_DOCENTETE.DO_TxEscompte),
						new SqlParameter("@CA_Num", f_DOCENTETE.CA_Num),
						new SqlParameter("@DO_Coord01", f_DOCENTETE.DO_Coord01),
						new SqlParameter("@DO_DateLivr", SqlDbType.DateTime) { Value = f_DOCENTETE.DO_DateLivr },
						new SqlParameter("@DO_Condition", f_DOCENTETE.DO_Condition),
						new SqlParameter("@DO_Tarif", f_DOCENTETE.DO_Tarif),
						new SqlParameter("@DO_Transaction", f_DOCENTETE.DO_Transaction),
						new SqlParameter("@DO_Langue", f_DOCENTETE.DO_Langue),
						new SqlParameter("@DO_Regime", f_DOCENTETE.DO_Regime),
						new SqlParameter("@N_CatCompta", f_DOCENTETE.N_CatCompta),
						new SqlParameter("@CG_Num", f_DOCENTETE.CG_Num),
						new SqlParameter("@DO_Heure", f_DOCENTETE.DO_Heure),
						new SqlParameter("@CA_No", f_DOCENTETE.CA_No),
						new SqlParameter("@cbCA_No ", f_DOCENTETE.cbCA_No ?? (object)DBNull.Value),
						new SqlParameter("@CO_NoCaissier", f_DOCENTETE.CO_NoCaissier),
						new SqlParameter("@cbCO_NoCaissier ", f_DOCENTETE.cbCO_NoCaissier ?? (object)DBNull.Value),
						new SqlParameter("@CA_NumIFRS", f_DOCENTETE.CA_NumIFRS),
						new SqlParameter("@DO_TypeFrais", f_DOCENTETE.DO_TypeFrais),
						new SqlParameter("@DO_ValFrais", f_DOCENTETE.DO_ValFrais),
						new SqlParameter("@DO_TypeLigneFrais", f_DOCENTETE.DO_TypeLigneFrais),
						new SqlParameter("@DO_TypeFranco", f_DOCENTETE.DO_TypeFranco),
						new SqlParameter("@DO_ValFranco", f_DOCENTETE.DO_ValFranco),
						new SqlParameter("@DO_TypeLigneFranco", f_DOCENTETE.DO_TypeLigneFranco),
						new SqlParameter("@DO_Taxe1", f_DOCENTETE.DO_Taxe1),
						new SqlParameter("@DO_DateLivrRealisee", SqlDbType.DateTime) { Value = f_DOCENTETE.DO_DateLivrRealisee },
						new SqlParameter("@cbCreationUser", f_DOCENTETE.cbCreationUser)

                    );
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        // ==============================================================================
        // FIN INSERT ===================================================================
        // ==============================================================================










        // ==============================================================================
        // DEBUT UPDATE =================================================================
        // ==============================================================================

        public void UpdateDO_Totaux_HT_Net_TTC_Repo(string DO_Piece, decimal? DO_TotalHT, decimal? DO_TotalHTNet, decimal? DO_TotalTTC)
        {
			try {
                string query = @"		
				DISABLE TRIGGER [dbo].[TG_CBUPD_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				DISABLE TRIGGER [dbo].[TG_UPD_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				DISABLE TRIGGER [dbo].[TG_UPD_CPTAF_DOCENTETE] ON [dbo].[F_DOCENTETE];
				
				UPDATE F_DOCENTETE 
				SET 
					DO_TotalHT = @DO_TotalHT,
					DO_TotalHTNet = @DO_TotalHTNet,
					DO_TotalTTC = @DO_TotalTTC,
					DO_NetAPayer = @DO_TotalTTC
				
				WHERE DO_Piece LIKE @DO_Piece;
				
				ENABLE TRIGGER[dbo].[TG_CBUPD_F_DOCENTETE] ON[dbo].[F_DOCENTETE];
				ENABLE TRIGGER [dbo].[TG_UPD_F_DOCENTETE] ON [dbo].[F_DOCENTETE];
				ENABLE TRIGGER [dbo].[TG_UPD_CPTAF_DOCENTETE] ON [dbo].[F_DOCENTETE];
			";

                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        query,
                        new SqlParameter("@DO_TotalHT", DO_TotalHT),
                        new SqlParameter("@DO_TotalHTNet", DO_TotalHTNet),
                        new SqlParameter("@DO_TotalTTC", DO_TotalTTC),
                        new SqlParameter("@DO_Piece", DO_Piece)
                    );
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }           
        }

        public void UpdateDO_MontantRegle(decimal RC_Montant, string DO_Piece)
		{
            string queryFDocEntete = @"
				DISABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;

				UPDATE F_DOCENTETE SET DO_MontantRegle = @DO_MontantRegle WHERE DO_Piece = @DO_Piece;

				ENABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;
            ";

            using (var context = new AppDbContext())
			{
				context.Database.ExecuteSqlCommand(
					queryFDocEntete,
					new SqlParameter("@DO_MontantRegle", RC_Montant),
					new SqlParameter("@DO_Piece", DO_Piece)
				);
            }
               
        }





		public void UpdateProprietesF_DOCENTETE(F_DOCENTETE f_DOCENTETEToUpdate)
		{
			try {
                string queryUpdateProprietesF_DOCENTETE = @"
				DISABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;

				UPDATE F_DOCENTETE
				SET					
					DO_Type = @DO_Type,
					DO_Piece = @DO_Piece,
					DO_Tiers = @DO_Tiers,
					DO_Date = @DO_Date,
					DO_DateLivr = @DO_DateLivr,
					DO_DateLivrRealisee = @DO_DateLivrRealisee,
					DO_Ref = @DO_Ref,
					CA_Num = @CA_Num,
					CO_No = @CO_No,
					cbCO_No = @cbCO_No,
					DO_Statut = @DO_Statut,
					CT_NumPayeur = @CT_NumPayeur,
					DO_Expedit = @DO_Expedit,
					DO_ValFranco = @DO_ValFranco,
					DO_ValFrais = @DO_ValFrais,
					DO_Coord01 = @DO_Coord01,
					--Commentaires = @Commentaires,
					--Divers = @Divers,
					DE_No = @DE_No
				--WHERE DO_Piece = @DO_Piece;
					WHERE cbMarq = @cbMarq;
				ENABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;
			";


                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        queryUpdateProprietesF_DOCENTETE,
                        new SqlParameter("@cbMarq", f_DOCENTETEToUpdate.cbMarq),
                        new SqlParameter("@DO_Type", f_DOCENTETEToUpdate.DO_Type),
                        new SqlParameter("@DO_Tiers", f_DOCENTETEToUpdate.DO_Tiers),
                        new SqlParameter("@DO_Date", f_DOCENTETEToUpdate.DO_Date),
                        new SqlParameter("@DO_DateLivr", f_DOCENTETEToUpdate.DO_DateLivr),
                        new SqlParameter("@DO_DateLivrRealisee", f_DOCENTETEToUpdate.DO_DateLivrRealisee),
                        new SqlParameter("@DO_Ref", f_DOCENTETEToUpdate.DO_Ref),
                        new SqlParameter("@CA_Num", f_DOCENTETEToUpdate.CA_Num),
                        new SqlParameter("@CT_NumPayeur", f_DOCENTETEToUpdate.CT_NumPayeur),
                        new SqlParameter("@CO_No", f_DOCENTETEToUpdate.CO_No),
                        new SqlParameter("@cbCO_No", f_DOCENTETEToUpdate.cbCO_No),
                        new SqlParameter("@DO_Statut", f_DOCENTETEToUpdate.DO_Statut),
                        new SqlParameter("@DO_Expedit", f_DOCENTETEToUpdate.DO_Expedit),
                        new SqlParameter("@DO_ValFranco", f_DOCENTETEToUpdate.DO_ValFranco),
                        new SqlParameter("@DO_ValFrais", f_DOCENTETEToUpdate.DO_ValFrais),
                        new SqlParameter("@DO_Coord01", f_DOCENTETEToUpdate.DO_Coord01),
                        new SqlParameter("@DE_No", f_DOCENTETEToUpdate.DE_No),
                        new SqlParameter("@DO_Piece", f_DOCENTETEToUpdate.DO_Piece)
                    );
                }
            }
            catch (System.Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }





		public void ValiderDocument(string DO_Piece, short? DO_Valide)
		{
            string queryValiderDocument = @"
				DISABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;

				UPDATE F_DOCENTETE
				SET
					DO_Valide = @DO_Valide
				WHERE DO_Piece = @DO_Piece;

				ENABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;
			";

			using(var context = new AppDbContext())
			{
                context.Database.ExecuteSqlCommand(
				    queryValiderDocument,
				    new SqlParameter("@DO_Valide", DO_Valide),
				    new SqlParameter("@DO_Piece", DO_Piece)
				);
            }
        }




		public void Update_DO_Imprim_F_DOCENTETE(string DO_Piece, short? DO_Imprim)
		{
			string queryImprimerDocument = @"
				DISABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;

				UPDATE F_DOCENTETE
				SET
					DO_Imprim = @DO_Imprim
				WHERE DO_Piece = @DO_Piece;

				ENABLE TRIGGER TG_CBUPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_UPD_CPTAF_DOCENTETE ON F_DOCENTETE;
			";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryImprimerDocument,
                    new SqlParameter("@DO_Imprim", DO_Imprim),
                    new SqlParameter("@DO_Piece", DO_Piece)
                );
            }
        }

        // ==============================================================================
        // FIN UPDATE ===================================================================
        // ==============================================================================










        // ==============================================================================
        // DEBUT DELETE =================================================================
        // ==============================================================================
        public void Delete(string DO_Piece)
		{
            string query = @"
				DISABLE TRIGGER TG_CBDEL_F_DOCENTETE ON F_DOCENTETE;
				DISABLE TRIGGER TG_DEL_F_DOCENTETE ON F_DOCENTETE;

				DELETE FROM F_DOCENTETE WHERE DO_Piece = @DO_Piece;

				ENABLE TRIGGER TG_CBDEL_F_DOCENTETE ON F_DOCENTETE;
				ENABLE TRIGGER TG_DEL_F_DOCENTETE ON F_DOCENTETE;
			";

			using(var context = new AppDbContext())
			{
                context.Database.ExecuteSqlCommand(query, new SqlParameter("@DO_Piece", DO_Piece));
            }
        }
        // ==============================================================================
        // FIN DELETE ===================================================================
        // ==============================================================================
    }
}
