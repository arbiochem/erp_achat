using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    internal class F_DOCLIGNEEMPLRepository
    {
        private readonly AppDbContext _context;

        public F_DOCLIGNEEMPLRepository(AppDbContext context)
        {
            _context = context;
        }




        public void InsertDOCLIGNEEMPL(int DP_No, decimal? DL_Qte, string AR_Ref)
        {
            int? DL_No = _context.F_DOCLIGNE.Select(dl => dl.DL_No).Max();

            string queryInsertF_DOCLIGNEEMPL = @"
                DISABLE TRIGGER [TG_INS_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                DISABLE TRIGGER [TG_CBINS_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];

                INSERT INTO [dbo].[F_DOCLIGNEEMPL] (
                	[DL_No]
                    ,[DP_No]
                    ,[DL_Qte]
                    ,[DL_QteAControler]
                    ,[cbProt]
                    ,[cbCreateur]
                    ,[cbModification]
                    ,[cbReplication]
                    ,[cbFlag]
                    ,[cbCreation]
                )
                VALUES (
                	{0}, {1}, {2}, 0, 0, 'COLS', SYSUTCDATETIME(), 0, 0, SYSUTCDATETIME()
                );

                ENABLE TRIGGER [TG_INS_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                ENABLE TRIGGER [TG_CBINS_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
            ";

            if (DP_No == 0)
            {
                DP_No = _context.F_ARTSTOCKEMPL.Where(fdle => fdle.AR_Ref == AR_Ref).Select(fdle => fdle.DP_No).FirstOrDefault();
            }

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryInsertF_DOCLIGNEEMPL,
                    DL_No,
                    DP_No,
                    DL_Qte
                );
            }
        }




        public void UpdateDL_Qte(string typeDocument, string DO_Piece, int? DL_Ligne, int? DL_Qte)
        {
            F_DOCLIGNE f_DOCLIGNE = _context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.DL_Ligne == DL_Ligne).FirstOrDefault();

            if (typeDocument == "Devis" || typeDocument == "Bon d'avoir finanicier" || typeDocument == "Facture d'avoir" || typeDocument == "Bon de commande")
            {
                // Aucun interaction avec l'emplacement des stock pour ces types de documents
            }
            else
            {
                if (typeDocument == "Préparation de livraison" || typeDocument == "Bon de livraison" || typeDocument == "Facture")
                {
                    // Ne rien faire
                }
                else // else if (typeDocument == "Facture de retour" || typeDocument == "Bon de retour")
                {
                    DL_Qte = -DL_Qte;
                }

                string queryUpdateF_DOCLIGNEEMPL = @"
                    DISABLE TRIGGER [TG_UPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                    DISABLE TRIGGER [TG_CBUPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                    DISABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];

                    UPDATE F_DOCLIGNEEMPL
                    SET
                    	DL_Qte = @DL_Qte
                    WHERE DL_No = @DL_No;

                    ENABLE TRIGGER [TG_UPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                    ENABLE TRIGGER [TG_CBUPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                    ENABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];
                ";

                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        queryUpdateF_DOCLIGNEEMPL,
                        new SqlParameter("@DL_Qte", DL_Qte),
                        new SqlParameter("@DL_No", f_DOCLIGNE.DL_No)
                    );
                }
            }
        }



        public void DeleteF_DOCLIGNEEMPL(string DO_Piece, int? DL_Ligne)
        {
            F_DOCLIGNE f_DOCLIGNE = _context.F_DOCLIGNE.Where(dl => dl.DO_Piece == DO_Piece && dl.DL_Ligne == DL_Ligne).FirstOrDefault();

            string queryDeleteF_DOCLIGNEEMPL = @"
                DISABLE TRIGGER [TG_UPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                DISABLE TRIGGER [TG_CBUPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                DISABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];

                DELETE FROM [dbo].[F_DOCLIGNEEMPL] WHERE DL_No = @DL_No;

                ENABLE TRIGGER [TG_UPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                ENABLE TRIGGER [TG_CBUPD_F_DOCLIGNEEMPL] ON [dbo].[F_DOCLIGNEEMPL];
                ENABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];
            ";

            if (f_DOCLIGNE != null)
            {
                using (var context = new AppDbContext())
                {
                    context.Database.ExecuteSqlCommand(
                        queryDeleteF_DOCLIGNEEMPL,
                        new SqlParameter("@DL_No", f_DOCLIGNE.DL_No)
                    );
                }
            }
        }

        

    }
}
