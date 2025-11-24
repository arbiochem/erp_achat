using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace arbioApp.Repositories.ModelsRepository
{
    public class F_ARTGAMMERepository
    {
        // =======================================================================================================================================
        // =================================================== DEBUT DECLARATION DES VARIABLES ===================================================
        // =======================================================================================================================================
        //private readonly AppDbContext _context;
        // =====================================================================================================================================
        // =================================================== FIN DECLARATION DES VARIABLES ===================================================
        // =====================================================================================================================================






        // ====================================================================================================================================
        // ======================================================== DEBUT CONSTRUCTEUR ========================================================
        // ====================================================================================================================================
        public F_ARTGAMMERepository(AppDbContext context)
        {
            //_context = context;
        }
        // =====================================================================================================================================
        // =================================================== FIN DECLARATION DES VARIABLES ===================================================
        // =====================================================================================================================================





        // ======================================================================================================================================
        // =================================================== DEBUT DECLARATION DES METHODES ===================================================
        // ======================================================================================================================================
        public F_ARTGAMME GetByEG_Enumere(string EG_Enumere)
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTGAMME.Where(eg => eg.EG_Enumere == EG_Enumere).FirstOrDefault();
            }
        }



        public F_ARTGAMME GetByAG_No(int? AG_No)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_ARTGAMME.Where(ag => ag.AG_No == AG_No).FirstOrDefault();
            }
        }


        public int? GetLastAG_No1()
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTGAMME.Where(ag => ag.AG_Type == 0).OrderByDescending(ag => ag.AG_No).Select(ag => ag.AG_No).FirstOrDefault();
            }
        }




        public int? GetLastAG_No2()
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTGAMME.Where(ag => ag.AG_Type == 1).OrderByDescending(ag => ag.AG_No).Select(ag => ag.AG_No).FirstOrDefault();
            }
        }





        public void Create(F_ARTGAMME f_ARTGAMME)
        {
            string queryCreateF_ARTGAMME = @"
                BEGIN TRANSACTION;

                BEGIN TRY
                    DECLARE @Next_AG_No INT;
                
                    -- Désactivation des triggers
                    DISABLE TRIGGER [TG_CBINS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                    DISABLE TRIGGER [TG_INS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                
                    -- Récupération du prochain AG_No
                    SELECT @Next_AG_No = ISNULL(MAX(AG_No), 0) + 1
                    FROM [dbo].[F_ARTGAMME]
                
                    -- Insertion des données
                    INSERT INTO [dbo].[F_ARTGAMME] (
                        AR_Ref, AG_No, EG_Enumere, AG_Type, cbCreateur, cbCreationUser
                    )
                    VALUES (
                        @AR_Ref, @Next_AG_No, @EG_Enumere, @AG_Type, 'COLS', NULL
                    );
                
                    -- Réactivation des triggers
                    ENABLE TRIGGER [TG_CBINS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                    ENABLE TRIGGER [TG_INS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                
                    COMMIT TRANSACTION;
                END TRY
                BEGIN CATCH
                    ROLLBACK TRANSACTION;
                
                    -- Réactivation des triggers en cas d'erreur
                    ENABLE TRIGGER [TG_CBINS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                    ENABLE TRIGGER [TG_INS_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                
                    -- Affichage de l'erreur pour débogage
                    DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
                    RAISERROR(@ErrorMessage, 16, 1);
                END CATCH;
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryCreateF_ARTGAMME,
                    new SqlParameter("@AR_Ref", f_ARTGAMME.AR_Ref),
                    new SqlParameter("@EG_Enumere", f_ARTGAMME.EG_Enumere),
                    new SqlParameter("@AG_Type", f_ARTGAMME.AG_Type)
                );
            }

        }





        public void UpdateEG_Enumere(int cbMarq, string EG_Enumere)
        {
            string queryUpdateF_ARTGAMME = @"
                DISABLE TRIGGER [dbo].[TG_UPD_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                
                UPDATE F_ARTGAMME
                SET
                    EG_Enumere = @EG_Enumere
                WHERE cbMarq = @cbMarq;
                
                ENABLE TRIGGER [dbo].[TG_UPD_F_ARTGAMME] ON [dbo].[F_ARTGAMME];
                ENABLE TRIGGER [dbo].[TG_CBUPD_F_ARTGAMME] ON [dbo].[F_ARTGAMME]
            ";

            using (AppDbContext context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                queryUpdateF_ARTGAMME,
                    new SqlParameter("@EG_Enumere", EG_Enumere),
                    new SqlParameter("@cbMarq", cbMarq)
                );
            }       
        }





        public void Delete(int cbMarq)
        {
            string queryDeleteDansF_ARTGAMME = @"
                IF EXISTS (SELECT 1 FROM [dbo].[F_ARTGAMME] WHERE cbMarq = @cbMarq)
                BEGIN
                    BEGIN TRANSACTION; -- Début de la transaction
                
                    BEGIN TRY
                        -- Désactiver les triggers
                        DISABLE TRIGGER ALL ON [dbo].[F_ARTGAMME];
                
                        -- Suppression des enregistrements
                        DELETE FROM [dbo].[F_ARTGAMME] 
                        WHERE cbMarq = @cbMarq;
                
                        -- Valider la transaction si tout s'est bien passé
                        COMMIT TRANSACTION;
                
                        -- Réactiver les triggers après la validation de la transaction
                        ENABLE TRIGGER ALL ON [dbo].[F_ARTGAMME];
                    END TRY
                    BEGIN CATCH
                        -- Annuler la transaction en cas d'erreur
                        IF @@TRANCOUNT > 0
                            ROLLBACK TRANSACTION;
                
                        -- Réactiver les triggers même en cas d'échec
                        ENABLE TRIGGER ALL ON [dbo].[F_ARTGAMME];
                
                        -- Afficher l'erreur
                        PRINT 'Erreur lors de la suppression de F_ARTGAMME. Transaction annulée.';
                        PRINT ERROR_MESSAGE();
                    END CATCH
                END
                ELSE
                BEGIN
                    PRINT 'Aucun enregistrement trouvé avec cbMarq = ' + CAST(@cbMarq AS VARCHAR);
                END
            ";

            using (AppDbContext context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryDeleteDansF_ARTGAMME,
                    new SqlParameter("@cbMarq", cbMarq)
                );
            }
        }
        // ====================================================================================================================================
        // =================================================== FIN DECLARATION DES METHODES ===================================================
        // ====================================================================================================================================


    }
}
