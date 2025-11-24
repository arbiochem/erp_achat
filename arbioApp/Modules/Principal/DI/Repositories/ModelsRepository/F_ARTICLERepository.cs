using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    public class F_ARTICLERepository
    {
        private readonly AppDbContext _context;





        public F_ARTICLERepository(AppDbContext context)
        {
            _context = context;
        }





        public List<DTO.Article> GetAll()
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTICLE
                .Select(a => new DTO.Article
                {
                    AR_Ref = a.AR_Ref,
                    AR_Design = a.AR_Design,
                    FA_CodeFamille = a.FA_CodeFamille,
                    AR_UnitePoids = a.AR_UnitePoids,
                    AR_PrixVen = a.AR_PrixVen,
                    AR_PrixTTC = a.AR_PrixTTC,
                    AR_UniteVen = a.AR_UniteVen
                }).ToList();
            }

        }





        public List<F_ARTICLE> GetAllF_ARTICLE()
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTICLE.ToList();
            }
        }




        public F_ARTICLE GetF_ARTICLEByAR_Ref(string AR_Ref)
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTICLE.FirstOrDefault(a => a.AR_Ref == AR_Ref);
            }
        }



        public F_ARTICLE GetF_ARTICLEByAR_Design(string AR_Design)
        {
            using (var context = new AppDbContext())
            {
                return context.F_ARTICLE.FirstOrDefault(a => a.AR_Design == AR_Design);
            }
        }

        List<F_ARTICLE> _listeArticles;
        public void UpdateDateModifArticle(string arRef, decimal prixAch, decimal puNet, short uniteVen)
        {

            using (var context = new AppDbContext())
            {
                var article = context.F_ARTICLE
                    .Include(a => a.artStock) // Propriété de navigation définie dans ton modèle
                    .FirstOrDefault(a => a.AR_Ref == arRef);


                if (article != null)
                {
                    article.AR_PrixAch = prixAch;
                    article.AR_PUNet = puNet;
                    article.AR_UniteVen = uniteVen;
                    article.AR_DateModif = DateTime.Now;

                    if (article.artStock != null)
                    {
                        //article.artStock.AS_QteMini = qteMini;
                        //article.artStock.AS_QteMaxi = qteMaxi;
                    }

                    _context.SaveChanges();
                }
            }

            
       


        //var article = _context.F_ARTICLE.FirstOrDefault(a => a.AR_Ref == arRef);
        //if (article != null)
        //{
        //    article.AR_PrixAch = prixAch;
        //    article.AR_PUNet = puNet;
        //    article.AR_UniteVen = uniteVen;
        //    article.AR_DateModif = DateTime.Now;

        //    _context.SaveChanges();
        //}
        //else
        //{
        //    throw new Exception($"Article {arRef} introuvable !");
        //}
        //using (var context = new AppDbContext())
        //{
        //    string queryUpdateDateModifArticle = @"
        //        BEGIN TRANSACTION;

        //        BEGIN TRY                    
        //            -- Désactivation des triggers pour la table F_ARTICLE
        //            DISABLE TRIGGER [TG_CBUPD_F_ARTICLE] ON [dbo].[F_ARTICLE];
        //            DISABLE TRIGGER [TG_UPD_F_ARTICLE] ON [dbo].[F_ARTICLE];

        //            -- Mise à jour de l'article avec le cbMarq spécifié
        //            UPDATE [dbo].[F_ARTICLE] 
        //            SET AR_DateModif = GETDATE(), 
        //                cbCreateur = 'COLS' 
        //            WHERE cbMarq = @cbMarq;

        //            -- Réactivation des triggers après la mise à jour
        //            ENABLE TRIGGER [TG_CBUPD_F_ARTICLE] ON [dbo].[F_ARTICLE];
        //            ENABLE TRIGGER [TG_UPD_F_ARTICLE] ON [dbo].[F_ARTICLE];

        //            -- Validation de la transaction
        //            COMMIT TRANSACTION;
        //        END TRY
        //        BEGIN CATCH
        //            -- Annulation de la transaction en cas d'erreur
        //            ROLLBACK TRANSACTION;

        //            -- Réactivation des triggers même en cas d'erreur
        //            ENABLE TRIGGER [TG_CBUPD_F_ARTICLE] ON [dbo].[F_ARTICLE];
        //            ENABLE TRIGGER [TG_UPD_F_ARTICLE] ON [dbo].[F_ARTICLE];

        //            -- Récupération et remontée du message d'erreur
        //            DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        //            RAISERROR(@ErrorMessage, 16, 1);
        //        END CATCH;
        //    ";

        //    context.Database.ExecuteSqlCommand(
        //        queryUpdateDateModifArticle,
        //        new SqlParameter("@cbMarq", cbMarq)
        //    );
        //}
    }


    }
}
