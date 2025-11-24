using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    public class F_ARTSTOCKRepository
    {
        private readonly AppDbContext _context;







        public F_ARTSTOCKRepository(AppDbContext context)
        {
            _context = context;
        }








        public F_ARTSTOCK GetF_ARTSTOCK_By_AR_Ref_DP_NoPrincipal(string AR_Ref, int? DP_NoPrincipal)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && artStck.DP_NoPrincipal == DP_NoPrincipal).FirstOrDefault();
            }
        }









        public void Update(F_ARTSTOCK f_ARTSTOCK)
        {
            string queryUpdateStock = @"
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                DISABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];

                UPDATE F_ARTSTOCK
                SET
                    AR_Ref = @AR_Ref,
                    AS_MontSto = @AS_MontSto,
                    AS_QteSto = @AS_QteSto,
                    AS_QteRes = @AS_QteRes,
                    AS_QteCom = @AS_QteCom,
                    AS_QteResCM = @AS_QteResCM,
                    AS_QteComCM = @AS_QteComCM,
                    AS_QtePrepa = @AS_QtePrepa,
                    AS_QteAControler = @AS_QteAControler,
                    cbModification = @cbModification
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                ENABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                queryUpdateStock,
                    new SqlParameter("@AR_Ref", f_ARTSTOCK.AR_Ref),
                    new SqlParameter("@AS_MontSto", f_ARTSTOCK.AS_MontSto),
                    new SqlParameter("@AS_QteSto", f_ARTSTOCK.AS_QteSto),
                    new SqlParameter("@AS_QteRes", f_ARTSTOCK.AS_QteRes),
                    new SqlParameter("@AS_QteCom", f_ARTSTOCK.AS_QteCom),
                    new SqlParameter("@AS_QteResCM", f_ARTSTOCK.AS_QteResCM),
                    new SqlParameter("@AS_QteComCM", f_ARTSTOCK.AS_QteComCM),
                    new SqlParameter("@AS_QtePrepa", f_ARTSTOCK.AS_QtePrepa),
                    new SqlParameter("@AS_QteAControler", f_ARTSTOCK.AS_QteAControler),
                    new SqlParameter("@cbModification", f_ARTSTOCK.cbModification),
                    new SqlParameter("@cbMarq", f_ARTSTOCK.cbMarq)
                );
            }


        }





        public void UpdateMontantEtQuantiteStock(string AR_Ref, int? DP_NoPrincipal, decimal? AS_MontSto, decimal? AS_QteSto)
        {
            int nombreObjetsArtStock = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref).Count();
            F_ARTSTOCK f_ARTSTOCKToUpdate = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && (nombreObjetsArtStock > 1 ? artStck.DP_NoPrincipal == DP_NoPrincipal : true)).FirstOrDefault();

            string queryUpdateMontantEtQuantiteStock = @"
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                DISABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];

                UPDATE F_ARTSTOCK
                SET
                    AS_MontSto = @AS_MontSto,
                    AS_QteSto = @AS_QteSto
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                ENABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryUpdateMontantEtQuantiteStock,
                    new SqlParameter("@AS_MontSto", AS_MontSto),
                    new SqlParameter("@AS_QteSto", AS_QteSto),
                    new SqlParameter("@cbMarq", f_ARTSTOCKToUpdate.cbMarq)
                );
            }


        }





        public void UpdateQuantiteReserve(string AR_Ref, int? DP_NoPrincipal, decimal? AS_QteRes)
        {
            int nombreObjetsArtStock = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref).Count();
            F_ARTSTOCK f_ARTSTOCKToUpdate = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && (nombreObjetsArtStock > 1 ? artStck.DP_NoPrincipal == DP_NoPrincipal : true)).FirstOrDefault();

            string queryUpdateMontantEtQuantiteStock = @"
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                DISABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];

                UPDATE F_ARTSTOCK
                SET
                    AS_QteRes = @AS_QteRes
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                ENABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
            ";



            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryUpdateMontantEtQuantiteStock,
                    new SqlParameter("@AS_QteRes", AS_QteRes),
                    new SqlParameter("@cbMarq", f_ARTSTOCKToUpdate.cbMarq)
                );
            }
        }





        public void UpdateQuantitePrepare(string AR_Ref, int? DP_NoPrincipal, decimal? AS_QtePrepa)
        {
            int nombreObjetsArtStock = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref).Count();
            F_ARTSTOCK f_ARTSTOCKToUpdate = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && (nombreObjetsArtStock > 1 ? artStck.DP_NoPrincipal == DP_NoPrincipal : true)).FirstOrDefault();

            string queryUpdateMontantEtQuantiteStock = @"
                DISABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                DISABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];

                UPDATE F_ARTSTOCK
                SET
                    AS_QtePrepa = @AS_QtePrepa
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [dbo].[TG_CBUPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
                ENABLE TRIGGER [dbo].[TG_UPD_F_ARTSTOCK] ON [dbo].[F_ARTSTOCK];
            ";


            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryUpdateMontantEtQuantiteStock,
                    new SqlParameter("@AS_QtePrepa", AS_QtePrepa),
                    new SqlParameter("@cbMarq", f_ARTSTOCKToUpdate.cbMarq)
                );
            }
        }
    }
}
