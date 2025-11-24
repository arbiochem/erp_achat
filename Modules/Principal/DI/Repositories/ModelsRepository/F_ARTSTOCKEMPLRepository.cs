using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    public class F_ARTSTOCKEMPLRepository
    {
        //private readonly AppDbContext _context;

        public F_ARTSTOCKEMPLRepository(AppDbContext context)
        {
            //_context = context;
        }



        public void UpdateAE_QteSto(decimal? AE_QteSto, int cbMarq)
        {
            string queryUpdateF_ARTSTOCKEMPLAE_QteSto = @"
                DISABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];

                UPDATE F_ARTSTOCKEMPL
                SET
                	AE_QteSto = @AE_QteSto
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryUpdateF_ARTSTOCKEMPLAE_QteSto,
                    new SqlParameter("@AE_QteSto", AE_QteSto),
                    new SqlParameter("@cbMarq", cbMarq)
                );
            }


        }




        public void UpdateAE_QtePrepa(decimal? AE_QtePrepa, int cbMarq)
        {
            string queryUpdateF_ARTSTOCKEMPLAE_QtePrepa = @"
                DISABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];

                UPDATE F_ARTSTOCKEMPL
                SET
                	AE_QtePrepa = @AE_QtePrepa
                WHERE cbMarq = @cbMarq;

                ENABLE TRIGGER [TG_CBUPD_F_ARTSTOCKEMPL] ON [dbo].[F_ARTSTOCKEMPL];
            ";

            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryUpdateF_ARTSTOCKEMPLAE_QtePrepa,
                    new SqlParameter("@AE_QtePrepa", AE_QtePrepa),
                    new SqlParameter("@cbMarq", cbMarq)
                );
            }
        }
    }
}
