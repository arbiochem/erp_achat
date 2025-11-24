using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;

namespace arbioApp.Repositories.ModelsRepository
{
    public class F_AGENDARepository
    {
        //private readonly AppDbContext _context;

        public F_AGENDARepository(AppDbContext context)
        {
            //_context = context;
        }

        public void DeleteF_AGENDA(F_AGENDA f_AGENDAToDelete)
        {
            string queryDeleteF_AGENDA = @"
                DISABLE TRIGGER [dbo].[TG_CBDEL_F_AGENDA] ON [dbo].[F_AGENDA];

                DELETE FROM F_AGENDA WHERE DL_No = @DL_No;

                ENABLE TRIGGER [dbo].[TG_CBDEL_F_AGENDA] ON [dbo].[F_AGENDA];
            ";


            using (var context = new AppDbContext())
            {
                context.Database.ExecuteSqlCommand(
                    queryDeleteF_AGENDA,
                    new SqlParameter("@DL_No", f_AGENDAToDelete.DL_No)
                );
            }

            
        }
    }
}
