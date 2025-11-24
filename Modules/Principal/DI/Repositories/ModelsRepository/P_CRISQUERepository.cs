using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Repositories.ModelsRepository
{
    public class P_CRISQUERepository
    {
        private readonly AppDbContext _context;
        public P_CRISQUERepository(AppDbContext context)
        {
            _context = context;
        }


        public List<P_CRISQUE> GetAll()
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_CRISQUE.ToList();
            }
        }



        public P_CRISQUE GetById(int cbMarq)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_CRISQUE.FirstOrDefault(x => x.cbMarq == cbMarq);
            }
        }



    }
}
