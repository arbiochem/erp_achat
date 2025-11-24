using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Repositories.ModelsRepository
{
    public class P_DEVISERepository
    {
        private readonly AppDbContext _context;

        public P_DEVISERepository(AppDbContext context)
        {
            //_context = context;
        }
        public P_DEVISE Get_P_DEVISEBy_D_Intitule(string D_Intitule)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_DEVISE.Where(exp => exp.D_Intitule == D_Intitule).FirstOrDefault();
            }
        }


        public P_DEVISE Get_P_DEVISE_By_cbMarq(int cbMarq)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_DEVISE.Where(exp => exp.cbMarq == cbMarq).FirstOrDefault();
            }
        }


        public List<P_DEVISE> GetAll_P_DEVISE_Not_Empty_String()
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_DEVISE.Where(expedit => expedit.D_Intitule != "").ToList();
            }
        }


    }
}
