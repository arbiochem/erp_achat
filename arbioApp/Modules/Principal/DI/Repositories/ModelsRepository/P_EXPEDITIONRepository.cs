using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Repositories.ModelsRepository
{
    public class P_EXPEDITIONRepository
    {
        private readonly AppDbContext _context;

        public P_EXPEDITIONRepository(AppDbContext context)
        {
            //_context = context;
        }
        public P_EXPEDITION Get_P_EXPEDITIONBy_E_Intitule(string E_Intitule)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_EXPEDITION.Where(exp => exp.E_Intitule == E_Intitule).FirstOrDefault();
            }
        }


        public P_EXPEDITION Get_P_EXPEDITION_By_cbMarq(int cbMarq)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_EXPEDITION.Where(exp => exp.cbMarq == cbMarq).FirstOrDefault();
            }
        }


        public List<P_EXPEDITION> GetAll_P_EXPEDITION_Not_Empty_String()
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.P_EXPEDITION.Where(expedit => expedit.E_Intitule != "").ToList();
            }
        }


    }
}
