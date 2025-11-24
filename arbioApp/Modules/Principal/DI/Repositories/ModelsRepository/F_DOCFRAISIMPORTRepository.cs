using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace arbioApp.Repositories.ModelsRepository
{
    internal class F_DOCFRAISIMPORTRepository
    {
		
        private readonly AppDbContext _context;
      
        public F_DOCFRAISIMPORTRepository(AppDbContext context)
        {
            _context = context;
        }
		
		public List<F_DOCFRAISIMPORT> GetAll()
		{
			using (AppDbContext context = new AppDbContext())
			{
				return context.F_DOCFRAISIMPORT.ToList();
			}
		}


        public F_DOCFRAISIMPORT GetBy_FRNum(int FR_Num)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_DOCFRAISIMPORT
                    .FirstOrDefault(doc => doc.FI_TypeFraisId == FR_Num);
            }
        }

       
        
    }
}
