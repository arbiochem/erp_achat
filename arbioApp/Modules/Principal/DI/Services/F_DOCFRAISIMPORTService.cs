using Objets100cLib;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.CodeParser;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Models;

namespace arbioApp.Services
{
    internal class F_DOCFRAISIMPORTService
    {
        private readonly AppDbContext _context;
        private readonly F_DOCFRAISIMPORTRepository _F_DOCFRAISIMPORTRepository;
        public F_DOCFRAISIMPORTService(F_DOCFRAISIMPORTRepository F_DOCFRAISIMPORTRepository)
        {
            _context = new AppDbContext();

            _F_DOCFRAISIMPORTRepository = F_DOCFRAISIMPORTRepository;

        }

        public F_DOCFRAISIMPORT GetNameByFRNum(int FRNum, List<F_DOCFRAISIMPORT> listeFrais)
        {
            return listeFrais
                .FirstOrDefault(doc => doc.FI_TypeFraisId == FRNum);
        }
    }
}
