//using Objets100cLib;
using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using DevExpress.CodeParser;
using arbioApp.Modules.Principal.DI._2_Documents;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;

namespace arbioApp.Services
{
    internal class F_CREGLEMENTService
    {
        private readonly F_CREGLEMENTRepository _repository;

        public F_CREGLEMENTService(F_CREGLEMENTRepository repository)
        {
            _repository = repository;
        }

        //public List<F_CREGLEMENT> GetAllReglements()
        //{
        //    return _repository.GetAll();
        //}

        //public F_CREGLEMENT GetReglementById(int cbMarq)
        //{
        //    return _repository.GetById(cbMarq);
        //}

        //public List<F_CREGLEMENT> GetReglementsByFournisseur(string ctNum)
        //{
        //    return _repository.GetByFournisseur(ctNum);
        //}

        public void AddReglement(F_CREGLEMENT reglement)
        {
            _repository.Add(reglement);
        }

        public void UpdateReglement(F_CREGLEMENT reglement)
        {
            _repository.Update(reglement);
        }

        public void DeleteReglement(int cbMarq)
        {
            //_repository.Delete(cbMarq);
        }
        

    }


}
