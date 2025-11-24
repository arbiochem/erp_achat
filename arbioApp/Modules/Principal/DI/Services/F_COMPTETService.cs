using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using arbioApp.Repositories.ModelsRepository;
using arbioApp.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace arbioApp.Modules.Principal.DI.Services
{
    internal class F_COMPTETService
    {
        private readonly F_COMPTETRepository _f_COMPTETRepository;

        public F_COMPTETService(F_COMPTETRepository repository)
        {
            _f_COMPTETRepository = repository;
        }

        // Récupère tous les comptes de type client (CT_Type = 1)
        public List<F_COMPTET> GetClients()
        {
            return _f_COMPTETRepository.GetAll_F_COMPTET_Zero();
        }

        // Récupère un compte par son numéro
        public F_COMPTET GetCompteByNumero(string numero)
        {
            return _f_COMPTETRepository.GetByCT_Num(numero);
        }

        // Récupère tous les numéros de comptes clients
        public List<string> GetNumerosClients()
        {
            return _f_COMPTETRepository.GetCTNumF_CompteT();
        }

        // Récupère le cours associé à un compte
        public decimal? GetCoursDeviseByCompte(string numero)
        {
            return _f_COMPTETRepository.GetF_COMPTET_Cours_N_Devise(numero);
        }

        // Récupère le compte général principal
        public string GetComptePrincipal(string numero)
        {
            return _f_COMPTETRepository.GetF_COMPTET_CG_NumPrinc(numero);
        }

        // Exemple de logique métier : vérifier si un client a une devise attribuée
        public bool CompteAvecDevise(string numero)
        {
            var compte = _f_COMPTETRepository.GetByCT_Num(numero);
            return compte != null && compte.N_Devise.HasValue;
        }
        public void CreateCompte(F_COMPTET newCompte)
        {
            if (string.IsNullOrWhiteSpace(newCompte.CT_Num))
                throw new ArgumentException("Le numéro de compte est obligatoire.");

            // Tu peux ajouter d'autres validations ici

            _f_COMPTETRepository.Add(newCompte);
        }
        public void SaveOrUpdate(F_COMPTET compte)
        {
            var existing = _f_COMPTETRepository.GetByCT_Num(compte.CT_Num);
            if (existing != null)
            {
                _f_COMPTETRepository.Update(compte);
            }
            else
            {
                _f_COMPTETRepository.Add(compte);
            }
        }
    }
}
