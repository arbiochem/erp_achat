using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace arbioApp.Repositories.ModelsRepository
{
    public class F_COMPTETRepository
    {
        // =====================================================================================
        // DEBUT DECLARATION DES VARIABLES =====================================================
        // =====================================================================================
        private readonly AppDbContext _context;
        // =====================================================================================
        // FIN DECLARATION DES VARIABLES =======================================================
        // =====================================================================================


        // =====================================================================================
        // DEBUT CONSTRUCTEUR ==================================================================
        // =====================================================================================
        public F_COMPTETRepository(AppDbContext context)
        {
            _context = context;
        }
        // =====================================================================================
        // FIN CONSTRUCTEUR ====================================================================
        // =====================================================================================

        // =====================================================================================
        // DEBUT METHODES GET ==================================================================
        // =====================================================================================
        public List<F_COMPTET> GetAll_F_COMPTET_Zero()
        {
            using (AppDbContext context = new AppDbContext())
            {
                var allRecords = context.F_COMPTET.ToList();
                MessageBox.Show($"Nombre total d'enregistrements dans F_COMPTET : {allRecords.Count}");

                foreach (var item in allRecords)
                {
                    MessageBox.Show($"CT_Num: {item.CT_Num}, CT_Type: {item.CT_Type}");
                }

                var filteredRecords = allRecords.Where(ct => ct.CT_Type == 1).ToList();
                MessageBox.Show($"Nombre d'enregistrements après filtrage : {filteredRecords.Count}");

                return filteredRecords;
            }
        }

        public void Add(F_COMPTET newCompte)
        {
            using (AppDbContext context = new AppDbContext())
            {
                context.F_COMPTET.Add(newCompte);
                context.SaveChanges();
            }
        }



        public F_COMPTET GetByCT_Num(string CT_Num)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_COMPTET.FirstOrDefault(ct => ct.CT_Num.Trim().ToUpper() == CT_Num.Trim().ToUpper());

            }
        }
        public F_COMPTET GetDeviseNameByCT_Num(string CT_Num)
        {
            using (AppDbContext context = new AppDbContext())
            {
                var compte = GetByCT_Num(CT_Num);
                if (compte != null)
                {
                    short? N_Devise = compte.N_Devise;
                    return new F_COMPTET { N_Devise = N_Devise };
                }
                return null;
            }
        }


        public List<string> GetCTNumF_CompteT()
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_COMPTET.Where(cpt => cpt.CT_Type == 1).Select(u => u.CT_Num).ToList();
            }
        }



        //public short? Get_N_Period_By_CT_Num(string CT_Num)
        //{
        //    using (AppDbContext context = new AppDbContext())
        //    {
        //        return context.F_COMPTET.Where(cpt => cpt.CT_Num == CT_Num).Select(u => u.N_Period).FirstOrDefault();
        //    }
        //}



        public decimal? GetF_COMPTET_Cours_N_Devise(string CT_Num)
        {
            try
            {             
                short? N_Devise = GetByCT_Num(CT_Num).N_Devise;
                using (AppDbContext context = new AppDbContext())
                {
                    P_DEVISE p_DEVISE = context.P_DEVISE.FirstOrDefault(dv => dv.cbMarq == N_Devise);
                    if (p_DEVISE != null)
                    {
                        return p_DEVISE.D_Cours;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        // =====================================================================================
        // FIN METHODES GET ====================================================================
        // =====================================================================================
        public string GetF_COMPTET_CG_NumPrinc(string CT_Num)
        {
            try
            {
                return GetByCT_Num(CT_Num).CG_NumPrinc;             
            }
            catch (Exception ex)
            {
                MethodBase m = MethodBase.GetCurrentMethod();
                MessageBox.Show($"Une erreur est survenue : {ex.Message}, {m}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public void Update(F_COMPTET updatedCompte)
        {
            using (var context = new AppDbContext())
            {
                var existing = context.F_COMPTET.FirstOrDefault(c => c.CT_Num == updatedCompte.CT_Num);
                if (existing != null)
                {
                    existing.CT_Intitule = updatedCompte.CT_Intitule;
                    existing.CT_Type = updatedCompte.CT_Type;
                    existing.CT_Qualite = updatedCompte.CT_Qualite;
                    existing.CT_Classement = updatedCompte.CT_Classement;
                    existing.CT_Adresse = updatedCompte.CT_Adresse;
                    existing.CT_Complement = updatedCompte.CT_Complement;
                    existing.CT_CodePostal = updatedCompte.CT_CodePostal;
                    existing.CT_Ville = updatedCompte.CT_Ville;
                    existing.CT_Pays = updatedCompte.CT_Pays;
                    existing.CT_CodeRegion = updatedCompte.CT_CodeRegion;
                    existing.CT_Identifiant = updatedCompte.CT_Identifiant;
                    existing.CT_Siret = updatedCompte.CT_Siret;
                    existing.CT_Commentaire = updatedCompte.CT_Commentaire;
                    existing.CT_NumPayeur = updatedCompte.CT_NumPayeur;
                    existing.CT_Sommeil = updatedCompte.CT_Sommeil;
                    existing.CT_Telephone = updatedCompte.CT_Telephone;
                    existing.CT_Telecopie = updatedCompte.CT_Telecopie;
                    existing.CT_Facebook = updatedCompte.CT_Facebook;
                    existing.CT_LinkedIn = updatedCompte.CT_LinkedIn;
                    existing.CT_EMail = updatedCompte.CT_EMail;
                    existing.CT_Site = updatedCompte.CT_Site;
                    existing.N_Devise = updatedCompte.N_Devise;
                    existing.DE_No = updatedCompte.DE_No;
                    existing.CG_NumPrinc = updatedCompte.CG_NumPrinc;
                    existing.CT_DateMAJ = DateTime.Now;

                    context.SaveChanges();
                }
            }
        }


    }
}
