using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace arbioApp.Repositories.ModelsRepository
{
    public class F_COLLABORATEURRepository
    {
        // =============================================================================
        // DEBUT DECLARATION DES VARIABLES =============================================
        // =============================================================================
        private readonly AppDbContext _context;
        // =============================================================================
        // FIN DECLARATION DES VARIABLES ===============================================
        // =============================================================================



        public void Dispose()
        {
            _context.Dispose(); // libère la connexion à la BDD
        }






        // =============================================================================
        // DEBUT CONSTRUCTEUR ==========================================================
        // =============================================================================
        public F_COLLABORATEURRepository(AppDbContext context)
        {
            _context = context;
        }
        // =============================================================================
        // FIN CONSTRUCTEUR ============================================================
        // =============================================================================










        // =============================================================================
        // DEBUT METHODES GET ==========================================================
        // =============================================================================
        public List<F_COLLABORATEUR> GetAll()
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_COLLABORATEUR.ToList();
            }
        }



        public F_COLLABORATEUR GetBy_CO_Nom_And_CO_Prenom(string CO_Nom_Prenom)
        {
            using (AppDbContext context = new AppDbContext())
            {
                return context.F_COLLABORATEUR.Where(coll => coll.CO_Nom + " " + coll.CO_Prenom == CO_Nom_Prenom).FirstOrDefault();
            }
        }



        //public F_COLLABORATEUR GetBy_CO_No(int? CO_No)
        //{
        //    using (AppDbContext context = new AppDbContext())
        //    {
        //        Console.WriteLine($"Recherche collaborateur avec CO_No={CO_No}");
        //        var result = context.F_COLLABORATEUR.FirstOrDefault(u => u.CO_No == CO_No);
        //        return result;
        //    }
        //}

        public F_COLLABORATEUR GetBy_CO_No(int? CO_No)
        {
            if (CO_No == null)
                throw new ArgumentNullException(nameof(CO_No));

            using (AppDbContext context = new AppDbContext())
            {
                var collaborateur = context.F_COLLABORATEUR.FirstOrDefault(u => u.CO_No == CO_No);

                if (collaborateur == null)
                    throw new KeyNotFoundException($"Aucun collaborateur trouvé avec CO_No = {CO_No}");

                return collaborateur;
            }
        }



        // =============================================================================
        // FIN METHODES GET ============================================================
        // =============================================================================










        // =============================================================================
        // DEBUT INSERT ================================================================
        // =============================================================================

        // =============================================================================
        // FIN INSERT ==================================================================
        // =============================================================================

    }
}
