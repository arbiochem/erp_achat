using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Linq;

namespace arbioApp.Repositories.ModelsRepository
{
    internal class F_ARTFOURNISSRepository : IRepository<F_ARTFOURNISS>
    {
        private readonly AppDbContext _context;

        public F_ARTFOURNISSRepository(AppDbContext context)
        {
            _context = context;
        }

        public void Add(F_ARTFOURNISS entity)
        {
            // Générer un nouveau cbMarq (clé primaire)
            entity.cbMarq = _context.F_ARTFOURNISS.Any() ?
                _context.F_ARTFOURNISS.Max(x => x.cbMarq) + 1 : 1;

            // Définir les valeurs par défaut
            entity.cbCreation = DateTime.Now;
            entity.cbModification = DateTime.Now;

            _context.F_ARTFOURNISS.Add(entity);
            _context.SaveChanges();
        }

        public void Delete(int id)
        {
            var entity = _context.F_ARTFOURNISS.FirstOrDefault(x => x.cbMarq == id);
            if (entity != null)
            {
                _context.F_ARTFOURNISS.Remove(entity);
                _context.SaveChanges();
            }
        }

        public List<F_ARTFOURNISS> GetAll()
        {
            return _context.F_ARTFOURNISS.ToList();
        }

        public F_ARTFOURNISS GetById(int id)
        {
            return _context.F_ARTFOURNISS.FirstOrDefault(x => x.cbMarq == id);
        }
        public F_ARTFOURNISS GetByArref(string arref)
        {
            return _context.F_ARTFOURNISS.FirstOrDefault(x => x.AR_Ref == arref);
        }

        public void Update(F_ARTFOURNISS entity)
        {
            var existing = _context.F_ARTFOURNISS.FirstOrDefault(x => x.cbMarq == entity.cbMarq);
            if (existing != null)
            {
                // Mettre à jour les propriétés
                _context.Entry(existing).CurrentValues.SetValues(entity);
                existing.cbModification = DateTime.Now;

                _context.SaveChanges();
            }
        }

        //public F_ARTFOURNISS GetByARRef(string ArRef)
        //{
        //    return _context.F_ARTFOURNISS
        //        .Where(artFr => artFr.AR_Ref == ArRef)
        //        .FirstOrDefault();
        //}
        public F_ARTFOURNISS GetByARRefAndPrincipal(string ArRef)
        {
            F_ARTFOURNISS fournisseurDeLArticle = _context.F_ARTFOURNISS.Where(artFr => artFr.AR_Ref == ArRef && artFr.AF_Principal == 1).FirstOrDefault();
            return fournisseurDeLArticle;
        }
    }
}
