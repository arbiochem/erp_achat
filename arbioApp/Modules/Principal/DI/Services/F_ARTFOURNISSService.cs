using arbioApp.Models;
using arbioApp.Repositories.ModelsRepository;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Services
{
    internal class F_ARTFOURNISSService
    {
        private readonly F_ARTFOURNISSRepository _artFournissRepo;

        public F_ARTFOURNISSService(F_ARTFOURNISSRepository artFournissRepo)
        {
            _artFournissRepo = artFournissRepo;
        }

        //public F_ARTFOURNISS GetByARRef(string arRef)
        //{
        //    return _artFournissRepo.GetByARRef(arRef);
        //}
        public F_ARTFOURNISS GetByARRefAndPrincipal(string arRef)
        {
            return _artFournissRepo.GetByARRefAndPrincipal(arRef);
        }

        public List<F_ARTFOURNISS> GetAll()
        {
            return _artFournissRepo.GetAll();
        }

        public void Add(F_ARTFOURNISS entity)
        {
            _artFournissRepo.Add(entity);
        }

        public void Update(F_ARTFOURNISS entity)
        {
            _artFournissRepo.Update(entity);
        }

        public void Delete(int id)
        {
            _artFournissRepo.Delete(id);
        }

        public void SaveChanges(List<F_ARTFOURNISS> entities)
        {
            foreach (var entity in entities)
            {
                if (entity.cbMarq == 0) // Nouvel enregistrement
                {
                    Add(entity);
                }
                else // Modification
                {
                    Update(entity);
                }
            }
        }
    }
}
