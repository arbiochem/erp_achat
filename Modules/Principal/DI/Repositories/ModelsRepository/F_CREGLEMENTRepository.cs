using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace arbioApp.Repositories.ModelsRepository
{
    public class F_CREGLEMENTRepository
    {
        private readonly AppDbContext _context;

        public F_CREGLEMENTRepository(AppDbContext context)
        {
            _context = context;
        }

        // Récupérer un enregistrement par son ID (clé primaire)
        public F_CREGLEMENT GetByDLNo(int DL_No)
        {
            return _context.F_CREGLEMENT.Find(DL_No);
        }

        // Ajouter un enregistrement
        public void Add(F_CREGLEMENT entity)
        {
            _context.F_CREGLEMENT.Add(entity);
        }

        // Mettre à jour un enregistrement
        public void Update(F_CREGLEMENT entity)
        {
            _context.F_CREGLEMENT.Attach(entity);
            _context.Entry(entity).State = EntityState.Modified;
        }

        // Supprimer un enregistrement
        public void Delete(F_CREGLEMENT entity)
        {
            if (_context.Entry(entity).State == EntityState.Detached)
            {
                _context.F_CREGLEMENT.Attach(entity);
            }
            _context.F_CREGLEMENT.Remove(entity);
        }

        // Sauvegarder les changements
        public void Save()
        {
            _context.SaveChanges();
        }
    }


}
