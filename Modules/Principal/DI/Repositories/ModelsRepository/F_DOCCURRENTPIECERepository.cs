using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace arbioApp.Repositories
{
    public class F_DOCCURRENTPIECERepository
    {
        // ====================================================================================================================================================================
        // ============================================================= DEBUT DECLARATION DES VARIABLES ======================================================================
        // ====================================================================================================================================================================
        private readonly AppDbContext _context;
        // ====================================================================================================================================================================
        // ============================================================= FIN DECLARATION DES VARIABLES ======================================================================
        // ====================================================================================================================================================================








        // ====================================================================================================================================================================
        // =================================================================== DEBUT CONSTRUCTEUR =============================================================================
        // ====================================================================================================================================================================
        public F_DOCCURRENTPIECERepository(AppDbContext context)
        {
            _context = context;
        }
        // ====================================================================================================================================================================
        // ==================================================================== FIN CONSTRUCTEUR ==============================================================================
        // ====================================================================================================================================================================









        //public List<ArtClient> GetAll()
        //{
        //    return _context.F_ARTCLIENT
        //        .Select(a => new ArtClient
        //        {
        //            AR_Ref = a.AR_Ref,
        //            AC_PrixVen = a.AC_PrixVen,
        //            AC_Categorie = a.AC_Categorie,
        //        }).ToList();
        //}




        // ====================================================================================================================================================================
        // ====================================================================== DEBUT UPDATE ================================================================================
        // ====================================================================================================================================================================
        public void Update(string typeDoc, string currentPieceNo)
        {
            _context.Database.ExecuteSqlCommand("DISABLE TRIGGER [dbo].[TG_CBUPD_F_DOCCURRENTPIECE] ON [dbo].[F_DOCCURRENTPIECE]");

            if (typeDoc == "Devis")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 1).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Bon de commande")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 2).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Préparation de livraison")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 3).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Bon de livraison")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 4).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Bon de retour")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 5).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Bon d'avoir finanicier")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 6).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Facture")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 74).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Facture de retour")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 75).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }
            else if (typeDoc == "Facture d'avoir")
            {
                F_DOCCURRENTPIECE fDocCurrent = _context.F_DOCCURRENTPIECE.Where(dc => dc.cbMarq == 76).FirstOrDefault();
                if (fDocCurrent.DC_Piece != currentPieceNo)
                {
                    fDocCurrent.DC_Piece = currentPieceNo;
                    fDocCurrent.cbModification = DateTime.Now;
                    _context.SaveChanges();
                }
            }

            _context.Database.ExecuteSqlCommand("ENABLE TRIGGER [dbo].[TG_CBUPD_F_DOCCURRENTPIECE] ON [dbo].[F_DOCCURRENTPIECE]");
        }
        // ====================================================================================================================================================================
        // ====================================================================== DEBUT UPDATE ================================================================================
        // ====================================================================================================================================================================
    }
}
