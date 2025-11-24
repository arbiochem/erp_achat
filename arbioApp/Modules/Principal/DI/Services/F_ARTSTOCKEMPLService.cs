using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;
namespace arbioApp.Modules.Principal.DI.Services
{
    internal class F_ARTSTOCKEMPLService
    {
        private readonly AppDbContext _context;
        private readonly F_ARTSTOCKEMPLRepository _f_ARTSTOCKEMPLRepository;



        public F_ARTSTOCKEMPLService(AppDbContext context, F_ARTSTOCKEMPLRepository f_ARTSTOCKEMPLRepository)
        {
            _context = context;
            _f_ARTSTOCKEMPLRepository = f_ARTSTOCKEMPLRepository;
        }



        public void UpdateArtstockEmpl(string typeDocument, string DO_Piece, int? DL_Ligne, string AR_Ref, int? prevQte, int? nouvQte, int? DE_No)
        {
            F_DEPOTRepository f_DEPOTRepository = new F_DEPOTRepository(_context);
            int? DP_No = f_DEPOTRepository.GetDP_NoF_ARTSTOCKEMPL(AR_Ref, DE_No);

            if (DP_No != null)
            {
                F_ARTSTOCKEMPL f_ARTSTOCKEMPLToUpdate = _context.F_ARTSTOCKEMPL.Where(artStckEmpl => artStckEmpl.AR_Ref == AR_Ref && artStckEmpl.DP_No == DP_No).FirstOrDefault();

                if (typeDocument == "Devis" || typeDocument == "Bon d'avoir finanicier" || typeDocument == "Facture d'avoir" || typeDocument == "Bon de commande")
                {
                    // Aucun interaction avec l'emplacement des stock pour ces types de documents
                }
                else if (typeDocument == "Préparation de livraison")
                {
                    decimal? AE_QtePrepa = f_ARTSTOCKEMPLToUpdate.AE_QtePrepa - prevQte + nouvQte;
                    _f_ARTSTOCKEMPLRepository.UpdateAE_QtePrepa(AE_QtePrepa, f_ARTSTOCKEMPLToUpdate.cbMarq);
                }
                else if (typeDocument == "Bon de livraison" || typeDocument == "Facture")
                {
                    decimal? AE_QteSto = f_ARTSTOCKEMPLToUpdate.AE_QteSto + prevQte - nouvQte;
                    _f_ARTSTOCKEMPLRepository.UpdateAE_QteSto(AE_QteSto, f_ARTSTOCKEMPLToUpdate.cbMarq);
                }
                else // else if (typeDocument == "Facture de retour" || typeDocument == "Bon de retour")
                {
                    decimal? AE_QteSto = f_ARTSTOCKEMPLToUpdate.AE_QteSto - prevQte + nouvQte;
                    _f_ARTSTOCKEMPLRepository.UpdateAE_QteSto(AE_QteSto, f_ARTSTOCKEMPLToUpdate.cbMarq);
                }
            }
        }




    }
}
