using arbioApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using arbioApp.Modules.Principal.DI.Repositories.ModelsRepository;

namespace arbioApp.Modules.Principal.DI.Services
{
    internal class F_ARTSTOCKService
    {
        // ====================================================================================================================================================================================================================================
        // ==================================================================================================== DECLARATION DES VARIABLES =====================================================================================================
        // ====================================================================================================================================================================================================================================
        private readonly AppDbContext _context;
        private readonly F_ARTSTOCKRepository _f_ARTSTOCKRepository;
        // ====================================================================================================================================================================================================================================
        // ==================================================================================================== DECLARATION DES VARIABLES =====================================================================================================
        // ====================================================================================================================================================================================================================================









        // ====================================================================================================================================================================================================================================
        // ======================================================================================================= DEBUT CONSTRUCTEUR =========================================================================================================
        // ====================================================================================================================================================================================================================================
        public F_ARTSTOCKService(AppDbContext context, F_ARTSTOCKRepository f_ARTSTOCKRepository)
        {
            _f_ARTSTOCKRepository = f_ARTSTOCKRepository;
            _context = context;
        }
        // ====================================================================================================================================================================================================================================
        // ======================================================================================================== FIN CONSTRUCTEUR ==========================================================================================================
        // ====================================================================================================================================================================================================================================









        // ====================================================================================================================================================================================================================================
        // ========================================================================================================== DEBUT UPDATE ============================================================================================================
        // ====================================================================================================================================================================================================================================

        public void UpdateMontantEtQuantiteStock(string typeDocument, string AR_Ref, decimal nouvQte, decimal previousQte, int? DE_No)
        {
            var DP_NoPrincipal = _context.F_DEPOT
                                        .Where(depot => depot.DE_No == DE_No)
                                        .Select(depot => _context.F_ARTSTOCK
                                            .Where(artStock => artStock.DE_No == depot.DE_No && artStock.AR_Ref == AR_Ref)
                                            .Select(artStock => artStock.DP_NoPrincipal > 0 ? artStock.DP_NoPrincipal : depot.DP_NoDefaut)
                                            .FirstOrDefault())
                                        .FirstOrDefault();

            int nombreObjetsArtStock = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref).Count();
            F_ARTSTOCK f_ARTSTOCKToUpdate = _context.F_ARTSTOCK.Where(artStck => artStck.AR_Ref == AR_Ref && (nombreObjetsArtStock > 1 ? artStck.DP_NoPrincipal == DP_NoPrincipal : true)).FirstOrDefault();
            
            if (f_ARTSTOCKToUpdate != null)
            {
                //decimal? cmup = f_ARTSTOCKToUpdate.AS_MontSto / (f_ARTSTOCKToUpdate.AS_QteSto == 0 ? 1 : f_ARTSTOCKToUpdate.AS_QteSto);
                if (typeDocument == "Devis" || typeDocument == "Bon d'avoir financier" || typeDocument == "Facture d'avoir")
                {
                    // Aucun interaction avec le stock pour ces types de documents
                }
                else if (typeDocument == "Bon de commande")
                {
                    decimal? AS_QteRes = f_ARTSTOCKToUpdate.AS_QteRes - previousQte + nouvQte;
                    _f_ARTSTOCKRepository.UpdateQuantiteReserve(AR_Ref, DP_NoPrincipal, AS_QteRes);
                }
                else if (typeDocument == "Préparation de livraison")
                {
                    decimal? AS_QtePrepa = f_ARTSTOCKToUpdate.AS_QtePrepa - previousQte + nouvQte;
                    _f_ARTSTOCKRepository.UpdateQuantitePrepare(AR_Ref, DP_NoPrincipal, AS_QtePrepa);
                }
                else if (typeDocument == "Bon de livraison" || typeDocument == "Facture")
                {
                    decimal? AS_QteSto = f_ARTSTOCKToUpdate.AS_QteSto + previousQte - nouvQte;
                    //decimal? AS_MontSto = AS_QteSto * cmup;
                    //_f_ARTSTOCKRepository.UpdateMontantEtQuantiteStock(AR_Ref, DP_NoPrincipal, AS_MontSto, AS_QteSto);
                }
                else // else if (typeDocument == "Facture de retour" || typeDocument == "Bon de retour")
                {
                    decimal? AS_QteSto = f_ARTSTOCKToUpdate.AS_QteSto - previousQte + nouvQte;
                    //decimal? AS_MontSto = AS_QteSto * cmup;
                    //_f_ARTSTOCKRepository.UpdateMontantEtQuantiteStock(AR_Ref, DP_NoPrincipal, AS_MontSto, AS_QteSto);
                }
            }
            else
            {
                //decimal? cmup = f_ARTSTOCKToUpdate.AS_MontSto / (f_ARTSTOCKToUpdate.AS_QteSto == 0 ? 1 : f_ARTSTOCKToUpdate.AS_QteSto);
                if (typeDocument == "projet d'achat" || typeDocument == "Facture d'avoir")
                {
                    // Aucun interaction avec le stock pour ces types de documents
                }
                else if (typeDocument == "Bon de commande")
                {
                    //decimal? AS_QteRes = f_ARTSTOCKToUpdate.AS_QteRes - previousQte + nouvQte;
                    //_f_ARTSTOCKRepository.UpdateQuantiteReserve(AR_Ref, DP_NoPrincipal, AS_QteRes);
                }
                else if (typeDocument == "Préparation de livraison")
                {
                    //decimal? AS_QtePrepa = f_ARTSTOCKToUpdate.AS_QtePrepa - previousQte + nouvQte;
                    //_f_ARTSTOCKRepository.UpdateQuantitePrepare(AR_Ref, DP_NoPrincipal, AS_QtePrepa);
                }
                else if (typeDocument == "Bon de livraison" || typeDocument == "Facture")
                {
                    //decimal? AS_QteSto = (f_ARTSTOCKToUpdate.AS_QteSto ?? 0) + previousQte - nouvQte;
                    //decimal? AS_MontSto = AS_QteSto * cmup;
                    //_f_ARTSTOCKRepository.UpdateMontantEtQuantiteStock(AR_Ref, DP_NoPrincipal, AS_MontSto, AS_QteSto);
                    var newArtStock = new F_ARTSTOCK
                    {
                        AR_Ref = AR_Ref,
                        DE_No = (int)DE_No, // dépôt actuel
                        DP_NoPrincipal = (int)DE_No, // ou une valeur par défaut
                       // AS_QteSto = AS_QteSto,
                        AS_MontSto = 0,
                        AS_QteRes = 0,
                        AS_QtePrepa = 0,

                    };
                    _context.F_ARTSTOCK.Add(newArtStock);
                    _context.SaveChanges();
                }
                else // else if (typeDocument == "Facture de retour" || typeDocument == "Bon de retour")
                {
                    //decimal? AS_QteSto = f_ARTSTOCKToUpdate.AS_QteSto - previousQte + nouvQte;
                    //decimal? AS_MontSto = AS_QteSto * cmup;
                    //_f_ARTSTOCKRepository.UpdateMontantEtQuantiteStock(AR_Ref, DP_NoPrincipal, AS_MontSto, AS_QteSto);
                }
                
            }
        }


        // ====================================================================================================================================================================================================================================
        // ========================================================================================================== FIN UPDATE ============================================================================================================
        // ====================================================================================================================================================================================================================================
    }
}
