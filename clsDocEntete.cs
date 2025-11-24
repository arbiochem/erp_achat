using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI._2_Documents
{
    public class clsDocEntete
    {
        public int DoDomaine { get; set; }
        public int DoType { get; set; }
        public string DoIntitule { get; set; }
        public string DoPiece { get; set; }
        public DateTime DoDate { get; set; }
        public string DORef { get; set; }
        public string DoTiers { get; set; }
        public int CONo { get; set; }
        public int DoPeriod { get; set; }
        public int DODevise { get; set; }
        public double DoCours { get; set; }
        public int DENo { get; set; }
        public int LiNo { get; set; }
        public string CTNumPayeur { get; set; }
        public int DOExpedit { get; set; }
        public int DONbFacture { get; set; }
        public int BlFac { get; set; }
        public string DOTxEscompte { get; set; }
        public string DOReliquat { get; set; }
        public int DOImprim { get; set; }
        public string CANum { get; set; }
        public string DOCoord01 { get; set; }
        public string DOCoord02 { get; set; }
        public string DOCoord03 { get; set; }
        public string DOCoord04 { get; set; }
        public string DOSouche { get; set; }
        public DateTime DoDateLivr { get; set; }
        public int DOCondition { get; set; }
        public int DOTarif { get; set; }
        public int DOColisage { get; set; }
        public int DOTransaction { get; set; }
        public int DOLangue { get; set; }
        public double DOEcart { get; set; }
        public int DoRegime { get; set; }
        public int NCatCompta { get; set; }
        public string DOVentile { get; set; }
        public int ABNo { get; set; }
        public DateTime DODebutAbo { get; set; }
        public DateTime DOFinAbo { get; set; }
        public DateTime DODebutPeriod { get; set; }
        public DateTime DOFinPeriod { get; set; }
        public string CGNum { get; set; }
        public int DOStatut { get; set; }
        public TimeOnly DoHeure { get; set; }
        public int CANo { get; set; }
        public int CONoCaissier { get; set; }
        public int DOTransfert { get; set; }
        public int DOCloture { get; set; }
        public string DONoWeb { get; set; }
        public int DOAttente { get; set; }
        public int DoPrevenance { get; set; }//0 à 4
                                                //0 = Normale
                                                //1 = Facture de retour
                                                //2 = Facture d’avoir
                                                //3 = Ticket
                                               //     Spécifique version
                                               //     Espagnole
                                              //  4 = Rectificatif

        public string CANumIFRS { get; set; }
        public int MRNo { get; set; }
        public int DOTypeFrais { get; set; }//0 à 3
                                            // 0 = Montant Forfaitaire
                                            // 1 = Quantité
                                            // 2 = Poids net
                                            // 3 = Poids brut 
        public double DOValFrais { get; set; }
        public int DOTypeLigneFrais { get; set; } //0= HT, 1= TTC
        public int DOTypeFranco { get; set; }//type franco de port 
                                            //0 ou 1
                                            // 0 = Montant forfaitaire
                                            // 1 = Quantité 
        public double DOValFranco { get; set; }
        public int DOTypeLigneFranco { get; set; }//0 ou 1
                                                  // 0 = HT
                                                  // 1 = TTC 
        public double DoTaxe1 { get; set; }
        public int DOTypeTaux1 { get; set; }//0 à 2
                                            // 0 = Taux %
                                            // 1 = Montant F
                                            // 2 = Quantité U 
        public int DOTypeTaxe1 { get; set; }//0 à 4
                                            // 0 = TVA/Débit
                                            // 1 = TVA/Encaissement
                                            // 2 = TP/HT
                                            // 3 = TP/TTC
                                            // 4 = TP/Poids 
        public double DoTaxe2 { get; set; }
        public int DOTypeTaux2 { get; set; }//0 à 2
                                            // 0 = Taux %
                                            // 1 = Montant F
                                            // 2 = Quantité U 
        public int DOTypeTaxe2 { get; set; }//0 à 4
                                            // 0 = TVA/Débit
                                            // 1 = TVA/Encaissement
                                            // 2 = TP/HT
                                            // 3 = TP/TTC
                                            // 4 = TP/Poids 

        public double DoTaxe3 { get; set; }
        public int DOTypeTaux3 { get; set; }//0 à 2
                                            // 0 = Taux %
                                            // 1 = Montant F
                                            // 2 = Quantité U 
        public int DOTypeTaxe3 { get; set; }//0 à 4
                                            // 0 = TVA/Débit
                                            // 1 = TVA/Encaissement
                                            // 2 = TP/HT
                                            // 3 = TP/TTC
                                            // 4 = TP/Poids 
        public int DOMajCpta { get; set; }//0 ou 1 
                                          // 0 = Non 
                                          // 1 = Oui 
        public string DOMotif { get; set; }//69 caractères max. 

        public string CTNumCentrale { get; set; }//Client centrale d’achat 17 caractères max. 
        public string DOContact { get; set; }//35 caractères max. 
        public int DOTypetTransac { get; set; }//0 à 17 
                                               // 0 = Classique
                                               // 1 = Regroupement de
                                               // factures
                                               // 2 = Regroupement de
                                               // tickets
                                               // 3 = Factures avec divers
                                               // types de taxes
                                               // 4 = Facture rectificative
                                               // 5 = TVA à payer par
                                               // l’émetteur facture
                                               // 6 = Régime des agences
                                               // de voyage
                                               // 7 = Régime des groupes
                                               // de sociétés
                                               // 8 = Régime des
                                               // transactions sur l’or
                                               // 9 = Inversion du sujet
                                               // passif (ISF)
                                               // 10 = Ticket
                                               // 11 = Rectification d’erreur
                                               // de registre
                                               // 12 = Régime des
                                               // commerçants à l’IGIC
                                               // 13 = Facture dont taxe à
                                               // percevoir
                                               // 14 = Facture des agences
                                               // de voyage
                                               // 15 = Factures de tickets
                                               // 16 = Acquisitions
                                               // intracommunautaires
                                               // 17 = Opération spéciale
                                               // (Objets d’art…) 
        public DateTime DODateLivrRealisee { get; set; }
        public DateTime DODateExpedition { get; set; }
        public string DOFactureFrs { get; set; }
        public string DOPieceOrig { get; set; }
        public string DOTypeColis { get; set; }
        public Guid DO_FactureGUID { get; set; }
        public int DOStatutFacture { get; set; }
        public int DODemandeRegul { get; set; }
        public int ETNo { get; set; }
        public int DOValide { get; set; }
        public int DOCoffre { get; set; }
        public int cbMarq { get; set; }
    }
}
