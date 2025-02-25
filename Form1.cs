﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TopSolid.Cad.Design.Automating;
using TopSolid.Kernel.Automating;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using System.Management;
using System.Linq;
using System.Net.NetworkInformation;
using System.Diagnostics;
using System.Threading;








namespace Folder_Creator_Tool_V3
{
  




    public partial class Form1 : Form
        {



        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        

        







        // Classe pour récupérer le chemin réseau complet
        public class NetworkPathHelper
        {
            [DllImport("mpr.dll")]
            private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);
            public static string GetNetworkPath(string driveLetter)
            {
                StringBuilder remoteName = new StringBuilder(256);
                int length = remoteName.Capacity;

                int result = WNetGetConnection(driveLetter, remoteName, ref length);
                if (result == 0)
                {
                    return remoteName.ToString();
                }
                else
                {
                    return null;
                }

            }
            public static string GetMappedDriveName(string driveLetter) 
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher($"SELECT VolumeName FROM Win32_LogicalDisk WHERE DeviceID='{driveLetter}'")) 
                {
                    foreach (ManagementObject disk in searcher.Get()) 
                    { 
                        return disk["VolumeName"]?.ToString(); 
                    } 
                }
                return null;
            }
        }







        PdmObjectId CurrentProjectPdmId = new PdmObjectId(); //Id du projet courant
        string CurrentProjectName = ""; //Nom du projet courent

        DocumentId CurrentDocumentId = new DocumentId(); //Id du document courant
        PdmObjectId PdmObjectIdCurrentDocumentId = new PdmObjectId();
        DocumentId CurrentDocumentIdLastRev;

        string TextCurrentDocumentName = ""; //Nom du document courant

        List<PdmObjectId> FolderIds = new List<PdmObjectId>(); //Liste des dossiers contenu dans le projet courant
        List<PdmObjectId> DocumentsIds = new List<PdmObjectId>(); //Liste des documents contenu dans le projet courant

        ElementId CurrentDocumentCommentaireId = new ElementId(); //Id du commentaire du document courant
        string TextCurrentDocumentCommentaire; //Texte du commentaire du document courant
        ElementId CurrentDocumentDesignationId = new ElementId();//Id de la designation du document courant
        string TextCurrentDocumentDesignation = "";//Texte de la designation du document courant

        string TextBoxProjectName =""; //Nom du projet affiché dans la textbox

        string TextBoxDesignationValue="";//Texte afficher dans la textbox designation
        string TextBoxCommentaireValue=""; //Texte afficher dans la textbox Commentaire
        string TextBoxIndiceValue = "";//Texte afficher dans la textbox Indice
        string TextBoxNomMouleValue;//Texte afficher dans la textbox NomMoule

        string ConstituantFolderName = "";
        List<string> ConstituantFolderNames = new List<string>();

        List<string> ListFoldersNames = new List<string>();

        string TexteDossierRep = ""; //Nom du dossier repere a creer

        bool test00; //Creation bool pour tester la presence des dossiers a creer dans le projet
        bool test01;
        bool test02;
        bool test03;


        List<PdmObjectId> DocumentsInIndiceFolder = new List<PdmObjectId>(); //Liste document fictive pour recuperation dossier indice
        List<PdmObjectId> IndiceFolderIds = new List<PdmObjectId>(); //Liste des Id de dossier indice

        PdmObjectId DossierExistantId =new PdmObjectId(); //Id du dossier existant

        string IndiceFolderName = "";

        PdmObjectId AtelierFolderId = new PdmObjectId(); //Id du dossier atelier



        string CommentaireTxtFormat00 = ""; //Different format de commantaire
        string CommentaireTxtFormat01 = "";
        string IndiceTxtFormat00 = ""; //Different format d'indice
        string IndiceTxtFormat01 = "";

        List<PdmObjectId> AtelierFolderIds = new List<PdmObjectId>();

        PdmObjectId DossierRepId = new PdmObjectId(); //Recuperation de l'Id du dossier rep pour creation du rep indice




        string TexteIndiceFolder = "";
        List<PdmObjectId> DocuDossierIndiceIds = new List<PdmObjectId>();


        PdmObjectId AuteurPdmObjectId = new PdmObjectId();

        PdmObjectId DocumentModeleDerivation = new PdmObjectId("19_5af816ad - b4b1 - 402d - 8914 - a4c95a895d88 & 3_6038"); //Recupération du PdmObjectId du document modele de dérivation

        DocumentId DerivéDocumentId = new DocumentId(); //recuperation de l'Id de document du document dérivé
        List<DocumentId> DerivéDocumentIds = new List<DocumentId>(); //recuperation de l'Id de document du document dérivé

        PdmObjectId Dossier3DPdmObjectId = new PdmObjectId();
        List<PdmObjectId> Dossier3DPdmObjectIds = new List<PdmObjectId>();


        PdmObjectId DerivéDocumentPdmObjectId = new PdmObjectId(); //Recupération du PdmObjectId du document dérivé

        List<PdmObjectId> DerivéDocumentPdmObjectIds = new List<PdmObjectId>(); //Creation liste du PdmObjectId du document dérivé

        PdmObjectId DossierIndiceId = new PdmObjectId(); //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        String nomDocu = "";
        List<PdmObjectId> nomDocuIds = new List<PdmObjectId>();

        PdmObjectId dossier3DGenereId = new PdmObjectId();

        ElementId CurrentNameParameterId = new ElementId();
        bool areDirectionsEqual = false;

        int X_TExporterIndex = new int();
        List<KeyValue> options = new List<KeyValue>();

        //------------------------------------------------------------------


        //Fonction de création des dossier apres ind
        static PdmObjectId creationAutreDossiers(PdmObjectId DossierIndiceIdFonction)
        {
            PdmObjectId dossier3DFonctionId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers
                                             //PdmObjectId DossierIndiceIdFonction; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers
            PdmObjectId DossierElectrodeId; //Recuperation de l'Id du dossier electrode pour creation des dossiers
            PdmObjectId DossierFraisageId; //Recuperation de l'Id du dossier fraisage pour creation des dossiers utilisateursP
            PdmObjectId DossierMethodeId;  //Recuperation de l'Id du dossier Methode pour creation des dossiers controle et tournage
            PdmObjectId DossierFLFA; //Recuperation de l'Id du dossier FLFA pour creation du dossier OP1
            PdmObjectId DossierBEHE; //Recuperation de l'Id du dossier FLFA pour creation du dossier OP1
            PdmObjectId DossierSETE; //Recuperation de l'Id du dossier FLFA pour creation du dossier OP1


            try
            {

                //Creation du reste des dossiers
                dossier3DFonctionId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "3D");
                DossierElectrodeId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Electrode");
                DossierFraisageId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Fraisage");
                DossierMethodeId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Methode");
                TSH.Pdm.CreateFolder(DossierMethodeId, "Contrôle");
                TSH.Pdm.CreateFolder(DossierMethodeId, "Tournage");

                //Cration des dossier utilisateur dans le dossier fraisage

                DossierBEHE = TSH.Pdm.CreateFolder(DossierFraisageId, "BEHE");
                DossierFLFA = TSH.Pdm.CreateFolder(DossierFraisageId, "FLFA");
                DossierSETE = TSH.Pdm.CreateFolder(DossierFraisageId, "SETE");

                //Creation des dossier OP1
                TSH.Pdm.CreateFolder(DossierBEHE, "OP1");
                TSH.Pdm.CreateFolder(DossierFLFA, "OP1");
                TSH.Pdm.CreateFolder(DossierSETE, "OP1");

                //Creation des dossiers dans le dossier Electrode

                TSH.Pdm.CreateFolder(DossierElectrodeId, "Air projetée");
                TSH.Pdm.CreateFolder(DossierElectrodeId, "Parallélisée");
                TSH.Pdm.CreateFolder(DossierElectrodeId, "Plan brut");
                TSH.Pdm.CreateFolder(DossierElectrodeId, "Usinage");
                //MessageBox.Show("Succés de la creation des dossiers");

            }

            catch (Exception ex)
            {

                MessageBox.Show("Erreur dans la fonction autre dossier " + ex.Message);

            }
            return (dossier3DFonctionId);

        }

        //------------Activation mode modification----------------------
        void modifActif (DocumentId DocuCourent = new DocumentId())
        {
            // Vérification de la connexion à TSH   
            if (!TSH.IsConnected) return;

            // Vérification de l'ID du document
            if (DocuCourent.IsEmpty) return;

            // Démarrage de la modification du document
            if (!TSH.Application.StartModification("My Action", false)) return;

            // Marquage du document comme modifié et récupération de l'ID du document en cours de modification
            TSH.Documents.EnsureIsDirty(ref DocuCourent);
            
        }

        //----------------------------------------Fonction treeview-----------------------------------
        // Cette fonction ajoute les nœuds cochés à une liste.
        void AddCheckedNodesToList(TreeNodeCollection nodes, List<PdmObjectId> list)
        {
            // Parcourir tous les nœuds dans la collection de nœuds donnée.
            foreach (TreeNode node in nodes)
            {
                // Si le nœud est coché...
                if (node.Checked)
                {
                    // Ajouter l'identifiant de l'objet PDM associé au nœud à la liste.
                    list.Add((PdmObjectId)node.Tag);
                }
                // Appeler cette fonction de manière récursive pour tous les nœuds enfants du nœud actuel.
                AddCheckedNodesToList(node.Nodes, list);
            }
        }


        //-------------------------------Fonction de recupéaration des pdf-------------------
        TreeNode rootFolderNode = new TreeNode();
        string targetName = "";
        void listePdf(string indice)
        {
            List<PdmObjectId> dossiers2Ds = new List<PdmObjectId>();
            PdmObjectId dossiers2D = new PdmObjectId();
            List<PdmObjectId> FoldersInPDFFolder = new List<PdmObjectId>();
            List<PdmObjectId> PDFIds = new List<PdmObjectId>();

            try
            {
                dossiers2Ds = TSH.Pdm.SearchFolderByName(CurrentProjectPdmId, "01-2D");
                dossiers2D = dossiers2Ds[0];
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Dossier '01-2D' introuvable" + ex.Message);
                return;
            }

            string rootFolderName = TSH.Pdm.GetName(dossiers2D);
            rootFolderNode = new TreeNode(rootFolderName);

            TSH.Pdm.GetConstituents(dossiers2D, out FoldersInPDFFolder, out PDFIds);

            bool isEventTriggeredByCode = false;
            treeView1.AfterCheck += (s, e) =>
            {
                if (isEventTriggeredByCode)
                    return;

                if (e.Node.Nodes.Count > 0)
                {
                    isEventTriggeredByCode = true;
                    e.Node.Checked = false;
                    isEventTriggeredByCode = false;
                }
            };

            foreach (PdmObjectId folderId in FoldersInPDFFolder)
            {
                ProcessFolder(folderId, rootFolderNode, indice);
            }

            treeView1.Nodes.Add(rootFolderNode);

           

            targetName = "Ind " +indice;

            
        }

        void ProcessFolder(PdmObjectId folderId, TreeNode parentNode, string indice)
        {
            string folderName = TSH.Pdm.GetName(folderId);
            TreeNode folderNode = new TreeNode(folderName);

            List<PdmObjectId> subFolderIds;
            List<PdmObjectId> fileIdsInFolder;
            TSH.Pdm.GetConstituents(folderId, out subFolderIds, out fileIdsInFolder);

            // Liste pour les sous-dossiers "Ind" à trier
            List<TreeNode> indFolderNodes = new List<TreeNode>();

            // Parcours des sous-dossiers
            foreach (PdmObjectId subFolderId in subFolderIds)
            {
                string subFolderName = TSH.Pdm.GetName(subFolderId);

                // Si le nom du sous-dossier commence par "Ind"
                if (subFolderName.StartsWith("Ind", StringComparison.OrdinalIgnoreCase))
                {
                    // Création du nœud pour ce sous-dossier
                    TreeNode subFolderNode = new TreeNode(subFolderName);
                    indFolderNodes.Add(subFolderNode); // Ajout du nœud à la liste
                }
                else
                {
                    // Traitement récursif pour les autres sous-dossiers
                    ProcessFolder(subFolderId, folderNode, indice);
                }
            }

            // Tri des sous-dossiers "Ind" par ordre alphabétique
            indFolderNodes.Sort((node1, node2) => string.Compare(node1.Text, node2.Text, StringComparison.OrdinalIgnoreCase));

            // Ajout des sous-dossiers "Ind" triés au dossier actuel
            foreach (TreeNode indFolderNode in indFolderNodes)
            {
                folderNode.Nodes.Add(indFolderNode);
            }

            // Ajouter les fichiers dans le dossier
            foreach (PdmObjectId fileId in fileIdsInFolder)
            {
                string fileName = TSH.Pdm.GetName(fileId);
                TreeNode fileNode = new TreeNode(fileName);
                fileNode.Tag = fileId;
                folderNode.Nodes.Add(fileNode);
            }

            // Ajouter le dossier actuel à l'arbre
            parentNode.Nodes.Add(folderNode);
        }

        void ExpandTargetNode(TreeNode rootNode, string targetName)
        {
            // Supprimer les espaces et convertir les deux chaînes en minuscules pour ne pas tenir compte de la casse et des espaces
            string normalizedTargetName = targetName.Replace(" ", "").ToLower();

            // Parcours récursif des nœuds enfants
            foreach (TreeNode node in rootNode.Nodes)
            {
                // Normalisation du nom du nœud
                string normalizedNodeName = node.Text.Replace(" ", "").ToLower();

                // Vérification si le nom du nœud correspond à la cible (en ignorant la casse et les espaces)
                if (normalizedNodeName == normalizedTargetName)
                {
                    // Déplie le nœud cible
                    node.Expand();

                    // Remonte la hiérarchie pour déplier les parents
                    TreeNode parent = node.Parent;
                    while (parent != null)
                    {
                        parent.Expand();
                        parent = parent.Parent;
                    }

                    // Tri des nœuds enfants après dépliage
                    SortTreeNodes(rootNode); // Ajouter le tri ici si nécessaire

                    return; // Sort de la fonction après avoir déplié le nœud
                }
                else
                {
                    // Appel récursif pour vérifier les sous-dossiers
                    ExpandTargetNode(node, targetName);
                }
            }
        }

        void SortTreeNodes(TreeNode rootNode)
        {
            // Tri des nœuds enfants par ordre alphabétique
            var sortedNodes = rootNode.Nodes.Cast<TreeNode>().OrderBy(n => n.Text, StringComparer.OrdinalIgnoreCase).ToList();

            // Réaffecter les nœuds triés à la collection de nœuds
            rootNode.Nodes.Clear();
            foreach (var node in sortedNodes)
            {
                rootNode.Nodes.Add(node);
            }
        }

        void ExpandNodeByName(TreeNodeCollection nodes, string targetName)
        {
            foreach (TreeNode node in nodes)
            {
                // Vérifie si le nom du nœud correspond au dossier recherché
                if (node.Text.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                {
                    // Déplie le nœud correspondant
                    node.Expand();
                    return;
                }

                // Appel récursif pour vérifier les sous-nœuds
                ExpandNodeByName(node.Nodes, targetName);
            }
        }

        //Fonction recuperation de l'indice pour valeur par defaut formulaire------------------------------------------------

        void ChercherDossierDocumentEnCours(PdmObjectId PdmObjectIdCurrentDocumentId, out string IndiceTxtBox)
        {
            IndiceTxtBox = "";

            List<PdmObjectId> dossiers3Ds = new List<PdmObjectId>();
            try
            {
                // Recherche du dossier "02-3D" dans le projet courant
                dossiers3Ds = TSH.Pdm.SearchFolderByName(CurrentProjectPdmId, "02-3D");
                if (dossiers3Ds.Count > 0)
                {
                    Dictionary<PdmObjectId, PdmObjectId> parentMapping = new Dictionary<PdmObjectId, PdmObjectId>();
                    if (ChercherDossierAvecParentMapping(dossiers3Ds[0], PdmObjectIdCurrentDocumentId, out IndiceTxtBox, parentMapping))
                    {
                        return;
                    }
                }
                else
                {
                    throw new Exception("Aucun dossier '02-3D' trouvé.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form { TopMost = true }, "Erreur : " + ex.Message);
            }
        }

        // Fonction récursive pour rechercher le dossier et construire la map des parents
        bool ChercherDossierAvecParentMapping(PdmObjectId dossierActuel, PdmObjectId documentIdRecherche, out string IndiceTxtBox, Dictionary<PdmObjectId, PdmObjectId> parentMapping)
        {
            IndiceTxtBox = "";

            List<PdmObjectId> sousDossiers = new List<PdmObjectId>();
            List<PdmObjectId> documents = new List<PdmObjectId>();
            TSH.Pdm.GetConstituents(dossierActuel, out sousDossiers, out documents);

            // Mappe les sous-dossiers à leur parent actuel
            foreach (PdmObjectId sousDossier in sousDossiers)
            {
                parentMapping[sousDossier] = dossierActuel;
            }

            // Si le document est trouvé
            if (documents.Contains(documentIdRecherche))
            {
                // Remonte jusqu'au dossier commençant par "ind"
                return RemonterJusquAuDossierInd(dossierActuel, out IndiceTxtBox, parentMapping);
            }

            // Recherche récursive dans les sous-dossiers
            foreach (PdmObjectId sousDossier in sousDossiers)
            {
                if (ChercherDossierAvecParentMapping(sousDossier, documentIdRecherche, out IndiceTxtBox, parentMapping))
                {
                    return true;
                }
            }

            return false;
        }

        // Fonction récursive pour remonter jusqu'à un dossier qui commence par "ind"
        bool RemonterJusquAuDossierInd(PdmObjectId dossierActuel, out string IndiceTxtBox, Dictionary<PdmObjectId, PdmObjectId> parentMapping)
        {
            IndiceTxtBox = "";

            // Vérifie le nom du dossier actuel
            string nomDossier = TSH.Pdm.GetName(dossierActuel).ToLower();
            if (nomDossier.StartsWith("ind"))
            {
                char derniereLettre = nomDossier[nomDossier.Length - 1];
                IndiceTxtBox = derniereLettre.ToString().ToUpper();
                return true;
            }

            // Vérifie si le parent existe et continue à remonter
            if (parentMapping.ContainsKey(dossierActuel))
            {
                return RemonterJusquAuDossierInd(parentMapping[dossierActuel], out IndiceTxtBox, parentMapping);
            }

            return false;
        }

        //-----------Fonction Récupération ID Document courant----------------------------------------------------------------------------------------------------------------------------
        void DocumentCourant(out PdmObjectId PdmObjectIdCurrentDocumentId, out DocumentId CurrentDocumentId, out DocumentId CurrentDocumentIdLastRev)
        {
           
            try
            {
                CurrentDocumentId = TSH.Documents.EditedDocument;  // Récupération ID Document courant
                PdmObjectIdCurrentDocumentId = TSH.Documents.GetPdmObject(CurrentDocumentId);
                PdmMinorRevisionId DerniereRev = TSH.Pdm.GetFinalMinorRevision(PdmObjectIdCurrentDocumentId);
                CurrentDocumentIdLastRev = TSH.Documents.GetMinorRevisionDocument(DerniereRev);


            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document courant. Ouvrez un document puis réessayez  " + ex.Message);
                Application.Exit();
                Environment.Exit(0);
                return;

            }
        }

        string indice; 
        void InitialiserTreeView()
        {
            //string indice;
            ChercherDossierDocumentEnCours(PdmObjectIdCurrentDocumentId, out indice);
            listePdf(indice);
        }

        //-----------Fonction Récupération Commentaire----------------------------------------------------------------------------------------------------------------------------
        void RecupCommentaire(in DocumentId CurrentDocumentId, out ElementId CurrentDocumentCommentaireId , out string TextCurrentDocumentCommentaire)
        {
            CurrentDocumentCommentaireId = new ElementId(); // Initialisation avant le bloc try
            TextCurrentDocumentCommentaire = "";
            try
            {
                CurrentDocumentCommentaireId = TSH.Parameters.GetCommentParameter(CurrentDocumentId);   // Récupération du commentaire (Repère)
                TextCurrentDocumentCommentaire = TSH.Parameters.GetTextLocalizedValue(CurrentDocumentCommentaireId);

                textBox2.Text = TextCurrentDocumentCommentaire; //Affichage du commentaire (Repère) dans la case texte
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du Commentaire " + ex.Message);
            }
        }

        class MyDocumentsEventsHost : IDocumentsEvents
        {
            // Variable pour suivre l'état d'édition du document
            private bool isEditing = false;
            // Variable pour la boîte de dialogue
            public Form myDialog;


            public void OnDocumentEditingStarted(DocumentId inDocumentId)
                    {
                        string name = TSH.Documents.GetName(inDocumentId);

                        // Le document est maintenant en mode d'édition
                        isEditing = true;

                        // Si le document est en mode d'édition, fermez la boîte de dialogue
                        if (isEditing && myDialog != null)
                        {
                            myDialog.Close();
                            myDialog = null;
                        }
            }

            public void OnDocumentEditingEnded(DocumentId inDocumentId)
            {
                string name = TSH.Documents.GetName(inDocumentId);

                // Le document n'est plus en mode d'édition
                isEditing = false;
            }

            // Méthode pour afficher la boîte de dialogue
            public void ShowDialog()
            {
                myDialog = new Form { TopMost = true };
                myDialog.Show();
            }
        }

        ///----------------------------Fonction verification dossier indice---------------------------------

        bool VerifDossierIndice(List<PdmObjectId> IndiceFolderIds, out bool FichierExiste)
        {
            // Initialisation de la variable de sortie
            FichierExiste = false;
            nomDocu = textBox2.Text + " Ind " + textBox8.Text + " " + textBox10.Text;

            // Vérification si la liste des IDs de dossier n'est pas vide
            if (IndiceFolderIds.Count != 0)
            {
                // Parcours de la liste des IDs de dossier
                for (int i = 0; i < IndiceFolderIds.Count; i++)
                {
                    // Obtention du nom du dossier à partir de l'ID
                    string IndiceFolderName = TSH.Pdm.GetName(IndiceFolderIds[i]);
                    // Vérification si le nom du dossier correspond à un format spécifique
                    bool test02 = IndiceFolderName.Equals(IndiceTxtFormat00, StringComparison.OrdinalIgnoreCase);
                    bool test03 = IndiceFolderName.Equals(IndiceTxtFormat01, StringComparison.OrdinalIgnoreCase);

                    // Stockage de l'ID du dossier courant
                    PdmObjectId IndiceFolderId = IndiceFolderIds[i];
                    // Si le nom du dossier correspond à l'un des formats spécifiés
                    if (test02 || test03)
                    {
                        // Recherche de documents par nom dans le projet actuel
                        List<PdmObjectId> nomDocuIds = TSH.Pdm.SearchDocumentByName(CurrentProjectPdmId, nomDocu);
                        // Si aucun document n'est trouvé, affichage d'un message
                        if (nomDocuIds.Count > 0)
                        {
                            MessageBox.Show(new Form { TopMost = true }, "Un fichier " + nomDocu + " existe déjà dans le dossier");
                         
                        }
                            return FichierExiste = true;
                        
                    }
                    // Si le fichier existe, la boucle continue sans modifier FichierExiste                      
                }
                   
            }

            return
                    !FichierExiste;
            
            // Retour de la valeur de FichierExiste après la fin de la boucle
            
        }
        
        private void FindParasolidExporterIndex(out int X_TExporterIndex)
        {
            //search the Parasolid exporter index

            X_TExporterIndex = -1;
            for (int i = 0; i < TopSolidHost.Application.ExporterCount; i++)
            {
                TopSolidHost.Application.GetExporterFileType(i, out string fileTypeName, out string[] outFileExtensions);
                if (fileTypeName != "Parasolid") { continue; }

                else
                {
                    X_TExporterIndex = i;
                    break;
                }
            }
        }
            
        private List<KeyValue> VersionX_T(int X_TExporterIndex, string version)
        {
            List<KeyValue> options = TSH.Application.GetExporterOptions(X_TExporterIndex);

            // Modification de la valeur associée à "SAVE_VERSION"
            for (int i = 0; i < options.Count; i++)
            {
                if (options[i].Key == "SAVE_VERSION")
                {
                    options[i] = new KeyValue("SAVE_VERSION", version); // Remplacement par un nouvel objet                                                      //break; // Sortie de la boucle après la modification
                }
            } 
                    return options;
        }

        private void FindNoConvertExporterIndex(out int NoConvertExporterIndex)
        {
            //search exporter sans conversion index

            NoConvertExporterIndex = -1;
            for (int i = 0; i < TopSolidHost.Application.ExporterCount; i++)
            {
                TopSolidHost.Application.GetExporterFileType(i, out string fileTypeName, out string[] outFileExtensions);
                if (fileTypeName != "") { continue; }

                else
                {
                    NoConvertExporterIndex = i;
                    break;
                }
            }
        }

        void CréaetionParam(ElementId parametrePubliedId, in ElementId ParamSytemElementId, in string NomParamTxt, in DocumentId document)
        {
            // Récupère la valeur du paramètre système en texte à partir de son identifiant
            string parametreValueTxt = TSH.Parameters.GetTextValue(ParamSytemElementId);

            // Crée un objet SmartText avec la valeur récupérée
            SmartText parametreValueSmartTxt = new SmartText(ParamSytemElementId);

            // Publie le paramètre texte dans le document spécifié avec le nom et la valeur fournis
            parametrePubliedId = TSH.Parameters.PublishText(document, NomParamTxt, parametreValueSmartTxt);

            // Attribue le nom spécifié à l'identifiant du paramètre publié
            TSH.Elements.SetName(parametrePubliedId, NomParamTxt);

        }

        public Form1()
        {
            InitializeComponent();

            // Charger le chemin sauvegardé et l'afficher dans la TextBox
            textBox4.Text = Properties.Settings.Default.FolderPath;

            //-----------Connexion a TopSolid-----------------------------------------------------------------------------------------------------------------

            bool TSConnected = TopSolidDesignHost.IsConnected;
            {
                try
                {
                    TSH.Connect("Folder Creator Tool");  // Connection à TopSolid
                    TopSolidDesignHost.Connect();    // Connection à TopSolidDesign
                }
                catch (Exception ex)
                {
                    this.TopMost = false;
                    MessageBox.Show("Impossible de se connecter à TopSolid " + ex.Message);
                    return;
                }
            }

            DocumentCourant(out PdmObjectIdCurrentDocumentId,out CurrentDocumentId, out DocumentId CurrentDocumentIdLastRev);

            // Variable pour stocker la dernière lettre du nom du dossier
            string IndiceTxtBox = "";

           //Recuperation l'exporteur parasolid
            FindParasolidExporterIndex(out X_TExporterIndex);

            string versionX_T = "31";
            //Configuartion version parasolid
           options = VersionX_T(X_TExporterIndex, (versionX_T+"0"));

                label7.Text = versionX_T; // Afficher la version dans le formulaire
  

            //----------- Récupération du nom du document courant----------------------------------------------------------------------------------------------------------------------------         
            try
            {
                string CurrentDocumentName = TSH.Documents.GetName(CurrentDocumentId);  // Récupération du nom du document courant
                TextCurrentDocumentName = CurrentDocumentName;

                textBox9.Text = TextCurrentDocumentName; //Affichage du nom du document courent dans la case texte
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du nom du document courant " + ex.Message);
            }

            // Appel de la méthode pour vérifier le chemin au démarrage
            VerifierCheminAuDemarrage();//-----------------------------------------------------------------------------////////////////////////////////////////---------------------------------------------------------

            // Restaurer le choix de matière au démarrage
            RestoreMaterialChoice();

            //-----------Récupération ID projet courant----------------------------------------------------------------------------------------------------------------------------
            try
            {
                CurrentProjectPdmId = TSH.Pdm.GetProject(PdmObjectIdCurrentDocumentId);   // Récupération ID projet courant
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du projet courant " + ex.Message);
                return;
            }

            //-----------Récupération Nom projet courant----------------------------------------------------------------------------------------------------------------------------
            try
            {
                CurrentProjectName = TSH.Pdm.GetName(CurrentProjectPdmId);  // Récupération Nom projet
                TextBoxProjectName = CurrentProjectName;

                textBox10.Text = TextBoxProjectName; //Affichage du nom du projet courent dans la case texte
                textBox1.Text = TextBoxProjectName; //Affichage le N° de moule dans la case texte

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du projet courant " + ex.Message);
            }
           
            // Appel de la fonction pour chercher le dossier contenant
            ChercherDossierDocumentEnCours(PdmObjectIdCurrentDocumentId, out IndiceTxtBox);

            //-------------Creation de la variable pour la recherche du dossier atelier-------------------------------------------------------------------------------------------------------------------

            try
            {
                AtelierFolderIds = TSH.Pdm.SearchFolderByName(CurrentProjectPdmId, "02-Atelier");
                AtelierFolderId = AtelierFolderIds[0];

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Dossier ''02-Atelier'' introuvable dans le projet " + ex.Message);
            }

            //------------- Récupération du commentaire (Repère) du document courant----------------------------------------------------------------------------------------------------------------------------

            RecupCommentaire(in CurrentDocumentId, out CurrentDocumentCommentaireId, out TextCurrentDocumentCommentaire);

            //----------- Récupération de la désignation du document courant----------------------------------------------------------------------------------------------------------------------------

            try
            {
                CurrentDocumentDesignationId = TSH.Parameters.GetDescriptionParameter(CurrentDocumentId);   // Récupération de la désignation
                TextCurrentDocumentDesignation = TSH.Parameters.GetTextLocalizedValue(CurrentDocumentDesignationId);

                textBox3.Text = TextCurrentDocumentDesignation; //Affichage du commentaire (Repère) dans la case texte


            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du Commentaire " + ex.Message);
            }

            //----------- Variable des differents façon de nommer le dossier indice----------------------------------------------------------------------------------------------------------------------------

            textBox8.Text = IndiceTxtBox; //Affichage de l'indice

            //Liste PDF--------------------------------

            listePdf(IndiceTxtBox);

            string DossierIndicePdf = "Ind " + IndiceTxtBox;

            // Trouver et déplier le nœud cible
            ExpandTargetNode(rootFolderNode, DossierIndicePdf);


            //treeView1.ExpandAll();
            // Fonction pour ajouter les nœuds cochés à la liste


        }

        //------------------------------Bouton click dossier-------------

        private void button2_Click_1(object sender, EventArgs e)
        {
            string celluleVideErreur = "Merci de remplir toute les cases";
            
            if (textBox2.Text == string.Empty)
            {
                MessageBox.Show ( celluleVideErreur);
                return;
            }
            if (textBox3.Text == string.Empty)
            {
                MessageBox.Show(celluleVideErreur);
                return;
            }
            if (textBox8.Text == string.Empty)
            {
                MessageBox.Show(celluleVideErreur);
                return;
            }

            // Récupère le texte de la TextBox
            string DossierAtelierServeur = textBox4.Text;

            // Vérifie si le chemin est valide et s'il existe
            if (Directory.Exists(DossierAtelierServeur))
            {
                        // Si le chemin est valide, continue avec le reste du programme
                        // Ton code ici

                    // Initialisation des listes
                    List<string> TxtCheckedItems = new List<string>();
                    List<string> CheckedItemsIdTxt = new List<string>();
                    List<PdmObjectId> CheckedItems = new List<PdmObjectId>();
                    List<PdmObjectId> CheckedItemListeCopie = new List<PdmObjectId>();
                    List<PdmObjectId> CheckedItemsliste = new List<PdmObjectId>();
                    List<PdmObjectId> CheckedItemCopieListe = new List<PdmObjectId>();

                    DialogResult resulta = MessageBox.Show("Voulez vous simplifier la piece", "Confirmation", MessageBoxButtons.YesNo);

                    if (resulta == DialogResult.Yes)
                    {
                    
                        TSH.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Shapes.Healing.HealCommand");
                    // Redémarre l'application
                    Application.Exit();
                    }

                    // Ajout des nœuds cochés à la liste CheckedItems
                    try
                    {
                        AddCheckedNodesToList(treeView1.Nodes, CheckedItems);

                        // Parcours de la liste CheckedItems
                        for (int i = 0; CheckedItems.Count > i; i++)
                        {
                            string ExtentionTxtPdf = "";
                            PdmObjectId CheckedItem = CheckedItems[i];
                            PdmObjectType TypePDF = TSH.Pdm.GetType(CheckedItem, out ExtentionTxtPdf);
                            // Si l'extension du fichier est .pdf, ajout à la liste CheckedItemsliste
                            if (ExtentionTxtPdf == ".pdf")
                            {
                                CheckedItemsliste.Add(CheckedItem);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.TopMost = false;
                        MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de la liste des PDF " + ex.Message);
                    }

                    ConstituantFolderNames.Clear();
                    ListFoldersNames.Clear();

                    bool recommencer;
                    recommencer = false; // Réinitialisez recommencer à false à chaque début de boucle
                                         //Récuperation des noms de dossiers

                    do
                    {
                        //Recuperation du texte modifié par l'utilisateur pour nommer les dossiers.
                        TextBoxCommentaireValue = textBox2.Text; //Repere de la piece
                        TextBoxIndiceValue = textBox8.Text; //Indice de la piece
                        TextBoxDesignationValue = textBox3.Text; //Designation de la piece
                        TextBoxNomMouleValue = textBox10.Text; //Numero du moule
                        TexteIndiceFolder = "Ind " + TextBoxIndiceValue;
                        TexteDossierRep = textBox2.Text + " - " + textBox3.Text; //Nom du dossier repere

                        IndiceTxtFormat00 = "Ind" + textBox8.Text;
                        IndiceTxtFormat01 = "Ind " + textBox8.Text;


                        CommentaireTxtFormat00 = textBox2.Text + "-";
                        CommentaireTxtFormat01 = textBox2.Text + " ";

                        nomDocu = textBox2.Text + " Ind " + textBox8.Text + " " + textBox10.Text;


                        try
                        {
                            //Activation des modifications
                            modifActif(CurrentDocumentId);

                            // Récupération à nouveau des informations du document actuel
                            DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                            //------------- Récupération du commentaire (Repère) du document courant----------------------------------------------------------------------------------------------------------------------------
                            CurrentDocumentCommentaireId = TSH.Parameters.GetCommentParameter(CurrentDocumentId);   // Récupération du commentaire (Repère)
                            TextCurrentDocumentCommentaire = TSH.Parameters.GetTextLocalizedValue(CurrentDocumentCommentaireId);
                            //RecupCommentaire(in CurrentDocumentId, out CurrentDocumentCommentaireId, out TextCurrentDocumentCommentaire);

                            //----------- Récupération de la désignation du document courant----------------------------------------------------------------------------------------------------------------------------
                            try
                            {
                                CurrentDocumentDesignationId = TSH.Parameters.GetDescriptionParameter(CurrentDocumentId);   // Récupération de la désignation
                                TextCurrentDocumentDesignation = TSH.Parameters.GetTextLocalizedValue(CurrentDocumentDesignationId);

                                textBox3.Text = TextCurrentDocumentDesignation; //Affichage du commentaire (Repère) dans la case texte


                            }
                            catch (Exception ex)
                            {
                                this.TopMost = false;
                                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du Commentaire " + ex.Message);
                            }


                            // Mise à jour des valeurs de commentaire et de désignation du document
                            TSH.Parameters.SetTextValue(CurrentDocumentCommentaireId, TextBoxCommentaireValue);
                            TSH.Parameters.SetTextValue(CurrentDocumentDesignationId, TextBoxDesignationValue);

                            // Appeler la méthode de vérification pendant l'initialisation ou un événement

                            try
                            {
                                // Initialisation
                                string domain = "jbtecnics";
                                string uName1 = "acierjbt";
                                string uName2 = "aciertrempejbt";

                                PdmObjectId MaterialAcierJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName1);
                                DocumentId MaterialAcierJbtId = TSH.Documents.GetDocument(MaterialAcierJbt);

                                PdmObjectId MaterialAcierTrempeJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName2);
                                DocumentId MaterialAcierTrempeJbtId = TSH.Documents.GetDocument(MaterialAcierTrempeJbt);

                                // Appeler la vérification pour s'assurer qu'une matière est sélectionnée dès le départ
                                CheckAndSetMaterial(MaterialAcierJbtId, MaterialAcierTrempeJbtId);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur pendant l'initialisation : " + ex.Message);
                            }





                        // Fin des modifications avec sauvegarde des changements
                        TSH.Application.EndModification(true, true);
                        }
                        catch (Exception ex)
                        {
                            this.TopMost = false;
                            // Annulation des modifications en cas d'erreur
                            TSH.Application.EndModification(true, true);
                            // Affichage d'un message d'erreur en cas d'échec de l'édition du commentaire et de la désignation
                            MessageBox.Show(new Form { TopMost = true }, "erreur lors de l'edition du commentaire et de la désignation du document " + ex.Message);
                            return;
                        }


                        //--------------Recuperation des dossier du projet--------------------------
                        TSH.Pdm.GetConstituents(AtelierFolderId, out FolderIds, out DocumentsIds);

                        try
                        {
                            bool aucunDossierProjet = true;
                            bool FichierExiste = false;
                            bool BesoinDeTousLesDossier = false;
                            if (FolderIds.Count != 0)
                            {
                                aucunDossierProjet = false;

                                for (int i = 0; i < FolderIds.Count; i++) //Boucle de décompte
                                {
                                    // Obtention du nom du dossier actuel dans la boucle
                                    ConstituantFolderName = TSH.Pdm.GetName(FolderIds[i]);

                                    // Vérification si le nom du dossier commence par un format de commentaire spécifique
                                    test00 = ConstituantFolderName.StartsWith(CommentaireTxtFormat00, StringComparison.OrdinalIgnoreCase);
                                    test01 = ConstituantFolderName.StartsWith(CommentaireTxtFormat01, StringComparison.OrdinalIgnoreCase);

                                    // Stockage de l'ID du dossier existant
                                    DossierExistantId = FolderIds[i];

                                    // Vérification si un seul des tests est vrai (opérateur XOR)
                                    if (test00 ^ test01)
                                    {
                                        // Si le nom du dossier à créer est différent du nom du dossier existant
                                        if (TexteDossierRep != ConstituantFolderName)
                                        {
                                            // Affichage d'un message d'avertissement à l'utilisateur concernant un doublon potentiel
                                            DialogResult result = MessageBox.Show("Un dossier existe avec le même repère mais avec une désignation différente." + Environment.NewLine + "Merci de vérifier et de corriger avant de continuer " + Environment.NewLine + "Nom du dossier détecté = " + ConstituantFolderName + Environment.NewLine + "Nom du dossier qui doit être créé : " + TexteDossierRep, "Doublon potentiel ", MessageBoxButtons.RetryCancel);

                                            // Si l'utilisateur choisit de réessayer, la variable 'recommencer' est définie sur true
                                            if (result == DialogResult.Retry)
                                            {
                                                recommencer = true; // La boucle while recommencera
                                            }
                                            // Si l'utilisateur annule, la fonction retourne et arrête l'exécution
                                            if (result == DialogResult.Cancel)
                                            {
                                                return;
                                            }
                                        }

                                        // Si le nom du dossier à créer est identique au nom du dossier existant
                                        if (TexteDossierRep == ConstituantFolderName)
                                        {
                                            // Affiche un message indiquant que le dossier existe déjà
                                            MessageBox.Show("le dossier " + TexteDossierRep + " existe déjà. Recherche du dossier d'indice");

                                            // Récupère les constituants du dossier existant
                                            TSH.Pdm.GetConstituents(DossierExistantId, out IndiceFolderIds, out DocumentsInIndiceFolder);

                                            // Vérifie l'existence du dossier indice
                                            VerifDossierIndice(IndiceFolderIds, out FichierExiste);
                                        }
                                        if (FichierExiste)
                                        {
                                            return;
                                        }
                                        // Si le fichier n'existe pas, crée un nouveau dossier
                                        if (!FichierExiste)
                                        {
                                            DossierIndiceId = TSH.Pdm.CreateFolder(DossierExistantId, TexteIndiceFolder);
                                            dossier3DGenereId = creationAutreDossiers(DossierIndiceId);
                                            DossierRepId = DossierExistantId;
                                            break;
                                        }
                                        else
                                            return;
                                    }

                                }

                            }
                            if ((!test00 && !test01) || BesoinDeTousLesDossier || aucunDossierProjet)
                            {
                                DossierRepId = TSH.Pdm.CreateFolder(AtelierFolderId, TexteDossierRep);
                                DossierIndiceId = TSH.Pdm.CreateFolder(DossierRepId, TexteIndiceFolder);
                                dossier3DGenereId = creationAutreDossiers(DossierIndiceId);
                            }
                        }
                        catch (Exception ex)
                        {
                            this.TopMost = false;
                            MessageBox.Show(new Form { TopMost = true }, "erreur" + ex.Message);
                        }


                    }
                    while (recommencer); // La boucle while recommencera si recommencer est true

                    try
                    {
                        try
                        {
                            // Récupération de l'ID du propriétaire du document actuel
                            AuteurPdmObjectId = TSH.Pdm.GetOwner(PdmObjectIdCurrentDocumentId);
                            // Création d'un document dérivé et récupération de son ID
                            DerivéDocumentId = TopSolidDesignHost.Tools.CreateDerivedDocument(AuteurPdmObjectId, CurrentDocumentId, false);

                            // Récupération de l'ID Pdm du document dérivé
                            DerivéDocumentPdmObjectId = TSH.Documents.GetPdmObject(DerivéDocumentId);
                            // Ajout de l'ID Pdm du document dérivé à la liste
                            DerivéDocumentPdmObjectIds.Add(DerivéDocumentPdmObjectId);

                            // Sauvegarde du document actuel
                            TSH.Documents.Save(CurrentDocumentId);
                            // Fermeture du document actuel
                            TSH.Documents.Close(CurrentDocumentId, false, false);
                            // Ouverture du document dérivé
                            TSH.Documents.Open(ref DerivéDocumentId);
                            PdmObjectId DerivéDocumentPdmId = TSH.Documents.GetPdmObject(DerivéDocumentId);


                        }
                        catch (Exception ex)
                        {
                            this.TopMost = false;
                            // Affichage d'un message d'erreur en cas d'échec de récupération de l'ID du document dérivé
                            MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document dérivé " + ex.Message);
                            return;
                        }

                        try
                        {

                            //DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                            modifActif(CurrentDocumentId);

                            DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                            List<ElementId> OtherSystemParameters = new List<ElementId>();
                            TopSolidDesignHost.Tools.SetDerivationInheritances(
                                               CurrentDocumentIdLastRev, // Identifiant de la dernière révision du document courant
                                                false, // inName
                                                true,  // inDescription
                                                false,  // inCode
                                                true,  // inPartNumber
                                                false,  // IinComplementaryPartNumber
                                                false,  // inManufacturer
                                                true,  //inManufacturerPartNumber
                                                true,  // inComment
                                                OtherSystemParameters,  // Inherit design parameters (null means default value)
                                                false,  // inNonSystemParameters
                                                false,  // inPoints
                                                false,  // inAxes
                                                false,  // inPlanes
                                                false,  // inFrames
                                                false,  // inSketches
                                                true,  // inShapes
                                                true,  // inPublishings
                                                false,  // inFunctions
                                                false,  // inSymmetries
                                                false,  // inUnsectionabilities
                                                true,  // inRepresentations
                                                false,  // inSets
                                                false  // inCameras
                                            );

                            TSH.Application.EndModification(true, true);
                        }
                        catch (Exception ex)
                        {
                            this.TopMost = false;
                            // Affichage d'un message d'erreur en cas d'échec de la dérivation
                            MessageBox.Show(new Form { TopMost = true }, "Erreur a l'edition des parametre de derivation" + ex.Message);
                            // Annulation des modifications en cas d'erreur
                            TSH.Application.EndModification(false, false);
                            return;
                        }

                        try
                        {
                            // Copie des PDF dans le dossier si la liste n'est pas vide
                            if (CheckedItemsliste.Count > 0)
                            {
                                // Copie des PDF sélectionnés dans le dossier
                                CheckedItemListeCopie = TSH.Pdm.CopySeveral(CheckedItemsliste, AuteurPdmObjectId);
                                // Déplacement du document dérivé et des PDF
                                TSH.Pdm.MoveSeveral(CheckedItemListeCopie, dossier3DGenereId);
                            }
                            // Déplacement des ID Pdm du document dérivé
                            TSH.Pdm.MoveSeveral(DerivéDocumentPdmObjectIds, dossier3DGenereId);
                        }
                        catch (Exception ex)
                        {
                            // Affichage d'un message d'erreur en cas d'échec de la copie ou du déplacement des PDF
                            MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la copie ou du deplacement du PDF dans le dossier 3D" + ex.Message);
                            return;
                        }

                        // Récupération des informations du document actuel et activation des modifications
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                        // Appelle une méthode personnalisée pour récupérer les identifiants du document actuel, 
                        // y compris l'identifiant de la dernière révision du document.

                        modifActif(CurrentDocumentId);
                        // Active les modifications sur le document actuel.

                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                        // Répète l'appel à la méthode pour mettre à jour les informations du document.

                        TSH.Documents.SetName(CurrentDocumentIdLastRev, nomDocu);
                        // Change le nom du document (dernière révision) avec la nouvelle valeur fournie par 'nomDocu'.

                        ElementId Indice3DNomParamId = TSH.Parameters.CreateTextParameter(CurrentDocumentId, TextBoxIndiceValue);
                        // Crée un paramètre texte dans le document actuel avec la valeur spécifiée et stocke son identifiant.
                        string Indice3DNomParamTxt = "Indice 3D";
                        // Définit une chaîne de caractères pour le nom du paramètre texte.
                        SmartText Indice3DNomParam = new SmartText(Indice3DNomParamId);
                        // Crée un objet SmartText à partir de la valeur contenue dans 'TextBoxIndiceValue'.
                        TSH.Elements.SetName(Indice3DNomParamId, Indice3DNomParamTxt);
                        // Attribue un nom au paramètre texte créé précédemment.
                        ElementId publishedIndice3DNomParamId = TSH.Parameters.PublishText(CurrentDocumentId, Indice3DNomParamTxt, Indice3DNomParam);
                        // Publie le paramètre texte 'Indice 3D' dans le document actuel et récupère l'identifiant de l'entité publiée.
                        TSH.Elements.SetName(publishedIndice3DNomParamId, Indice3DNomParamTxt);
                        // Attribue le nom 'Indice 3D' à l'entité publiée.



                         ElementId matierePlanParamId = TSH.Parameters.CreateTextParameter(CurrentDocumentId, textBox5.Text);
                        // Crée un paramètre texte dans le document actuel avec la valeur spécifiée et stocke son identifiant.
                        string matiereNomParamTxt = "Matiére plan";
                        // Définit une chaîne de caractères pour le nom du paramètre texte.
                        SmartText matierePlanParam = new SmartText(matierePlanParamId);
                        // Crée un objet SmartText à partir de la valeur contenue dans 'TextBoxIndiceValue'.
                        TSH.Elements.SetName(matierePlanParamId, matiereNomParamTxt);
                        // Attribue un nom au paramètre texte créé précédemment.
                        ElementId publishedMatierePlanParamId = TSH.Parameters.PublishText(CurrentDocumentId, matiereNomParamTxt, matierePlanParam);
                        // Publie le paramètre texte 'Indice 3D' dans le document actuel et récupère l'identifiant de l'entité publiée.
                        TSH.Elements.SetName(publishedMatierePlanParamId, matiereNomParamTxt);
                        // Attribue le nom 'Indice 3D' à l'entité publiée.

                         ElementId traitementParamId = TSH.Parameters.CreateTextParameter(CurrentDocumentId, textBox6.Text);
                        // Crée un paramètre texte dans le document actuel avec la valeur spécifiée et stocke son identifiant.
                        string traitementNomParamTxt = "Traitement";
                        // Définit une chaîne de caractères pour le nom du paramètre texte.
                        SmartText traitementParam = new SmartText(traitementParamId);
                        // Crée un objet SmartText à partir de la valeur contenue dans 'TextBoxIndiceValue'.
                        TSH.Elements.SetName(traitementParamId, traitementNomParamTxt);
                        // Attribue un nom au paramètre texte créé précédemment.
                        ElementId publishedTraitementParamId = TSH.Parameters.PublishText(CurrentDocumentId, traitementNomParamTxt, traitementParam);
                        // Publie le paramètre texte 'Indice 3D' dans le document actuel et récupère l'identifiant de l'entité publiée.
                        TSH.Elements.SetName(publishedTraitementParamId, traitementNomParamTxt);
                        // Attribue le nom 'Indice 3D' à l'entité publiée.

                         ElementId nbrPiecesParamId = TSH.Parameters.CreateTextParameter(CurrentDocumentId, textBox7.Text);
                        // Crée un paramètre texte dans le document actuel avec la valeur spécifiée et stocke son identifiant.
                        string nbrPiecesNomParamTxt = "Nombre de piéces";
                        // Définit une chaîne de caractères pour le nom du paramètre texte.
                        SmartText nbrPiecesParam = new SmartText(nbrPiecesParamId);
                        // Crée un objet SmartText à partir de la valeur contenue dans 'TextBoxIndiceValue'.
                        TSH.Elements.SetName(nbrPiecesParamId, nbrPiecesNomParamTxt);
                        // Attribue un nom au paramètre texte créé précédemment.
                        ElementId publishedNbrPiecesParamId = TSH.Parameters.PublishText(CurrentDocumentId, nbrPiecesNomParamTxt, nbrPiecesParam);
                        // Publie le paramètre texte 'Indice 3D' dans le document actuel et récupère l'identifiant de l'entité publiée.
                        TSH.Elements.SetName(publishedNbrPiecesParamId, nbrPiecesNomParamTxt);
                        // Attribue le nom 'Indice 3D' à l'entité publiée.

                        // Récupération de l'identifiant du paramètre de description du système
                        ElementId DesignationSystemeId = TSH.Parameters.GetDescriptionParameter(CurrentDocumentId);
                        // Définition du nom du paramètre de description
                        string DesignationNomParam = "Designation";
                        // Création du paramètre de description avec l'identifiant récupéré
                        ElementId PubliedDesignationSystemeId = new ElementId();
                        CréaetionParam(PubliedDesignationSystemeId, in DesignationSystemeId, in DesignationNomParam, in CurrentDocumentId);

                        // Récupération de l'identifiant du paramètre de commentaire du système
                        ElementId CommentaireSystemeId = TSH.Parameters.GetCommentParameter(CurrentDocumentId);
                        // Définition du nom du paramètre de commentaire
                        string CommentaireNomParam = "Commentaire";
                        // Création du paramètre de commentaire avec l'identifiant récupéré
                        ElementId PubliedCommentaireSystemeId = new ElementId();
                        CréaetionParam(PubliedCommentaireSystemeId, in CommentaireSystemeId, in CommentaireNomParam, in CurrentDocumentId);

                        // Récupération de l'identifiant du paramètre de nom du document
                        ElementId Nom_docu = TSH.Parameters.GetNameParameter(CurrentDocumentId);
                        // Définition du nom du paramètre de nom du document
                        string NomDocuNomParam = "Nom_docu";
                        // Création du paramètre de nom du document avec l'identifiant récupéré
                        ElementId PubliedNom_docu = new ElementId();

                        //Modifiction tolérence de visualisation
                        CréaetionParam(PubliedNom_docu, in Nom_docu, in NomDocuNomParam, in CurrentDocumentId);
                        double LinearTol = 0.00001;
                        double AngularTol = 0.08726646259971647;
                        TSH.Options.SetVisualizationTolerances(CurrentDocumentIdLastRev, LinearTol, AngularTol);

                        

                        // Fin des modifications avec sauvegarde des changements
                        TSH.Application.EndModification(true, true);
                    }
                    catch (Exception ex)
                    {
                        this.TopMost = false;
                        // Affichage d'un message d'erreur en cas d'échec de la dérivation
                        MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la dérivation" + ex.Message);
                        // Annulation des modifications en cas d'erreur
                        TSH.Application.EndModification(false, false);
                        return;
                    }


                    //-------------- transfo sur rep abs ------------

                    this.TopMost = false;
                    this.WindowState = FormWindowState.Minimized;

                    // Initialisation de la variable qui recevra le repère sélectionné par l'utilisateur
                    SmartFrame3D ReponseRepereUser = null;

                    try
                    {
                        //DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                        modifActif(CurrentDocumentId);
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);


                        // Boucle demandant à l'utilisateur de sélectionner le plan XY du repère jusqu'à obtenir une réponse
                        while (ReponseRepereUser == null)
                        {
                            string titre = "Repére";
                            string label = "Créer repére";
                            UserQuestion QuestionPlan = new UserQuestion(titre, label);
                            QuestionPlan.AllowsCreation = true;
                            TopSolidHost.User.AskFrame3D(QuestionPlan, true, null, out ReponseRepereUser);
                        }

                        // Fin de la modification du document
                        TopSolidHost.Application.EndModification(true, true);
                    }
                    catch (Exception ex)
                    {
                        // En cas d'erreur lors de la transformation, affichage d'un message d'erreur
                        this.TopMost = false;
                        TopSolidHost.Application.EndModification(false, false);
                        MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la creation du repere " + ex.Message);
                        return;
                    }


                    ElementId RepereUserId = ReponseRepereUser.ElementId;

                    // Création d'un Frame3D pour stocker le repère utilisateur
                    Frame3D RepereUser = TopSolidHost.Geometries3D.GetFrameGeometry(RepereUserId);

                    // Récupération du repère absolu
                    ElementId AbsRepId = TopSolidHost.Geometries3D.GetAbsoluteFrame(DerivéDocumentId);
                    Frame3D AbsRep = TopSolidHost.Geometries3D.GetFrameGeometry(AbsRepId);

                    Direction3D dxABS = AbsRep.XDirection;
                    Direction3D dyABS = AbsRep.YDirection;
                    Direction3D dzABS = AbsRep.ZDirection;
                    //Direction3D dzABS = new Direction3D(AbsRep.ZDirection.X, AbsRep.ZDirection.Y, AbsRep.ZDirection.Z);

                    // Récupération des directions des axes du repère utilisateur
                    Direction3D dx = RepereUser.XDirection;
                    Direction3D dy = RepereUser.YDirection;
                    Direction3D dz = RepereUser.ZDirection;
                    //Direction3D dz = new Direction3D(RepereUser.ZDirection.X, RepereUser.ZDirection.X, RepereUser.ZDirection.X);

                    bool isXParallelAndSameDirection = false;
                    bool isYParallelAndSameDirection = false;
                    bool isZParallelAndSameDirection = false;
                    bool areDirectionsParallelAndSameDirection = false;

                    // Vérification si les directions du repère utilisateur sont parallèles et dans le même sens que les directions du repère absolu
                    isXParallelAndSameDirection = dx.IsParallelTo(dxABS, true);
                    isYParallelAndSameDirection = dy.IsParallelTo(dyABS, true);
                    isZParallelAndSameDirection = dz.IsParallelTo(dzABS, true);

                    areDirectionsParallelAndSameDirection = isXParallelAndSameDirection && isYParallelAndSameDirection && isZParallelAndSameDirection;


                    // Calcul de la différence de position entre les repères
                    Point3D originUser = RepereUser.Origin;
                    Point3D originAbs = AbsRep.Origin;

                    double Tx = originAbs.X - originUser.X;
                    double Ty = originAbs.Y - originUser.Y;
                    double Tz = originAbs.Z - originUser.Z;

                    // Vérification si les origines sont identiques
                    bool areOriginsIdentical = (originUser.X == originAbs.X) && (originUser.Y == originAbs.Y) && (originUser.Z == originAbs.Z);


                    // Translation pour ramener l'origine du repère utilisateur à l'origine du repère absolu
                    Transform3D translation = new Transform3D(
                        dxABS.X, 0, 0, Tx,
                        0, dyABS.Y, 0, Ty,
                        0, 0, dzABS.Z, Tz,
                        0, 0, 0, 1
                        );

                    // Rotation pour ramener les axes paralelles
                    Transform3D Rotation = new Transform3D(
                        dx.X, dx.Y, dx.Z, 0,
                        dy.X, dy.Y, dy.Z, 0,
                        dz.X, dz.Y, dz.Z, 0,
                        0, 0, 0, 1
                    );

                    //--------------------------------------------------------------------------------------------------

                    // Recherche du dossier Formes dans le document
                    ElementId DossierForme = TSH.Elements.SearchByName(DerivéDocumentId, "$TopSolid.Kernel.DB.D3.Shapes.Documents.ElementName.Shapes");

                    // Récupération de tous les éléments du dossier Formes
                    List<ElementId> FormesList = TSH.Elements.GetConstituents(DossierForme);

                    try
                    {
                        // Activation de la modification du document courant
                        modifActif(CurrentDocumentId);
                        // Récupération des identifiants du document courant
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                        // Si les directions sont parallèles et dans le même sens mais que les origines ne sont pas identiques
                        if (areDirectionsParallelAndSameDirection && !areOriginsIdentical)
                        {
                            // Appliquer une translation à chaque forme de la liste
                            for (int i = 0; i < FormesList.Count; i++)
                            {
                                TSH.Entities.Transform(FormesList[i], translation);
                            }
                        }
                        // Si les directions ne sont pas parallèles mais que les origines sont identiques
                        if (!areDirectionsParallelAndSameDirection && areOriginsIdentical)
                        {
                            // Appliquer une rotation à chaque forme de la liste
                            for (int i = 0; i < FormesList.Count; i++)
                            {
                                TSH.Entities.Transform(FormesList[i], Rotation);
                            }
                        }
                        // Si les directions ne sont pas parallèles et que les origines ne sont pas identiques
                        if (!areDirectionsParallelAndSameDirection && !areOriginsIdentical)
                        {
                            // Appliquer une translation suivie d'une rotation à chaque forme de la liste
                            for (int i = 0; i < FormesList.Count; i++)
                            {
                                TSH.Entities.Transform(FormesList[i], translation);
                                TSH.Entities.Transform(FormesList[i], Rotation);
                            }
                        }

                        // Terminer la modification du document avec succès
                        TSH.Application.EndModification(true, true);
                    }
                    catch (Exception ex)
                    {
                        // En cas d'erreur lors de la transformation, terminer la modification sans sauvegarder et afficher un message d'erreur
                        this.TopMost = false;
                        TSH.Application.EndModification(false, false);
                        MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la transformation : " + ex.Message);
                        return;
                    }


                    try
                    {
                        // Active la modification du document actuel
                        modifActif(CurrentDocumentId);

                        // Récupère les identifiants du document actuel
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                        // Récupère tous les identifiants d'éléments du document actuel
                        List<ElementId> TotalElementIds = TSH.Elements.GetElements(CurrentDocumentId);

                        // Cache tous les éléments du document
                        for (int i = 0; i < TotalElementIds.Count; i++)
                        {
                            TSH.Elements.Hide(TotalElementIds[i]);
                        }

                        // Affiche uniquement les éléments spécifiés dans FormesList
                        for (int i = 0; i < FormesList.Count; i++)
                        {
                            TSH.Elements.Show(FormesList[i]);
                        }

                        // Termine la modification avec succès
                        TSH.Application.EndModification(true, true);
                    }
                    catch (Exception ex)
                    {
                        // En cas d'erreur lors de la transformation
                        this.TopMost = false;
                        MessageBox.Show(new Form { TopMost = true }, "Erreur lors de voir seulement " + ex.Message);

                        // Termine la modification sans sauvegarder les changements
                        TSH.Application.EndModification(false, false);
                        return;
                    }


                    try
                    {
                        // Active la modification du document actuel
                        modifActif(CurrentDocumentId);

                        // Récupère les identifiants du document actuel
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                        // Obtient l'identifiant de la vue graphique active
                        int VueActiveInt = TSH.Visualization3D.GetActiveView(CurrentDocumentId);

                        // Récupère les paramètres actuels de la caméra de perspective
                        ElementId PerspectiveCamera = TSH.Visualization3D.GetPerspectiveCamera(CurrentDocumentId);
                        Point3D outEyePosition = new Point3D();
                        Direction3D outLookAtDirection = new Direction3D();
                        Direction3D outUpDirection = new Direction3D();
                        double outFieldAngle = new double();
                        double outFieldRadius = new double();
                        TSH.Visualization3D.GetCameraDefinition(PerspectiveCamera,
                                                                out outEyePosition,
                                                                out outLookAtDirection,
                                                                out outUpDirection,
                                                                out outFieldAngle,
                                                                out outFieldRadius);

                        // Définit la caméra de la vue active avec les paramètres récupérés
                        TSH.Visualization3D.SetViewCamera(CurrentDocumentId,
                                                            VueActiveInt,
                                                            outEyePosition,
                                                            outLookAtDirection,
                                                            outUpDirection,
                                                            outFieldAngle,
                                                            outFieldRadius);

                        // Effectue un zoom pour adapter la vue à l'écran
                        TSH.Visualization3D.ZoomToFitView(CurrentDocumentId, VueActiveInt);

                        // Termine la modification avec succès
                        TSH.Application.EndModification(true, true);
                    }
                    catch (Exception ex)
                    {
                        // En cas d'erreur lors de la récupération ou de la mise à jour de la vue
                        this.TopMost = false;
                        MessageBox.Show(new Form { TopMost = true }, "Erreur camera " + ex.Message);

                        // Termine la modification sans sauvegarder les changements
                        TSH.Application.EndModification(false, false);
                        return;
                    }




                    try
                    {
                        // Affiche le document dans l'arborescence du projet TopSolid
                        TSH.Pdm.ShowInProjectTree(PdmObjectIdCurrentDocumentId);

                        // Active la modification du document actuel
                        modifActif(CurrentDocumentId);


                        // Récupère les identifiants du document actuel
                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                        // Termine la modification avec succès
                        TSH.Application.EndModification(true, true);
                        PdmMinorRevisionId minorRevisionId = TSH.Pdm.GetFinalMinorRevision(PdmObjectIdCurrentDocumentId);

                        TSH.Documents.Open(ref CurrentDocumentIdLastRev);

                        // Sauvegarde du document actuel
                        TSH.Documents.Save(CurrentDocumentIdLastRev);

                        // Enregistre le document dans le dossier de stockage
                        TSH.Pdm.CheckIn(DossierRepId, true);

                        // Définit l'état du cycle de vie du document sur "Validé"
                        TSH.Pdm.SetLifeCycleMainState(PdmObjectIdCurrentDocumentId, PdmLifeCycleMainState.Validated);

                        // Quitte l'application
                        //Environment.Exit(0);
                    }
                    catch (Exception ex)
                    {
                        // En cas d'erreur, affiche un message et termine la modification sans sauvegarder
                        this.TopMost = false;
                        MessageBox.Show(new Form { TopMost = true }, "Erreur Update, Mise au coffre ou validation " + ex.Message);

                        // Termine la modification sans sauvegarder les changements
                        TSH.Application.EndModification(false, false);
                        return;
                    }

                    /////////Export////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    ///


                    // Initialisation de la variable qui indique si le dossier Atelier existe sur le serveur
                    bool DossierAtelierServeurExiste = false;

                    // Initialisation du tableau qui contiendra les chemins des dossiers trouvés
                    string[] directories = new string[0];

                    // Initialisation du tableau qui contiendra les chemins des dossiers commençant par DossierRep
                    string[] startDirectories = new string[0];

                    // Affichage d'une boîte de dialogue demandant à l'utilisateur s'il souhaite exporter les fichiers
                    DialogResult dialogResult = MessageBox.Show("Souhaitez-vous exporter les fichiers dans le dossier Atelier",
                                                    "Confirmation",
                                                    MessageBoxButtons.YesNo,
                                                    MessageBoxIcon.Information,
                                                    MessageBoxDefaultButton.Button1,
                                                    MessageBoxOptions.DefaultDesktopOnly);



                    // Récupération du nom du dossier à partir de la TextBox
                    string folderName = TextBoxNomMouleValue;

                    bool ExporterFichiers = false;
                    string path3D = "";
                    try
                    {
                        if (dialogResult == DialogResult.Yes)
                        {
                            // Si l'utilisateur choisit "Oui", le code d'exportation des fichiers est exécuté

                            // Recherche du dossier dans le chemin spécifié
                            directories = System.IO.Directory.GetDirectories(DossierAtelierServeur, folderName, System.IO.SearchOption.TopDirectoryOnly);

                            if (directories.Length > 0)
                            {
                                // Si le dossier est trouvé, DossierAtelierServeurExiste est défini sur true
                                DossierAtelierServeurExiste = true;
                            }
                            else
                            {
                                //Si non si aucun dossier existe
                                // Création du dossier principal
                                System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName);

                                // Création du dossier repere
                                System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep);

                                // Création du sous-dossier "TexteIndiceFolder"
                                System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder);

                                // Création du sous-dossier "3D" dans "TexteIndiceFolder"
                                path3D = DossierAtelierServeur + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder + "\\3D";
                                System.IO.Directory.CreateDirectory(path3D);

                            //// Ouverture du dossier "3D"
                            //System.Diagnostics.Process.Start("explorer.exe", path3D);

                            ExporterFichiers = true;
                            }
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                        // Si l'utilisateur choisit "Non", le programme est terminé
                        Application.Exit();
                    }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erreur lors de la recherche des dossiers : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }



                    bool DossierRepExiste = false;
                    string cheminComplet = "";
                    string cheminCompletPDF = "";
                    string DossierRep = "";

                    if (DossierAtelierServeurExiste)
                    {
                        // Le dossier a été trouvé
                        DossierRep = TextBoxCommentaireValue;

                        // Recherche du dossier qui commence par DossierRep dans le dossier trouvé précédemment
                        startDirectories = System.IO.Directory.GetDirectories(directories[0], DossierRep + "*", System.IO.SearchOption.TopDirectoryOnly);

                        if (startDirectories.Length > 0)
                        {
                            // Si le dossier est trouvé, DossierRepExiste est défini sur true
                            DossierRepExiste = true;
                        }
                        else
                        {
                            //Si non si dossier moule existe mais pas dossier rep

                            // Création du dossier repere
                            System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep);

                            // Création du sous-dossier "TexteIndiceFolder"
                            System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder);

                            // Création du sous-dossier "3D" dans "TexteIndiceFolder"
                            path3D = DossierAtelierServeur + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder + "\\3D";
                            System.IO.Directory.CreateDirectory(path3D);

                            //// Ouverture du dossier "3D"
                            //System.Diagnostics.Process.Start("explorer.exe", path3D);

                            ExporterFichiers = true;
                        }
                    }
                    string nomDocument = TSH.Documents.GetName(CurrentDocumentId);
                    string nomFichier = $"{nomDocument}.x_t"; // Ajoutez l'extension souhaitée
                    if (DossierRepExiste)
                    {
                        // Le dossier DossierRep a été trouvé
                        string DossierIndice = TexteIndiceFolder;

                        // Recherche du dossier qui a le même nom que DossierIndice dans le dossier DossierRep trouvé précédemment
                        string[] indiceDirectories = System.IO.Directory.GetDirectories(startDirectories[0], DossierIndice, System.IO.SearchOption.TopDirectoryOnly);

                        if (indiceDirectories.Length > 0)
                        {
                        // Le dossier avec le nom DossierIndice a été trouvé
                        //Tous les dossiers existe
                        // Chemin du dossier à ouvrir
                        string folderPath = indiceDirectories[0];

                        // Ouvre le dossier dans l'explorateur de fichiers
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = folderPath,
                            UseShellExecute = true,
                            Verb = "open"
                        });

                            // Ajoute un délai pour s'assurer que le dossier est ouvert
                            Thread.Sleep(1000); // 1 seconde de délai

                            // Affiche une boîte de message avec un message d'information
                            MessageBox.Show("Il semble que le dossier existe", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly); // Ouvre le dossier dans l'explorateur de fichiers
                                Application.Exit();
                        }
                        else
                        {
                            //Si non si dossier moule et le dossier rep existe mais pas le dossier ind
                            // Le dossier avec le nom DossierIndice n'a pas été trouvé
                            MessageBox.Show(this, $"Aucun dossier avec le nom '{DossierIndice}' n'a été trouvé", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                            // Création du sous-dossier "TexteIndiceFolder"
                            System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder);

                            // Création du sous-dossier "3D" dans "TexteIndiceFolder"
                            path3D = DossierAtelierServeur + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder + "\\3D";
                            System.IO.Directory.CreateDirectory(path3D);


                            ExporterFichiers = true;

                        }
                    }

                    



                    PdmMinorRevisionId PDFRev = new PdmMinorRevisionId();

                    if (ExporterFichiers)
                    {
                        try
                        {
                            cheminComplet = System.IO.Path.Combine(path3D, nomFichier);
                            TSH.Documents.ExportWithOptions(X_TExporterIndex, options, CurrentDocumentId, cheminComplet);

                            for (int i = 0; i < CheckedItemListeCopie.Count; i++)
                            {
                                try
                                {
                                    // Récupérer l'ID de l'objet PDM de la liste
                                    PdmObjectId pdmObjectId = CheckedItemListeCopie[i];

                                    // Récupérer le nom de l'objet
                                    string PDFName = TSH.Pdm.GetName(pdmObjectId);

                                    PDFName = $"{PDFName}.pdf"; // Ajoutez l'extension souhaitée

                                    PdmMajorRevisionId PDFMajRev = TSH.Pdm.GetLastMajorRevision(pdmObjectId);


                                    // Récupérer la dernière révision de l'objet
                                    PDFRev = TSH.Pdm.GetLastMinorRevision(PDFMajRev);


                                    cheminCompletPDF = System.IO.Path.Combine(path3D, PDFName);
                                    TSH.Pdm.ExportMinorRevisionFile(PDFRev, cheminCompletPDF);





                                }
                                catch (Exception ex)
                                {
                                    // Gestion des erreurs
                                    MessageBox.Show($"Erreur pour l'ID {CheckedItemListeCopie[i]}: {ex.Message}");
                                }
                            }



                        Process.Start(new ProcessStartInfo
                        {
                            FileName = "explorer.exe",
                            Arguments = path3D,
                            UseShellExecute = true,
                            Verb = "open"
                        });

                            // Ajoute un délai pour s'assurer que le dossier est ouvert
                            Thread.Sleep(1000); // 1 seconde de délai

                            // Affiche une boîte de message avec un message d'information
                            MessageBox.Show("Exportation réussie.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    
                            // Ferme l'application après que l'utilisateur ait cliqué sur OK
                            Application.Exit();



                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Erreur lors de l'exportation : " + ex.Message);
                        }
                    }
            }
            else
            {
              // Affiche un message d'erreur si le chemin n'est pas valide
              MessageBox.Show("Veuillez entrer un chemin de dossier atelier valide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
            }
        }

        // Fonction appelée lorsque le bouton 'button1' est cliqué
        private void button1_Click_1(object sender, EventArgs e)
        {
            // Redémarre l'application
            Application.Restart();
        }

        // Fonction appelée lorsque l'option de menu 'quitterToolStripMenuItem' est cliquée
        private void quitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Ferme l'application
            Application.Exit();
            
        }


        //fonction pour selectionner le chemin du dossier atelier
        private void button3_Click_1(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                folderBrowserDialog.Description = "Sélectionnez un dossier";
                folderBrowserDialog.ShowNewFolderButton = true;

                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    string driveLetter = Path.GetPathRoot(selectedPath).TrimEnd('\\');
                    string networkPath = NetworkPathHelper.GetNetworkPath(driveLetter);

                    string EnsureTrailingSlash(string path)
                    {
                        return path.EndsWith(Path.DirectorySeparatorChar.ToString()) ? path : path + Path.DirectorySeparatorChar;
                    }

                    if (networkPath != null)
                    {
                        // Récupérer le nom du lecteur mappé
                        string driveName = NetworkPathHelper.GetMappedDriveName(driveLetter);

                        // Construire correctement le chemin complet réseau
                        string fullPath = EnsureTrailingSlash(selectedPath.Replace(driveLetter, networkPath));

                        // Afficher le chemin réseau complet
                        string displayPath = fullPath; // [$"{fullPath} [{driveLetter}({driveName})]"];
                        textBox4.Text = displayPath;

                        // Sauvegarder le chemin réseau complet
                        Properties.Settings.Default.FolderPath = fullPath;
                    }
                    else
                    {
                        // Si ce n'est pas un chemin réseau, ajouter le slash si nécessaire
                        string displayPath = EnsureTrailingSlash(Path.GetFullPath(selectedPath));
                        textBox4.Text = displayPath;

                        // Sauvegarder le chemin local
                        Properties.Settings.Default.FolderPath = displayPath;
                    }

                    // Sauvegarder le chemin dans les paramètres de l'application
                    Properties.Settings.Default.Save();
                }
            }
        }



        private void VerifierCheminAuDemarrage()
        {
            // Récupérer le chemin enregistré dans les paramètres de l'application
            string savedPath = Properties.Settings.Default.FolderPath;

            // Vérifier si le chemin est vide ou null
            if (string.IsNullOrEmpty(savedPath))
            {
                // Si aucun chemin n'est configuré, utiliser le texte actuel dans TextBox comme texte par défaut
                textBox4.Text = "Chemin du dossier atelier";
                return;
            }

            // Vérifier si le chemin enregistré existe réellement
            if (Directory.Exists(savedPath) || File.Exists(savedPath))
            {
                // Si le chemin existe, afficher le chemin dans textBox4
                textBox4.Text = savedPath;
            }
            else
            {
                // Si le chemin n'existe pas, réinitialiser le chemin dans les paramètres
                Properties.Settings.Default.FolderPath = string.Empty;
                Properties.Settings.Default.Save();

                // Utiliser le texte actuel dans TextBox comme texte par défaut si le chemin est invalide
                textBox4.Text = "Chemin du dossier atelier";
            }
        }

        // Sauvegarder le choix de matière dans les paramètres
        private void SaveMaterialChoice()
        {
            if (matiereButton1.Checked)
            {
                // Sauvegarder le choix de matière dans les paramètres de l'application
                Properties.Settings.Default.SelectedMaterial = "acierjbt";
            }
            else if (matiereButton2.Checked)
            {
                Properties.Settings.Default.SelectedMaterial = "aciertrempejbt";
            }
            else
            {
                Properties.Settings.Default.SelectedMaterial = string.Empty; // Aucun choix
            }

            // Sauvegarder les paramètres
            Properties.Settings.Default.Save();
        }

        // Restaurer le choix de matière lors du démarrage
        private void RestoreMaterialChoice()
        {
            // Récupérer le choix de matière enregistré dans les paramètres
            string savedMaterial = Properties.Settings.Default.SelectedMaterial;

            // Vérifier le choix sauvegardé et cocher le bon bouton radio
            if (savedMaterial == "acierjbt")
            {
                matiereButton1.Checked = true;
            }
            else if (savedMaterial == "aciertrempejbt")
            {
                matiereButton2.Checked = true;
            }
            else
            {
                // Aucun choix sauvegardé ou choix vide, laisser les boutons décochés
                matiereButton1.Checked = false;
                matiereButton2.Checked = false;
            }
        }

        private void buttonSetMaterial_Click(object sender, EventArgs e)
        {
            try
            {
                string domain = "jbtecnics";
                string uName1 = "acierjbt";
                string uName2 = "aciertrempejbt";

                // Rechercher les documents de matériau dans le PDM
                PdmObjectId MaterialAcierJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName1);
                DocumentId MaterialAcierJbtId = TSH.Documents.GetDocument(MaterialAcierJbt);

                PdmObjectId MaterialAcierTrempeJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName2);
                DocumentId MaterialAcierTrempeJbtId = TSH.Documents.GetDocument(MaterialAcierTrempeJbt);

                // Appeler la vérification de la matière
                CheckAndSetMaterial(MaterialAcierJbtId, MaterialAcierTrempeJbtId);

                // Sauvegarder le choix de matière
                SaveMaterialChoice();

                // Message de confirmation (facultatif)
                MessageBox.Show("La matière a été définie pour la pièce.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur lors de la définition de la matière : " + ex.Message);
            }
        }

        // Nouvelle méthode pour vérifier les boutons et appliquer la matière
        private void CheckAndSetMaterial(DocumentId MaterialAcierJbtId, DocumentId MaterialAcierTrempeJbtId)
        {
            if (matiereButton1.Checked)
            {
                // Si matiereButton1 est coché, assigner MaterialAcierJbtId
                TopSolidDesignHost.Parts.SetMaterial(CurrentDocumentIdLastRev, MaterialAcierJbtId);
            }
            else if (matiereButton2.Checked)
            {
                // Si matiereButton2 est coché, assigner MaterialAcierTrempeJbtId
                TopSolidDesignHost.Parts.SetMaterial(CurrentDocumentIdLastRev, MaterialAcierTrempeJbtId);
            }
            else
            {
                // Aucun bouton n'est sélectionné, afficher un message d'erreur
                MessageBox.Show("Veuillez sélectionner une matière avant de continuer.");
                throw new InvalidOperationException("Aucune matière sélectionnée.");
            }
        }

        // Appeler la fonction RestoreMaterialChoice lors de l'initialisation de l'interface utilisateur (par exemple, dans le Form_Load)
        private void Form_Load(object sender, EventArgs e)
        {
            // Restaurer le choix de matière au démarrage
            RestoreMaterialChoice();
        }

        
    }
}
