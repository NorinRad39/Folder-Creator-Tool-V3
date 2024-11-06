using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Kernel.Automating;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Xml.Linq;
using System.Diagnostics.Eventing.Reader;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using TS = TopSolid.Kernel.Automating;
using TSH = TopSolid.Kernel.Automating.TopSolidHost;
using System.Windows.Media;
using System.IO;







namespace Folder_Creator_Tool_V3
{




    public partial class Form1 : Form
        {

            

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

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


        //------------------------------------------------------------------
        

        //Fonction de création des dossier apres ind
        static PdmObjectId creationAutreDossiers(PdmObjectId DossierIndiceIdFonction)
        {
            PdmObjectId dossier3DFonctionId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers
                                             //PdmObjectId DossierIndiceIdFonction; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers
            PdmObjectId DossierElectrodeId; //Recuperation de l'Id du dossier electrode pour creation des dossiers
            PdmObjectId DossierFraisageId; //Recuperation de l'Id du dossier fraisage pour creation des dossiers utilisateursP


            try
            {

                //Creation du reste des dossiers
                dossier3DFonctionId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "3D");
                DossierElectrodeId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Electrode");
                DossierFraisageId = TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Fraisage");
                TSH.Pdm.CreateFolder(DossierIndiceIdFonction, "Methode");


                //Cration des dossier utilisateur dans le dossier fraisage

                TSH.Pdm.CreateFolder(DossierFraisageId, "BEHE");
                TSH.Pdm.CreateFolder(DossierFraisageId, "FLFA");
                TSH.Pdm.CreateFolder(DossierFraisageId, "SETE");
                

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

        // Cette fonction liste les fichiers PDF dans un dossier spécifique.
        void listePdf()
        {
            // Initialisation des listes pour stocker les identifiants des objets PDM.
            List<PdmObjectId> dossiers2Ds = new List<PdmObjectId>();
            PdmObjectId dossiers2D = new PdmObjectId();
            List<PdmObjectId> FoldersInPDFFolder = new List<PdmObjectId>();
            List<PdmObjectId> PDFIds = new List<PdmObjectId>();

            try
            {
                // Recherche du dossier "01-2D" dans le projet courant.
                dossiers2Ds = TSH.Pdm.SearchFolderByName(CurrentProjectPdmId, "01-2D");
                dossiers2D = dossiers2Ds[0];
            }
            catch (Exception ex)
            {
                // Affichage d'un message d'erreur si le dossier "01-2D" n'est pas trouvé.
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Dossier '01-2D' introuvable" + ex.Message);
            }

            // Obtention du nom du dossier racine et création d'un TreeNode pour le dossier racine.
            string rootFolderName = TSH.Pdm.GetName(dossiers2D);
            TreeNode rootFolderNode = new TreeNode(rootFolderName);

            // Obtention des dossiers et des fichiers dans le dossier "01-2D".
            TSH.Pdm.GetConstituents(dossiers2D, out FoldersInPDFFolder, out PDFIds);

            // Initialisation de la variable pour gérer l'événement AfterCheck.
            bool isEventTriggeredByCode = false;

            // Gestion de l'événement AfterCheck pour empêcher la coche des dossiers.
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

            // Parcours de chaque dossier dans le dossier "01-2D".
            foreach (PdmObjectId folderId in FoldersInPDFFolder)
            {
                // Appel de la fonction récursive pour chaque dossier.
                ProcessFolder(folderId, rootFolderNode);
            }

            // Ajout du TreeNode racine au TreeView.
            treeView1.Nodes.Add(rootFolderNode);
        }


        // Cette fonction traite un dossier en obtenant ses fichiers et ses sous-dossiers.
        void ProcessFolder(PdmObjectId folderId, TreeNode parentNode)
        {
            // Obtention du nom du dossier et création d'un TreeNode pour le dossier.
            string folderName = TSH.Pdm.GetName(folderId);
            TreeNode folderNode = new TreeNode(folderName);

            // Obtention des fichiers et des sous-dossiers dans le dossier.
            List<PdmObjectId> subFolderIds;
            List<PdmObjectId> fileIdsInFolder;
            TSH.Pdm.GetConstituents(folderId, out subFolderIds, out fileIdsInFolder);

            // Parcours de chaque fichier dans le dossier.
            foreach (PdmObjectId fileId in fileIdsInFolder)
            {
                // Obtention du nom du fichier et création d'un TreeNode pour le fichier.
                string fileName = TSH.Pdm.GetName(fileId);
                TreeNode fileNode = new TreeNode(fileName);

                // Stockage de l'identifiant de l'objet PDM dans le Tag du TreeNode.
                fileNode.Tag = fileId;

                // Ajout du TreeNode du fichier au TreeNode du dossier.
                folderNode.Nodes.Add(fileNode);
            }

            // Parcours de chaque sous-dossier dans le dossier.
            foreach (PdmObjectId subFolderId in subFolderIds)
            {
                // Appel récursif pour chaque sous-dossier.
                ProcessFolder(subFolderId, folderNode);
            }

            // Ajout du TreeNode du dossier au TreeNode parent.
            parentNode.Nodes.Add(folderNode);
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
                Environment.Exit(0);
                //return;

            }
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
            bool existePas = true;
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
            
                existePas = true;
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







        void CréaetionParam(ElementId parametrePubliedId, in ElementId ParamSytemElementId, in string NomParamTxt, in DocumentId document)
        {
            // Récupère la valeur du paramètre système en texte à partir de son identifiant
            string parametreValueTxt = TSH.Parameters.GetTextValue(ParamSytemElementId);

            // Crée un objet SmartText avec la valeur récupérée
            SmartText parametreValueSmartTxt = new SmartText(parametreValueTxt);

            // Publie le paramètre texte dans le document spécifié avec le nom et la valeur fournis
            parametrePubliedId = TSH.Parameters.PublishText(document, NomParamTxt, parametreValueSmartTxt);

            // Attribue le nom spécifié à l'identifiant du paramètre publié
            TSH.Elements.SetName(parametrePubliedId, NomParamTxt);

        }











        public Form1()
        {
            

            InitializeComponent();
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

            listePdf();
            treeView1.ExpandAll();
            // Fonction pour ajouter les nœuds cochés à la liste


        }

        //------------------------------Bouton click dossier-------------

        private void button2_Click_1(object sender, EventArgs e)
        {
            // Initialisation des listes
            List<string> TxtCheckedItems = new List<string>();
            string TxtCheckedItem = null;
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
                Environment.Exit(0);
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


                    // Mise à jour des valeurs de commentaire et de désignation du document
                    TSH.Parameters.SetTextValue(CurrentDocumentCommentaireId, TextBoxCommentaireValue);
                    TSH.Parameters.SetTextValue(CurrentDocumentDesignationId, TextBoxDesignationValue);

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

                string Indice3DNomParamTxt = "Indice 3D";
                // Définit une chaîne de caractères pour le nom du paramètre texte.

                SmartText Indice3DNomParam = new SmartText(TextBoxIndiceValue);
                // Crée un objet SmartText à partir de la valeur contenue dans 'TextBoxIndiceValue'.

                ElementId Indice3DNomParamId = TSH.Parameters.CreateTextParameter(CurrentDocumentId, TextBoxIndiceValue);
                // Crée un paramètre texte dans le document actuel avec la valeur spécifiée et stocke son identifiant.

                TSH.Elements.SetName(Indice3DNomParamId, Indice3DNomParamTxt);
                // Attribue un nom au paramètre texte créé précédemment.

                ElementId publishedIndice3DNomParamId = TSH.Parameters.PublishText(CurrentDocumentId, Indice3DNomParamTxt, Indice3DNomParam);
                // Publie le paramètre texte 'Indice 3D' dans le document actuel et récupère l'identifiant de l'entité publiée.

                TSH.Elements.SetName(publishedIndice3DNomParamId, Indice3DNomParamTxt);
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
                double LinearTol=0.00001;
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
                Environment.Exit(0);
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


            //// Initialisation de la variable qui indique si le dossier Atelier existe sur le serveur
            //bool DossierAtelierServeurExiste = false;

            //// Initialisation du tableau qui contiendra les chemins des dossiers trouvés
            //string[] directories = new string[0];

            //// Initialisation du tableau qui contiendra les chemins des dossiers commençant par DossierRep
            //string[] startDirectories = new string[0];

            //// Affichage d'une boîte de dialogue demandant à l'utilisateur s'il souhaite exporter les fichiers
            //DialogResult dialogResult = MessageBox.Show("Souhaitez vous exporter les fichiers dans le dossier Atelier", "Confirmation", MessageBoxButtons.YesNo);

            //    // Définition du chemin du dossier Atelier sur le serveur
            //    string DossierAtelierServeur = @"\\jbtec-be\meca$\atelier";

            //    // Récupération du nom du dossier à partir de la TextBox
            //    string folderName = TextBoxNomMouleValue;

            //bool ExporterFichiers = false;
            //string path3D = "";
            //if (dialogResult == DialogResult.Yes)
            //{
            //    // Si l'utilisateur choisit "Oui", le code d'exportation des fichiers est exécuté


            //    // Recherche du dossier dans le chemin spécifié
            //    directories = System.IO.Directory.GetDirectories(DossierAtelierServeur, folderName, System.IO.SearchOption.AllDirectories);

            //    if (directories.Length > 0)
            //    {
            //        // Si le dossier est trouvé, DossierAtelierServeurExiste est défini sur true
            //        DossierAtelierServeurExiste = true;
            //    }
            //    else
            //    {
            //        // Création du dossier principal
            //        System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName);

            //        // Création du dossier repere
            //        System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep);

            //        // Création du sous-dossier "TexteIndiceFolder"
            //        System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder);

            //        // Création du sous-dossier "3D" dans "TexteIndiceFolder"
            //        path3D = DossierAtelierServeur + "\\" + folderName + "\\" + TexteDossierRep + "\\" + TexteIndiceFolder + "\\3D";
            //        System.IO.Directory.CreateDirectory(path3D);

            //        // Ouverture du dossier "3D"
            //        System.Diagnostics.Process.Start("explorer.exe", path3D);

            //        ExporterFichiers = true;
            //    }
            //}
            //else if (dialogResult == DialogResult.No)
            //{
            //    // Si l'utilisateur choisit "Non", le programme est terminé
            //    Environment.Exit(0);
            //}


            //bool DossierRepExiste = false;
            //if (DossierAtelierServeurExiste)
            //{
            //    // Le dossier a été trouvé
            //    string DossierRep = TextBoxCommentaireValue;

            //    // Recherche du dossier qui commence par DossierRep dans le dossier trouvé précédemment
            //    startDirectories = System.IO.Directory.GetDirectories(directories[0], DossierRep + "*", System.IO.SearchOption.AllDirectories);

            //    if (startDirectories.Length > 0)
            //    {
            //        // Si le dossier est trouvé, DossierRepExiste est défini sur true
            //        DossierRepExiste = true;
            //    }
            //    else
            //    {
            //        // Création du sous-dossier "TexteIndiceFolder"
            //        System.IO.Directory.CreateDirectory(DossierAtelierServeur + "\\" + folderName + "\\" + TexteIndiceFolder);

            //        // Création du sous-dossier "3D" dans "TexteIndiceFolder"
            //        path3D = DossierAtelierServeur + "\\" + folderName + "\\" + TexteIndiceFolder + "\\3D";
            //        System.IO.Directory.CreateDirectory(path3D);

            //        // Ouverture du dossier "3D"
            //        System.Diagnostics.Process.Start("explorer.exe", path3D);

            //        ExporterFichiers = false;
            //    }



            //}

            //bool indiceDirectoriesExiste = false;
            //if (DossierRepExiste)
            //{
            //    // Le dossier DossierRep a été trouvé
            //    string DossierIndice = TexteIndiceFolder;

            //    // Recherche du dossier qui a le même nom que DossierIndice dans le dossier DossierRep trouvé précédemment
            //    string[] indiceDirectories = System.IO.Directory.GetDirectories(startDirectories[0], DossierIndice, System.IO.SearchOption.AllDirectories);

            //    if (indiceDirectories.Length > 0)
            //    {
            //        // Le dossier avec le nom DossierIndice a été trouvé
            //        MessageBox.Show(this, "Il semble que le dossier existe", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //        System.Diagnostics.Process.Start(indiceDirectories[0]); // Ouvre le dossier dans l'explorateur de fichiers
            //    }
            //    else
            //    {
            //        // Le dossier avec le nom DossierIndice n'a pas été trouvé
            //        MessageBox.Show(this, $"Aucun dossier avec le nom '{DossierIndice}' n'a été trouvé", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    }
            //}


            //FindParasolidExporterIndex(out int X_TExporterIndex);

            //if (ExporterFichiers)
            //{
            //    try
            //    {
            //        TSH.Documents.Export(X_TExporterIndex, CurrentDocumentId, path3D);

            //        Console.WriteLine("Exportation réussie.");
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine("Erreur lors de l'exportation : " + ex.Message);
            //    }
            //}









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
            Environment.Exit(0);
            
        }



    }
}
