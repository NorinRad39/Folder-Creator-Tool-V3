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
using System.Windows.Media;
//using System.Windows.Media.Media3D;






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
                                                                                                                            //DocumentId DocumentModeleDerivationDocumentId = TopSolidHost.Pdm.(DocumentModeleDerivation);

        DocumentId DerivéDocumentId = new DocumentId(); //recuperation de l'Id de document du document dérivé
        List<DocumentId> DerivéDocumentIds = new List<DocumentId>(); //recuperation de l'Id de document du document dérivé

        PdmObjectId Dossier3DPdmObjectId = new PdmObjectId();
        List<PdmObjectId> Dossier3DPdmObjectIds = new List<PdmObjectId>();


        PdmObjectId DerivéDocumentPdmObjectId = new PdmObjectId(); //Recupération du PdmObjectId du document dérivé

        List<PdmObjectId> DerivéDocumentPdmObjectIds = new List<PdmObjectId>(); //Creation liste du PdmObjectId du document dérivé

        PdmObjectId DossierIndiceId = new PdmObjectId(); //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        //PdmObjectId dossier3DId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        String nomDocu = "";
        List<PdmObjectId> nomDocuIds = new List<PdmObjectId>();

        PdmObjectId dossier3DGenereId = new PdmObjectId();

        ElementId CurrentNameParameterId = new ElementId();


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
                dossier3DFonctionId = TopSolidHost.Pdm.CreateFolder(DossierIndiceIdFonction, "3D");
                DossierElectrodeId = TopSolidHost.Pdm.CreateFolder(DossierIndiceIdFonction, "Electrode");
                DossierFraisageId = TopSolidHost.Pdm.CreateFolder(DossierIndiceIdFonction, "Fraisage");
                TopSolidHost.Pdm.CreateFolder(DossierIndiceIdFonction, "Methode");


                //Cration des dossier utilisateur dans le dossier fraisage

                TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "BEHE");
                TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "FLFA");
                TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "SETE");
                TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "THHE");

                //Creation des dossiers dans le dossier Electrode

                TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Parallélisée");
                TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Plan brut");
                TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Usinage");
                //MessageBox.Show("Succés de la creation des dossiers");

            }

            catch (Exception ex)
            {

                MessageBox.Show("Erreur dans la fonction autre dossier " + ex.Message);

            }
            return (dossier3DFonctionId);

        }







        //Fonction de recupéaration des pdf

        void listePdf()
        {
            // Initialisation des listes pour stocker les dossiers et les fichiers
            List<PdmObjectId> dossiers2Ds = new List<PdmObjectId>();
            PdmObjectId dossiers2D = new PdmObjectId();
            List<PdmObjectId> FoldersInPDFFolder = new List<PdmObjectId>();
            List<PdmObjectId> PDFIds = new List<PdmObjectId>();

            try
            {
                // Recherche du dossier "01-2D" dans le projet courant
                dossiers2Ds = TopSolidHost.Pdm.SearchFolderByName(CurrentProjectPdmId, "01-2D");
                dossiers2D = dossiers2Ds[0];
            }
            catch (Exception ex)
            {
                // Affichage d'un message d'erreur si le dossier "01-2D" n'est pas trouvé
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Dossier '01-2D' introuvable" + ex.Message);
            }

            // Obtention du nom du dossier racine et création d'un TreeNode pour le dossier racine
            string rootFolderName = TopSolidHost.Pdm.GetName(dossiers2D);
            TreeNode rootFolderNode = new TreeNode(rootFolderName);

            // Obtention des dossiers et des fichiers dans le dossier "01-2D"
            TopSolidHost.Pdm.GetConstituents(dossiers2D, out FoldersInPDFFolder, out PDFIds);

            // Initialisation de la variable pour gérer l'événement AfterCheck
            bool isEventTriggeredByCode = false;

            // Gestion de l'événement AfterCheck pour empêcher la coche des dossiers
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

            // Parcours de chaque dossier dans le dossier "01-2D"
            foreach (PdmObjectId folderId in FoldersInPDFFolder)
            {
                // Obtention du nom du dossier et création d'un TreeNode pour le dossier
                string folderName = TopSolidHost.Pdm.GetName(folderId);
                TreeNode folderNode = new TreeNode(folderName);

                // Obtention des fichiers dans le dossier
                List<PdmObjectId> fileIdsInFolder;
                TopSolidHost.Pdm.GetConstituents(folderId, out _, out fileIdsInFolder);

                // Parcours de chaque fichier dans le dossier
                foreach (PdmObjectId fileId in fileIdsInFolder)
                {
                    // Obtention du nom du fichier et création d'un TreeNode pour le fichier
                    string fileName = TopSolidHost.Pdm.GetName(fileId);
                    TreeNode fileNode = new TreeNode(fileName);

                    // Stockage du PdmObjectId dans le Tag du TreeNode
                    fileNode.Tag = fileId;

                    // Ajout du TreeNode du fichier au TreeNode du dossier
                    folderNode.Nodes.Add(fileNode);
                }

                // Ajout du TreeNode du dossier au TreeNode racine
                rootFolderNode.Nodes.Add(folderNode);
            }

            // Ajout du TreeNode racine au TreeView
            treeView1.Nodes.Add(rootFolderNode);
        }


        //-----------Fonction Récupération ID Document courant----------------------------------------------------------------------------------------------------------------------------
        void DocumentCourant(out PdmObjectId PdmObjectIdCurrentDocumentId, out DocumentId CurrentDocumentId, out DocumentId CurrentDocumentIdLastRev)
        {
           
            try
            {
                CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                PdmObjectIdCurrentDocumentId = TopSolidHost.Documents.GetPdmObject(CurrentDocumentId);
                PdmMinorRevisionId DerniereRev = TopSolidHost.Pdm.GetFinalMinorRevision(PdmObjectIdCurrentDocumentId);
                CurrentDocumentIdLastRev = TopSolidHost.Documents.GetMinorRevisionDocument(DerniereRev);


            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document courant. Ouvrez un document puis réessayez  " + ex.Message);
                Environment.Exit(0);
                //return;

            }
        }

        void RecupCommentaire(in DocumentId CurrentDocumentId, out ElementId CurrentDocumentCommentaireId , out string TextCurrentDocumentCommentaire)
        {
            CurrentDocumentCommentaireId = new ElementId(); // Initialisation avant le bloc try
            TextCurrentDocumentCommentaire = "";
            try
            {
                CurrentDocumentCommentaireId = TopSolidHost.Parameters.GetCommentParameter(CurrentDocumentId);   // Récupération du commentaire (Repère)
                TextCurrentDocumentCommentaire = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentCommentaireId);

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
                                string name = TopSolidHost.Documents.GetName(inDocumentId);

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
                                string name = TopSolidHost.Documents.GetName(inDocumentId);

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

        public Form1()
        {
            

            InitializeComponent();
            //-----------Connexion a TopSolid-----------------------------------------------------------------------------------------------------------------

            bool TSConnected = TopSolidDesignHost.IsConnected;
            {
                /*if (TSConnected == false)
                    MessageBox.Show("Tentative de connexion a TopSolid");
                else

                    MessageBox.Show("Deja connecté");*/

                try
                {
                    TopSolidHost.Connect("Folder Creator Tool");  // Connection à TopSolid
                    TopSolidDesignHost.Connect();    // Connection à TopSolidDesign
                }
                catch (Exception ex)
                {
                    this.TopMost = false;
                    MessageBox.Show("Impossible de se connecter à TopSolid " + ex.Message);
                    return;
                }
                /*if (TSConnected == false)
                    MessageBox.Show("Connexion réussi");
                else
                    MessageBox.Show("Connexion échouée");*/
            }



            
            

            DocumentCourant(out PdmObjectIdCurrentDocumentId,out CurrentDocumentId, out DocumentId CurrentDocumentIdLastRev);


            //----------- Récupération du nom du document courant----------------------------------------------------------------------------------------------------------------------------         
            try
            {
                string CurrentDocumentName = TopSolidHost.Documents.GetName(CurrentDocumentId);  // Récupération du nom du document courant
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
                CurrentProjectPdmId = TopSolidHost.Pdm.GetProject(PdmObjectIdCurrentDocumentId);   // Récupération ID projet courant
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
                CurrentProjectName = TopSolidHost.Pdm.GetName(CurrentProjectPdmId);  // Récupération Nom projet
                TextBoxProjectName = CurrentProjectName;

                textBox10.Text = TextBoxProjectName; //Affichage du nom du projet courent dans la case texte
                textBox1.Text = TextBoxProjectName; //Affichage le N° de moule dans la case texte

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du projet courant " + ex.Message);
            }

            //-------------Creation de la variable pour la recherche du dossier atelier-------------------------------------------------------------------------------------------------------------------

            try
            {
                AtelierFolderIds = TopSolidHost.Pdm.SearchFolderByName(CurrentProjectPdmId, "02-Atelier");
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
                CurrentDocumentDesignationId = TopSolidHost.Parameters.GetDescriptionParameter(CurrentDocumentId);   // Récupération de la désignation
                TextCurrentDocumentDesignation = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentDesignationId);

                textBox3.Text = TextCurrentDocumentDesignation; //Affichage du commentaire (Repère) dans la case texte


            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du Commentaire " + ex.Message);
            }

            //----------- Variable des differents façon de nommer le dossier indice----------------------------------------------------------------------------------------------------------------------------


            textBox8.Text = "A"; //Affichage de l'indice


            //Liste PDF--------------------------------

            listePdf();
            treeView1.ExpandAll();
        }

        private double Dot(Direction3D a, Direction3D b)
        {
            return a.X * b.X + a.Y * b.Y + a.Z * b.Z;
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

            // Fonction pour ajouter les nœuds cochés à la liste
            void AddCheckedNodesToList(TreeNodeCollection nodes, List<PdmObjectId> list)
            {
                foreach (TreeNode node in nodes)
                {
                    // Si le nœud est coché, ajoutez son PdmObjectId à la liste
                    if (node.Checked)
                    {
                        list.Add((PdmObjectId)node.Tag);
                    }

                    // Appel récursif pour les nœuds enfants
                    AddCheckedNodesToList(node.Nodes, list);
                }
            }

            try
            {
                // Ajout des nœuds cochés à la liste CheckedItems
                AddCheckedNodesToList(treeView1.Nodes, CheckedItems);

                // Parcours de la liste CheckedItems
                for (int i = 0; CheckedItems.Count > i; i++)
                {
                    string ExtentionTxtPdf = "";
                    PdmObjectId CheckedItem = CheckedItems[i];
                    PdmObjectType TypePDF = TopSolidHost.Pdm.GetType(CheckedItem, out ExtentionTxtPdf);
                    // Si l'extension du fichier est .pdf, ajoutez-le à la liste CheckedItemsliste
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


                                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out DocumentId CurrentDocumentIdLastRev);

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

                                        if (!TopSolidHost.IsConnected) return;

                                        if (CurrentDocumentId.IsEmpty) return;
                                        // Start modification.
                                        if (!TopSolidHost.Application.StartModification("My Action", false)) return;
                                        // Modify document.

                                        ////Recuperation du PdmObjectId de la nouvelle revision du document apres passage a l'etat modification
                                        //DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId);

                                        TopSolidHost.Documents.EnsureIsDirty(ref CurrentDocumentId);

                                        DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);
                                        RecupCommentaire(in CurrentDocumentIdLastRev, out CurrentDocumentCommentaireId, out TextCurrentDocumentCommentaire);
                                        // Récupération ID Document courant
                                        //PdmObjectIdCurrentDocumentId = TopSolidHost.Documents.GetPdmObject(CurrentDocumentId); // Récupération PdmObjectId Document courant
                                        //CurrentDocumentId = TopSolidHost.Documents.GetDocument(PdmObjectIdCurrentDocumentId);


                                        TopSolidHost.Parameters.SetTextValue(CurrentDocumentCommentaireId, TextBoxCommentaireValue);
                                        TopSolidHost.Parameters.SetTextValue(CurrentDocumentDesignationId, TextBoxDesignationValue);

                                        TopSolidHost.Application.EndModification(true, true);
                                     }
                                        catch (Exception ex)
                                     {
                                            this.TopMost = false;
                                            TopSolidHost.Application.EndModification(false, false);
                                            MessageBox.Show(new Form { TopMost = true }, "erreur lors de l'edition du commentaire et de la désignation du document " + ex.Message);
                                        return;
                                     }
                                       

                                        TopSolidHost.Pdm.GetConstituents(AtelierFolderId, out FolderIds, out DocumentsIds);

                             try
                             {
                                    int i = 0; // index
                                    if (FolderIds.Count != 0)
                                    {
                                        for (i = 0; i < FolderIds.Count; i++) //Boucle de décompte
                                        {
                                            ConstituantFolderName = TopSolidHost.Pdm.GetName(FolderIds[i]); //Obtention des noms de dossier
                                            test00 = ConstituantFolderName.StartsWith(CommentaireTxtFormat00, StringComparison.OrdinalIgnoreCase);
                                            test01 = ConstituantFolderName.StartsWith(CommentaireTxtFormat01, StringComparison.OrdinalIgnoreCase);


                                            DossierExistantId = FolderIds[i];

                                            if (test00 ^ test01)
                                            {
                                                if (TexteDossierRep != ConstituantFolderName)
                                                {
                                                    DialogResult result = MessageBox.Show("Un dossier existe avec le même repère mais avec une désignation différente." + Environment.NewLine + "Merci de vérifier et de corriger avant de continuer " + Environment.NewLine + "Nom du dossier détecté = " + ConstituantFolderName + Environment.NewLine + "Nom du dossier qui doit être créé : " + TexteDossierRep, "Doublon potentiel ", MessageBoxButtons.RetryCancel);

                                                    if (result == DialogResult.Retry)
                                                    {
                                                        recommencer = true; // Si l'utilisateur clique sur "Retry", recommencer sera true et la boucle while recommencera
                                                    }
                                                    if (result == DialogResult.Cancel)
                                                    {
                                                        return;
                                    
                                                    }
                                                }

                                                if (TexteDossierRep == ConstituantFolderName)
                                                {
                                                    MessageBox.Show("le dossier " + TexteDossierRep + " existe déjà. Recherche du dossier d'indice");
                                                    TopSolidHost.Pdm.GetConstituents(DossierExistantId, out IndiceFolderIds, out DocumentsInIndiceFolder);

                                                    PdmObjectId IndiceFolderId;

                                                    if (IndiceFolderIds.Count != 0)
                                                    {

                                                        for (int i3 = 0; i3 < IndiceFolderIds.Count; i3++)
                                                        {
                                                            IndiceFolderName = TopSolidHost.Pdm.GetName(IndiceFolderIds[i3]);
                                                            test02 = IndiceFolderName.Equals(IndiceTxtFormat00, StringComparison.OrdinalIgnoreCase);
                                                            test03 = IndiceFolderName.Equals(IndiceTxtFormat01, StringComparison.OrdinalIgnoreCase);

                                                            IndiceFolderId = IndiceFolderIds[i3];
                                                            if (test02 || test03)
                                                            {
                                                                MessageBox.Show("les dossiers existe deja");
                                                                nomDocuIds = TopSolidHost.Pdm.SearchDocumentByName(CurrentProjectPdmId,nomDocu);
                                                                 if (nomDocuIds.Count==0)
                                                                 {
                                                                    MessageBox.Show(new Form { TopMost = true }, "Un fichier " + nomDocu + " existe deja dans le dossier");
                                                    
                                                                 }
                                                                 return;
                                        
                                                            }
                                                        }
                                                            DossierRepId = DossierExistantId;
                                                            break;
                                                    }

                                                }
                                            }

                                        }

                                    }
                                    else
                                    //Creation du dossier repere
                                    DossierRepId = TopSolidHost.Pdm.CreateFolder(AtelierFolderId, TexteDossierRep);






                             }
                                catch (Exception ex)
                             {
                                this.TopMost = false;
                                TopSolidHost.Application.EndModification(false, false);
                                MessageBox.Show(new Form { TopMost = true }, "erreur" + ex.Message);
                             }
                    }
                    while (recommencer); // La boucle while recommencera si recommencer est true
                    

                        //Creation du dosser indice
                        DossierIndiceId = TopSolidHost.Pdm.CreateFolder(DossierRepId, TexteIndiceFolder);

                        dossier3DGenereId = creationAutreDossiers(DossierIndiceId);

            //--------------------- Dérivation et déplacement du fichier dérivé dans le dossier 3D -----------------------


            try
            {
                try
                {
                    DocumentCourant(out PdmObjectIdCurrentDocumentId, out CurrentDocumentId, out CurrentDocumentIdLastRev);

                }
                catch (Exception ex)
                {
                    this.TopMost = false;
                    MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document dérivé " + ex.Message);

                    return;
                }

                AuteurPdmObjectId = TopSolidHost.Pdm.GetOwner(PdmObjectIdCurrentDocumentId);
                DerivéDocumentId = TopSolidDesignHost.Tools.CreateDerivedDocument(AuteurPdmObjectId, CurrentDocumentId,true); //Creation de la derivé et recuperation de document Id
            
                DerivéDocumentPdmObjectId = TopSolidHost.Documents.GetPdmObject(DerivéDocumentId); //recuperation du PdmObectId du document derivé

                DerivéDocumentPdmObjectIds.Add(DerivéDocumentPdmObjectId); //ajout du PdmObectId du document derivé a la liste

                TopSolidHost.Documents.Save(CurrentDocumentId); //Sauvegarde du document courant
                TopSolidHost.Documents.Close(CurrentDocumentId,false,false); //fermeture du document courant
                TopSolidHost.Documents.Open(ref DerivéDocumentId); //Ouverture du document dérivé

                try
                {

                    //Copie des pdf dans le dossier
                    if (CheckedItemsliste.Count > 0)
                    {
                            CheckedItemListeCopie = TopSolidHost.Pdm.CopySeveral(CheckedItemsliste, AuteurPdmObjectId); //Copie des pdf dans le dossier
                        TopSolidHost.Pdm.MoveSeveral(CheckedItemListeCopie, dossier3DGenereId); //Déplacement du document dérivé et des PDF
                    }
                        TopSolidHost.Pdm.MoveSeveral(DerivéDocumentPdmObjectIds, dossier3DGenereId);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la copie ou du deplacement du PDF dans le dossier 3D" + ex.Message);
                    return;
                }
                    
                    if (!TopSolidHost.Application.StartModification("My Action", false)) return;
                // Modify document.

                TopSolidHost.Documents.EnsureIsDirty(ref DerivéDocumentId);
                DerivéDocumentId = TopSolidHost.Documents.EditedDocument;
                ElementId Indice3DParamId = TopSolidHost.Elements.SearchByName(DerivéDocumentId, "Indice 3D");
                TopSolidHost.Parameters.SetTextValue(Indice3DParamId, TextBoxIndiceValue);

                CurrentNameParameterId = TopSolidHost.Parameters.GetNameParameter(DerivéDocumentId);
                TopSolidHost.Parameters.SetTextValue(CurrentNameParameterId, nomDocu) ;
                TopSolidHost.Application.EndModification(true, true);

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                TopSolidHost.Application.EndModification(false, false);
                MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la dérivation" + ex.Message);
                return;
            }

            //-------------- transfo sur rep abs ------------

            this.TopMost = false;
            this.WindowState = FormWindowState.Minimized;


            // Création des objets nécessaires pour définir un repère
            Plane3D PlanOrigineRep = new Plane3D();
            TopSolid.Kernel.Automating.Point3D PointOrigineRep = new TopSolid.Kernel.Automating.Point3D ();
            Axis3D intersectionAxisRep = new Axis3D();
            Direction3D XDirectionRep = new Direction3D();

            // Initialisation de la variable qui recevra le repère sélectionné par l'utilisateur
            SmartFrame3D ReponseRepereUser = null;

            try
            {
                // Vérification de la connexion à TopSolidHost
                if (!TopSolidHost.IsConnected) return;

                // Vérification de l'ID du document
                if (DerivéDocumentId.IsEmpty) return;

                // Démarrage de la modification du document
                if (!TopSolidHost.Application.StartModification("My Action", false)) return;

                // Marquage du document comme modifié et récupération de l'ID du document en cours de modification
                TopSolidHost.Documents.EnsureIsDirty(ref DerivéDocumentId);
                DerivéDocumentId = TopSolidHost.Documents.EditedDocument;

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
                MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la transformation " + ex.Message);
                return;
            }
            //****************************************************************************************
          
            ElementId RepereUserId = ReponseRepereUser.ElementId;

            // Création d'un Frame3D pour stocker le repère utilisateur
            Frame3D RepereUser = new Frame3D();
            RepereUser = TopSolidHost.Geometries3D.GetFrameGeometry(RepereUserId);

            // Récupération du repère absolu et de ses axes
            ElementId AbsRepId = TopSolidHost.Geometries3D.GetAbsoluteFrame(DerivéDocumentId);
            Frame3D AbsRep = TopSolidHost.Geometries3D.GetFrameGeometry(AbsRepId);

            // Récupération des directions des axes du repère absolu
            Direction3D ox = Direction3D.DX;
            Direction3D oy = Direction3D.DY;
            Direction3D oz = Direction3D.DZ;

            // Récupération des coordonnées de l'origine du repère utilisateur
            Point3D origin = RepereUser.Origin;
            double x = RepereUser.Origin.X;
            double y = RepereUser.Origin.Y;
            double z = RepereUser.Origin.Z;
            x=-x; y=-y; z=-z;

            // Récupération des directions des axes du repère utilisateur
            Direction3D dx = RepereUser.XDirection;
            Direction3D dy = RepereUser.YDirection;
            Direction3D dz = RepereUser.ZDirection;

            string dxTxt = dx.X.ToString();
            string dyTxt = dy.Y.ToString();
            string dzTxt = dz.Z.ToString();

            // Utilisation des informations sur les axes
            // Affichage des informations dans une MessageBox
            string message = $"Origine: {origin}\n" +
                             $"Direction X: {dx.X.CompareTo(ox.X)}\n" +
                             $"Direction Y: {dy.Y.CompareTo(oy.Y)}\n" +
                             $"Direction Z: {dz.Y.CompareTo(oy.Z)}";

            //Transform3D transfo1 = new Transform3D(
            //1, 0, 0, x,  // Translation sur l'axe X absolu
            //0, 1, 0, y,  // Translation sur l'axe Y absolu
            //0, 0, 1, z,  // Translation sur l'axe Z absolu
            //0, 0, 0, 1
            // );

            //// Création de la matrice de rotation
            //Transform3D transfo2 = new Transform3D(
            //dx.X, dy.X, dz.X, 0,
            //dx.Y, dy.Y, dz.Y, 0,
            //dx.Z, dy.Z, dz.Z, 0,
            //0, 0, 0, 1
            //);

            Transform3D transfoFinal = new Transform3D(
            ox.X * dx.X + oy.X * dy.X + oz.X * dz.X,
            ox.X * dx.Y + oy.X * dy.Y + oz.X * dz.Y,
            ox.X * dx.Z + oy.X * dy.Z + oz.X * dz.Z,
            x,

            ox.Y * dx.X + oy.Y * dy.X + oz.Y * dz.X,
            ox.Y * dx.Y + oy.Y * dy.Y + oz.Y * dz.Y,
            ox.Y * dx.Z + oy.Y * dy.Z + oz.Y * dz.Z,
            y,

            ox.Z * dx.X + oy.Z * dy.X + oz.Z * dz.X,
            ox.Z * dx.Y + oy.Z * dy.Y + oz.Z * dz.Y,
            ox.Z * dx.Z + oy.Z * dy.Z + oz.Z * dz.Z,
            z,

            0, 0, 0, 1
        );

            // Recherche du dossier Formes dans le document
            ElementId DossierForme = TopSolidHost.Elements.SearchByName(DerivéDocumentId, "$TopSolid.Kernel.DB.D3.Shapes.Documents.ElementName.Shapes");

            // Récupération de tous les éléments du dossier Formes
            List<ElementId> FormesList = TopSolidHost.Elements.GetConstituents(DossierForme);

            try
            {
                // Vérification de la connexion à TopSolidHost   
                if (!TopSolidHost.IsConnected) return;

                // Vérification de l'ID du document
                if (DerivéDocumentId.IsEmpty) return;

                // Démarrage de la modification du document
                if (!TopSolidHost.Application.StartModification("My Action", false)) return;

                // Marquage du document comme modifié et récupération de l'ID du document en cours de modification
                TopSolidHost.Documents.EnsureIsDirty(ref DerivéDocumentId);
                DerivéDocumentId = TopSolidHost.Documents.EditedDocument;

                // Boucle sur chaque forme dans la liste FormesList et application de la transformation
                for (int i = 0; i < FormesList.Count; i++)
                {
                   //ElementId transformedElement1 = TopSolidHost.Entities.Transform(FormesList[i], transfo1);
                   //ElementId transformedElement2 = TopSolidHost.Entities.Transform(FormesList[i], transfo2);
                    ElementId transformedElement = TopSolidHost.Entities.Transform(FormesList[i], transfoFinal);
                }

                // Fin de la modification du document
                TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                // En cas d'erreur lors de la transformation, affichage d'un message d'erreur
                this.TopMost = false;
                TopSolidHost.Application.EndModification(false, false);
                MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la transformation " + ex.Message);
                return;
            }
            Application.Restart();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Application.Restart();
        }

        private void quitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit(); 
        }

       
    }
}



class DocumentsEventsHost : IDocumentsEvents
{
    public void OnDocumentEditingStarted(DocumentId inDocumentId)
    {
        string name = TopSolidHost.Documents.GetName(inDocumentId); MessageBox.Show(name, "Start Editing");
    }

    public void OnDocumentEditingEnded(DocumentId inDocumentId)
    {
        string name = TopSolidHost.Documents.GetName(inDocumentId); MessageBox.Show(name, "End Editing");
    }
}
