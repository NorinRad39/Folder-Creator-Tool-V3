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



namespace Folder_Creator_Tool_V3
{

    public partial class Form1 : Form
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        PdmObjectId CurrentProjectPdmId; //Id du projet courant
        string CurrentProjectName; //Nom du projet courent

        DocumentId CurrentDocumentId; //Id du document courant
        string TextCurrentDocumentName; //Nom du document courant

        List<PdmObjectId> FolderIds = new List<PdmObjectId>(); //Liste des dossiers contenu dans le projet courant
        List<PdmObjectId> DocumentsIds = new List<PdmObjectId>(); //Liste des documents contenu dans le projet courant

        ElementId CurrentDocumentCommentaireId; //Id du commentaire du document courant
        string TextCurrentDocumentCommentaire; //Texte du commentaire du document courant
        ElementId CurrentDocumentDesignationId;//Id de la designation du document courant
        string TextCurrentDocumentDesignation;//Texte de la designation du document courant

        string TextBoxProjectName; //Nom du projet affiché dans la textbox

        string TextBoxDesignationValue;//Texte afficher dans la textbox designation
        string TextBoxCommentaireValue; //Texte afficher dans la textbox Commentaire
        string TextBoxIndiceValue;//Texte afficher dans la textbox Indice
        string TextBoxNomMouleValue;//Texte afficher dans la textbox NomMoule

        string ConstituantFolderName;
        List<string> ConstituantFolderNames = new List<string>();

        List<string> ListFoldersNames = new List<string>();

        string TexteDossierRep; //Nom du dossier repere a creer

        bool test00; //Creation bool pour tester la presence des dossiers a creer dans le projet
        bool test01;
        bool test02;
        bool test03;


        List<PdmObjectId> DocumentsInIndiceFolder = new List<PdmObjectId>(); //Liste document fictive pour recuperation dossier indice
        List<PdmObjectId> IndiceFolderIds = new List<PdmObjectId>(); //Liste des Id de dossier indice

        PdmObjectId DossierExistantId; //Id du dossier existant

        string IndiceFolderName;

        PdmObjectId AtelierFolderId; //Id du dossier atelier



        string CommentaireTxtFormat00; //Different format de commantaire
        string CommentaireTxtFormat01;
        string IndiceTxtFormat00; //Different format d'indice
        string IndiceTxtFormat01;

        List<PdmObjectId> AtelierFolderIds;

        PdmObjectId DossierRepId; //Recuperation de l'Id du dossier rep pour creation du rep indice




        string TexteIndiceFolder;
        List<PdmObjectId> DocuDossierIndiceIds = new List<PdmObjectId>();

        PdmObjectId PdmObjectIdCurrentDocumentId = new PdmObjectId();

        PdmObjectId AuteurPdmObjectId = new PdmObjectId();

        PdmObjectId DocumentModeleDerivation = new PdmObjectId("19_5af816ad - b4b1 - 402d - 8914 - a4c95a895d88 & 3_6038"); //Recupération du PdmObjectId du document modele de dérivation
        //DocumentId DocumentModeleDerivationDocumentId = TopSolidHost.Pdm.(DocumentModeleDerivation);
        
        DocumentId DerivéDocumentId = new DocumentId(); //recuperation de l'Id de document du document dérivé
        List<DocumentId> DerivéDocumentIds = new List<DocumentId> (); //recuperation de l'Id de document du document dérivé

        PdmObjectId Dossier3DPdmObjectId = new PdmObjectId();
        List<PdmObjectId> Dossier3DPdmObjectIds = new List<PdmObjectId> ();


        PdmObjectId DerivéDocumentPdmObjectId = new PdmObjectId (); //Recupération du PdmObjectId du document dérivé

        List<PdmObjectId> DerivéDocumentPdmObjectIds = new List<PdmObjectId>(); //Creation liste du PdmObjectId du document dérivé

        PdmObjectId DossierIndiceId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        //PdmObjectId dossier3DId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        String nomDocu = "";
        List <PdmObjectId> nomDocuIds ;

        PdmObjectId dossier3DGenereId = new PdmObjectId();

        ElementId CurrentNameParameterId = new ElementId ();


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
                MessageBox.Show("Succés de la creation des dossiers");

            }
            
            catch (Exception ex)
            {
                
                MessageBox.Show("Erreur dans la fonction autre dossier " + ex.Message);
                
            }
            return (dossier3DFonctionId);

        }
        //empeche le demarrage de 2 applications en meme temps
        
    


        public Form1()
            {
            InitializeComponent();

            string appId = "Folder_Creator_Tool_V3";

            bool nouvelleInstance;
            using (Mutex mutex = new Mutex(true, appId, out nouvelleInstance))
            {
                if (nouvelleInstance)
                {
                    Application.Run(new Form1());
                }
                else
                {
                    // Récupère l'instance existante et la met au premier plan.
                    Process current = Process.GetCurrentProcess();
                    foreach (Process process in Process.GetProcessesByName(current.ProcessName))
                    {
                        if (process.Id != current.Id)
                        {
                            SetForegroundWindow(process.MainWindowHandle);
                            break;
                        }
                    }
                }
            }

            //Current project

            this.WindowState = FormWindowState.Normal;
            this.Show();
            this.TopMost = true;


            //-----------Connexion a TopSolid-----------------------------------------------------------------------------------------------------------------

            bool TSConnected = TopSolidDesignHost.IsConnected;
            {
                /*if (TSConnected == false)
                    MessageBox.Show("Tentative de connexion a TopSolid");
                else

                    MessageBox.Show("Deja connecté");*/

                try
                {
                    TopSolidHost.Connect();  // Connection à TopSolid
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



            //-----------Récupération ID Document----------------------------------------------------------------------------------------------------------------------------

            try
            {
                CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                PdmObjectIdCurrentDocumentId = TopSolidHost.Documents.GetPdmObject(CurrentDocumentId);

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document courant. Ouvrez un document puis cliquez sur Ok  " + ex.Message);
                

                return;
            }

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




        }


        //------------------------------Bouton click dossier-------------

        private void button2_Click_1(object sender, EventArgs e)
        {

                ConstituantFolderNames.Clear();
                ListFoldersNames.Clear(); 

                bool recommencer;
                recommencer = false; // Réinitialisez recommencer à false à chaque début de boucle
                                     //Récuperation des noms de dossiers

                    do
                    {


                                        CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                                        //CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                                        
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
                                        //TopSolidHost.Application.StartModification("My Action", true);

                                        //Recuperation du PdmObjectId de la nouvelle revision du document apres passage a l'etat modification

                                        CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                                        PdmObjectIdCurrentDocumentId = TopSolidHost.Documents.GetPdmObject(CurrentDocumentId); // Récupération PdmObjectId Document courant
                                        
                                        TopSolidHost.Documents.EnsureIsDirty(ref CurrentDocumentId);
                                       CurrentDocumentId = TopSolidHost.Documents.GetDocument(PdmObjectIdCurrentDocumentId);
                                        //CurrentDocumentCommentaireId = TopSolidHost.Parameters.GetCommentParameter(CurrentDocumentId);   // Récupération du commentaire (Repère)
                                        //CurrentDocumentDesignationId = TopSolidHost.Parameters.GetDescriptionParameter(CurrentDocumentId);   // Récupération de la désignation




                                        TopSolidHost.Parameters.SetTextValue(CurrentDocumentCommentaireId, TextBoxCommentaireValue);
                                        TopSolidHost.Parameters.SetTextValue(CurrentDocumentDesignationId, TextBoxDesignationValue);

                                        //TopSolidHost.Pdm.SetComment(PdmObjectIdCurrentDocumentId, TextBoxCommentaireValue); //Edition du parametre commentaire
                                        //TopSolidHost.Pdm.SetDescription(PdmObjectIdCurrentDocumentId, TextBoxDesignationValue); //Edition du designation commentaire

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
                                                 
                                        
                                                            }
                                                        }
                                                        return;
                                            

                                                    }

                                                }
                                

                                            }

                                        }

                                    }

                    
                             }
                                catch (Exception ex)
                             {
                                this.TopMost = false;
                                TopSolidHost.Application.EndModification(false, false);
                                MessageBox.Show(new Form { TopMost = true }, "erreur" + ex.Message);

                             }
                    }
                    while (recommencer); // La boucle while recommencera si recommencer est true
                    
                        //Creation du dossier repere
                        DossierRepId = TopSolidHost.Pdm.CreateFolder(AtelierFolderId, TexteDossierRep);


                        //Creation du dosser indice
                        DossierIndiceId = TopSolidHost.Pdm.CreateFolder(DossierRepId, TexteIndiceFolder);

            dossier3DGenereId = creationAutreDossiers(DossierIndiceId);
                    
                    //break;




            //--------------------- Dérivation et déplacement du fichier dérivé dans le dossier 3D -----------------------

            
            try
            {
                try
                {
                    CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document courant
                    PdmObjectIdCurrentDocumentId = TopSolidHost.Documents.GetPdmObject(CurrentDocumentId);

                }
                catch (Exception ex)
                {
                    this.TopMost = false;
                    MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du document dérivé " + ex.Message);

                    return;
                }

                //Dossier3DPdmObjectIds = TopSolidHost.Pdm.SearchFolderByName(CurrentProjectPdmId,"3D"); //Recuperation du PdmObjectId (liste) du dossier 3D

                AuteurPdmObjectId = TopSolidHost.Pdm.GetOwner(PdmObjectIdCurrentDocumentId);
                DerivéDocumentId = TopSolidDesignHost.Tools.CreateDerivedDocument(AuteurPdmObjectId, CurrentDocumentId,true); //Creation de la derivé et recuperation de document Id
            
                DerivéDocumentPdmObjectId = TopSolidHost.Documents.GetPdmObject(DerivéDocumentId); //recuperation du PdmObectId du document derivé

                DerivéDocumentPdmObjectIds.Add(DerivéDocumentPdmObjectId); //ajout du PdmObectId du document derivé a la liste

                TopSolidHost.Documents.Save(CurrentDocumentId); //fermeture du document courant
                TopSolidHost.Documents.Close(CurrentDocumentId,false,false); //fermeture du document courant
                TopSolidHost.Documents.Open(ref DerivéDocumentId); //Ouverture du document dérivé

                TopSolidHost.Pdm.MoveSeveral(DerivéDocumentPdmObjectIds, dossier3DGenereId);
                
                //TopSolidHost.Pdm.CheckIn(DerivéDocumentPdmObjectId,true);
                if (!TopSolidHost.Application.StartModification("My Action", false)) return;
                // Modify document.
                //--------------------------------TopSolidHost.Application.InvokeCommand();

                

                //TopSolidHost.Pdm.SetComment(PdmObjectIdCurrentDocumentId, TextBoxCommentaireValue); //Edition du parametre commentaire
                //TopSolidHost.Pdm.SetDescription(PdmObjectIdCurrentDocumentId, TextBoxDesignationValue); //Edition du designation commentaire
                //PdmLifeCycleMainState PdmLifeCycleDerivéDocument = (PdmLifeCycleMainState)2; //mise au coffre


                TopSolidHost.Documents.EnsureIsDirty(ref DerivéDocumentId);
                DerivéDocumentId = TopSolidHost.Documents.EditedDocument;
                ElementId Indice3DParamId = TopSolidHost.Elements.SearchByName(DerivéDocumentId, "Indice 3D");
                TopSolidHost.Parameters.SetTextValue(Indice3DParamId, TextBoxIndiceValue);

                CurrentNameParameterId = TopSolidHost.Parameters.GetNameParameter(DerivéDocumentId);
                TopSolidHost.Parameters.SetTextValue(CurrentNameParameterId, nomDocu) ;


                //TopSolidHost.Documents.SetName(DerivéDocumentId, nomDocu);

                TopSolidHost.Application.EndModification(true, true);




                //TopSolidHost.Pdm.SetLifeCycleMainState(DerivéDocumentPdmObjectId, PdmLifeCycleDerivéDocument);

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                TopSolidHost.Application.EndModification(false, false);
                MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la dérivation" + ex.Message);
                return;
            }
            this.TopMost = false;
            this.WindowState = FormWindowState.Minimized;

            TopSolidHost.Pdm.ShowInProjectTree(DerivéDocumentPdmObjectId);

            //TopSolidHost.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Frames.SmartFrameCommand"); 


            //MessageBox.Show(new Form { TopMost = true }, "Veuillez sélectionner ou créer le repère absolu, puis cliquez sur OK.");
            //List<ElementId> UserRepABSs = TopSolidHost.Geometries3D.GetFrames(DerivéDocumentId);
            

            TopSolidHost.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Transforms.PositioningTransformCommand");
            MessageBox.Show(new Form { TopMost = true }, "Veuillez sélectionner ou créer le repère absolu puis selectionner le repère que vous avez créé comme repère d’origine et le repère absolu comme repère de destination, puis cliquez sur OK.");

            TopSolidHost.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Transforms.TransformCommand");
            MessageBox.Show(new Form { TopMost = true }, "Veuillez sélectionner la pièce et la transformation, puis cliquez sur OK.");

            this.WindowState = FormWindowState.Normal;
            this.Show();
            this.TopMost = true;
            MessageBox.Show("Opération réussie.");







            /* ElementId UserRepABS = UserRepABSs.Last();
             ElementId DocRepABS = TopSolidHost.Geometries3D.GetAbsoluteFrame(DerivéDocumentId);

             Transform3D RepSurRep = new Transform3D();

             SmartPoint3D PointUserRepABS; //= TopSolidHost.Geometries3D.GetPointGeometry(UserRepABS);
             SmartDirection3D DirectionUserRepABS;
             SmartReal DistanceUserRepABS;

            // TopSolidHost.Geometries3D.GetOffsetPointCreation(UserRepABS, out PointUserRepABS, out DirectionUserRepABS, out DistanceUserRepABS);

             */




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











































 


     

