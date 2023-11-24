using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TopSolid.Cad.Design.Automating;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Kernel.Automating;


namespace Folder_Creator_Tool_V3
{
    public partial class Form1 : Form
    {
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

        PdmObjectId DossierIndiceId; //Recuperation de l'Id du dossier Indice pour creation du reste des dossiers

        PdmObjectId DossierFraisageId; //Recuperation de l'Id du dossier fraisage pour creation des dossiers utilisateurs

        PdmObjectId DossierElectrodeId; //Recuperation de l'Id du dossier electrode pour creation des dossiers

        string TexteIndiceFolder;
        List<PdmObjectId> DocuDossierIndiceIds = new List<PdmObjectId>();

        PdmObjectId PdmObjectIdCurrentDocumentId = new PdmObjectId();

        public Form1()
        {
            InitializeComponent();
            //Current project




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
                MessageBox.Show("Echec de la récupération de l'id du document courant " + ex.Message);

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
                MessageBox.Show("Echec de la récupération du nom du document courant " + ex.Message);
            }


            //-----------Récupération ID projet courant----------------------------------------------------------------------------------------------------------------------------
            try
            {
                CurrentProjectPdmId = TopSolidHost.Pdm.GetProject(PdmObjectIdCurrentDocumentId);   // Récupération ID projet courant
            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération de l'id du projet courant " + ex.Message);
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
                MessageBox.Show("Echec de la récupération de l'id du projet courant " + ex.Message);
            }

            //-------------Creation de la variable pour la recherche du dossier atelier-------------------------------------------------------------------------------------------------------------------

            try
            {
                AtelierFolderIds = TopSolidHost.Pdm.SearchFolderByName(CurrentProjectPdmId, "02-Atelier");
                AtelierFolderId = AtelierFolderIds[0];

            }
            catch (Exception ex)
            {
                MessageBox.Show("Dossier ''02-Atelier'' introuvable dans le projet " + ex.Message);
            }


            //------------- Récupération du commentaire (Repère) du document courant----------------------------------------------------------------------------------------------------------------------------

            try
            {
                CurrentDocumentCommentaireId = TopSolidHost.Parameters.GetCommentParameter(CurrentDocumentId); ;  // Récupération du commentaire (Repère)
                TextCurrentDocumentCommentaire = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentCommentaireId);

                textBox2.Text = TextCurrentDocumentCommentaire; //Affichage du commentaire (Repère) dans la case texte


                //----------- Récupération Récupération de la désignation du document courant----------------------------------------------------------------------------------------------------------------------------
            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération du Commentaire " + ex.Message);
            }

            //----------- Récupération de la désignation du document courant----------------------------------------------------------------------------------------------------------------------------

            try
            {
                CurrentDocumentDesignationId = TopSolidHost.Parameters.GetDescriptionParameter(CurrentDocumentId); ;  // Récupération de la désignation
                TextCurrentDocumentDesignation = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentDesignationId);

                textBox3.Text = TextCurrentDocumentDesignation; //Affichage du commentaire (Repère) dans la case texte


            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération du Commentaire " + ex.Message);
            }

            //----------- Variable des differents façon de nommer le dossier indice----------------------------------------------------------------------------------------------------------------------------


            textBox8.Text = "A"; //Affichage de l'indice




        }


        //------------------------------Bouton click dossier-------------

        private void button2_Click_1(object sender, EventArgs e)
        {






            bool recommencer;
            do
            {
                ConstituantFolderNames.Clear();
                ListFoldersNames.Clear(); ;

                recommencer = false; // Réinitialisez recommencer à false à chaque début de boucle
                                     //Récuperation des noms de dossiers
                try
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

                    TopSolidHost.Application.StartModification("", true);
                    TopSolidHost.Elements.SetComment(CurrentDocumentCommentaireId, TextBoxCommentaireValue); //Edition du parametre commentaire
                    TopSolidHost.Elements.SetComment(CurrentDocumentDesignationId, TextBoxDesignationValue); //Edition du parametre commentaire
                    TopSolidHost.Application.EndModification(true, true);
                    TopSolidHost.Pdm.GetConstituents(AtelierFolderId, out FolderIds, out DocumentsIds);

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
                                    else if (result == DialogResult.Cancel)
                                    {
                                        break;
                                    }
                                }


                                else if (TexteDossierRep == ConstituantFolderName)
                                {
                                    MessageBox.Show("le dossier " + TexteDossierRep + " existe déjà. Recherche du dossier d'indice");
                                    TopSolidHost.Pdm.GetConstituents(DossierExistantId, out IndiceFolderIds, out DocumentsInIndiceFolder);
                                    if (IndiceFolderIds.Count != 0)
                                    {

                                        for (int i3 = 0; i3 < IndiceFolderIds.Count; i3++)
                                        {
                                            IndiceFolderName = TopSolidHost.Pdm.GetName(IndiceFolderIds[i3]);
                                            test02 = IndiceFolderName.Equals(IndiceTxtFormat00, StringComparison.OrdinalIgnoreCase);
                                            test03 = IndiceFolderName.Equals(IndiceTxtFormat01, StringComparison.OrdinalIgnoreCase);

                                            if (test02 ^ test03)
                                            {
                                                MessageBox.Show("les dossiers existe deja");
                                                return;
                                            }

                                        }

                                    }

                                    MessageBox.Show("Création du dossier Ind");
                                    DossierIndiceId = TopSolidHost.Pdm.CreateFolder(DossierExistantId, TexteIndiceFolder);
                                    System.Threading.Thread.Sleep(1000);


                                    //Creation du reste des dossiers
                                    TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "3D");
                                    DossierElectrodeId = TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Electrode");
                                    DossierFraisageId = TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Fraisage");
                                    TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Methode");
                                    System.Threading.Thread.Sleep(1000);

                                    //Cration des dossier utilisateur dans le dossier fraisage

                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "BEHE");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "FLFA");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "SETE");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "THHE");
                                    System.Threading.Thread.Sleep(1000);

                                    //Creation des dossiers dans le dossier Electrode

                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Parallélisée");
                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Plan brut");
                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Usinage");
                                    MessageBox.Show("Succés de la creation des dossiers");


                                }
                                else
                                {
                                    //Creation du dossier repere
                                    DossierRepId = TopSolidHost.Pdm.CreateFolder(AtelierFolderId, TexteDossierRep);


                                    //Creation du dosser indice
                                    DossierIndiceId = TopSolidHost.Pdm.CreateFolder(DossierRepId, TexteIndiceFolder);




                                    //Creation du dossier reste des dossiers
                                    TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "3D");
                                    DossierElectrodeId = TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Electrode");
                                    DossierFraisageId = TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Fraisage");
                                    TopSolidHost.Pdm.CreateFolder(DossierIndiceId, "Methode");

                                    //Creation des dossiers dans le dossier Electrode

                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Parallélisée");
                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Plan brut");
                                    TopSolidHost.Pdm.CreateFolder(DossierElectrodeId, "Usinage");

                                    //Cration des dossier utilisateur dans le dossier fraisage

                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "BEHE");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "FLFA");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "SETE");
                                    TopSolidHost.Pdm.CreateFolder(DossierFraisageId, "THHE");
                                    MessageBox.Show("Succés de la creation des dossiers");
                                    break;
                                }
                            }

                        }

                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show("erreur" + ex.Message);
                }

            }

            while (recommencer); // La boucle while recommencera si recommencer est true







            



           





        }
    }
}











































 


     

