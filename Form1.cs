using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
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
        public Form1()
        {
            InitializeComponent();
            //Current project
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PdmObjectId CurrentProjectPdmId; //Id du projet courent
            string CurrentProjectName; //Nom du projet courent

            DocumentId CurrentDocumentId; //Id du document courent
            string TextCurrentDocumentName; //Nom du document courent

            List<PdmObjectId> FolderIds = new List<PdmObjectId>(); //Liste des dossiers contenu dans le projet courent
            List<PdmObjectId> DocumentsIds = new List<PdmObjectId>(); //Liste des documents contenu dans le projet courent

            ElementId CurrentDocumentCommentaireId; //Id du commentaire du document courent
            string TextCurrentDocumentCommentaire; //Texte du commentaire du document courent
            ElementId CurrentDocumentDesignationId;//Id de la designation du document courent
            string TextCurrentDocumentDesignation;//Texte de la designation du document courent

            string TextBoxProjectName; //Nom du projet affiché dans la textbox

            string TextBoxDesignationValue;//Texte afficher dans la textbox designation
            string TextBoxCommentaireValue; //Texte afficher dans la textbox Commentaire
            string TextBoxIndiceValue;//Texte afficher dans la textbox Indice
            string TextBoxNomMouleValue;//Texte afficher dans la textbox NomMoule

            string ConstituantFolderName;
            List<string> ConstituantFolderNames = new List<string>();

            string ConstituantDocumentName;
            List<string> ConstituantDocumentNames = new List<string>();

            List<string> ListDocumentsFoldersNames = new List<string>();

            string NomDossierRep; //Nom du dossier repere a creer

            bool test00; //Creation bool pour tester la presence des dossiers a creer dans le projet
            bool test01;
            bool test03;


            string NomDossierExistant; //Variable qui contient le nom du dossier si existant

            List < PdmObjectId > IndiceFolderIds = new List<PdmObjectId >();

            PdmObjectId DossierExistant; //Id du dossier existant

            string IndiceFolderName;






            bool TSConnected = TopSolidDesignHost.IsConnected;
            {
                if (TSConnected == false)
                    MessageBox.Show("Tentative de connexion a TopSolid");
                else
                    MessageBox.Show("Deja connecté");

                try
                {
                    TopSolidHost.Connect();  // Connection à TopSolid
                    TopSolidDesignHost.Connect();    // Connection à TopSolidDesign
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Impossible de se connecter à TopSolid" + ex.Message);
                    return;
                }
                if (TSConnected == false)
                    MessageBox.Show("Connexion réussi");
                else
                    MessageBox.Show("Connexion échouée");
            }

            try
            {
                CurrentProjectPdmId = TopSolidHost.Pdm.GetCurrentProject();   // Récupération ID projet courent
            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération de l'id du projet courant" + ex.Message);
                return;
            }

            try
            {
                CurrentProjectName = TopSolidHost.Pdm.GetName(CurrentProjectPdmId);  // Récupération Nom projet
                TextBoxProjectName = CurrentProjectName;

                textBox10.Text = TextBoxProjectName; //Affichage du nom du projet courent dans la case texte
                textBox1.Text = TextBoxProjectName; //Affichage le N° de moule dans la case texte

            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération de l'id du projet courant" + ex.Message);
            }

            try
            {
                CurrentDocumentId = TopSolidHost.Documents.EditedDocument;  // Récupération ID Document
            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération de l'id du document courant" + ex.Message);

                return;
            }

            try
            {
                string CurrentDocumentName = TopSolidHost.Documents.GetName(CurrentDocumentId);  // Récupération du nom du document courent
                TextCurrentDocumentName = CurrentDocumentName;

                textBox9.Text = TextCurrentDocumentName; //Affichage du nom du document courent dans la case texte
                


            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération du nom du document courent" + ex.Message);
            }


            try
            {
                CurrentDocumentCommentaireId = TopSolidHost.Parameters.GetCommentParameter(CurrentDocumentId); ;  // Récupération du commentaire (Repère)
                TextCurrentDocumentCommentaire = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentCommentaireId);

                textBox2.Text = TextCurrentDocumentCommentaire; //Affichage du commenataire (Repère) dans la case texte


            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération du Commentaire" + ex.Message);
            }


            try
            {
                CurrentDocumentDesignationId = TopSolidHost.Parameters.GetDescriptionParameter(CurrentDocumentId); ;  // Récupération de la désignation
                TextCurrentDocumentDesignation = TopSolidHost.Parameters.GetTextLocalizedValue(CurrentDocumentDesignationId);

                textBox3.Text = TextCurrentDocumentDesignation; //Affichage du commenataire (Repère) dans la case texte


            }
            catch (Exception ex)
            {
                MessageBox.Show("Echec de la récupération du Commentaire" + ex.Message);
            }


            textBox8.Text = "A"; //Affichage de l'indice
            

            //Recuperation du texte modifié par l'utilisateur pour nommer les dossiers.
            TextBoxCommentaireValue = textBox2.Text; //Repere de la piece
            TextBoxIndiceValue = textBox8.Text; //Indice de la piece
            TextBoxDesignationValue = textBox3.Text; //Designation de la piece
            TextBoxNomMouleValue = textBox10.Text; //Numero du moule

            MessageBox.Show("dossier" + TextBoxCommentaireValue + TextBoxIndiceValue + TextBoxDesignationValue + TextBoxNomMouleValue);

            NomDossierRep = textBox2.Text + " - " + textBox3.Text; //Nom du dossier repere


            bool recommencer;
            do
            {
                ConstituantFolderNames.Clear();
                ConstituantDocumentNames.Clear(); 

                ListDocumentsFoldersNames.Clear(); ;

                recommencer = false; // Réinitialisez recommencer à false à chaque début de boucle
                //Récuperation des noms de dossiers et documents
                try
                {
                    TopSolidHost.Pdm.GetConstituents(CurrentProjectPdmId, out FolderIds, out DocumentsIds);

                    int i; //1er index

                    for (i = 0; i < FolderIds.Count; i++) //Boucle de déconte
                    {
                        ConstituantFolderName = TopSolidHost.Pdm.GetName(FolderIds[i]); //Obtention des noms de dossier

                        ConstituantFolderNames.Add(ConstituantFolderName); // Ajoutez le nom de dossier à la liste
                    }
                    DossierExistant = FolderIds[i];

                    int i2; //2eme index

                    for (i2 = 0; i2 < DocumentsIds.Count; i++) //Boucle de déconte
                    {
                        ConstituantDocumentName = TopSolidHost.Pdm.GetName(DocumentsIds[i2]); //Obtention des noms de documents
                                                                                              // Ajoutez le nom des documents à la liste
                        ConstituantDocumentNames.Add(ConstituantDocumentName);
                        // Retournez la liste des noms de dossiers et de documents
                    }

                    ListDocumentsFoldersNames.AddRange(ConstituantFolderNames); //Concatenation des 2 listes dans "ListDocumentsFoldersNames"
                    ListDocumentsFoldersNames.AddRange(ConstituantDocumentNames);



                }
                catch (Exception ex)
                {
                    MessageBox.Show("Echec de la récupération des nom de dossier constituant" + ex.Message);
                }

                try
                {
                    //recommencer = false; // Réinitialisez recommencer à false à chaque début de boucle
                    for (int i = 0; i < ListDocumentsFoldersNames.Count; i++)
                    {
                        string CommentaireTxtFormat00 = textBox2.Text + "-";
                        string CommentaireTxtFormat01 = textBox2.Text + " ";

                        test00 = ListDocumentsFoldersNames[i].StartsWith(CommentaireTxtFormat00, StringComparison.OrdinalIgnoreCase);
                        test01 = ListDocumentsFoldersNames[i].StartsWith(CommentaireTxtFormat01, StringComparison.OrdinalIgnoreCase);
                        if (test00 ^ test01)
                        {
                            NomDossierExistant = ListDocumentsFoldersNames[i];
                            if (NomDossierRep != NomDossierExistant)
                            {
                                DialogResult result = MessageBox.Show("Un dossier existe avec le même repère mais avec une désignation différente." + Environment.NewLine + "Merci de vérifier et de corriger avant de continuer" + Environment.NewLine + "Nom du dossier détecté = " + NomDossierExistant + Environment.NewLine + "Nom du dossier qui doit être créé : " + NomDossierRep, "Doublon potentiel", MessageBoxButtons.RetryCancel);

                                if (result == DialogResult.Retry)
                                {
                                    recommencer = true; // Si l'utilisateur clique sur "Retry", recommencer sera true et la boucle while recommencera
                                }
                                else if (result == DialogResult.Cancel)
                                {
                                    break;
                                }
                            }
                            else if (NomDossierRep == NomDossierExistant)
                            {
                                MessageBox.Show("le dossier " + NomDossierRep + " existe déjà. Recherche du dossier d'indice");

                                IndiceFolderIds = TopSolidHost.Pdm.SearchFolderByName(DossierExistant, NomDossierExistant);
                               
                                for (int i3 = 0; i3 < IndiceFolderIds.Count; i3++)
                                {
                                    IndiceFolderName=TopSolidHost.Pdm.GetName(IndiceFolderIds[i3]);

                                    if (IndiceFolderName == "Ind " + textBox8.Text);
                                    {
                                        MessageBox.Show("les dossiers existe deja");

                                    }
                                 
                                    

                                
                                
                                
                                
                                }





                            }





                        }

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Echec de la récupération des noms de dossier constituant" + ex.Message);
                }
            }
            while (recommencer) ; // La boucle while recommencera si recommencer est true













        }
        private void button3_Click(object sender, EventArgs e)
        {
            TopSolidDesignHost.Disconnect();
            TopSolidHost.Disconnect();

            {
                bool TSDisConnected = TopSolidDesignHost.IsConnected;
                if (TSDisConnected == false)
                {
                    MessageBox.Show("Tentative de deconnexion");
                    MessageBox.Show("Déconnexion réussi");
                }
                else
                {
                    MessageBox.Show("Déconnexion échouée. Toujours connecté");
                }


            }

        }

       static private void TopSolid_Exited(object sender, EventArgs e)
       {
           Application.Exit();
       }


    }
            
        
} 











































 


     

