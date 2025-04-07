using System;
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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using TopSolid.Cad.Drafting.Automating;
using TopSolid.Cam.NC.Kernel.Automating;
using TSHD = TopSolid.Cad.Design.Automating.TopSolidDesignHost;
using TSCH = TopSolid.Cam.NC.Kernel.Automating.TopSolidCamHost;
using S = System.Collections.Generic;
using System.Security;
using System.Security.Cryptography;
using System.Xml.Linq;
using System.Drawing.Text;
using Folder_Creator_Tool_V3;
using static Folder_Creator_Tool_V3.FolderCreatorTool;



namespace Folder_Creator_Tool_V3
{
    public partial class Form1 : Form
        {
        private Document currentDoc;
        private Document derivedCurrentDoc;
        private StartConnect startConnect;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        #region Initialisation formulaire
        public Form1()
        {
            InitializeComponent();
            startConnect = new StartConnect();
            startConnect.ConnectionTopsolid();

            // Initialisation de currentDoc
            currentDoc = new Document();
            currentDoc.DocId = TSH.Documents.EditedDocument;
            derivedCurrentDoc = new Document();

            // Charger le chemin sauvegardé et l'afficher dans la TextBox
            textBox4.Text = Properties.Settings.Default.FolderPath;

            // Variable pour stocker la dernière lettre du nom du dossier
            string IndiceTxtBox = "";

           //Recuperation l'exporteur parasolid
            FindParasolidExporterIndex(out X_TExporterIndex);


            string versionX_T = "31";
                   

            //Configuartion version parasolid
            List<KeyValue> options = new List<KeyValue>();

            options = VersionX_T(X_TExporterIndex, (versionX_T+"0"));

                label7.Text = versionX_T; // Afficher la version dans le formulaire
  

            //----------- Récupération du nom du document courant----------------------------------------------------------------------------------------------------------------------------         
            try
            {
                //string CurrentDocumentName = TSH.Documents.GetName(CurrentDocumentId);  // Récupération du nom du document courant
                //TextCurrentDocumentName = currentDoc.Nom;

                textBox9.Text = currentDoc.Nom; //Affichage du nom du document courent dans la case texte
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

            //-----------Récupération Nom projet courant----------------------------------------------------------------------------------------------------------------------------
            try
            {
                //CurrentProjectName = TSH.Pdm.GetName(currentDoc.ProjetId);  // Récupération Nom projet
                string TextBoxProjectName = currentDoc.Projet.NomProjet;

                textBox10.Text = TextBoxProjectName; //Affichage du nom du projet courent dans la case texte
                textBox1.Text = TextBoxProjectName; //Affichage le N° de moule dans la case texte

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération de l'id du projet courant " + ex.Message);
            }
           
            // Appel de la fonction pour chercher le dossier contenant
            ChercherDossierDocumentEnCours(currentDoc.PdmObject, out IndiceTxtBox);

            //-------------Creation de la variable pour la recherche du dossier atelier-------------------------------------------------------------------------------------------------------------------
            try
            {
                List<PdmObjectId> AtelierFolderIds = TSH.Pdm.SearchFolderByName(currentDoc.Projet.ProjetId, "02-Atelier");
                AtelierFolderId = AtelierFolderIds[0];

            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Dossier ''02-Atelier'' introuvable dans le projet " + ex.Message);
            }

            

            //------------- Récupération du commentaire (Repère) du document courant----------------------------------------------------------------------------------------------------------------------------

            //RecupCommentaire(in CurrentDocumentId, out CurrentDocumentCommentaireId, out TextCurrentDocumentCommentaire);

            //----------- Récupération de la désignation du document courant----------------------------------------------------------------------------------------------------------------------------

            try
            {
                //CurrentDocumentDesignationId = TSH.Parameters.GetDescriptionParameter(CurrentDocumentId);   // Récupération de la désignation
                //TextCurrentDocumentDesignation = TSH.Parameters.GetTextLocalizedValue(currentDoc.CommentaireId);
                textBox2.Text = TSH.Parameters.GetTextLocalizedValue(currentDoc.CommentaireId);
                textBox3.Text = TSH.Parameters.GetTextLocalizedValue(currentDoc.DesignationId); //Affichage du commentaire (Repère) dans la case texte
            }
            catch (Exception ex)
            {
                this.TopMost = false;
                MessageBox.Show(new Form { TopMost = true }, "Echec de la récupération du Commentaire " + ex.Message);
            }

            //----------- Variable des differents façon de nommer le dossier indice----------------------------------------------------------------------------------------------------------------------------

            textBox8.Text = IndiceTxtBox; //Affichage de l'indice

            //Liste PDF--------------------------------

            TreeNode rootFolderNode = listePdf(IndiceTxtBox);

            string DossierIndicePdf = "Ind " + IndiceTxtBox;

            // Trouver et déplier le nœud cible
            ExpandTargetNode(rootFolderNode, DossierIndicePdf);

        }
        #endregion Fin initialisation formulaire

        #region Bouton formulaire
        private void button2_Click_1(object sender, EventArgs e)
        {
            string celluleVideErreur = "Merci de remplir toute les cases";

            if (string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox3.Text) || string.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show(celluleVideErreur);
                return;
            }

            string dossierAtelierServeur = textBox4.Text;
            if (!Directory.Exists(dossierAtelierServeur))
            {
                MessageBox.Show("Veuillez entrer un chemin de dossier atelier valide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (MessageBox.Show("Voulez vous simplifier la piece", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TSH.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Shapes.Healing.HealCommand");
                    Application.Exit();
                }

                List<PdmObjectId> checkedItems = GetCheckedItems(treeView1.Nodes);
                List<PdmObjectId> pdfItems = checkedItems
                    .Where(item =>
                    {
                        TSH.Pdm.GetType(item, out string extension);
                        return extension == ".pdf";
                    })
                    .ToList();
                
                bool recommencer;
                do
                {
                    recommencer = false;
                    string texteDossierRep = $"{textBox2.Text} - {textBox3.Text}";
                    string texteIndiceFolder = $"Ind {textBox8.Text}";

                    List<string> folderNames = GetAllProjectFolderNames(currentDoc.Projet.AtelierFolderId);

                    PdmObjectId dossier3D = new PdmObjectId();

                    string dossierRepExiste = DossierExiste(textBox2.Text, textBox3.Text, folderNames);
                    PdmObjectId dossierRepId = new PdmObjectId();

                    if (string.Empty!= dossierRepExiste)
                    {
                        PdmObjectId dossierIndiceId = CreateFolderIfNotExists(dossierRepId, texteIndiceFolder);
                    }
                    else
                    {
                        dossierRepId = CreateFolderIfNotExists(AtelierFolderId, texteDossierRep);
                        PdmObjectId dossierIndiceId = CreateFolderIfNotExists(dossierRepId, texteIndiceFolder);
                        dossier3D = CreationAutreDossiers(dossierIndiceId);
                    }
                    //else
                    //{
                    //    MessageBox.Show(
                    //     "Un dossier avec le même repère ou avec la même désignation existe déjà.\n" +
                    //     "Merci de vérifier et de relancer l'application si nécessaire.\n" +
                    //     folderName,
                    //     "Erreur",
                    //     MessageBoxButtons.OK,
                    //     MessageBoxIcon.Error
                    //    );
                    //    return;

                    //}


                    PdmObjectId auteurPdmObjectId = new PdmObjectId();

                    if (currentDoc.DocId.IsEmpty) return;

                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        try
                        {
                            currentDoc.DocId = TSH.Documents.EditedDocument;
                            DocumentId docId = currentDoc.DocId;
                            TopSolidHost.Documents.EnsureIsDirty(ref docId);
                            currentDoc.DocId = TSH.Documents.EditedDocument;

                            UpdateDocumentParameters(textBox2.Text, textBox3.Text, textBox8.Text);

                            //TopSolidHost.Application.EndModification(true, true);
                        }
                        catch
                        {
                            TopSolidHost.Application.EndModification(false, false);
                        }

                            auteurPdmObjectId = TSH.Pdm.GetOwner(currentDoc.PdmObject);
                            derivedCurrentDoc.DocId = TSHD.Tools.CreateDerivedDocument(auteurPdmObjectId, currentDoc.DocId, false);

                            SaveAndCloseDocument(currentDoc.DocId);

                            MovePdfItems(pdfItems, dossier3D, auteurPdmObjectId);
                            TSH.Pdm.MoveSeveral(new List<PdmObjectId> { derivedCurrentDoc.PdmObject }, dossier3D);
                    }

                    DocumentId derivedDocId = derivedCurrentDoc.DocId;
                    TSH.Documents.Open(ref derivedDocId);

                    if (derivedCurrentDoc.DocId.IsEmpty) return;

                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        try
                        {
                            derivedCurrentDoc.DocId = TSH.Documents.EditedDocument;
                            DocumentId docId = derivedCurrentDoc.DocId;
                            TopSolidHost.Documents.EnsureIsDirty(ref docId);
                            derivedCurrentDoc.DocId = TSH.Documents.EditedDocument;

                            UpdateDerivedDocumentParameters(derivedCurrentDoc.DocId, textBox5.Text, textBox6.Text, textBox7.Text);

                            TransformShapesToAbsoluteFrame(derivedCurrentDoc.DocId);
                            HideAllExceptShapes(derivedCurrentDoc.DocId);
                            SetCameraView(derivedCurrentDoc.DocId);

                            TSH.Pdm.ShowInProjectTree(derivedCurrentDoc.PdmObject);
                            TSH.Pdm.CheckIn(dossier3D, true);
                            TSH.Pdm.SetLifeCycleMainState(derivedCurrentDoc.PdmObject, PdmLifeCycleMainState.Validated); 
                            //TSH.Application.EndModification(true, true);
                        }
                        catch
                        {
                            TopSolidHost.Application.EndModification(false, false);
                        }
                    }

                    if (MessageBox.Show("Souhaitez-vous exporter les fichiers dans le dossier Atelier", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ExportFiles(dossierAtelierServeur, textBox10.Text, texteDossierRep, texteIndiceFolder, pdfItems);
                    }

                } while (recommencer);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<PdmObjectId> GetCheckedItems(TreeNodeCollection nodes)
        {
            List<PdmObjectId> checkedItems = new List<PdmObjectId>();
            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                {
                    checkedItems.Add((PdmObjectId)node.Tag);
                }
                checkedItems.AddRange(GetCheckedItems(node.Nodes));
            }
            return checkedItems;
        }

        private PdmObjectId CreateFolderIfNotExists(PdmObjectId parentFolderId, string folderName)
        {
            List<PdmObjectId> folderIds = new List<PdmObjectId>();
            TSH.Pdm.GetConstituents(parentFolderId, out folderIds, out _);

            PdmObjectId folderId = folderIds.FirstOrDefault(id => TSH.Pdm.GetName(id) == folderName);

            if (folderId.IsEmpty)
            {
                folderId = TSH.Pdm.CreateFolder(parentFolderId, folderName);
            }

            return folderId;
        }

        private void UpdateDocumentParameters(string commentaire, string designation, string indice)
        {
            TSH.Parameters.SetTextValue(currentDoc.CommentaireId, commentaire);
            TSH.Parameters.SetTextValue(currentDoc.DesignationId, designation);
            TSH.Application.EndModification(true, true);
        }

        private void SaveAndCloseDocument(DocumentId docId)
        {
            TSH.Documents.Save(docId);
            TSH.Documents.Close(docId, false, false);
        }

        private void MovePdfItems(List<PdmObjectId> pdfItems, PdmObjectId targetFolderId, PdmObjectId ownerId)
        {
            if (pdfItems.Any())
            {
                List<PdmObjectId> copiedItems = TSH.Pdm.CopySeveral(pdfItems, ownerId);
                TSH.Pdm.MoveSeveral(copiedItems, targetFolderId);
            }
        }

        private void UpdateDerivedDocumentParameters(DocumentId DocId, string matiere, string traitement, string nbrPieces)
        {
            // Activer les modifications sur le document
            modifActif(DocId);
            DocId = TSH.Documents.EditedDocument;

            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Indice 3D", textBox8.Text);
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Matiére plan", matiere);
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Traitement", traitement);
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Nombre de piéces", nbrPieces);
            TSH.Application.EndModification(true, true);
        }

        private void CreateAndPublishTextParameter(DocumentId docId, string paramName, string value)
        {
            ElementId paramId = TSH.Parameters.CreateTextParameter(docId, value);
            TSH.Elements.SetName(paramId, paramName);
            TSH.Parameters.PublishText(docId, paramName, new SmartText(paramId));
        }

        private void TransformShapesToAbsoluteFrame(DocumentId docId)
        {
            // Initialisation de la variable qui recevra le repère sélectionné par l'utilisateur
            SmartFrame3D repereUser = null;

            try
            {
                // Activer les modifications sur le document
                modifActif(docId);
                docId = TSH.Documents.EditedDocument;

                // Boucle demandant à l'utilisateur de sélectionner le plan XY du repère jusqu'à obtenir une réponse
                while (repereUser == null)
                {
                    string titre = "Repère";
                    string label = "Créer repère";
                    UserQuestion questionPlan = new UserQuestion(titre, label);
                    questionPlan.AllowsCreation = true;
                    TopSolidHost.User.AskFrame3D(questionPlan, true, null, out repereUser);
                }

                // Fin de la modification du document
                //TopSolidHost.Application.EndModification(true, true);
            }
            catch (Exception ex)
            {
                // En cas d'erreur lors de la transformation, affichage d'un message d'erreur
                TopSolidHost.Application.EndModification(false, false);
                MessageBox.Show(new Form { TopMost = true }, "Erreur lors de la création du repère : " + ex.Message);
                return;
            }

            ElementId repereUserId = repereUser.ElementId;

            // Création d'un Frame3D pour stocker le repère utilisateur
            Frame3D userRep = TopSolidHost.Geometries3D.GetFrameGeometry(repereUserId);

            // Récupération du repère absolu
            ElementId absRepId = TopSolidHost.Geometries3D.GetAbsoluteFrame(docId);
            Frame3D absRep = TopSolidHost.Geometries3D.GetFrameGeometry(absRepId);

            // Calcul de la translation pour aligner les origines
            Transform3D translation = new Transform3D(
                absRep.XDirection.X, 0, 0, absRep.Origin.X - userRep.Origin.X,
                0, absRep.YDirection.Y, 0, absRep.Origin.Y - userRep.Origin.Y,
                0, 0, absRep.ZDirection.Z, absRep.Origin.Z - userRep.Origin.Z,
                0, 0, 0, 1);

            // Calcul de la rotation pour aligner les axes
            Transform3D rotation = new Transform3D(
                userRep.XDirection.X, userRep.XDirection.Y, userRep.XDirection.Z, 0,
                userRep.YDirection.X, userRep.YDirection.Y, userRep.YDirection.Z, 0,
                userRep.ZDirection.X, userRep.ZDirection.Y, userRep.ZDirection.Z, 0,
                0, 0, 0, 1);

            // Récupération des formes à transformer
            ElementId dossierForme = TSH.Elements.SearchByName(docId, "$TopSolid.Kernel.DB.D3.Shapes.Documents.ElementName.Shapes");
            List<ElementId> formesList = TSH.Elements.GetConstituents(dossierForme);

            // Application des transformations
            foreach (ElementId formeId in formesList)
            {
                TSH.Entities.Transform(formeId, translation);
                TSH.Entities.Transform(formeId, rotation);
            }

            // Fin des modifications
            TSH.Application.EndModification(true, true);
        }

        private void HideAllExceptShapes(DocumentId docId)
        {
            // Activer les modifications sur le document
            modifActif(docId);
            docId = TSH.Documents.EditedDocument;

            List<ElementId> allElements = TSH.Elements.GetElements(docId);
            foreach (ElementId elementId in allElements)
            {
                TSH.Elements.Hide(elementId);
            }
            ElementId dossierForme = TSH.Elements.SearchByName(docId, "$TopSolid.Kernel.DB.D3.Shapes.Documents.ElementName.Shapes");
            List<ElementId> formesList = TSH.Elements.GetConstituents(dossierForme);
            foreach (ElementId formeId in formesList)
            {
                TSH.Elements.Show(formeId);
            }
            TSH.Application.EndModification(true, true);
        }

        private void SetCameraView(DocumentId docId)
        {
            // Activer les modifications sur le document
            modifActif(docId);
            docId = TSH.Documents.EditedDocument;

            int vueActiveInt = TSH.Visualization3D.GetActiveView(docId);
            ElementId perspectiveCamera = TSH.Visualization3D.GetPerspectiveCamera(docId);
            TSH.Visualization3D.GetCameraDefinition(perspectiveCamera, out Point3D eyePosition, out Direction3D lookAtDirection, out Direction3D upDirection, out double fieldAngle, out double fieldRadius);
            TSH.Visualization3D.SetViewCamera(docId, vueActiveInt, eyePosition, lookAtDirection, upDirection, fieldAngle, fieldRadius);
            TSH.Visualization3D.ZoomToFitView(docId, vueActiveInt);
            TSH.Application.EndModification(true, true);
        }

        private void ExportFiles(string dossierAtelierServeur, string folderName, string texteDossierRep, string texteIndiceFolder, List<PdmObjectId> pdfItems)
        {
            string path3D = Path.Combine(dossierAtelierServeur, folderName, texteDossierRep, texteIndiceFolder, "3D");
            Directory.CreateDirectory(path3D);

            string nomFichier = $"{derivedCurrentDoc.Nom}.x_t";
            string cheminComplet = Path.Combine(path3D, nomFichier);
            TSH.Documents.ExportWithOptions(X_TExporterIndex, options, derivedCurrentDoc.DocId, cheminComplet);

            foreach (PdmObjectId pdfItem in pdfItems)
            {
                string pdfName = $"{TSH.Pdm.GetName(pdfItem)}.pdf";
                string cheminCompletPDF = Path.Combine(path3D, pdfName);
                PdmMinorRevisionId pdfRev = TSH.Pdm.GetLastMinorRevision(TSH.Pdm.GetLastMajorRevision(pdfItem));
                TSH.Pdm.ExportMinorRevisionFile(pdfRev, cheminCompletPDF);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = path3D,
                UseShellExecute = true,
                Verb = "open"
            });
            MessageBox.Show("Exportation réussie.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Application.Exit();
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

        #endregion Fin bouton formulaire

        #region Variable divers

        PdmObjectId AtelierFolderId = new PdmObjectId(); //Id du dossier atelier


        #endregion Fin variable divers

        #region Fonction de verification de l'existence du dossier repere dans le dossier atelier du projet courant
        string DossierExiste(string repere, string designation, List<string> folderNames)
        {
            foreach (string folderName in folderNames)
            {
                string normalizedDossierName = folderName.Replace(" ", "").ToLower();
                string normalizedRepere = repere.Replace(" ", "").ToLower();
                string normalizedDesignation = designation.Replace(" ", "").ToLower();

                bool startsWithRepere = normalizedDossierName.StartsWith(normalizedRepere);
                bool endsWithDesignation = normalizedDossierName.EndsWith(normalizedDesignation);

                if (startsWithRepere || (startsWithRepere && endsWithDesignation))
                {
                    return (folderName);
                }
            }
            return (string.Empty);
        }
        #endregion Fin fonction de verification de l'existence du dossier repere dans le dossier atelier du projet courant

        #region Fonction de verification de l'existence du dossier Indice dans le dossier du repére
        bool DossierIndExiste(string indice, List<string> folderNames)
        {
            foreach (string folderName in folderNames)
            {
                if (folderName.EndsWith(indice))
                {
                    return true;
                }
            }
            return false;
        }
        #endregion Fin fonction de verification de l'existence du dossier repere dans le dossier atelier du projet courant

        #region  Récupération de la liste des dossiers dans le dossier atelier du projet----------------------------------------------------------------------------------------------------------------------------
        /// <summary>
        /// Récupère la liste des dossiers situés dans un atelier donné.
        /// </summary>
        /// <param name="dossierAtelierPdmId">L'ID du dossier d'atelier à explorer.</param>
        /// <returns>Une liste d'IDs de dossiers PDM trouvés dans l'atelier.</returns>
        private List<PdmObjectId> DossierPdmInAtelier(PdmObjectId dossierAtelierPdmId)
        {
            // Liste qui contiendra les IDs des dossiers trouvés dans l'atelier
            List<PdmObjectId> folderIds = new List<PdmObjectId>();

            // Liste temporaire pour stocker les IDs des documents (non utilisée ici)
            List<PdmObjectId> documentIds = new List<PdmObjectId>();

            try
            {
                // Récupération des dossiers et documents contenus dans le dossier d'atelier
                TopSolidHost.Pdm.GetConstituents(dossierAtelierPdmId, out folderIds, out documentIds);
            }
            catch (Exception ex)
            {
                // Affichage d'un message en cas d'erreur lors de la récupération des données
                Console.WriteLine("Échec de la récupération des dossiers racine: " + ex.Message);
            }

            // Retourne la liste des dossiers trouvés dans l'atelier
            return folderIds;
        }


        /// <summary>
        /// Récupère récursivement tous les dossiers et sous-dossiers à partir d'un dossier parent donné.
        /// </summary>
        /// <param name="parentFolderId">L'ID du dossier parent à explorer.</param>
        /// <returns>Une liste contenant tous les dossiers et sous-dossiers trouvés.</returns>
        public List<PdmObjectId> GetAllSubFolders(PdmObjectId parentFolderId)
        {
            // Liste pour stocker tous les dossiers trouvés
            List<PdmObjectId> allFolders = new List<PdmObjectId>();

            // Pile pour gérer l'exploration des dossiers sans récursion (évite un dépassement de pile)
            Stack<PdmObjectId> stack = new Stack<PdmObjectId>();
            stack.Push(parentFolderId); // Ajoute le dossier parent dans la pile

            // Tant qu'il reste des dossiers à explorer
            while (stack.Count > 0)
            {
                // Récupère et enlève le dernier dossier ajouté à la pile
                PdmObjectId currentFolder = stack.Pop();

                // Ajoute ce dossier à la liste des dossiers récupérés
                allFolders.Add(currentFolder);

                // Listes temporaires pour stocker les sous-dossiers et documents contenus dans le dossier courant
                List<PdmObjectId> subFolderIds = new List<PdmObjectId>();
                List<PdmObjectId> subDocumentIds = new List<PdmObjectId>();

                try
                {
                    // Récupère les constituants du dossier courant (dossiers et documents)
                    TopSolidHost.Pdm.GetConstituents(currentFolder, out subFolderIds, out subDocumentIds);
                }
                catch (Exception ex)
                {
                    // En cas d'erreur, affiche un message et continue le traitement
                    Console.WriteLine("Erreur lors de la récupération des sous-dossiers: " + ex.Message);
                    continue;
                }

                // Ajoute tous les sous-dossiers trouvés dans la pile pour les explorer ensuite
                foreach (PdmObjectId subFolderId in subFolderIds)
                {
                    stack.Push(subFolderId);
                }
            }

            // Retourne la liste complète des dossiers et sous-dossiers récupérés
            return allFolders;
        }


        /// <summary>
        /// Récupère tous les noms des dossiers et sous-dossiers d'un projet à partir d'un dossier d'atelier donné.
        /// </summary>
        /// <param name="dossierAtelierPdmId">L'ID du dossier d'atelier du projet.</param>
        /// <returns>Une liste contenant les noms de tous les dossiers et sous-dossiers.</returns>
        public List<string> GetAllProjectFolderNames(PdmObjectId dossierAtelierPdmId)
        {
            // Vérifie si l'ID du dossier est valide
            if (dossierAtelierPdmId.IsEmpty)
            {
                Console.WriteLine("Projet introuvable.");
                return new List<string>(); // Retourne une liste vide si aucun projet n'est trouvé
            }

            // Récupère tous les dossiers et sous-dossiers du projet
            List<PdmObjectId> allFolders = GetAllSubFolders(dossierAtelierPdmId);

            // Convertit la liste d'IDs en une liste de noms de dossiers
            List<string> folderNames = new List<string>();

            foreach (PdmObjectId folderId in allFolders)
            {
                try
                {
                    string folderName = TopSolidHost.Pdm.GetName(folderId);
                    folderNames.Add(folderName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur lors de la récupération du nom du dossier : " + ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            return folderNames;
        }

        #endregion Fin Récupération de la liste des dossiers dans le dossier atelier du projet

        #region Fonction de creation des dossier dans le dossier PDM Indice
        /// <summary>
        /// Vérifie si le chemin du dossier est valide et s'il existe
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static PdmObjectId CreateFolderWithCheck(PdmObjectId parent, string name)
        {
            PdmObjectId folder = TSH.Pdm.CreateFolder(parent, name);
            if (folder.IsEmpty)
                throw new Exception($"Échec de la création du dossier {name} sous {parent}");
            return folder;
        }



        /// <summary>
        /// Crée les dossiers après le dossier "Ind" dans le dossier parent spécifié.
        /// </summary>
        /// <param name="dossierIndicePdmId"></param>
        /// <returns>Dossier 3D dans le PDM</returns>
        static PdmObjectId CreationAutreDossiers(PdmObjectId dossierIndicePdmId)
        {
            Dictionary<string, PdmObjectId> dossiers = new Dictionary<string, PdmObjectId>();

            try
            {
                dossiers["3D"] = CreateFolderWithCheck(dossierIndicePdmId, "3D");
                dossiers["Electrode"] = CreateFolderWithCheck(dossierIndicePdmId, "Electrode");
                dossiers["Fraisage"] = CreateFolderWithCheck(dossierIndicePdmId, "Fraisage");
                dossiers["Methode"] = CreateFolderWithCheck(dossierIndicePdmId, "Methode");

                CreateFolderWithCheck(dossiers["Methode"], "Contrôle");
                CreateFolderWithCheck(dossiers["Methode"], "Tournage");

                dossiers["BEHE"] = CreateFolderWithCheck(dossiers["Fraisage"], "BEHE");
                dossiers["FLFA"] = CreateFolderWithCheck(dossiers["Fraisage"], "FLFA");
                dossiers["SETE"] = CreateFolderWithCheck(dossiers["Fraisage"], "SETE");

                CreateFolderWithCheck(dossiers["BEHE"], "OP1");
                CreateFolderWithCheck(dossiers["FLFA"], "OP1");
                CreateFolderWithCheck(dossiers["SETE"], "OP1");

                CreateFolderWithCheck(dossiers["Electrode"], "Air projetée");
                CreateFolderWithCheck(dossiers["Electrode"], "Parallélisée");
                CreateFolderWithCheck(dossiers["Electrode"], "Plan brut");
                CreateFolderWithCheck(dossiers["Electrode"], "Usinage");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur dans la fonction autre dossier : " + ex.Message);
            }

            return dossiers["3D"];
        }

        #endregion Fin fonction de creation des dossier dans le dossier PDM Indice

        #region Activation modification document
        /// <summary>
        /// Active le mode modification pour le document spécifié. 
        /// </summary>
        /// <param name="DocuCourent"></param>
        void modifActif (DocumentId DocuCourent)
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
        #endregion Fin Activation modification document

        #region Fonction treeview liste PDF----------------------------------------------------------------------------------------------------------------------------
        
        /// <summary>
        /// Crée un arbre de nœuds à partir des dossiers et des fichiers PDF dans le dossier "01-2D".
        /// </summary>
        /// <param name="indice"></param>
        /// <returns></returns>
        private TreeNode listePdf(string indice)
        {
            if (!TSH.Pdm.SearchFolderByName(currentDoc.Projet.ProjetId, "01-2D").Any())
            {
                MessageBox.Show("Dossier '01-2D' introuvable.");
                return null;
            }

            PdmObjectId dossiers2D = TSH.Pdm.SearchFolderByName(currentDoc.Projet.ProjetId, "01-2D").First();
            TreeNode rootFolderNode = new TreeNode(TSH.Pdm.GetName(dossiers2D));

            List<PdmObjectId> FoldersInPDFFolder = new List<PdmObjectId>();
            List<PdmObjectId> PDFIds = new List<PdmObjectId>();
            TSH.Pdm.GetConstituents(dossiers2D, out FoldersInPDFFolder, out PDFIds);

            foreach (PdmObjectId folderId in FoldersInPDFFolder)
            {
                ProcessFolder(folderId, rootFolderNode, indice);
            }

            treeView1.Nodes.Add(rootFolderNode);
            return rootFolderNode;
        }

        /// <summary>
        /// Fonction treeview. Ajoute les nœuds cochés à une liste.
        /// </summary>
        /// <param name="nodes"></param>
        /// <param name="list"></param>
        void AddCheckedNodesToList(TreeNodeCollection nodes, List<PdmObjectId> list)
        {
            // Parcourir tous les nœuds dans la collection de nœuds donnée.
            foreach (TreeNode node in nodes)
            {
                // Si le nœud est coché...
                if (node.Checked && node.Tag is PdmObjectId folderId)
                {
                    // Ajouter l'identifiant de l'objet PDM associé au nœud à la liste.
                    list.Add((PdmObjectId)node.Tag);
                }
                // Appeler cette fonction de manière récursive pour tous les nœuds enfants du nœud actuel.
                AddCheckedNodesToList(node.Nodes, list);
            }
        }

        void ProcessFolder(PdmObjectId folderId, TreeNode parentNode, string indice)
        {
            string folderName = TSH.Pdm.GetName(folderId);
            TreeNode folderNode = new TreeNode(folderName);

            List<PdmObjectId> subFolderIds;
            List<PdmObjectId> fileIdsInFolder;
            TSH.Pdm.GetConstituents(folderId, out subFolderIds, out fileIdsInFolder);

            // Ajout des fichiers (par exemple, les PDF)
            foreach (PdmObjectId fileId in fileIdsInFolder)
            {
                string fileName = TSH.Pdm.GetName(fileId);    
                folderNode.Nodes.Add(new TreeNode(fileName) { Tag = fileId });
            }

            // Appel récursif sur les sous-dossiers filtrés (par exemple, ceux qui commencent par "Ind")
            foreach (PdmObjectId subFolderId in subFolderIds)
            {
                string subFolderName = TSH.Pdm.GetName(subFolderId);
                if (subFolderName.StartsWith("Ind", StringComparison.OrdinalIgnoreCase))
                {
                    ProcessFolder(subFolderId, folderNode, indice);
                }
            }

            parentNode.Nodes.Add(folderNode);
        }

        void ExpandTargetNode(TreeNode rootNode, string targetName)
        {
            string normalizedTargetName = targetName.Replace(" ", "").ToLower();

            foreach (TreeNode node in rootNode.Nodes)
            {
                string normalizedNodeName = node.Text.Replace(" ", "").ToLower();

                if (normalizedNodeName == normalizedTargetName)
                {
                    node.Expand();
                    for (TreeNode parent = node.Parent; parent != null; parent = parent.Parent)
                    {
                        parent.Expand();
                    }

                    SortTreeNodes(rootNode); // Placer ici avant de sortir
                    return;
                }

                ExpandTargetNode(node, targetName);
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

        // Fonction récursive pour rechercher le dossier et construire la map des parents
        bool ChercherDossierAvecParentMapping(PdmObjectId dossierActuel, PdmObjectId documentIdRecherche, out string IndiceTxtBox, Dictionary<PdmObjectId, PdmObjectId> parentMapping)
        {
            IndiceTxtBox = "";
            List<PdmObjectId> sousDossiers = new List<PdmObjectId>();
            List<PdmObjectId> documents = new List<PdmObjectId>();
            TSH.Pdm.GetConstituents(dossierActuel, out sousDossiers, out documents);

            // Ajout des sous-dossiers dans le dictionnaire
            foreach (PdmObjectId sousDossier in sousDossiers)
            {
                if (!parentMapping.ContainsKey(sousDossier))
                {
                    parentMapping[sousDossier] = dossierActuel;
                }
            }

            // Vérification si le document est présent
            if (documents.Contains(documentIdRecherche))
            {
                return RemonterJusquAuDossierInd(dossierActuel, out IndiceTxtBox, parentMapping);
            }

            // Recherche récursive dans chaque sous-dossier
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
            string nomDossier = TSH.Pdm.GetName(dossierActuel);

            if (nomDossier.StartsWith("Ind", StringComparison.OrdinalIgnoreCase))
            {
                char derniereLettre = nomDossier[nomDossier.Length - 1];  
                IndiceTxtBox = derniereLettre.ToString().ToUpper();
                return true;
            }

            if (parentMapping.ContainsKey(dossierActuel))
            {
                return RemonterJusquAuDossierInd(parentMapping[dossierActuel], out IndiceTxtBox, parentMapping);
            }

            return false;
        }

        #endregion Fin fonction treeview----------------------------------------------------------------------------------------------------------------------

        #region Fonction recuperation de l'indice pour valeur par defaut formulaire
        void ChercherDossierDocumentEnCours(PdmObjectId PdmObjectIdCurrentDocumentId, out string IndiceTxtBox)
        {
            IndiceTxtBox = "";

            List<PdmObjectId> dossiers3Ds = new List<PdmObjectId>();
            try
            {
                // Recherche du dossier "02-3D" dans le projet courant
                dossiers3Ds = TSH.Pdm.SearchFolderByName(currentDoc.Projet.ProjetId, "02-3D");
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
        #endregion Fin fonction recuperation de l'indice pour valeur par defaut formulaire

        #region Fonction Récupération Commentaire
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
        #endregion Fin Fonction Récupération Commentaire

        #region Fonction verification dossier indice---------------------------------
        string IndiceTxtFormat00 = ""; //Different format d'indice
        string IndiceTxtFormat01 = "";
        string nomDocu = string.Empty;

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
                        List<PdmObjectId> nomDocuIds = TSH.Pdm.SearchDocumentByName(currentDoc.Projet.ProjetId, nomDocu);
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
        #endregion Fin Fonction verification dossier indice

        #region Fonction de configuration de l'exportateur Parasolid
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
                    options[i] = new KeyValue("SAVE_VERSION", version); // Remplacement par un nouvel objet                                                      
                    //break; // Sortie de la boucle après la modification
                }
            } 
                    return options;
        }

        int X_TExporterIndex = new int();
        //Configuartion version parasolid
        List<KeyValue> options = new List<KeyValue>();

        #endregion Fin Fonction de configuration de l'exportateur Parasolid

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

        #region Fonction de vérification du chemin atelier au démarrage

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
        #endregion Fin Fonction de vérification du chemin atelier au démarrage


        #region choix matiere
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

        // Nouvelle méthode pour vérifier les boutons et appliquer la matière
        private void CheckAndSetMaterial(DocumentId MaterialAcierJbtId, DocumentId MaterialAcierTrempeJbtId)
        {
            if (matiereButton1.Checked)
            {
                // Si matiereButton1 est coché, assigner MaterialAcierJbtId
                TopSolidDesignHost.Parts.SetMaterial(currentDoc.DocId, MaterialAcierJbtId);
            }
            else if (matiereButton2.Checked)
            {
                // Si matiereButton2 est coché, assigner MaterialAcierTrempeJbtId
                TopSolidDesignHost.Parts.SetMaterial(currentDoc.DocId, MaterialAcierTrempeJbtId);
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

        #endregion  fin choix matiere

    }
}
