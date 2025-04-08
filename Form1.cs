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
        /// <summary>
        /// Gère l'événement de clic sur le bouton pour exécuter une série d'opérations sur le document.
        /// </summary>
        private void button2_Click_1(object sender, EventArgs e)
        {
            string celluleVideErreur = "Merci de remplir toute les cases";

            // Vérifier que tous les champs nécessaires sont remplis
            if (string.IsNullOrEmpty(textBox2.Text) || string.IsNullOrEmpty(textBox3.Text) || string.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show(celluleVideErreur);
                return;
            }

            // Vérifier que le chemin du dossier atelier est valide
            string dossierAtelierServeur = textBox4.Text;
            if (!Directory.Exists(dossierAtelierServeur))
            {
                MessageBox.Show("Veuillez entrer un chemin de dossier atelier valide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Demander à l'utilisateur s'il souhaite simplifier la pièce
                if (MessageBox.Show("Voulez vous simplifier la piece", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    TSH.Application.InvokeCommand("TopSolid.Kernel.UI.D3.Shapes.Healing.HealCommand");
                    Application.Exit();
                }

                // Récupérer les éléments cochés dans l'arborescence
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

                    PdmObjectId dossier3D = new PdmObjectId();

                    // Vérifier si le dossier repère existe déjà
                    PdmObjectId dossierRepId = DossierExiste(textBox2.Text, textBox3.Text, currentDoc.Projet.AtelierConstituantFolderIds);

                    if (dossierRepId.IsEmpty)
                    {
                        // Créer le dossier repère s'il n'existe pas
                        dossierRepId = CreateFolderIfNotExists(currentDoc.Projet.AtelierFolderId, texteDossierRep);
                        PdmObjectId dossierIndiceId = CreateFolderIfNotExists(dossierRepId, texteIndiceFolder);
                        dossier3D = CreationAutreDossiers(dossierIndiceId);
                    }
                    else if (!dossierRepId.IsEmpty)
                    {
                        // Afficher un message d'erreur si le dossier repère existe déjà
                        string folderName = TSH.Pdm.GetName(dossierRepId);
                        MessageBox.Show(
                             "Un dossier avec le même repère existe déjà.\n" +
                             "Merci de vérifier et de relancer l'application si nécessaire.\n" +
                             folderName,
                             "Erreur",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Error
                            );
                        return;
                    }
                    else
                    {
                        // Vérifier si le dossier indice existe déjà
                        List<PdmObjectId> dossierRepIds = new List<PdmObjectId>();
                        TSH.Pdm.GetConstituents(dossierRepId, out List<PdmObjectId> outFolderIds, out List<PdmObjectId> outDocumentObjectIds);

                        PdmObjectId dossierIndId = DossierExiste("Ind", textBox8.Text, outFolderIds);

                        if (dossierIndId.IsEmpty)
                        {
                            dossierIndId = CreateFolderIfNotExists(dossierRepId, texteIndiceFolder);
                            dossier3D = CreationAutreDossiers(dossierIndId);
                        }
                        else if (!dossierIndId.IsEmpty)
                        {
                            // Afficher un message d'erreur si le dossier indice existe déjà
                            string folderName = TSH.Pdm.GetName(dossierIndId);
                            MessageBox.Show(
                             "Un dossier pour cette indice existe déjà.\n" +
                             "Merci de vérifier et de relancer l'application si nécessaire.\n" +
                             folderName,
                             "Erreur",
                             MessageBoxButtons.OK,
                             MessageBoxIcon.Error
                            );
                            return;
                        }
                    }

                    PdmObjectId auteurPdmObjectId = new PdmObjectId();

                    // Vérifier si le document courant est valide
                    if (currentDoc.DocId.IsEmpty) return;

                    // Démarrer les modifications sur le document
                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        try
                        {
                            currentDoc.DocId = TSH.Documents.EditedDocument;
                            DocumentId docId = currentDoc.DocId;
                            TopSolidHost.Documents.EnsureIsDirty(ref docId);
                            currentDoc.DocId = TSH.Documents.EditedDocument;

                            // Mettre à jour les paramètres du document
                            UpdateDocumentParameters(textBox2.Text, textBox3.Text, textBox8.Text);
                        }
                        catch
                        {
                            TopSolidHost.Application.EndModification(false, false);
                        }

                        // Récupérer l'ID de l'auteur du document PDM
                        auteurPdmObjectId = TSH.Pdm.GetOwner(currentDoc.PdmObject);
                        derivedCurrentDoc.DocId = TSHD.Tools.CreateDerivedDocument(auteurPdmObjectId, currentDoc.DocId, false);

                        // Sauvegarder et fermer le document
                        SaveAndCloseDocument(currentDoc.DocId);

                        // Déplacer les éléments PDF vers le dossier 3D
                        MovePdfItems(pdfItems, dossier3D, auteurPdmObjectId);
                        TSH.Pdm.MoveSeveral(new List<PdmObjectId> { derivedCurrentDoc.PdmObject }, dossier3D);
                    }

                    DocumentId derivedDocId = derivedCurrentDoc.DocId;
                    TSH.Documents.Open(ref derivedDocId);

                    // Vérifier si le document dérivé est valide
                    if (derivedCurrentDoc.DocId.IsEmpty) return;

                    // Démarrer les modifications sur le document dérivé
                    if (TopSolidHost.Application.StartModification("My Modification", false))
                    {
                        try
                        {
                            derivedCurrentDoc.DocId = TSH.Documents.EditedDocument;
                            string nom_Docu = textBox2.Text + " Ind " + textBox8.Text + " " + textBox1.Text;
                            TSH.Documents.SetName(derivedCurrentDoc.DocId, nom_Docu);
                            DocumentId docId = derivedCurrentDoc.DocId;
                            TopSolidHost.Documents.EnsureIsDirty(ref docId);
                            derivedCurrentDoc.DocId = TSH.Documents.EditedDocument;

                            // Mettre à jour les paramètres du document dérivé
                            UpdateDerivedDocumentParameters(derivedCurrentDoc.DocId, textBox5.Text, textBox6.Text, textBox7.Text);

                            // Transformer les formes vers le repère absolu
                            TransformShapesToAbsoluteFrame(derivedCurrentDoc.DocId);
                            // Masquer tous les éléments sauf les formes
                            HideAllExceptShapes(derivedCurrentDoc.DocId);
                            // Définir la vue de la caméra
                            SetCameraView(derivedCurrentDoc.DocId);

                            // Afficher le document dans l'arbre du projet et valider le cycle de vie
                            TSH.Pdm.ShowInProjectTree(derivedCurrentDoc.PdmObject);
                            TSH.Pdm.CheckIn(dossier3D, true);
                            TSH.Pdm.SetLifeCycleMainState(derivedCurrentDoc.PdmObject, PdmLifeCycleMainState.Validated);
                        }
                        catch
                        {
                            TopSolidHost.Application.EndModification(false, false);
                        }
                    }

                    // Demander à l'utilisateur s'il souhaite exporter les fichiers dans le dossier Atelier
                    if (MessageBox.Show("Souhaitez-vous exporter les fichiers dans le dossier Atelier", "Confirmation", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        ExportFiles(dossierAtelierServeur, textBox10.Text, texteDossierRep, texteIndiceFolder, pdfItems);
                    }

                } while (recommencer);
            }
            catch (Exception ex)
            {
                // Afficher un message d'erreur en cas d'exception
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Récupère les éléments cochés dans une collection de nœuds d'arbre.
        /// </summary>
        /// <param name="nodes">La collection de nœuds d'arbre à parcourir.</param>
        /// <returns>Une liste des IDs des éléments PDM cochés.</returns>
        private List<PdmObjectId> GetCheckedItems(TreeNodeCollection nodes)
        {
            // Liste pour stocker les IDs des éléments cochés
            List<PdmObjectId> checkedItems = new List<PdmObjectId>();

            // Parcourir chaque nœud dans la collection
            foreach (TreeNode node in nodes)
            {
                // Si le nœud est coché, ajouter son ID à la liste
                if (node.Checked)
                {
                    checkedItems.Add((PdmObjectId)node.Tag);
                }

                // Ajouter récursivement les éléments cochés des nœuds enfants
                checkedItems.AddRange(GetCheckedItems(node.Nodes));
            }

            // Retourner la liste des éléments cochés
            return checkedItems;
        }

        /// <summary>
        /// Crée un dossier dans le PDM s'il n'existe pas déjà.
        /// </summary>
        /// <param name="parentFolderId">L'ID du dossier parent où le nouveau dossier doit être créé.</param>
        /// <param name="folderName">Le nom du dossier à créer.</param>
        /// <returns>L'ID du dossier créé ou existant.</returns>
        private PdmObjectId CreateFolderIfNotExists(PdmObjectId parentFolderId, string folderName)
        {
            // Liste pour stocker les IDs des dossiers constituants
            List<PdmObjectId> folderIds = new List<PdmObjectId>();

            // Récupérer les constituants du dossier parent
            TSH.Pdm.GetConstituents(parentFolderId, out folderIds, out _);

            // Rechercher un dossier existant avec le nom spécifié
            PdmObjectId folderId = folderIds.FirstOrDefault(id => TSH.Pdm.GetName(id) == folderName);

            // Si le dossier n'existe pas, le créer
            if (folderId.IsEmpty)
            {
                folderId = TSH.Pdm.CreateFolder(parentFolderId, folderName);
            }

            // Retourner l'ID du dossier créé ou existant
            return folderId;
        }

        /// <summary>
        /// Met à jour les paramètres d'un document avec les valeurs spécifiées.
        /// </summary>
        /// <param name="commentaire">Le commentaire à définir pour le document.</param>
        /// <param name="designation">La désignation à définir pour le document.</param>
        /// <param name="indice">L'indice à définir pour le document.</param>
        private void UpdateDocumentParameters(string commentaire, string designation, string indice)
        {
            // Définir la valeur du commentaire pour le document
            TSH.Parameters.SetTextValue(currentDoc.CommentaireId, commentaire);

            // Définir la valeur de la désignation pour le document
            TSH.Parameters.SetTextValue(currentDoc.DesignationId, designation);

            // Terminer les modifications sur le document
            TSH.Application.EndModification(true, true);
        }

        /// <summary>
        /// Sauvegarde et ferme le document spécifié.
        /// </summary>
        /// <param name="docId">L'ID du document à sauvegarder et fermer.</param>
        private void SaveAndCloseDocument(DocumentId docId)
        {
            // Sauvegarder le document
            TSH.Documents.Save(docId);

            // Fermer le document sans forcer la fermeture ni sauvegarder à nouveau
            TSH.Documents.Close(docId, false, false);
        }

        /// <summary>
        /// Déplace les éléments PDF vers un dossier cible dans le PDM.
        /// </summary>
        /// <param name="pdfItems">La liste des IDs des éléments PDF à déplacer.</param>
        /// <param name="targetFolderId">L'ID du dossier cible où les éléments PDF doivent être déplacés.</param>
        /// <param name="ownerId">L'ID du propriétaire des éléments PDF.</param>
        private void MovePdfItems(List<PdmObjectId> pdfItems, PdmObjectId targetFolderId, PdmObjectId ownerId)
        {
            // Vérifier si la liste des éléments PDF n'est pas vide
            if (pdfItems.Any())
            {
                // Copier les éléments PDF avec le propriétaire spécifié
                List<PdmObjectId> copiedItems = TSH.Pdm.CopySeveral(pdfItems, ownerId);

                // Déplacer les éléments copiés vers le dossier cible
                TSH.Pdm.MoveSeveral(copiedItems, targetFolderId);
            }
        }


        /// <summary>
        /// Met à jour les paramètres d'un document dérivé.
        /// </summary>
        /// <param name="DocId">L'ID du document dérivé à mettre à jour.</param>
        /// <param name="matiere">La matière à définir pour le document.</param>
        /// <param name="traitement">Le traitement à définir pour le document.</param>
        /// <param name="nbrPieces">Le nombre de pièces à définir pour le document.</param>
        private void UpdateDerivedDocumentParameters(DocumentId DocId, string matiere, string traitement, string nbrPieces)
        {
            // Activer les modifications sur le document
            modifActif(DocId);

            // Récupérer l'ID du document en cours d'édition
            DocId = TSH.Documents.EditedDocument;

            // Créer et publier le paramètre "Indice 3D" avec la valeur du TextBox
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Indice 3D", textBox8.Text);

            // Créer et publier le paramètre "Matière plan" avec la valeur spécifiée
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Matiére plan", matiere);

            // Créer et publier le paramètre "Traitement" avec la valeur spécifiée
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Traitement", traitement);

            // Créer et publier le paramètre "Nombre de pièces" avec la valeur spécifiée
            CreateAndPublishTextParameter(derivedCurrentDoc.DocId, "Nombre de piéces", nbrPieces);

            // Terminer les modifications sur le document
            TSH.Application.EndModification(true, true);
        }


        /// <summary>
        /// Crée et publie un paramètre de texte dans le document spécifié.
        /// </summary>
        /// <param name="docId">L'ID du document dans lequel le paramètre doit être créé.</param>
        /// <param name="paramName">Le nom du paramètre à créer.</param>
        /// <param name="value">La valeur du paramètre de texte.</param>
        private void CreateAndPublishTextParameter(DocumentId docId, string paramName, string value)
        {
            // Créer un paramètre de texte avec la valeur spécifiée
            ElementId paramId = TSH.Parameters.CreateTextParameter(docId, value);

            // Définir le nom du paramètre
            TSH.Elements.SetName(paramId, paramName);

            // Publier le paramètre de texte
            ElementId publishedParameter = TSH.Parameters.PublishText(docId, paramName, new SmartText(paramId));

            // Définir le nom du paramètre publié
            TSH.Elements.SetName(publishedParameter, paramName);
        }


        /// <summary>
        /// Transforme les formes du document pour les aligner avec le repère absolu.
        /// </summary>
        /// <param name="docId">L'ID du document contenant les formes à transformer.</param>
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
                // TopSolidHost.Application.EndModification(true, true);
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


        /// <summary>
        /// Masque tous les éléments du document sauf les formes.
        /// </summary>
        /// <param name="docId">L'ID du document pour lequel les éléments doivent être masqués.</param>
        private void HideAllExceptShapes(DocumentId docId)
        {
            // Activer les modifications sur le document
            modifActif(docId);

            // Récupérer l'ID du document en cours d'édition
            docId = TSH.Documents.EditedDocument;

            // Obtenir la liste de tous les éléments du document
            List<ElementId> allElements = TSH.Elements.GetElements(docId);

            // Masquer chaque élément du document
            foreach (ElementId elementId in allElements)
            {
                TSH.Elements.Hide(elementId);
            }

            // Rechercher le dossier contenant les formes
            ElementId dossierForme = TSH.Elements.SearchByName(docId, "$TopSolid.Kernel.DB.D3.Shapes.Documents.ElementName.Shapes");

            // Obtenir la liste des formes dans le dossier
            List<ElementId> formesList = TSH.Elements.GetConstituents(dossierForme);

            // Afficher chaque forme
            foreach (ElementId formeId in formesList)
            {
                TSH.Elements.Show(formeId);
            }

            // Terminer les modifications sur le document
            TSH.Application.EndModification(true, true);
        }


        /// <summary>
        /// Définit la vue de la caméra pour le document spécifié.
        /// </summary>
        /// <param name="docId">L'ID du document pour lequel la vue de la caméra doit être définie.</param>
        private void SetCameraView(DocumentId docId)
        {
            // Activer les modifications sur le document
            modifActif(docId);

            // Récupérer l'ID du document en cours d'édition
            docId = TSH.Documents.EditedDocument;

            // Obtenir l'ID de la vue active
            int vueActiveInt = TSH.Visualization3D.GetActiveView(docId);

            // Obtenir l'ID de la caméra perspective
            ElementId perspectiveCamera = TSH.Visualization3D.GetPerspectiveCamera(docId);

            // Obtenir la définition de la caméra
            TSH.Visualization3D.GetCameraDefinition(perspectiveCamera, out Point3D eyePosition, out Direction3D lookAtDirection, out Direction3D upDirection, out double fieldAngle, out double fieldRadius);

            // Définir la vue de la caméra avec les paramètres obtenus
            TSH.Visualization3D.SetViewCamera(docId, vueActiveInt, eyePosition, lookAtDirection, upDirection, fieldAngle, fieldRadius);

            // Ajuster la vue pour qu'elle s'adapte à la scène
            TSH.Visualization3D.ZoomToFitView(docId, vueActiveInt);

            // Terminer les modifications sur le document
            TSH.Application.EndModification(true, true);
        }


        /// <summary>
        /// Exporte les fichiers 3D et PDF vers le dossier spécifié sur le serveur de l'atelier.
        /// </summary>
        /// <param name="dossierAtelierServeur">Le chemin du dossier serveur de l'atelier.</param>
        /// <param name="folderName">Le nom du dossier principal.</param>
        /// <param name="texteDossierRep">Le texte représentant le dossier de répertoire.</param>
        /// <param name="texteIndiceFolder">Le texte représentant le dossier d'indice.</param>
        /// <param name="pdfItems">La liste des IDs des fichiers PDF à exporter.</param>
        private void ExportFiles(string dossierAtelierServeur, string folderName, string texteDossierRep, string texteIndiceFolder, List<PdmObjectId> pdfItems)
        {
            // Construire le chemin complet pour le dossier 3D
            string path3D = Path.Combine(dossierAtelierServeur, folderName, texteDossierRep, texteIndiceFolder, "3D");

            // Créer le dossier 3D s'il n'existe pas
            Directory.CreateDirectory(path3D);

            // Construire le nom du fichier 3D à exporter
            string nomFichier = $"{derivedCurrentDoc.Nom}.x_t";
            string cheminComplet = Path.Combine(path3D, nomFichier);

            // Exporter le fichier 3D avec les options spécifiées
            TSH.Documents.ExportWithOptions(X_TExporterIndex, options, derivedCurrentDoc.DocId, cheminComplet);

            // Exporter chaque fichier PDF dans le dossier 3D
            foreach (PdmObjectId pdfItem in pdfItems)
            {
                string pdfName = $"{TSH.Pdm.GetName(pdfItem)}.pdf";
                string cheminCompletPDF = Path.Combine(path3D, pdfName);
                PdmMinorRevisionId pdfRev = TSH.Pdm.GetLastMinorRevision(TSH.Pdm.GetLastMajorRevision(pdfItem));
                TSH.Pdm.ExportMinorRevisionFile(pdfRev, cheminCompletPDF);
            }

            // Ouvrir le dossier 3D dans l'explorateur de fichiers
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = path3D,
                UseShellExecute = true,
                Verb = "open"
            });

            // Afficher un message de succès
            MessageBox.Show("Exportation réussie.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // Fermer l'application
            Application.Exit();
        }


        /// <summary>
        /// Gère l'événement de clic sur le bouton pour redémarrer l'application.
        /// </summary>
        /// <param name="sender">L'objet qui a déclenché l'événement.</param>
        /// <param name="e">Les arguments de l'événement.</param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            // Redémarre l'application
            Application.Restart();
        }


        /// <summary>
        /// Gère l'événement de clic sur l'option de menu pour quitter l'application.
        /// </summary>
        /// <param name="sender">L'objet qui a déclenché l'événement.</param>
        /// <param name="e">Les arguments de l'événement.</param>
        private void quitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Ferme l'application
            Application.Exit();
        }


        /// <summary>
        /// Gère l'événement de clic sur le bouton pour sélectionner le chemin du dossier atelier.
        /// </summary>
        /// <param name="sender">L'objet qui a déclenché l'événement.</param>
        /// <param name="e">Les arguments de l'événement.</param>
        private void button3_Click_1(object sender, EventArgs e)
        {
            // Utiliser un FolderBrowserDialog pour permettre à l'utilisateur de sélectionner un dossier
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                // Configurer le dialogue pour afficher une description et un bouton pour créer un nouveau dossier
                folderBrowserDialog.Description = "Sélectionnez un dossier";
                folderBrowserDialog.ShowNewFolderButton = true;

                // Afficher le dialogue et vérifier si l'utilisateur a sélectionné un dossier
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    // Récupérer le chemin sélectionné par l'utilisateur
                    string selectedPath = folderBrowserDialog.SelectedPath;

                    // Obtenir la lettre du lecteur racine du chemin sélectionné
                    string driveLetter = Path.GetPathRoot(selectedPath).TrimEnd('\\');

                    // Obtenir le chemin réseau correspondant à la lettre du lecteur
                    string networkPath = NetworkPathHelper.GetNetworkPath(driveLetter);

                    // Fonction pour s'assurer que le chemin se termine par un slash
                    string EnsureTrailingSlash(string path)
                    {
                        return path.EndsWith(Path.DirectorySeparatorChar.ToString()) ? path : path + Path.DirectorySeparatorChar;
                    }

                    // Vérifier si le chemin est un chemin réseau
                    if (networkPath != null)
                    {
                        // Récupérer le nom du lecteur mappé
                        string driveName = NetworkPathHelper.GetMappedDriveName(driveLetter);

                        // Construire correctement le chemin complet réseau
                        string fullPath = EnsureTrailingSlash(selectedPath.Replace(driveLetter, networkPath));

                        // Afficher le chemin réseau complet dans le TextBox
                        string displayPath = fullPath;
                        textBox4.Text = displayPath;

                        // Sauvegarder le chemin réseau complet dans les paramètres de l'application
                        Properties.Settings.Default.FolderPath = fullPath;
                    }
                    else
                    {
                        // Si ce n'est pas un chemin réseau, ajouter le slash si nécessaire
                        string displayPath = EnsureTrailingSlash(Path.GetFullPath(selectedPath));
                        textBox4.Text = displayPath;

                        // Sauvegarder le chemin local dans les paramètres de l'application
                        Properties.Settings.Default.FolderPath = displayPath;
                    }

                    // Sauvegarder le chemin dans les paramètres de l'application
                    Properties.Settings.Default.Save();
                }
            }
        }


        /// <summary>
        /// Gère l'événement de clic sur le bouton pour définir le matériau.
        /// </summary>
        /// <param name="sender">L'objet qui a déclenché l'événement.</param>
        /// <param name="e">Les arguments de l'événement.</param>
        private void buttonSetMaterial_Click(object sender, EventArgs e)
        {
            try
            {
                // Définir le domaine et les noms universels des matériaux
                string domain = "jbtecnics";
                string uName1 = "acierjbt";
                string uName2 = "aciertrempejbt";

                // Rechercher les documents de matériau dans le PDM pour le premier matériau
                PdmObjectId MaterialAcierJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName1);
                DocumentId MaterialAcierJbtId = TSH.Documents.GetDocument(MaterialAcierJbt);

                // Rechercher les documents de matériau dans le PDM pour le second matériau
                PdmObjectId MaterialAcierTrempeJbt = TSH.Pdm.SearchDocumentByUniversalId(PdmObjectId.Empty, domain, uName2);
                DocumentId MaterialAcierTrempeJbtId = TSH.Documents.GetDocument(MaterialAcierTrempeJbt);

                // Appeler la méthode pour vérifier et définir le matériau
                CheckAndSetMaterial(MaterialAcierJbtId, MaterialAcierTrempeJbtId);

                // Sauvegarder le choix de matériau
                SaveMaterialChoice();

                // Afficher un message de confirmation
                MessageBox.Show("La matière a été définie pour la pièce.");
            }
            catch (Exception ex)
            {
                // Afficher un message d'erreur en cas d'exception
                MessageBox.Show("Erreur lors de la définition de la matière : " + ex.Message);
            }
        }


        #endregion Fin bouton formulaire

        #region Fonction de verification de l'existence du dossier repere dans le dossier atelier du projet courant
        /// <summary>
        /// Vérifie l'existence d'un dossier en fonction d'un repère et d'une désignation.
        /// </summary>
        /// <param name="repere">Le repère à rechercher au début du nom du dossier.</param>
        /// <param name="designation">La désignation à rechercher à la fin du nom du dossier.</param>
        /// <param name="constituantAtelierFolderIds">La liste des IDs des dossiers à vérifier.</param>
        /// <returns>L'ID du dossier trouvé ou PdmObjectId.Empty si aucun dossier ne correspond.</returns>
        PdmObjectId DossierExiste(string repere, string designation, List<PdmObjectId> constituantAtelierFolderIds)
        {
            // Parcourir chaque dossier dans la liste des dossiers de l'atelier
            foreach (PdmObjectId constituantAtelierFolderId in constituantAtelierFolderIds)
            {
                // Récupérer le nom du dossier actuel
                string folderName = TSH.Pdm.GetName(constituantAtelierFolderId);

                // Normaliser les noms en supprimant les espaces et en convertissant en minuscules
                string normalizedDossierName = folderName.Replace(" ", "").ToLower();
                string normalizedRepere = repere.Replace(" ", "").ToLower();
                string normalizedDesignation = designation.Replace(" ", "").ToLower();
                string normalizedIndice = "Ind".ToLower();

                // Initialiser les variables de vérification
                bool startsWithRepere = false;
                bool endsWithDesignation = false;

                // Vérifier si le nom du dossier commence par le repère
                startsWithRepere = normalizedDossierName.StartsWith(normalizedRepere);

                // Vérifier si le nom du dossier se termine par la désignation
                endsWithDesignation = normalizedDossierName.EndsWith(normalizedDesignation);

                // Si le nom commence par le repère mais ne se termine pas par la désignation, retourner un ID vide
                if (startsWithRepere && !endsWithDesignation)
                {
                    return PdmObjectId.Empty;
                }

                // Si le nom commence par le repère et se termine par la désignation, retourner l'ID du dossier
                if (startsWithRepere && endsWithDesignation)
                {
                    return constituantAtelierFolderId;
                }

                // Si le nom commence par le repère mais ne se termine pas par la désignation, retourner un ID vide
                if (startsWithRepere && !endsWithDesignation)
                {
                    return PdmObjectId.Empty;
                }
            }

            // Si aucun dossier ne correspond, retourner un ID vide
            return PdmObjectId.Empty;
        }

        #endregion Fin fonction de verification de l'existence du dossier repere dans le dossier atelier du projet courant

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
