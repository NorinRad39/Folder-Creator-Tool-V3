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

namespace Folder_Creator_Tool_V3
{
    public class FolderCreatorTool
    {
        public class Projet
        {
            public PdmObjectId ProjetId { get; private set; }

            public Projet(PdmObjectId projetId)
            {
                ProjetId = projetId;
            }

            /// <summary>
            /// Nom du projet.
            /// </summary>
            public string NomProjet
            {
                get
                {
                    try
                    {
                        return TSH.Pdm.GetName(ProjetId);
                    }
                    catch (Exception ex)
                    {
                        HandleException(ex, "Impossible de récupérer le nom du projet.");
                        return string.Empty;
                    }
                }
            }

            /// <summary>
            /// Liste des dossiers du projet.
            /// </summary>
            public List<PdmObjectId> FolderIds
            {
                get
                {
                    List<PdmObjectId> folderIds = new List<PdmObjectId>();
                    GetConstituentsRecursive(ProjetId, folderIds, null);
                    return folderIds;
                }
            }

            /// <summary>
            /// Liste des documents du projet.
            /// </summary>
            public List<PdmObjectId> DocumentIds
            {
                get
                {
                    List<PdmObjectId> documentIds = new List<PdmObjectId>();
                    GetConstituentsRecursive(ProjetId, null, documentIds);
                    return documentIds;
                }
            }

            /// <summary>
            /// Dossier "02-Atelier" du projet.
            /// </summary>
            public PdmObjectId AtelierFolderId
            {
                get
                {
                    try
                    {
                        List<PdmObjectId> atelierFolderIds = TSH.Pdm.SearchFolderByName(ProjetId, "02-Atelier");
                        return atelierFolderIds.FirstOrDefault();
                    }
                    catch (Exception ex)
                    {
                        HandleException(ex, "Dossier '02-Atelier' introuvable dans le projet.");
                        return PdmObjectId.Empty;
                    }
                }
            }

            /// <summary>
            /// Méthode récursive pour obtenir les dossiers ou les documents d'un projet.
            /// </summary>
            private void GetConstituentsRecursive(PdmObjectId objectId, List<PdmObjectId> folderIds, List<PdmObjectId> documentIds)
            {
                try
                {
                    IPdm pdm = TSH.Pdm;
                    pdm.GetConstituents(objectId, out List<PdmObjectId> folders, out List<PdmObjectId> documents);

                    if (folderIds != null)
                    {
                        folderIds.AddRange(folders);
                    }

                    if (documentIds != null)
                    {
                        documentIds.AddRange(documents);
                    }

                    if (folderIds != null)
                    {
                        foreach (var folderId in folders)
                        {
                            GetConstituentsRecursive(folderId, folderIds, null);
                        }
                    }
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer les constituants du projet.");
                }
            }

            private void HandleException(Exception ex, string message)
            {
                LogError(message, ex);
                MessageBox.Show($"{message}\nDétails : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            private void LogError(string message, Exception ex)
            {
                string logPath = "errors.log";
                File.AppendAllText(logPath, $"{DateTime.Now}: {message} - {ex.Message}\n");
            }
        }

        public class Document
        {
            #region Propriétés

            /// <summary>
            /// Identifiant unique du document.
            /// </summary>
            public DocumentId DocId { get; set; }

            /// <summary>
            /// Liste des opérations associées au document.
            /// </summary>
            public List<ElementId> Operations => GetOperations(DocId);

            /// <summary>
            /// Liste des opérations CAM si le document est un document CAM.
            /// </summary>
            public List<ElementId> CamOperations => GetCamOperations(DocId);

            /// <summary>
            /// Liste des paramètres associés au document.
            /// </summary>
            public List<ElementId> Parametres => GetDocuParameters(DocId);

            /// <summary>
            /// Nom du document.
            /// </summary>
            public string Nom => GetNomDocu(DocId);

            /// <summary>
            /// Extension du document (ex: .TopPrt, .TopAsm).
            /// </summary>
            public string Extension => GetExtension(PdmObject);

            /// <summary>
            /// Vérifie si le document est un document CAM.
            /// </summary>
            public bool DocuCam => IsDocuCam(DocId);

            /// <summary>
            /// Récupère l'identifiant PDM du document.
            /// </summary>
            public PdmObjectId PdmObject => PdmObjectDocu(DocId);

            /// <summary>
            /// Numéro OP du document (si applicable).
            /// </summary>
            public string OP => NumOP(DocId);

            /// <summary>
            /// ElementId du commentaire système du document (si applicable).
            /// </summary>
            public ElementId CommentaireId => GetCommentaireId(DocId);

            /// <summary>
            /// ElementId de la désignation système du document (si applicable).
            /// </summary>
            public ElementId DesignationId => GetDesignationId(DocId);

            /// <summary>
            /// Projet associé au document.
            /// </summary>
            public Projet Projet => new Projet(GetProjectIdFromDocument(PdmObject));

            /// <summary>
            /// Liste des ElementIds du document.
            /// </summary>
            public List<ElementId> ElementIds => GetElementIds(DocId);

            /// <summary>
            /// Cache pour stocker les noms des documents déjà récupérés afin d'éviter des appels répétés à TopSolid.
            /// </summary>
            private static Dictionary<DocumentId, string> _cacheNomDocuments = new Dictionary<DocumentId, string>();

            #endregion Propriétés

            #region Méthodes

            // Constructeur par défaut
            public Document() { }

            // Constructeur avec initialisation de l'identifiant du document
            public Document(DocumentId docId)
            {
                DocId = docId;
            }

            /// <summary>
            /// Vérifie si un DocumentId est valide (non vide)
            /// </summary>
            private bool IsValidDocument(DocumentId document)
            {
                if (document == DocumentId.Empty)
                {
                    MessageBox.Show("Erreur : DocumentId vide.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                return true;
            }

            /// <summary>
            /// Gère les exceptions et affiche un message d'erreur
            /// </summary>
            private void HandleException(Exception ex, string message)
            {
                LogError(message, ex);
                MessageBox.Show($"{message}\nDétails : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            /// <summary>
            /// Enregistre les erreurs dans un fichier log
            /// </summary>
            private void LogError(string message, Exception ex)
            {
                string logPath = "errors.log";
                File.AppendAllText(logPath, $"{DateTime.Now}: {message} - {ex.Message}\n");
            }

            /// <summary>
            /// Récupère la liste des opérations d'un document
            /// </summary>
            private List<ElementId> GetOperations(DocumentId document)
            {
                if (!IsValidDocument(document)) return new List<ElementId>();

                try
                {
                    return TSH.Operations.GetOperations(document);
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer les opérations.");
                    return new List<ElementId>();
                }
            }

            /// <summary>
            /// Récupère les opérations CAM uniquement si le document est un document CAM
            /// </summary>
            private List<ElementId> GetCamOperations(DocumentId document)
            {
                if (!IsValidDocument(document) || !TSCH.Documents.IsCam(document))
                    return new List<ElementId>();

                try
                {
                    return TSCH.Operations.GetOperations(document);
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer les opérations CAM.");
                    return new List<ElementId>();
                }
            }

            /// <summary>
            /// Récupère la liste des paramètres associés au document
            /// </summary>
            private List<ElementId> GetDocuParameters(DocumentId document)
            {
                if (!IsValidDocument(document)) return new List<ElementId>();

                try
                {
                    return TSH.Parameters.GetParameters(document);
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer les paramètres.");
                    return new List<ElementId>();
                }
            }

            /// <summary>
            /// Récupère le nom du document et le met en cache pour éviter les appels répétitifs
            /// </summary>
            private string GetNomDocu(DocumentId document)
            {
                if (!IsValidDocument(document)) return string.Empty;

                // Vérifie si le nom du document est déjà en cache
                if (_cacheNomDocuments.ContainsKey(document))
                    return _cacheNomDocuments[document];

                try
                {
                    string nom = TSH.Documents.GetName(document);
                    _cacheNomDocuments[document] = nom; // Stocke en cache
                    return nom;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer le nom du document.");
                    return string.Empty;
                }
            }

            /// <summary>
            /// Récupère l'extension du document (ex: .TopPrt, .TopAsm)
            /// </summary>
            private string GetExtension(PdmObjectId pdmObject)
            {
                string extension = string.Empty;
                try
                {
                    TSH.Pdm.GetType(pdmObject, out extension);
                    return extension;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer l'extension du document.");
                    return string.Empty;
                }
            }

            /// <summary>
            /// Vérifie si le document est un document CAM
            /// </summary>
            private bool IsDocuCam(DocumentId document)
            {
                if (!IsValidDocument(document)) return false;

                return TSCH.Documents.IsCam(document);
            }

            /// <summary>
            /// Récupère l'identifiant PDM du document
            /// </summary>
            private PdmObjectId PdmObjectDocu(DocumentId document)
            {
                if (!IsValidDocument(document)) return new PdmObjectId();

                try
                {
                    return TSH.Documents.GetPdmObject(document);
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer l'ID PDM du document.");
                    return new PdmObjectId();
                }
            }

            /// <summary>
            /// Récupère le numéro OP du document (si applicable)
            /// </summary>
            private string NumOP(DocumentId document)
            {
                if (!IsValidDocument(document)) return string.Empty;

                try
                {
                    ElementId OP = TSH.Elements.SearchByName(document, "OP");
                    return OP != ElementId.Empty ? TSH.Parameters.GetTextValue(OP) : string.Empty;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer le numéro OP.");
                    return string.Empty;
                }
            }

            /// <summary>
            /// Récupère le paramètre Commentaire du document (si applicable)
            /// </summary>
            private ElementId GetCommentaireId(DocumentId document)
            {
                if (!IsValidDocument(document))
                    return ElementId.Empty;

                try
                {
                    ElementId commentaireId = TSH.Parameters.GetCommentParameter(document);
                    return commentaireId;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer le numéro OP.");
                    return ElementId.Empty;
                }
            }

            /// <summary>
            /// Récupère le paramètre Désignation du document
            /// </summary>
            private ElementId GetDesignationId(DocumentId document)
            {
                if (!IsValidDocument(document))
                    return ElementId.Empty;

                try
                {
                    ElementId designationId = TSH.Parameters.GetDescriptionParameter(document);
                    return designationId;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer le numéro OP.");
                    return ElementId.Empty;
                }
            }

            /// <summary>
            /// Récupère le PdmObjectId du projet du document
            /// </summary>
            private PdmObjectId GetProjectIdFromDocument(PdmObjectId document)
            {
                try
                {
                    return TSH.Pdm.GetProject(document);
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer le projet du document courant.");
                    return PdmObjectId.Empty;
                }
            }

            /// <summary>
            /// Récupère les ElementIds du document
            /// </summary>
            private List<ElementId> GetElementIds(DocumentId document)
            {
                List<ElementId> elementsId = new List<ElementId>();

                if (!IsValidDocument(document))
                    return elementsId;

                try
                {
                    elementsId = TSH.Elements.GetElements(document);
                    return elementsId;
                }
                catch (Exception ex)
                {
                    HandleException(ex, "Impossible de récupérer les ElementIds du document courant.");
                    return elementsId;
                }
            }

            #endregion Méthodes
        }
    }
}