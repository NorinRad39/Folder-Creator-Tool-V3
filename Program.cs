using System;
using System.IO;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;
using AutoUpdaterDotNET;

namespace Folder_Creator_Tool_V3
{
    /// <summary>
    /// Modele WinForms : controle de version obligatoire avant la premiere fenetre.
    /// </summary>
    /// <remarks>
    /// A adapter : le namespace, <see cref="DescripteurMiseAJour"/>, <see cref="NomApplication"/>
    /// et la fenetre passee a Application.Run.
    ///
    /// Prealable : paquet NuGet « AutoUpdater.NET.Official ».
    /// </remarks>
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Controle de version avant toute fenetre : une version en retard ne doit meme pas
            // s'ouvrir sur un document.
            if (!PeutDemarrer()) return;

            Application.Run(new Form1());
        }

        /// <summary>Adresse du descripteur de mise a jour, sur le partage reseau.</summary>
        /// <remarks>
        /// Chemin UNC et non une lettre de lecteur, qui designerait pourtant le meme dossier : un
        /// mappage est propre a la session, et sur un poste ou il manque les mises a jour
        /// cesseraient sans que personne ne s'en apercoive.
        ///
        /// Deux antislashs en tete : c'est une chaine verbatim (arobase-guillemet), rien n'y est
        /// echappe. Avec un seul, System.Uri leve UriFormatException, l'exception est avalee par le
        /// catch plus bas, et l'application demarre sans jamais se mettre a jour — sans le moindre
        /// signe.
        ///
        /// Cette adresse ne doit plus bouger. Chaque poste la porte en dur dans l'executable qu'il
        /// a installe : la deplacer coupe des mises a jour tous ceux deja en service.
        /// </remarks>
        private const string DescripteurMiseAJour = @"\\jbtec-be\meca$\topsolid\Folder-Creator-Tool-V3\update.xml";

        /// <summary>Nom affiche a l'utilisateur dans les messages de mise a jour.</summary>
        private const string NomApplication = "Folder-Creator-Tool-V3";

        /// <summary>
        /// Indique si l'application est autorisee a demarrer, apres controle de sa version.
        /// </summary>
        /// <remarks>
        /// La mise a jour peut etre refusee, mais l'application ne demarre pas tant qu'elle n'est
        /// pas faite : tous les postes doivent tourner sur la meme version.
        ///
        /// Le controle est fait ici, et non par <c>AutoUpdater.Start</c> : celui-ci mene le deroule
        /// de bout en bout et ne dit pas si l'utilisateur a refuse. Ses boites de dialogue
        /// s'affichent tres bien sans boucle de messages, elles pompent la leur.
        /// </remarks>
        private static bool PeutDemarrer()
        {
            UpdateInfoEventArgs descripteur;

            try
            {
                descripteur = LireDescripteurMiseAJour();
            }
            catch (Exception ex)
            {
                // Partage injoignable, poste hors reseau : on laisse travailler plutot que
                // d'immobiliser sur un incident de reseau. Ne pas savoir n'est pas la meme chose
                // que savoir qu'une version manque.
                Console.WriteLine($"[Mise a jour] Verification impossible : {ex.Message}");
                return true;
            }

            Version installee = Assembly.GetExecutingAssembly().GetName().Version;
            if (descripteur == null || string.IsNullOrWhiteSpace(descripteur.CurrentVersion)) return true;
            if (new Version(descripteur.CurrentVersion) <= installee) return true;

            DialogResult reponse = MessageBox.Show(
                $"La version {descripteur.CurrentVersion} est disponible ; ce poste utilise la {installee}.\n\n"
                + NomApplication + " ne peut pas s'ouvrir tant que la mise a jour n'est pas faite.",
                "Mise a jour requise",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);

            if (reponse != DialogResult.OK)
            {
                Console.WriteLine("[Mise a jour] Refusee : l'application ne demarre pas.");
                return false;
            }

            // L'installation se fait par utilisateur, dans %LOCALAPPDATA% : aucune elevation n'est
            // necessaire. Sans ce reglage, AutoUpdater lance le programme d'installation avec le
            // verbe « runas » et declenche une invite UAC pour rien.
            AutoUpdater.RunUpdateAsAdmin = false;

            // Rend la main une fois le programme d'installation lance : celui-ci remplace les
            // fichiers puis relance l'application, il ne faut donc pas la laisser demarrer ici.
            if (AutoUpdater.DownloadUpdate(descripteur)) return false;

            MessageBox.Show(
                "Le telechargement de la mise a jour n'a pas abouti.\n\n"
                + "Verifiez l'acces au reseau, puis relancez " + NomApplication + ".",
                "Mise a jour",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);

            return false;
        }

        /// <summary>
        /// Lit le descripteur publie sur le partage.
        /// </summary>
        /// <remarks>
        /// Meme format et meme mecanisme que ceux d'AutoUpdater — WebClient sait lire une URI
        /// file://, donc un chemin UNC — mais lu ici pour garder la decision de demarrer.
        /// </remarks>
        private static UpdateInfoEventArgs LireDescripteurMiseAJour()
        {
            using (WebClient client = new WebClient())
            {
                // Sans cela, un update.xml fraichement publie peut rester masque par le cache le
                // temps que les postes le voient.
                client.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore);

                string xml = client.DownloadString(new Uri(DescripteurMiseAJour));

                XmlSerializer serialiseur = new XmlSerializer(typeof(UpdateInfoEventArgs));
                using (StringReader lecteur = new StringReader(xml))
                {
                    return (UpdateInfoEventArgs)serialiseur.Deserialize(lecteur);
                }
            }
        }
    }
}
