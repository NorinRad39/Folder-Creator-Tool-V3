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
    public class StartConnect
    {

        //Connection topsolid
        private void ConnectToTopSolid()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TSH.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TSH.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TSH.IsConnected)
                    {
                        //MessageBox.Show("Connexion réussie à TopSolid.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Connexion échouée à TopSolid.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("TopSolid est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                MessageBox.Show($"Erreur lors de la connexion à TopSolid : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Connection topsolid design
        private void ConnectToTopSolidDesignHost()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TopSolidDesignHost.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TopSolidDesignHost.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TopSolidDesignHost.IsConnected)
                    {
                        //MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //Connection topsolid cam
        private void ConnectToTopSolidCamHost()
        {
            try
            {
                // Vérifier si la connexion est déjà établie
                if (!TopSolidCamHost.IsConnected)
                {
                    // Connexion à TopSolid avec un paramètre d'initialisation (si nécessaire)
                    TopSolidCamHost.Connect();

                    // Vérifier à nouveau si la connexion est réussie
                    if (TopSolidCamHost.IsConnected)
                    {
                        //MessageBox.Show("Connexion réussie à TopSolid module design.", "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Connexion échouée à TopSolid module design.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("TopSolid module design est déjà connecté.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (InvalidOperationException ex)
            {
                // Gérer une exception spécifique si nécessaire
                MessageBox.Show($"Problème opérationnel : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // Gérer d'autres exceptions
                MessageBox.Show($"Erreur lors de la connexion à TopSolid module design : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Méthode pour connecter tous les modules TopSolid
        public void ConnectionTopsolid()
        {
            ConnectToTopSolid();
            ConnectToTopSolidDesignHost();
            //ConnectToTopSolidCamHost();
        }

    }
}
