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
            PdmObjectId CurrentProjectPdmId;
            string CurrentProjectName;
            List<PdmObjectId> FolderIDs = new List<PdmObjectId>();

            {

                bool TSConnected = TopSolidDesignHost.IsConnected;
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

                 try
                 {
                      CurrentProjectPdmId = TopSolidHost.Pdm.GetCurrentProject();   // Récupération ID projet
                 }
                 catch (Exception ex)
                 {
                     MessageBox.Show("Echec de la récupération de l'id du projet courant" + ex.Message);
                     return;
                 }

                 try
                 {
                   CurrentProjectName = TopSolidHost.Pdm.GetName(CurrentProjectPdmId);  // Récupération Nom projet
                   string TextBoxProjectName = CurrentProjectName;
                 
                    MessageBox.Show (TextBoxProjectName);
                 }
                 catch (Exception ex)
                 {
                   MessageBox.Show("Echec de la récupération de l'id du projet courant" + ex.Message);
                 }

                  TopSolidHost.Pdm.GetConstituents(CurrentProjectPdmId, FolderIDs);




            }
        }
    }
}
    



     

