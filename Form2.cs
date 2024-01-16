﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;


namespace Folder_Creator_Tool_V3
{
    public partial class Launcher : Form
    {
        public Launcher()
        {
            InitializeComponent();
            // Attachez les événements ici, dans le constructeur.
            checkBox3D.CheckedChanged += CheckBox3D_CheckedChanged;
            checkBoxPDF.CheckedChanged += CheckBoxPDF_CheckedChanged;

            this.FormClosing += new FormClosingEventHandler(Form_FormClosing);
            this.Load += new EventHandler(Form_Load);
        }

        private void CheckBox3D_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3D.Checked)
            {
                checkBoxPDF.Checked = true;
            }
        }

        private void CheckBoxPDF_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBoxPDF.Checked)
            {
                checkBox3D.Checked = false;
            }
        }


        private void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Sauvegarde de l'état de la case à cocher lors de la fermeture de l'application.
            Properties.Settings.Default.CheckBoxState = checkBox3D.Checked;
            Properties.Settings.Default.CheckBoxState = checkBoxPDF.Checked;

            Properties.Settings.Default.Save();
        }

        private void Form_Load(object sender, EventArgs e)
        {
            // Restauration de l'état de la case à cocher au démarrage de l'application.
            checkBox3D.Checked = Properties.Settings.Default.CheckBoxState;
            checkBoxPDF.Checked = Properties.Settings.Default.CheckBoxState;


        }

        private void quiitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Exit l'application.
            Application.Exit();
        }

        private void buttonDémarrer_Click(object sender, EventArgs e)
        {
            if (checkBox3D.Checked)
            {
                
            }
        }
    }
}
