namespace Folder_Creator_Tool_V3
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonDémarrer = new System.Windows.Forms.Button();
            this.checkBoxExport = new System.Windows.Forms.CheckBox();
            this.checkBoxPDF = new System.Windows.Forms.CheckBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.quiitterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.checkBox3D = new System.Windows.Forms.CheckBox();
            this.checkBoxSimplifier = new System.Windows.Forms.CheckBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonDémarrer
            // 
            this.buttonDémarrer.Location = new System.Drawing.Point(10, 194);
            this.buttonDémarrer.Name = "buttonDémarrer";
            this.buttonDémarrer.Size = new System.Drawing.Size(180, 99);
            this.buttonDémarrer.TabIndex = 0;
            this.buttonDémarrer.Text = "Démarrer";
            this.buttonDémarrer.UseVisualStyleBackColor = true;
            // 
            // checkBoxExport
            // 
            this.checkBoxExport.AutoSize = true;
            this.checkBoxExport.Location = new System.Drawing.Point(6, 19);
            this.checkBoxExport.Name = "checkBoxExport";
            this.checkBoxExport.Size = new System.Drawing.Size(135, 17);
            this.checkBoxExport.TabIndex = 3;
            this.checkBoxExport.Text = "Export PDM ==> atelier";
            this.checkBoxExport.UseVisualStyleBackColor = true;
            // 
            // checkBoxPDF
            // 
            this.checkBoxPDF.AutoSize = true;
            this.checkBoxPDF.Location = new System.Drawing.Point(6, 19);
            this.checkBoxPDF.Name = "checkBoxPDF";
            this.checkBoxPDF.Size = new System.Drawing.Size(98, 17);
            this.checkBoxPDF.TabIndex = 4;
            this.checkBoxPDF.Text = "Archivage PDF";
            this.checkBoxPDF.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox3);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Location = new System.Drawing.Point(10, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(180, 161);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Configuration";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.quiitterToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(200, 24);
            this.menuStrip1.TabIndex = 6;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // quiitterToolStripMenuItem
            // 
            this.quiitterToolStripMenuItem.Name = "quiitterToolStripMenuItem";
            this.quiitterToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.quiitterToolStripMenuItem.Text = "Quiitter";
            this.quiitterToolStripMenuItem.Click += new System.EventHandler(this.quiitterToolStripMenuItem_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.checkBoxPDF);
            this.groupBox2.Controls.Add(this.checkBox3D);
            this.groupBox2.Controls.Add(this.checkBoxSimplifier);
            this.groupBox2.Location = new System.Drawing.Point(6, 19);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(166, 88);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "PDM";
            // 
            // checkBox3D
            // 
            this.checkBox3D.AutoSize = true;
            this.checkBox3D.Location = new System.Drawing.Point(6, 40);
            this.checkBox3D.Name = "checkBox3D";
            this.checkBox3D.Size = new System.Drawing.Size(91, 17);
            this.checkBox3D.TabIndex = 1;
            this.checkBox3D.Text = "Archivage 3D";
            this.checkBox3D.UseVisualStyleBackColor = true;
            // 
            // checkBoxSimplifier
            // 
            this.checkBoxSimplifier.AutoSize = true;
            this.checkBoxSimplifier.Location = new System.Drawing.Point(20, 63);
            this.checkBoxSimplifier.Name = "checkBoxSimplifier";
            this.checkBoxSimplifier.Size = new System.Drawing.Size(67, 17);
            this.checkBoxSimplifier.TabIndex = 2;
            this.checkBoxSimplifier.Text = "Simplifier";
            this.checkBoxSimplifier.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.checkBoxExport);
            this.groupBox3.Location = new System.Drawing.Point(6, 113);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(166, 42);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Serveur atelier";
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(200, 304);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.buttonDémarrer);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form2";
            this.Text = "Form2";
            this.groupBox1.ResumeLayout(false);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonDémarrer;
        private System.Windows.Forms.CheckBox checkBoxExport;
        private System.Windows.Forms.CheckBox checkBoxPDF;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem quiitterToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBox3D;
        private System.Windows.Forms.CheckBox checkBoxSimplifier;
        private System.Windows.Forms.GroupBox groupBox3;
    }
}