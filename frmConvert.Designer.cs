namespace ConvertApp
{
    partial class frmConvert
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCnvWord = new System.Windows.Forms.Button();
            this.btnCnvExcel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCnvWord
            // 
            this.btnCnvWord.Location = new System.Drawing.Point(54, 17);
            this.btnCnvWord.Name = "btnCnvWord";
            this.btnCnvWord.Size = new System.Drawing.Size(176, 58);
            this.btnCnvWord.TabIndex = 0;
            this.btnCnvWord.Text = "Convert Word";
            this.btnCnvWord.UseVisualStyleBackColor = true;
            this.btnCnvWord.Click += new System.EventHandler(this.btnCnvWord_Click);
            // 
            // btnCnvExcel
            // 
            this.btnCnvExcel.Location = new System.Drawing.Point(54, 94);
            this.btnCnvExcel.Name = "btnCnvExcel";
            this.btnCnvExcel.Size = new System.Drawing.Size(176, 54);
            this.btnCnvExcel.TabIndex = 1;
            this.btnCnvExcel.Text = "Convert Excel";
            this.btnCnvExcel.UseVisualStyleBackColor = true;
            this.btnCnvExcel.Click += new System.EventHandler(this.btnCnvExcel_Click);
            // 
            // frmConvert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 160);
            this.Controls.Add(this.btnCnvExcel);
            this.Controls.Add(this.btnCnvWord);
            this.Name = "frmConvert";
            this.Text = "Convert";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnCnvWord;
        private System.Windows.Forms.Button btnCnvExcel;
    }
}

