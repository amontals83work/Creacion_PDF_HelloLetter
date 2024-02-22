namespace Creacion_PDF_HelloLetter
{
    partial class Form1
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.Label3 = new System.Windows.Forms.Label();
            this.btnFichero = new System.Windows.Forms.Button();
            this.txtFichero = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(114, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "CARTERA";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(23, 91);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(262, 40);
            this.button1.TabIndex = 1;
            this.button1.Text = "Crear PDFs";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Label3
            // 
            this.Label3.AutoSize = true;
            this.Label3.Location = new System.Drawing.Point(20, 56);
            this.Label3.Name = "Label3";
            this.Label3.Size = new System.Drawing.Size(42, 13);
            this.Label3.TabIndex = 17;
            this.Label3.Text = "Fichero";
            // 
            // btnFichero
            // 
            this.btnFichero.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFichero.Location = new System.Drawing.Point(256, 51);
            this.btnFichero.Name = "btnFichero";
            this.btnFichero.Size = new System.Drawing.Size(29, 23);
            this.btnFichero.TabIndex = 16;
            this.btnFichero.Text = "...";
            this.btnFichero.UseVisualStyleBackColor = true;
            this.btnFichero.Click += new System.EventHandler(this.btnFichero_Click);
            // 
            // txtFichero
            // 
            this.txtFichero.Location = new System.Drawing.Point(68, 52);
            this.txtFichero.Name = "txtFichero";
            this.txtFichero.Size = new System.Drawing.Size(183, 20);
            this.txtFichero.TabIndex = 15;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(313, 151);
            this.Controls.Add(this.Label3);
            this.Controls.Add(this.btnFichero);
            this.Controls.Add(this.txtFichero);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Button btnFichero;
        internal System.Windows.Forms.TextBox txtFichero;
    }
}

