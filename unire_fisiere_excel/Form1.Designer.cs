
namespace unire_fisiere_excel
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.numeDirector = new System.Windows.Forms.Label();
            this.od1 = new System.Windows.Forms.OpenFileDialog();
            this.lista = new System.Windows.Forms.ListBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Cauta fisiere";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // numeDirector
            // 
            this.numeDirector.AutoSize = true;
            this.numeDirector.Location = new System.Drawing.Point(12, 47);
            this.numeDirector.Name = "numeDirector";
            this.numeDirector.Size = new System.Drawing.Size(35, 13);
            this.numeDirector.TabIndex = 1;
            this.numeDirector.Text = "label1";
            // 
            // od1
            // 
            this.od1.FileName = "openFileDialog1";
            this.od1.Filter = "Fisier Excel|*.xlsx";
            // 
            // lista
            // 
            this.lista.FormattingEnabled = true;
            this.lista.Location = new System.Drawing.Point(6, 19);
            this.lista.Name = "lista";
            this.lista.Size = new System.Drawing.Size(135, 303);
            this.lista.TabIndex = 2;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lista);
            this.groupBox1.Location = new System.Drawing.Point(12, 63);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 340);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Lista de fisiere";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.numeDirector);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Unire fisiere Excel";
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label numeDirector;
        private System.Windows.Forms.OpenFileDialog od1;
        private System.Windows.Forms.ListBox lista;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

