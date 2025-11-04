namespace Inventario
{
    partial class PantallaPrincipal
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            panelHeader = new Panel();
            lblTitulo = new Label();
            panelCentral = new Panel();
            btnContinuar = new Button();
            lblArchivoSeleccionado = new Label();
            btnCargarArchivo = new Button();
            lblInstruccion = new Label();
            lblBienvenida = new Label();
            panelFooter = new Panel();
            lblVersion = new Label();
            panelHeader.SuspendLayout();
            panelCentral.SuspendLayout();
            panelFooter.SuspendLayout();
            SuspendLayout();
            //
            // panelHeader
            //
            panelHeader.BackColor = Color.FromArgb(41, 128, 185);
            panelHeader.Controls.Add(lblTitulo);
            panelHeader.Dock = DockStyle.Top;
            panelHeader.Location = new Point(0, 0);
            panelHeader.Name = "panelHeader";
            panelHeader.Size = new Size(1000, 80);
            panelHeader.TabIndex = 0;
            //
            // lblTitulo
            //
            lblTitulo.AutoSize = true;
            lblTitulo.Font = new Font("Segoe UI", 24F, FontStyle.Bold);
            lblTitulo.ForeColor = Color.White;
            lblTitulo.Location = new Point(30, 20);
            lblTitulo.Name = "lblTitulo";
            lblTitulo.Size = new Size(448, 45);
            lblTitulo.TabIndex = 0;
            lblTitulo.Text = "📦 StockControl";
            //
            // panelCentral
            //
            panelCentral.BackColor = Color.White;
            panelCentral.Controls.Add(btnContinuar);
            panelCentral.Controls.Add(lblArchivoSeleccionado);
            panelCentral.Controls.Add(btnCargarArchivo);
            panelCentral.Controls.Add(lblInstruccion);
            panelCentral.Controls.Add(lblBienvenida);
            panelCentral.Location = new Point(150, 140);
            panelCentral.Name = "panelCentral";
            panelCentral.Size = new Size(700, 400);
            panelCentral.TabIndex = 1;
            //
            // btnContinuar
            //
            btnContinuar.BackColor = Color.FromArgb(46, 204, 113);
            btnContinuar.Cursor = Cursors.Hand;
            btnContinuar.Enabled = false;
            btnContinuar.FlatAppearance.BorderSize = 0;
            btnContinuar.FlatStyle = FlatStyle.Flat;
            btnContinuar.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnContinuar.ForeColor = Color.White;
            btnContinuar.Location = new Point(225, 320);
            btnContinuar.Name = "btnContinuar";
            btnContinuar.Size = new Size(250, 50);
            btnContinuar.TabIndex = 4;
            btnContinuar.Text = "Continuar";
            btnContinuar.UseVisualStyleBackColor = false;
            btnContinuar.Visible = false;
            btnContinuar.Click += btnContinuar_Click;
            //
            // lblArchivoSeleccionado
            //
            lblArchivoSeleccionado.Font = new Font("Segoe UI", 10F, FontStyle.Italic);
            lblArchivoSeleccionado.ForeColor = Color.FromArgb(46, 204, 113);
            lblArchivoSeleccionado.Location = new Point(50, 270);
            lblArchivoSeleccionado.Name = "lblArchivoSeleccionado";
            lblArchivoSeleccionado.Size = new Size(600, 30);
            lblArchivoSeleccionado.TabIndex = 3;
            lblArchivoSeleccionado.Text = "Archivo: ninguno";
            lblArchivoSeleccionado.TextAlign = ContentAlignment.MiddleCenter;
            lblArchivoSeleccionado.Visible = false;
            //
            // btnCargarArchivo
            //
            btnCargarArchivo.BackColor = Color.FromArgb(41, 128, 185);
            btnCargarArchivo.Cursor = Cursors.Hand;
            btnCargarArchivo.FlatAppearance.BorderSize = 0;
            btnCargarArchivo.FlatStyle = FlatStyle.Flat;
            btnCargarArchivo.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            btnCargarArchivo.ForeColor = Color.White;
            btnCargarArchivo.Location = new Point(175, 190);
            btnCargarArchivo.Name = "btnCargarArchivo";
            btnCargarArchivo.Size = new Size(350, 60);
            btnCargarArchivo.TabIndex = 2;
            btnCargarArchivo.Text = "📂  Cargar Archivo Excel";
            btnCargarArchivo.UseVisualStyleBackColor = false;
            btnCargarArchivo.Click += btnCargarArchivo_Click;
            //
            // lblInstruccion
            //
            lblInstruccion.Font = new Font("Segoe UI", 12F);
            lblInstruccion.ForeColor = Color.FromArgb(52, 73, 94);
            lblInstruccion.Location = new Point(50, 120);
            lblInstruccion.Name = "lblInstruccion";
            lblInstruccion.Size = new Size(600, 50);
            lblInstruccion.TabIndex = 1;
            lblInstruccion.Text = "Por favor, carga el archivo de Excel con los datos del inventario para comenzar.";
            lblInstruccion.TextAlign = ContentAlignment.MiddleCenter;
            //
            // lblBienvenida
            //
            lblBienvenida.AutoSize = true;
            lblBienvenida.Font = new Font("Segoe UI", 20F, FontStyle.Bold);
            lblBienvenida.ForeColor = Color.FromArgb(41, 128, 185);
            lblBienvenida.Location = new Point(210, 50);
            lblBienvenida.Name = "lblBienvenida";
            lblBienvenida.Size = new Size(280, 37);
            lblBienvenida.TabIndex = 0;
            lblBienvenida.Text = "¡Bienvenido/a!";
            //
            // panelFooter
            //
            panelFooter.BackColor = Color.FromArgb(52, 73, 94);
            panelFooter.Controls.Add(lblVersion);
            panelFooter.Dock = DockStyle.Bottom;
            panelFooter.Location = new Point(0, 580);
            panelFooter.Name = "panelFooter";
            panelFooter.Size = new Size(1000, 40);
            panelFooter.TabIndex = 2;
            //
            // lblVersion
            //
            lblVersion.AutoSize = true;
            lblVersion.Font = new Font("Segoe UI", 9F);
            lblVersion.ForeColor = Color.White;
            lblVersion.Location = new Point(30, 12);
            lblVersion.Name = "lblVersion";
            lblVersion.Size = new Size(175, 15);
            lblVersion.TabIndex = 0;
            lblVersion.Text = "StockControl v1.2.0 | Desarrollado por Fernando Carrasco";
            //
            // PantallaPrincipal
            //
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(240, 244, 248);
            ClientSize = new Size(1000, 620);
            Controls.Add(panelFooter);
            Controls.Add(panelCentral);
            Controls.Add(panelHeader);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "PantallaPrincipal";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "StockControl - Gestión de Inventarios";
            panelHeader.ResumeLayout(false);
            panelHeader.PerformLayout();
            panelCentral.ResumeLayout(false);
            panelCentral.PerformLayout();
            panelFooter.ResumeLayout(false);
            panelFooter.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panelHeader;
        private Label lblTitulo;
        private Panel panelCentral;
        private Label lblBienvenida;
        private Label lblInstruccion;
        private Button btnCargarArchivo;
        private Label lblArchivoSeleccionado;
        private Button btnContinuar;
        private Panel panelFooter;
        private Label lblVersion;
    }
}
