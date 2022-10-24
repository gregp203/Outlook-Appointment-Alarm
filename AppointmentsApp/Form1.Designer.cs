namespace AppointmentsApp {
    partial class Form1 {
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
            this.apptDGV = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.apptDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // apptDGV
            // 
            this.apptDGV.AllowUserToAddRows = false;
            this.apptDGV.AllowUserToDeleteRows = false;
            this.apptDGV.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.apptDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.apptDGV.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnKeystroke;
            this.apptDGV.Location = new System.Drawing.Point(12, 12);
            this.apptDGV.Name = "apptDGV";
            this.apptDGV.RowHeadersVisible = false;
            this.apptDGV.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.apptDGV.Size = new System.Drawing.Size(773, 169);
            this.apptDGV.TabIndex = 0;
            this.apptDGV.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.apptDGV_CellEndEdit);
            this.apptDGV.CurrentCellDirtyStateChanged += new System.EventHandler(this.apptDGV_CurrentCellDirtyStateChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(797, 193);
            this.Controls.Add(this.apptDGV);
            this.Name = "Form1";
            this.Text = "Calander Events For Today";
            ((System.ComponentModel.ISupportInitialize)(this.apptDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.DataGridView apptDGV;
    }
}

