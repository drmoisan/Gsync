namespace Gsync.Utilities.Dialogs
{
    partial class DelegateButtonTemplate
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.Button1 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.Button1);
            this.panel1.Location = new System.Drawing.Point(42, 35);
            this.panel1.Margin = new System.Windows.Forms.Padding(0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(218, 91);
            this.panel1.TabIndex = 0;
            // 
            // Button1
            // 
            this.Button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.Button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Button1.Location = new System.Drawing.Point(0, 0);
            this.Button1.Margin = new System.Windows.Forms.Padding(6);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(218, 91);
            this.Button1.TabIndex = 1;
            this.Button1.Text = "button1";
            this.Button1.UseVisualStyleBackColor = true;
            // 
            // DelegateButtonTemplate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(747, 250);
            this.Controls.Add(this.panel1);
            this.Name = "DelegateButtonTemplate";
            this.Text = "DelegateButtonTemplate";
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        internal System.Windows.Forms.Panel panel1;
        internal System.Windows.Forms.Button Button1;
    }
}