namespace Gsync.Utilities.Dialogs
{
    partial class InputBoxViewer
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
            this.L1h = new System.Windows.Forms.TableLayoutPanel();
            this.L1h1L2v = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.Cancel = new System.Windows.Forms.Button();
            this.Ok = new System.Windows.Forms.Button();
            this.Message = new System.Windows.Forms.TextBox();
            this.Input = new System.Windows.Forms.TextBox();
            this.L1h.SuspendLayout();
            this.L1h1L2v.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // L1h
            // 
            this.L1h.ColumnCount = 1;
            this.L1h.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1h.Controls.Add(this.L1h1L2v, 0, 0);
            this.L1h.Controls.Add(this.Input, 0, 1);
            this.L1h.Dock = System.Windows.Forms.DockStyle.Fill;
            this.L1h.Location = new System.Drawing.Point(0, 0);
            this.L1h.Name = "L1h";
            this.L1h.RowCount = 2;
            this.L1h.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1h.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.L1h.Size = new System.Drawing.Size(825, 316);
            this.L1h.TabIndex = 0;
            // 
            // L1h1L2v
            // 
            this.L1h1L2v.ColumnCount = 2;
            this.L1h1L2v.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1h1L2v.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 200F));
            this.L1h1L2v.Controls.Add(this.tableLayoutPanel1, 1, 0);
            this.L1h1L2v.Controls.Add(this.Message, 0, 0);
            this.L1h1L2v.Dock = System.Windows.Forms.DockStyle.Fill;
            this.L1h1L2v.Location = new System.Drawing.Point(3, 3);
            this.L1h1L2v.Name = "L1h1L2v";
            this.L1h1L2v.RowCount = 1;
            this.L1h1L2v.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1h1L2v.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 210F));
            this.L1h1L2v.Size = new System.Drawing.Size(819, 210);
            this.L1h1L2v.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Controls.Add(this.Cancel, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.Ok, 0, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(622, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 33.33333F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(194, 204);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // Cancel
            // 
            this.Cancel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Cancel.Font = new System.Drawing.Font("SF Pro Heavy", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Cancel.Location = new System.Drawing.Point(0, 75);
            this.Cancel.Margin = new System.Windows.Forms.Padding(0, 7, 25, 7);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(169, 54);
            this.Cancel.TabIndex = 2;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Ok
            // 
            this.Ok.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Ok.Font = new System.Drawing.Font("SF Pro Heavy", 10.125F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Ok.Location = new System.Drawing.Point(0, 7);
            this.Ok.Margin = new System.Windows.Forms.Padding(0, 7, 25, 7);
            this.Ok.Name = "Ok";
            this.Ok.Size = new System.Drawing.Size(169, 54);
            this.Ok.TabIndex = 1;
            this.Ok.Text = "OK";
            this.Ok.UseVisualStyleBackColor = true;
            this.Ok.Click += new System.EventHandler(this.Ok_Click);
            // 
            // Message
            // 
            this.Message.BackColor = System.Drawing.SystemColors.Control;
            this.Message.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Message.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Message.Font = new System.Drawing.Font("SF Pro Rounded", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Message.Location = new System.Drawing.Point(25, 8);
            this.Message.Margin = new System.Windows.Forms.Padding(25, 8, 10, 8);
            this.Message.Multiline = true;
            this.Message.Name = "Message";
            this.Message.Size = new System.Drawing.Size(584, 194);
            this.Message.TabIndex = 1;
            this.Message.TabStop = false;
            this.Message.Text = "[Message]";
            // 
            // Input
            // 
            this.Input.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Input.Font = new System.Drawing.Font("SF Pro Rounded", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Input.Location = new System.Drawing.Point(25, 241);
            this.Input.Margin = new System.Windows.Forms.Padding(25);
            this.Input.Name = "Input";
            this.Input.Size = new System.Drawing.Size(775, 46);
            this.Input.TabIndex = 0;
            this.Input.Text = "[Input]";
            // 
            // InputBoxViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(825, 316);
            this.Controls.Add(this.L1h);
            this.Name = "InputBoxViewer";
            this.Text = "InputBoxViewer";
            this.L1h.ResumeLayout(false);
            this.L1h.PerformLayout();
            this.L1h1L2v.ResumeLayout(false);
            this.L1h1L2v.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel L1h;
        private System.Windows.Forms.TableLayoutPanel L1h1L2v;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        internal System.Windows.Forms.TextBox Input;
        internal System.Windows.Forms.Button Cancel;
        internal System.Windows.Forms.Button Ok;
        internal System.Windows.Forms.TextBox Message;
    }
}