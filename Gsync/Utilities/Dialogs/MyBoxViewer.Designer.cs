namespace Gsync.Utilities.Dialogs
{
    partial class MyBoxViewer
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param _name="disposing">true if managed resources should be disposed; otherwise, false.</param>
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyBoxViewer));
            this.L1Vertical = new System.Windows.Forms.TableLayoutPanel();
            this.L2Bottom = new System.Windows.Forms.TableLayoutPanel();
            this.Button2 = new System.Windows.Forms.Button();
            this.Button1 = new System.Windows.Forms.Button();
            this.L1v1L2h = new System.Windows.Forms.TableLayoutPanel();
            this.TextMessage = new System.Windows.Forms.TextBox();
            this.SvgIcon = new SVGControl.PictureBoxSVG();
            this.L1Vertical.SuspendLayout();
            this.L2Bottom.SuspendLayout();
            this.L1v1L2h.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SvgIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // L1Vertical
            // 
            this.L1Vertical.ColumnCount = 1;
            this.L1Vertical.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1Vertical.Controls.Add(this.L2Bottom, 0, 1);
            this.L1Vertical.Controls.Add(this.L1v1L2h, 0, 0);
            this.L1Vertical.Dock = System.Windows.Forms.DockStyle.Fill;
            this.L1Vertical.Location = new System.Drawing.Point(0, 0);
            this.L1Vertical.Margin = new System.Windows.Forms.Padding(6);
            this.L1Vertical.Name = "L1Vertical";
            this.L1Vertical.RowCount = 2;
            this.L1Vertical.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1Vertical.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 115F));
            this.L1Vertical.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 38F));
            this.L1Vertical.Size = new System.Drawing.Size(490, 275);
            this.L1Vertical.TabIndex = 0;
            // 
            // L2Bottom
            // 
            this.L2Bottom.ColumnCount = 4;
            this.L2Bottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.L2Bottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 230F));
            this.L2Bottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 230F));
            this.L2Bottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.L2Bottom.Controls.Add(this.Button2, 2, 0);
            this.L2Bottom.Controls.Add(this.Button1, 1, 0);
            this.L2Bottom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.L2Bottom.Location = new System.Drawing.Point(6, 166);
            this.L2Bottom.Margin = new System.Windows.Forms.Padding(6);
            this.L2Bottom.Name = "L2Bottom";
            this.L2Bottom.RowCount = 1;
            this.L2Bottom.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L2Bottom.Size = new System.Drawing.Size(478, 103);
            this.L2Bottom.TabIndex = 0;
            // 
            // Button2
            // 
            this.Button2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.Button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Button2.Location = new System.Drawing.Point(245, 6);
            this.Button2.Margin = new System.Windows.Forms.Padding(6);
            this.Button2.Name = "Button2";
            this.Button2.Size = new System.Drawing.Size(218, 91);
            this.Button2.TabIndex = 1;
            this.Button2.Text = "button2";
            this.Button2.UseVisualStyleBackColor = true;
            this.Button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // Button1
            // 
            this.Button1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.Button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Button1.Location = new System.Drawing.Point(15, 6);
            this.Button1.Margin = new System.Windows.Forms.Padding(6);
            this.Button1.Name = "Button1";
            this.Button1.Size = new System.Drawing.Size(218, 91);
            this.Button1.TabIndex = 0;
            this.Button1.Text = "button1";
            this.Button1.UseVisualStyleBackColor = true;
            this.Button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // L1v1L2h
            // 
            this.L1v1L2h.ColumnCount = 2;
            this.L1v1L2h.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.L1v1L2h.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1v1L2h.Controls.Add(this.TextMessage, 1, 0);
            this.L1v1L2h.Controls.Add(this.SvgIcon, 0, 0);
            this.L1v1L2h.Dock = System.Windows.Forms.DockStyle.Fill;
            this.L1v1L2h.Location = new System.Drawing.Point(6, 6);
            this.L1v1L2h.Margin = new System.Windows.Forms.Padding(6);
            this.L1v1L2h.Name = "L1v1L2h";
            this.L1v1L2h.RowCount = 1;
            this.L1v1L2h.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.L1v1L2h.Size = new System.Drawing.Size(478, 148);
            this.L1v1L2h.TabIndex = 1;
            // 
            // TextMessage
            // 
            this.TextMessage.BackColor = System.Drawing.SystemColors.Control;
            this.TextMessage.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.TextMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TextMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TextMessage.Location = new System.Drawing.Point(86, 6);
            this.TextMessage.Margin = new System.Windows.Forms.Padding(6);
            this.TextMessage.Multiline = true;
            this.TextMessage.Name = "TextMessage";
            this.TextMessage.Size = new System.Drawing.Size(386, 136);
            this.TextMessage.TabIndex = 1;
            this.TextMessage.Text = "Message Contents";
            this.TextMessage.TextChanged += new System.EventHandler(this.TextMessage_TextChanged);
            // 
            // SvgIcon
            // 
            this.SvgIcon.Image = ((System.Drawing.Image)(resources.GetObject("SvgIcon.Image")));
            this.SvgIcon.Location = new System.Drawing.Point(6, 6);
            this.SvgIcon.Margin = new System.Windows.Forms.Padding(6);
            this.SvgIcon.Name = "SvgIcon";
            this.SvgIcon.Size = new System.Drawing.Size(68, 136);
            this.SvgIcon.TabIndex = 2;
            this.SvgIcon.TabStop = false;
            // 
            // MyBoxViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 275);
            this.Controls.Add(this.L1Vertical);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.MinimumSize = new System.Drawing.Size(496, 284);
            this.Name = "MyBoxViewer";
            this.Text = "FormName";
            this.L1Vertical.ResumeLayout(false);
            this.L2Bottom.ResumeLayout(false);
            this.L1v1L2h.ResumeLayout(false);
            this.L1v1L2h.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SvgIcon)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel L1Vertical;
        private System.Windows.Forms.Button Button2;
        private System.Windows.Forms.Button Button1;
        private System.Windows.Forms.TableLayoutPanel L1v1L2h;
        internal System.Windows.Forms.TableLayoutPanel L2Bottom;
        internal System.Windows.Forms.TextBox TextMessage;
        internal SVGControl.PictureBoxSVG SvgIcon;
    }
}