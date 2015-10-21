namespace ReaderFromExcel
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
            this.excelName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.browser = new System.Windows.Forms.Button();
            this.convertXML = new System.Windows.Forms.Button();
            this.saveXML = new System.Windows.Forms.Button();
            this.xmlView = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // excelName
            // 
            this.excelName.Location = new System.Drawing.Point(87, 12);
            this.excelName.Name = "excelName";
            this.excelName.Size = new System.Drawing.Size(199, 20);
            this.excelName.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Excel: ";
            // 
            // browser
            // 
            this.browser.Location = new System.Drawing.Point(308, 12);
            this.browser.Name = "browser";
            this.browser.Size = new System.Drawing.Size(75, 23);
            this.browser.TabIndex = 2;
            this.browser.Text = "Browser";
            this.browser.UseVisualStyleBackColor = true;
            this.browser.Click += new System.EventHandler(this.browser_Click);
            // 
            // convertXML
            // 
            this.convertXML.Location = new System.Drawing.Point(12, 50);
            this.convertXML.Name = "convertXML";
            this.convertXML.Size = new System.Drawing.Size(91, 23);
            this.convertXML.TabIndex = 3;
            this.convertXML.Text = "Convert to XML";
            this.convertXML.UseVisualStyleBackColor = true;
            //this.convertXML.Click += new System.EventHandler(this.convertXML_Click);
            // 
            // saveXML
            // 
            this.saveXML.Location = new System.Drawing.Point(12, 227);
            this.saveXML.Name = "saveXML";
            this.saveXML.Size = new System.Drawing.Size(75, 23);
            this.saveXML.TabIndex = 4;
            this.saveXML.Text = "Save XML";
            this.saveXML.UseVisualStyleBackColor = true;
            this.saveXML.Click += new System.EventHandler(this.saveXML_Click);
            // 
            // xmlView
            // 
            this.xmlView.Enabled = false;
            this.xmlView.Location = new System.Drawing.Point(87, 99);
            this.xmlView.Multiline = true;
            this.xmlView.Name = "xmlView";
            this.xmlView.Size = new System.Drawing.Size(199, 101);
            this.xmlView.TabIndex = 5;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 262);
            this.Controls.Add(this.xmlView);
            this.Controls.Add(this.saveXML);
            this.Controls.Add(this.convertXML);
            this.Controls.Add(this.browser);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.excelName);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox excelName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button browser;
        private System.Windows.Forms.Button convertXML;
        private System.Windows.Forms.Button saveXML;
        private System.Windows.Forms.TextBox xmlView;
    }
}

