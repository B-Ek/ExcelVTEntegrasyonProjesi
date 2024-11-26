namespace ExcelVTEntegrasyonProjesi
{
    partial class btnExceldenOku
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
            button1 = new Button();
            richTextBox1 = new RichTextBox();
            richTextBox2 = new RichTextBox();
            ExceldenOku = new Button();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(134, 12);
            button1.Name = "button1";
            button1.Size = new Size(239, 48);
            button1.TabIndex = 0;
            button1.Text = "Veri Tabanından Oku Excel'e Yaz";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(134, 66);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(473, 120);
            richTextBox1.TabIndex = 1;
            richTextBox1.Text = "";
            // 
            // richTextBox2
            // 
            richTextBox2.Location = new Point(134, 271);
            richTextBox2.Name = "richTextBox2";
            richTextBox2.Size = new Size(473, 120);
            richTextBox2.TabIndex = 2;
            richTextBox2.Text = "";
            // 
            // ExceldenOku
            // 
            ExceldenOku.Location = new Point(134, 205);
            ExceldenOku.Name = "ExceldenOku";
            ExceldenOku.Size = new Size(239, 60);
            ExceldenOku.TabIndex = 3;
            ExceldenOku.Text = " Excel'den oku Veri Tabanında Yaz";
            ExceldenOku.UseVisualStyleBackColor = true;
            ExceldenOku.Click += button2_Click;
            // 
            // btnExceldenOku
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.ControlDarkDark;
            ClientSize = new Size(800, 450);
            Controls.Add(ExceldenOku);
            Controls.Add(richTextBox2);
            Controls.Add(richTextBox1);
            Controls.Add(button1);
            Name = "btnExceldenOku";
            Text = "Veri Tabanı Excel Entegrasyon";
            ResumeLayout(false);
        }

        #endregion

        private Button button1;
        private RichTextBox richTextBox1;
        private RichTextBox richTextBox2;
        private Button ExceldenOku;
    }
}
