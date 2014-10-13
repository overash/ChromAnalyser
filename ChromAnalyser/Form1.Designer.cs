namespace WindowsFormsApplication1
{
    partial class Form1
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button5 = new System.Windows.Forms.Button();
            this.loadedChroms = new System.Windows.Forms.CheckedListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(31, 269);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(105, 23);
            this.button5.TabIndex = 13;
            this.button5.Text = "Выбрать папку";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // loadedChroms
            // 
            this.loadedChroms.CheckOnClick = true;
            this.loadedChroms.FormattingEnabled = true;
            this.loadedChroms.Location = new System.Drawing.Point(12, 12);
            this.loadedChroms.Name = "loadedChroms";
            this.loadedChroms.Size = new System.Drawing.Size(261, 244);
            this.loadedChroms.TabIndex = 15;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(142, 269);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(117, 23);
            this.button1.TabIndex = 16;
            this.button1.Text = "Построить отчет";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(287, 308);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.loadedChroms);
            this.Controls.Add(this.button5);
            this.Name = "Form1";
            this.Text = "ChromAnalyser";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.CheckedListBox loadedChroms;
        private System.Windows.Forms.Button button1;
    }
}

