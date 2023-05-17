using System.Drawing;

namespace PrintWord
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
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
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtPath = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.combPrintType = new System.Windows.Forms.ComboBox();
            this.btBrowse = new System.Windows.Forms.Button();
            this.btConvert = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Path to html:";
            // 
            // txtPath
            // 
            this.txtPath.Location = new System.Drawing.Point(9, 23);
            this.txtPath.Name = "txtPath";
            this.txtPath.Size = new System.Drawing.Size(285, 22);
            this.txtPath.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Type converter:";
            // 
            // combPrintType
            // 
            this.combPrintType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.combPrintType.FormattingEnabled = true;
            this.combPrintType.Location = new System.Drawing.Point(9, 63);
            this.combPrintType.Name = "combPrintType";
            this.combPrintType.Size = new System.Drawing.Size(121, 21);
            this.combPrintType.TabIndex = 3;
            // 
            // btBrowse
            // 
            this.btBrowse.BackColor = System.Drawing.SystemColors.Control;
            this.btBrowse.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btBrowse.Location = new System.Drawing.Point(259, 63);
            this.btBrowse.Name = "btBrowse";
            this.btBrowse.Size = new System.Drawing.Size(35, 21);
            this.btBrowse.TabIndex = 4;
            this.btBrowse.Text = "...";
            this.btBrowse.UseVisualStyleBackColor = false;
            this.btBrowse.Click += new System.EventHandler(this.BtBrowse_Click);
            // 
            // btConvert
            // 
            this.btConvert.BackColor = System.Drawing.SystemColors.Control;
            this.btConvert.FlatAppearance.BorderColor = System.Drawing.Color.Silver;
            this.btConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btConvert.Location = new System.Drawing.Point(146, 63);
            this.btConvert.Name = "btConvert";
            this.btConvert.Size = new System.Drawing.Size(107, 21);
            this.btConvert.TabIndex = 5;
            this.btConvert.Text = "Convert";
            this.btConvert.UseVisualStyleBackColor = false;
            this.btConvert.Click += new System.EventHandler(this.BtConvert_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(303, 92);
            this.Controls.Add(this.btConvert);
            this.Controls.Add(this.btBrowse);
            this.Controls.Add(this.combPrintType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtPath);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PrintWord";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtPath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox combPrintType;
        private System.Windows.Forms.Button btBrowse;
        private System.Windows.Forms.Button btConvert;
    }
}

