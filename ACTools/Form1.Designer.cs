namespace AcceptionTools
{
    partial class ACTools
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ProjNumber = new System.Windows.Forms.Label();
            this.buildNameLabel = new System.Windows.Forms.Label();
            this.excelLabel = new System.Windows.Forms.Label();
            this.ExcelPathBox = new System.Windows.Forms.TextBox();
            this.excelInputButton = new System.Windows.Forms.Button();
            this.generateButton = new System.Windows.Forms.Button();
            this.buildNameBox = new System.Windows.Forms.ComboBox();
            this.wordInputButton = new System.Windows.Forms.Button();
            this.WordPathBox = new System.Windows.Forms.TextBox();
            this.wordLabel = new System.Windows.Forms.Label();
            this.writeButton = new System.Windows.Forms.Button();
            this.checkButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // ProjNumber
            // 
            this.ProjNumber.AutoSize = true;
            this.ProjNumber.Location = new System.Drawing.Point(272, 104);
            this.ProjNumber.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.ProjNumber.Name = "ProjNumber";
            this.ProjNumber.Size = new System.Drawing.Size(120, 15);
            this.ProjNumber.TabIndex = 0;
            this.ProjNumber.Text = "放线/验收案号：";
            this.ProjNumber.UseWaitCursor = true;
            // 
            // buildNameLabel
            // 
            this.buildNameLabel.AutoSize = true;
            this.buildNameLabel.Location = new System.Drawing.Point(272, 144);
            this.buildNameLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.buildNameLabel.Name = "buildNameLabel";
            this.buildNameLabel.Size = new System.Drawing.Size(112, 15);
            this.buildNameLabel.TabIndex = 1;
            this.buildNameLabel.Text = "建设项目名称：";
            // 
            // excelLabel
            // 
            this.excelLabel.AutoSize = true;
            this.excelLabel.Location = new System.Drawing.Point(16, 22);
            this.excelLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.excelLabel.Name = "excelLabel";
            this.excelLabel.Size = new System.Drawing.Size(152, 15);
            this.excelLabel.TabIndex = 3;
            this.excelLabel.Text = "面积汇总表（.xlsx）";
            this.excelLabel.UseWaitCursor = true;
            // 
            // ExcelPathBox
            // 
            this.ExcelPathBox.Location = new System.Drawing.Point(179, 15);
            this.ExcelPathBox.Margin = new System.Windows.Forms.Padding(4);
            this.ExcelPathBox.Name = "ExcelPathBox";
            this.ExcelPathBox.Size = new System.Drawing.Size(280, 25);
            this.ExcelPathBox.TabIndex = 5;
            // 
            // excelInputButton
            // 
            this.excelInputButton.Location = new System.Drawing.Point(483, 13);
            this.excelInputButton.Margin = new System.Windows.Forms.Padding(4);
            this.excelInputButton.Name = "excelInputButton";
            this.excelInputButton.Size = new System.Drawing.Size(80, 30);
            this.excelInputButton.TabIndex = 6;
            this.excelInputButton.Text = "导入";
            this.excelInputButton.UseVisualStyleBackColor = true;
            this.excelInputButton.Click += new System.EventHandler(this.excelInputButton_Click);
            // 
            // generateButton
            // 
            this.generateButton.Location = new System.Drawing.Point(51, 282);
            this.generateButton.Margin = new System.Windows.Forms.Padding(4);
            this.generateButton.Name = "generateButton";
            this.generateButton.Size = new System.Drawing.Size(117, 30);
            this.generateButton.TabIndex = 7;
            this.generateButton.Text = "生成核实概况";
            this.generateButton.UseVisualStyleBackColor = true;
            this.generateButton.Click += new System.EventHandler(this.generateButton_Click);
            // 
            // buildNameBox
            // 
            this.buildNameBox.FormattingEnabled = true;
            this.buildNameBox.Location = new System.Drawing.Point(392, 139);
            this.buildNameBox.Margin = new System.Windows.Forms.Padding(4);
            this.buildNameBox.Name = "buildNameBox";
            this.buildNameBox.Size = new System.Drawing.Size(171, 23);
            this.buildNameBox.TabIndex = 10;
            // 
            // wordInputButton
            // 
            this.wordInputButton.Location = new System.Drawing.Point(483, 53);
            this.wordInputButton.Margin = new System.Windows.Forms.Padding(4);
            this.wordInputButton.Name = "wordInputButton";
            this.wordInputButton.Size = new System.Drawing.Size(80, 31);
            this.wordInputButton.TabIndex = 13;
            this.wordInputButton.Text = "导入";
            this.wordInputButton.UseVisualStyleBackColor = true;
            this.wordInputButton.Click += new System.EventHandler(this.wordInputButton_Click);
            // 
            // WordPathBox
            // 
            this.WordPathBox.Location = new System.Drawing.Point(179, 53);
            this.WordPathBox.Margin = new System.Windows.Forms.Padding(4);
            this.WordPathBox.Name = "WordPathBox";
            this.WordPathBox.Size = new System.Drawing.Size(280, 25);
            this.WordPathBox.TabIndex = 12;
            // 
            // wordLabel
            // 
            this.wordLabel.AutoSize = true;
            this.wordLabel.Location = new System.Drawing.Point(16, 60);
            this.wordLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.wordLabel.Name = "wordLabel";
            this.wordLabel.Size = new System.Drawing.Size(152, 15);
            this.wordLabel.TabIndex = 11;
            this.wordLabel.Text = "核实槪况表（.docx）";
            this.wordLabel.UseWaitCursor = true;
            // 
            // writeButton
            // 
            this.writeButton.Location = new System.Drawing.Point(254, 282);
            this.writeButton.Margin = new System.Windows.Forms.Padding(4);
            this.writeButton.Name = "writeButton";
            this.writeButton.Size = new System.Drawing.Size(100, 30);
            this.writeButton.TabIndex = 14;
            this.writeButton.Text = "选择项目名称";
            this.writeButton.UseVisualStyleBackColor = true;
            this.writeButton.Click += new System.EventHandler(this.writeButton_Click);
            // 
            // checkButton
            // 
            this.checkButton.Location = new System.Drawing.Point(448, 282);
            this.checkButton.Margin = new System.Windows.Forms.Padding(4);
            this.checkButton.Name = "checkButton";
            this.checkButton.Size = new System.Drawing.Size(100, 30);
            this.checkButton.TabIndex = 15;
            this.checkButton.Text = "数据检核";
            this.checkButton.UseVisualStyleBackColor = true;
            this.checkButton.Click += new System.EventHandler(this.checkButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 104);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 15);
            this.label1.TabIndex = 16;
            this.label1.Text = "建设单位:";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(97, 96);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(167, 25);
            this.textBox1.TabIndex = 17;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 143);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 15);
            this.label2.TabIndex = 18;
            this.label2.Text = "设计单位:";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(97, 181);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(167, 25);
            this.textBox6.TabIndex = 25;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 186);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(75, 15);
            this.label6.TabIndex = 24;
            this.label6.Text = "施工单位:";
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(187, 229);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(149, 25);
            this.textBox8.TabIndex = 29;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 236);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(165, 15);
            this.label8.TabIndex = 28;
            this.label8.Text = "建设工程规划许可证号:";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(97, 137);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(167, 25);
            this.textBox2.TabIndex = 19;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(352, 236);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(90, 15);
            this.label11.TabIndex = 36;
            this.label11.Text = "相关批文号:";
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(448, 229);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(115, 25);
            this.textBox11.TabIndex = 37;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(392, 96);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(170, 25);
            this.textBox3.TabIndex = 38;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(272, 186);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 15);
            this.label3.TabIndex = 39;
            this.label3.Text = "建设位置：";
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(360, 181);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(202, 25);
            this.textBox4.TabIndex = 40;
            // 
            // ACTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 325);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox11);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.textBox8);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkButton);
            this.Controls.Add(this.writeButton);
            this.Controls.Add(this.wordInputButton);
            this.Controls.Add(this.WordPathBox);
            this.Controls.Add(this.wordLabel);
            this.Controls.Add(this.buildNameBox);
            this.Controls.Add(this.excelInputButton);
            this.Controls.Add(this.ExcelPathBox);
            this.Controls.Add(this.excelLabel);
            this.Controls.Add(this.buildNameLabel);
            this.Controls.Add(this.ProjNumber);
            this.Controls.Add(this.generateButton);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ACTools";
            this.Text = "规划条件核实数据处理工具";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label ProjNumber;
        private System.Windows.Forms.Label buildNameLabel;
        private System.Windows.Forms.Label excelLabel;
        private System.Windows.Forms.TextBox ExcelPathBox;
        private System.Windows.Forms.Button excelInputButton;
        private System.Windows.Forms.Button generateButton;
        private System.Windows.Forms.ComboBox buildNameBox;
        private System.Windows.Forms.Button wordInputButton;
        private System.Windows.Forms.TextBox WordPathBox;
        private System.Windows.Forms.Label wordLabel;
        private System.Windows.Forms.Button writeButton;
        private System.Windows.Forms.Button checkButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox4;
    }
}

