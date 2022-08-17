namespace ExcelDataAnalyse
{
    partial class MainForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.textBox_zongchengji = new System.Windows.Forms.TextBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox_kemumanfen = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.textBox_chuqirenshu = new System.Windows.Forms.TextBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.button4 = new System.Windows.Forms.Button();
            this.textBox_renkejiaoshi = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.button6 = new System.Windows.Forms.Button();
            this.textBox_baobiaomoban = new System.Windows.Forms.TextBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.button5 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.textBox_zongchengji);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(519, 68);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "1.总成绩表";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(360, 26);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(143, 21);
            this.button1.TabIndex = 1;
            this.button1.Text = "请选择总成绩表";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // textBox_zongchengji
            // 
            this.textBox_zongchengji.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_zongchengji.Location = new System.Drawing.Point(25, 27);
            this.textBox_zongchengji.Multiline = true;
            this.textBox_zongchengji.Name = "textBox_zongchengji";
            this.textBox_zongchengji.ReadOnly = true;
            this.textBox_zongchengji.Size = new System.Drawing.Size(329, 20);
            this.textBox_zongchengji.TabIndex = 0;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog";
            this.openFileDialog.Filter = "excel文件(*.xls,*.xlsl)|*.xls;*.xlsx";
            this.openFileDialog.RestoreDirectory = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.textBox_kemumanfen);
            this.groupBox2.Location = new System.Drawing.Point(13, 87);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(519, 68);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "2.科目满分表";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(359, 27);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(143, 21);
            this.button2.TabIndex = 2;
            this.button2.Text = "请选择科目满分表";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox_kemumanfen
            // 
            this.textBox_kemumanfen.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_kemumanfen.Location = new System.Drawing.Point(24, 28);
            this.textBox_kemumanfen.Multiline = true;
            this.textBox_kemumanfen.Name = "textBox_kemumanfen";
            this.textBox_kemumanfen.ReadOnly = true;
            this.textBox_kemumanfen.Size = new System.Drawing.Size(329, 20);
            this.textBox_kemumanfen.TabIndex = 2;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button3);
            this.groupBox3.Controls.Add(this.textBox_chuqirenshu);
            this.groupBox3.Location = new System.Drawing.Point(13, 162);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(519, 68);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "3.初期人数表";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(359, 30);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(143, 21);
            this.button3.TabIndex = 3;
            this.button3.Text = "请选择期初人数表";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // textBox_chuqirenshu
            // 
            this.textBox_chuqirenshu.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_chuqirenshu.Location = new System.Drawing.Point(24, 31);
            this.textBox_chuqirenshu.Multiline = true;
            this.textBox_chuqirenshu.Name = "textBox_chuqirenshu";
            this.textBox_chuqirenshu.ReadOnly = true;
            this.textBox_chuqirenshu.Size = new System.Drawing.Size(329, 20);
            this.textBox_chuqirenshu.TabIndex = 3;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.button4);
            this.groupBox4.Controls.Add(this.textBox_renkejiaoshi);
            this.groupBox4.Location = new System.Drawing.Point(13, 237);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(519, 68);
            this.groupBox4.TabIndex = 3;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "4.任课教师表";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(359, 28);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(143, 21);
            this.button4.TabIndex = 4;
            this.button4.Text = "请选择任课教师名单表";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // textBox_renkejiaoshi
            // 
            this.textBox_renkejiaoshi.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_renkejiaoshi.Location = new System.Drawing.Point(24, 29);
            this.textBox_renkejiaoshi.Multiline = true;
            this.textBox_renkejiaoshi.Name = "textBox_renkejiaoshi";
            this.textBox_renkejiaoshi.ReadOnly = true;
            this.textBox_renkejiaoshi.Size = new System.Drawing.Size(329, 20);
            this.textBox_renkejiaoshi.TabIndex = 5;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.groupBox5);
            this.panel1.Location = new System.Drawing.Point(6, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(537, 381);
            this.panel1.TabIndex = 4;
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.button6);
            this.groupBox5.Controls.Add(this.textBox_baobiaomoban);
            this.groupBox5.Location = new System.Drawing.Point(7, 306);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(519, 68);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "5.报表模板表";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(359, 30);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(143, 21);
            this.button6.TabIndex = 6;
            this.button6.Text = "请选择报表模板";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // textBox_baobiaomoban
            // 
            this.textBox_baobiaomoban.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox_baobiaomoban.Location = new System.Drawing.Point(24, 30);
            this.textBox_baobiaomoban.Multiline = true;
            this.textBox_baobiaomoban.Name = "textBox_baobiaomoban";
            this.textBox_baobiaomoban.ReadOnly = true;
            this.textBox_baobiaomoban.Size = new System.Drawing.Size(329, 20);
            this.textBox_baobiaomoban.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.button5);
            this.panel2.Location = new System.Drawing.Point(6, 402);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(537, 73);
            this.panel2.TabIndex = 5;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(31, 21);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(115, 28);
            this.button5.TabIndex = 0;
            this.button5.Text = "生 成 报 表";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(556, 491);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.panel1);
            this.Name = "MainForm";
            this.Text = "XX初级中学成绩分析工具";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox_zongchengji;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox_kemumanfen;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox textBox_chuqirenshu;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox textBox_renkejiaoshi;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.TextBox textBox_baobiaomoban;
    }
}

