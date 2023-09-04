namespace ProjectClosureToolWinFormsNET6
{
    partial class Form1
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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            label1 = new Label();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            textBox1 = new TextBox();
            textBox2 = new TextBox();
            textBox3 = new TextBox();
            textBox4 = new TextBox();
            button1 = new Button();
            button2 = new Button();
            dataGridView1 = new DataGridView();
            button3 = new Button();
            button4 = new Button();
            button6 = new Button();
            button7 = new Button();
            button8 = new Button();
            button9 = new Button();
            button10 = new Button();
            button11 = new Button();
            dataGridView2 = new DataGridView();
            button5 = new Button();
            button12 = new Button();
            dataGridView3 = new DataGridView();
            dataGridView4 = new DataGridView();
            dataGridView5 = new DataGridView();
            button13 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView5).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(224, 15);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(96, 25);
            label1.TabIndex = 0;
            label1.Text = "board URL";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(222, 55);
            label2.Margin = new Padding(2, 0, 2, 0);
            label2.Name = "label2";
            label2.Size = new Size(104, 25);
            label2.TabIndex = 1;
            label2.Text = "board code";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(720, 11);
            label3.Margin = new Padding(2, 0, 2, 0);
            label3.Name = "label3";
            label3.Size = new Size(67, 25);
            label3.TabIndex = 2;
            label3.Text = "APIKey";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(720, 58);
            label4.Margin = new Padding(2, 0, 2, 0);
            label4.Name = "label4";
            label4.Size = new Size(124, 25);
            label4.TabIndex = 3;
            label4.Text = "MyTrelloToken";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(356, 6);
            textBox1.Margin = new Padding(2, 4, 2, 4);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(235, 31);
            textBox1.TabIndex = 4;
            textBox1.Text = "https://trello.com/1/boards/";
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(356, 51);
            textBox2.Margin = new Padding(2, 4, 2, 4);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(98, 31);
            textBox2.TabIndex = 5;
            textBox2.Text = "dXURQTbH";
            textBox2.TextChanged += textBox2_TextChanged;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(854, 6);
            textBox3.Margin = new Padding(2, 4, 2, 4);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(150, 31);
            textBox3.TabIndex = 6;
            textBox3.Text = "Default";
            textBox3.TextChanged += textBox3_TextChanged;
            // 
            // textBox4
            // 
            textBox4.Location = new Point(854, 51);
            textBox4.Margin = new Padding(2, 4, 2, 4);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(150, 31);
            textBox4.TabIndex = 7;
            textBox4.Text = "Default";
            textBox4.TextChanged += textBox4_TextChanged;
            // 
            // button1
            // 
            button1.Location = new Point(11, 11);
            button1.Margin = new Padding(2, 4, 2, 4);
            button1.Name = "button1";
            button1.Size = new Size(135, 43);
            button1.TabIndex = 8;
            button1.Text = "Обработка";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(11, 448);
            button2.Margin = new Padding(2, 4, 2, 4);
            button2.Name = "button2";
            button2.Size = new Size(465, 39);
            button2.TabIndex = 10;
            button2.Text = "Список ярлыков";
            button2.UseVisualStyleBackColor = true;
            button2.Visible = false;
            button2.Click += button2_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = SystemColors.Window;
            dataGridViewCellStyle1.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle1.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            dataGridView1.DefaultCellStyle = dataGridViewCellStyle1;
            dataGridView1.Location = new Point(548, 228);
            dataGridView1.Margin = new Padding(4);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 29;
            dataGridView1.Size = new Size(471, 403);
            dataGridView1.TabIndex = 11;
            dataGridView1.Visible = false;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            // 
            // button3
            // 
            button3.Location = new Point(8, 494);
            button3.Margin = new Padding(2, 4, 2, 4);
            button3.Name = "button3";
            button3.Size = new Size(468, 26);
            button3.TabIndex = 12;
            button3.Text = "Список блоков";
            button3.UseVisualStyleBackColor = true;
            button3.Visible = false;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(8, 541);
            button4.Margin = new Padding(2, 4, 2, 4);
            button4.Name = "button4";
            button4.Size = new Size(468, 26);
            button4.TabIndex = 13;
            button4.Text = "Список комбинаций ярлыков";
            button4.UseVisualStyleBackColor = true;
            button4.Visible = false;
            button4.Click += button4_Click;
            // 
            // button6
            // 
            button6.Location = new Point(8, 789);
            button6.Margin = new Padding(2, 4, 2, 4);
            button6.Name = "button6";
            button6.Size = new Size(468, 39);
            button6.TabIndex = 15;
            button6.Text = "Просмотр списка игнорируемых ярлыков";
            button6.UseVisualStyleBackColor = true;
            button6.Visible = false;
            button6.Click += button6_Click;
            // 
            // button7
            // 
            button7.Location = new Point(8, 654);
            button7.Margin = new Padding(2, 4, 2, 4);
            button7.Name = "button7";
            button7.Size = new Size(468, 39);
            button7.TabIndex = 16;
            button7.Text = "Очистка списка игнорируемых ярлыков";
            button7.UseVisualStyleBackColor = true;
            button7.Visible = false;
            button7.Click += button7_Click;
            // 
            // button8
            // 
            button8.Location = new Point(8, 588);
            button8.Margin = new Padding(2, 4, 2, 4);
            button8.Name = "button8";
            button8.Size = new Size(468, 56);
            button8.TabIndex = 17;
            button8.Text = "Список комбинаций ярлыков (игнорируемые не учитываются)";
            button8.UseVisualStyleBackColor = true;
            button8.Visible = false;
            button8.Click += button8_Click;
            // 
            // button9
            // 
            button9.Location = new Point(11, 836);
            button9.Margin = new Padding(2, 4, 2, 4);
            button9.Name = "button9";
            button9.RightToLeft = RightToLeft.Yes;
            button9.Size = new Size(468, 39);
            button9.TabIndex = 18;
            button9.Text = "Удаление ярлыков из списка игнорируемых";
            button9.UseVisualStyleBackColor = true;
            button9.Visible = false;
            button9.Click += button9_Click;
            // 
            // button10
            // 
            button10.Location = new Point(11, 199);
            button10.Margin = new Padding(2, 4, 2, 4);
            button10.Name = "button10";
            button10.Size = new Size(468, 74);
            button10.TabIndex = 19;
            button10.Text = "Суммарные оценки для выбранных блоков для всех комбинаций ярлыков";
            button10.UseVisualStyleBackColor = true;
            button10.Visible = false;
            button10.Click += button10_Click;
            // 
            // button11
            // 
            button11.Location = new Point(11, 302);
            button11.Margin = new Padding(2, 4, 2, 4);
            button11.Name = "button11";
            button11.Size = new Size(468, 88);
            button11.TabIndex = 20;
            button11.Text = "Суммарные оценки для выбранных блоков для выбранных комбинаций ярлыков";
            button11.UseVisualStyleBackColor = true;
            button11.Visible = false;
            button11.Click += button11_Click;
            // 
            // dataGridView2
            // 
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;
            dataGridView2.AllowUserToResizeColumns = false;
            dataGridView2.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView2.Location = new Point(8, 102);
            dataGridView2.Margin = new Padding(4);
            dataGridView2.Name = "dataGridView2";
            dataGridView2.RowHeadersVisible = false;
            dataGridView2.RowHeadersWidth = 51;
            dataGridView2.RowTemplate.Height = 29;
            dataGridView2.Size = new Size(372, 529);
            dataGridView2.TabIndex = 21;
            dataGridView2.CellContentClick += dataGridView2_CellContentClick;
            // 
            // button5
            // 
            button5.Location = new Point(11, 742);
            button5.Margin = new Padding(2, 4, 2, 4);
            button5.Name = "button5";
            button5.Size = new Size(468, 39);
            button5.TabIndex = 14;
            button5.Text = "Добавление ярлыков к списку игнорируемых";
            button5.UseVisualStyleBackColor = true;
            button5.Visible = false;
            button5.Click += button5_Click;
            // 
            // button12
            // 
            button12.Location = new Point(8, 701);
            button12.Margin = new Padding(2, 4, 2, 4);
            button12.Name = "button12";
            button12.Size = new Size(465, 39);
            button12.TabIndex = 22;
            button12.Text = "Сохранение списка игнорируемых ярлыков";
            button12.UseVisualStyleBackColor = true;
            button12.Visible = false;
            button12.Click += button12_Click;
            // 
            // dataGridView3
            // 
            dataGridView3.AllowUserToAddRows = false;
            dataGridView3.AllowUserToDeleteRows = false;
            dataGridView3.AllowUserToResizeColumns = false;
            dataGridView3.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView3.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView3.Location = new Point(388, 102);
            dataGridView3.Margin = new Padding(4);
            dataGridView3.Name = "dataGridView3";
            dataGridView3.RowHeadersVisible = false;
            dataGridView3.RowHeadersWidth = 51;
            dataGridView3.RowTemplate.Height = 29;
            dataGridView3.Size = new Size(349, 529);
            dataGridView3.TabIndex = 23;
            dataGridView3.CellContentClick += dataGridView3_CellContentClick;
            // 
            // dataGridView4
            // 
            dataGridView4.AllowUserToAddRows = false;
            dataGridView4.AllowUserToDeleteRows = false;
            dataGridView4.AllowUserToResizeColumns = false;
            dataGridView4.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView4.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dataGridView4.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = SystemColors.Window;
            dataGridViewCellStyle2.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle2.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.True;
            dataGridView4.DefaultCellStyle = dataGridViewCellStyle2;
            dataGridView4.Location = new Point(744, 102);
            dataGridView4.Margin = new Padding(4);
            dataGridView4.Name = "dataGridView4";
            dataGridView4.RowHeadersVisible = false;
            dataGridView4.RowHeadersWidth = 51;
            dataGridView4.RowTemplate.Height = 29;
            dataGridView4.Size = new Size(350, 529);
            dataGridView4.TabIndex = 24;
            dataGridView4.CellContentClick += dataGridView4_CellContentClick;
            // 
            // dataGridView5
            // 
            dataGridView5.AllowUserToAddRows = false;
            dataGridView5.AllowUserToDeleteRows = false;
            dataGridView5.AllowUserToResizeColumns = false;
            dataGridView5.AllowUserToResizeRows = false;
            dataGridView5.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView5.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders;
            dataGridView5.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = SystemColors.Window;
            dataGridViewCellStyle3.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            dataGridViewCellStyle3.ForeColor = SystemColors.ControlText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.True;
            dataGridView5.DefaultCellStyle = dataGridViewCellStyle3;
            dataGridView5.Location = new Point(1102, 102);
            dataGridView5.Margin = new Padding(4);
            dataGridView5.Name = "dataGridView5";
            dataGridView5.RowHeadersVisible = false;
            dataGridView5.RowHeadersWidth = 51;
            dataGridView5.RowTemplate.Height = 29;
            dataGridView5.Size = new Size(171, 529);
            dataGridView5.TabIndex = 25;
            dataGridView5.CellContentClick += dataGridView5_CellContentClick;
            // 
            // button13
            // 
            button13.Location = new Point(1125, 11);
            button13.Margin = new Padding(2, 4, 2, 4);
            button13.Name = "button13";
            button13.Size = new Size(148, 43);
            button13.TabIndex = 26;
            button13.Text = "Экспорт в Excel";
            button13.UseVisualStyleBackColor = true;
            button13.Click += button13_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1286, 646);
            Controls.Add(button13);
            Controls.Add(dataGridView5);
            Controls.Add(dataGridView4);
            Controls.Add(button1);
            Controls.Add(textBox4);
            Controls.Add(textBox3);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(dataGridView2);
            Controls.Add(dataGridView3);
            Controls.Add(dataGridView1);
            Controls.Add(button12);
            Controls.Add(button11);
            Controls.Add(button10);
            Controls.Add(button9);
            Controls.Add(button8);
            Controls.Add(button7);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(button2);
            Margin = new Padding(2, 4, 2, 4);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView2).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView3).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView4).EndInit();
            ((System.ComponentModel.ISupportInitialize)dataGridView5).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private TextBox textBox1;
        private TextBox textBox2;
        private TextBox textBox3;
        private TextBox textBox4;
        private Button button1;
        private Button button2;
        private DataGridView dataGridView1;
        private Button button3;
        private Button button4;
        private Button button6;
        private Button button7;
        private Button button8;
        private Button button9;
        private Button button10;
        private Button button11;
        private DataGridView dataGridView2;
        private Button button5;
        private Button button12;
        private DataGridView dataGridView3;
        private DataGridView dataGridView4;
        private DataGridView dataGridView5;
        private Button button13;
    }
}