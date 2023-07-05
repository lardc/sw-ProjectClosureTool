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
            button5 = new Button();
            button6 = new Button();
            button7 = new Button();
            button8 = new Button();
            button9 = new Button();
            button10 = new Button();
            button11 = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(11, 15);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(96, 25);
            label1.TabIndex = 0;
            label1.Text = "board URL";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(11, 56);
            label2.Margin = new Padding(2, 0, 2, 0);
            label2.Name = "label2";
            label2.Size = new Size(104, 25);
            label2.TabIndex = 1;
            label2.Text = "board code";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(11, 106);
            label3.Margin = new Padding(2, 0, 2, 0);
            label3.Name = "label3";
            label3.Size = new Size(67, 25);
            label3.TabIndex = 2;
            label3.Text = "APIKey";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(11, 156);
            label4.Margin = new Padding(2, 0, 2, 0);
            label4.Name = "label4";
            label4.Size = new Size(124, 25);
            label4.TabIndex = 3;
            label4.Text = "MyTrelloToken";
            // 
            // textBox1
            // 
            textBox1.Location = new Point(145, 17);
            textBox1.Margin = new Padding(2, 4, 2, 4);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(235, 31);
            textBox1.TabIndex = 4;
            textBox1.Text = "https://trello.com/1/boards/";
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // textBox2
            // 
            textBox2.Location = new Point(145, 56);
            textBox2.Margin = new Padding(2, 4, 2, 4);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(98, 31);
            textBox2.TabIndex = 5;
            textBox2.Text = "dXURQTbH";
            textBox2.TextChanged += textBox2_TextChanged;
            // 
            // textBox3
            // 
            textBox3.Location = new Point(145, 101);
            textBox3.Margin = new Padding(2, 4, 2, 4);
            textBox3.Name = "textBox3";
            textBox3.Size = new Size(150, 31);
            textBox3.TabIndex = 6;
            textBox3.Text = "Default";
            textBox3.TextChanged += textBox3_TextChanged;
            // 
            // textBox4
            // 
            textBox4.Location = new Point(145, 150);
            textBox4.Margin = new Padding(2, 4, 2, 4);
            textBox4.Name = "textBox4";
            textBox4.Size = new Size(150, 31);
            textBox4.TabIndex = 7;
            textBox4.Text = "Default";
            textBox4.TextChanged += textBox4_TextChanged;
            // 
            // button1
            // 
            button1.Location = new Point(325, 74);
            button1.Margin = new Padding(2, 4, 2, 4);
            button1.Name = "button1";
            button1.Size = new Size(128, 88);
            button1.TabIndex = 8;
            button1.Text = "Обработка";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Location = new Point(11, 211);
            button2.Margin = new Padding(2, 4, 2, 4);
            button2.Name = "button2";
            button2.Size = new Size(172, 39);
            button2.TabIndex = 10;
            button2.Text = "Список ярлыков";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(710, 15);
            dataGridView1.Margin = new Padding(4);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowHeadersWidth = 51;
            dataGridView1.RowTemplate.Height = 29;
            dataGridView1.Size = new Size(976, 836);
            dataGridView1.TabIndex = 11;
            // 
            // button3
            // 
            button3.Location = new Point(11, 258);
            button3.Margin = new Padding(2, 4, 2, 4);
            button3.Name = "button3";
            button3.Size = new Size(172, 39);
            button3.TabIndex = 12;
            button3.Text = "Список блоков";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button4
            // 
            button4.Location = new Point(11, 493);
            button4.Margin = new Padding(2, 4, 2, 4);
            button4.Name = "button4";
            button4.Size = new Size(279, 39);
            button4.TabIndex = 13;
            button4.Text = "Список комбинаций ярлыков";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // button5
            // 
            button5.Location = new Point(11, 305);
            button5.Margin = new Padding(2, 4, 2, 4);
            button5.Name = "button5";
            button5.Size = new Size(403, 39);
            button5.TabIndex = 14;
            button5.Text = "Добавление ярлыков к списку игнорируемых";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // button6
            // 
            button6.Location = new Point(11, 352);
            button6.Margin = new Padding(2, 4, 2, 4);
            button6.Name = "button6";
            button6.Size = new Size(371, 39);
            button6.TabIndex = 15;
            button6.Text = "Просмотр списка игнорируемых ярлыков";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // button7
            // 
            button7.Location = new Point(11, 446);
            button7.Margin = new Padding(2, 4, 2, 4);
            button7.Name = "button7";
            button7.Size = new Size(371, 39);
            button7.TabIndex = 16;
            button7.Text = "Очистка списка игнорируемых ярлыков";
            button7.UseVisualStyleBackColor = true;
            button7.Click += button7_Click;
            // 
            // button8
            // 
            button8.Location = new Point(11, 540);
            button8.Margin = new Padding(2, 4, 2, 4);
            button8.Name = "button8";
            button8.Size = new Size(541, 39);
            button8.TabIndex = 17;
            button8.Text = "Список комбинаций ярлыков (игнорируемые не учитываются)";
            button8.UseVisualStyleBackColor = true;
            button8.Click += button8_Click;
            // 
            // button9
            // 
            button9.Location = new Point(11, 399);
            button9.Margin = new Padding(2, 4, 2, 4);
            button9.Name = "button9";
            button9.RightToLeft = RightToLeft.Yes;
            button9.Size = new Size(384, 39);
            button9.TabIndex = 18;
            button9.Text = "Удаление ярлыков из списка игнорируемых";
            button9.UseVisualStyleBackColor = true;
            button9.Click += button9_Click;
            // 
            // button10
            // 
            button10.Location = new Point(11, 587);
            button10.Margin = new Padding(2, 4, 2, 4);
            button10.Name = "button10";
            button10.Size = new Size(627, 39);
            button10.TabIndex = 19;
            button10.Text = "Суммарные оценки для конкретнго блока для всех комбинаций ярлыков";
            button10.UseVisualStyleBackColor = true;
            button10.Click += button10_Click;
            // 
            // button11
            // 
            button11.Location = new Point(11, 634);
            button11.Margin = new Padding(2, 4, 2, 4);
            button11.Name = "button11";
            button11.Size = new Size(692, 39);
            button11.TabIndex = 20;
            button11.Text = "Суммарные оценки для конкретнго блока для конкретной комбинации ярлыков";
            button11.UseVisualStyleBackColor = true;
            button11.Click += button11_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1699, 866);
            Controls.Add(button11);
            Controls.Add(button10);
            Controls.Add(button9);
            Controls.Add(button8);
            Controls.Add(button7);
            Controls.Add(button6);
            Controls.Add(button5);
            Controls.Add(button4);
            Controls.Add(button3);
            Controls.Add(dataGridView1);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(textBox4);
            Controls.Add(textBox3);
            Controls.Add(textBox2);
            Controls.Add(textBox1);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(label1);
            Margin = new Padding(2, 4, 2, 4);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
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
        private Button button5;
        private Button button6;
        private Button button7;
        private Button button8;
        private Button button9;
        private Button button10;
        private Button button11;
    }
}