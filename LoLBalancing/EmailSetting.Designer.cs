namespace LoLBalancing
{
	partial class EmailSetting
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EmailSetting));
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.comboBox_Client = new System.Windows.Forms.ComboBox();
			this.textBox_EmailName = new System.Windows.Forms.TextBox();
			this.textBox_PW = new System.Windows.Forms.TextBox();
			this.textBox_Subject = new System.Windows.Forms.TextBox();
			this.richTextBox_Body = new System.Windows.Forms.RichTextBox();
			this.button_OK = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(12, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(95, 19);
			this.label1.TabIndex = 0;
			this.label1.Text = "Client:";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(12, 37);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(95, 19);
			this.label2.TabIndex = 1;
			this.label2.Text = "Email Name:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(12, 93);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(95, 19);
			this.label3.TabIndex = 3;
			this.label3.Text = "Email Subject:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(12, 65);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(95, 19);
			this.label4.TabIndex = 2;
			this.label4.Text = "Password:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(12, 121);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(95, 19);
			this.label5.TabIndex = 4;
			this.label5.Text = "Email Body:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// comboBox_Client
			// 
			this.comboBox_Client.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox_Client.FormattingEnabled = true;
			this.comboBox_Client.Items.AddRange(new object[] {
            "Gmail"});
			this.comboBox_Client.Location = new System.Drawing.Point(113, 8);
			this.comboBox_Client.Name = "comboBox_Client";
			this.comboBox_Client.Size = new System.Drawing.Size(252, 23);
			this.comboBox_Client.TabIndex = 5;
			// 
			// textBox_EmailName
			// 
			this.textBox_EmailName.Location = new System.Drawing.Point(113, 38);
			this.textBox_EmailName.Name = "textBox_EmailName";
			this.textBox_EmailName.Size = new System.Drawing.Size(252, 21);
			this.textBox_EmailName.TabIndex = 6;
			// 
			// textBox_PW
			// 
			this.textBox_PW.Location = new System.Drawing.Point(113, 65);
			this.textBox_PW.Name = "textBox_PW";
			this.textBox_PW.Size = new System.Drawing.Size(252, 21);
			this.textBox_PW.TabIndex = 7;
			// 
			// textBox_Subject
			// 
			this.textBox_Subject.Location = new System.Drawing.Point(113, 94);
			this.textBox_Subject.Name = "textBox_Subject";
			this.textBox_Subject.Size = new System.Drawing.Size(252, 21);
			this.textBox_Subject.TabIndex = 8;
			// 
			// richTextBox_Body
			// 
			this.richTextBox_Body.Location = new System.Drawing.Point(113, 121);
			this.richTextBox_Body.Name = "richTextBox_Body";
			this.richTextBox_Body.Size = new System.Drawing.Size(252, 179);
			this.richTextBox_Body.TabIndex = 9;
			this.richTextBox_Body.Text = "";
			// 
			// button_OK
			// 
			this.button_OK.Location = new System.Drawing.Point(290, 306);
			this.button_OK.Name = "button_OK";
			this.button_OK.Size = new System.Drawing.Size(75, 23);
			this.button_OK.TabIndex = 10;
			this.button_OK.Text = "Save";
			this.button_OK.UseVisualStyleBackColor = true;
			this.button_OK.Click += new System.EventHandler(this.button_OK_Click);
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(21, 306);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(263, 23);
			this.label6.TabIndex = 11;
			this.label6.Text = "Using \'#\' will signify the proper Team Number";
			this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// EmailSetting
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(375, 340);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.button_OK);
			this.Controls.Add(this.richTextBox_Body);
			this.Controls.Add(this.textBox_Subject);
			this.Controls.Add(this.textBox_PW);
			this.Controls.Add(this.textBox_EmailName);
			this.Controls.Add(this.comboBox_Client);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.Name = "EmailSetting";
			this.Text = "Email Set-Up";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ComboBox comboBox_Client;
		private System.Windows.Forms.TextBox textBox_EmailName;
		private System.Windows.Forms.TextBox textBox_PW;
		private System.Windows.Forms.TextBox textBox_Subject;
		private System.Windows.Forms.RichTextBox richTextBox_Body;
		private System.Windows.Forms.Button button_OK;
		private System.Windows.Forms.Label label6;
	}
}