using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LoLBalancing
{
	public partial class EmailSetting : Form
	{
		private bool button_pressed = false;

		public EmailSetting() {
			InitializeComponent();
		}

		public void Dialog_Init() {
			comboBox_Client.SelectedIndex = 0;
			textBox_EmailName.Text = MainForm.emailName;
			textBox_PW.Text = MainForm.emailPass;
			textBox_Subject.Text = MainForm.emailSubject;
			richTextBox_Body.Text = MainForm.emailBody;
			this.ShowDialog();
			if (button_pressed) {
				MainForm.emailName = textBox_EmailName.Text;
				MainForm.emailPass = textBox_PW.Text;
				MainForm.emailSubject = textBox_Subject.Text;
				MainForm.emailBody = richTextBox_Body.Text;
			}
		}

		private void button_OK_Click(object sender, EventArgs e) {
			button_pressed = true;
			this.Close();
		}
	}
}
