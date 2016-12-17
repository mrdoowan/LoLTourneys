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
	public partial class AddPlayer : Form
	{
		public AddPlayer() {
			InitializeComponent();
		}

		private bool button_pressed = false;

		public void AddDialog(ref DataGridView Players) {
			this.ShowDialog();
			if (button_pressed) {
				DataGridViewButtonColumn button = new DataGridViewButtonColumn();
				string roles = "";
				if (checkBox_Top.Checked) { roles += "T"; }
				if (checkBox_Jg.Checked) { roles += "J"; }
				if (checkBox_Mid.Checked) { roles += "M"; }
				if (checkBox_ADC.Checked) { roles += "A"; }
				if (checkBox_Supp.Checked) { roles += "S"; }
				string Duo = "";
				if (!string.IsNullOrWhiteSpace(textBox_Duo.Text)) { Duo = textBox_Duo.Text; }
				else { Duo = "N/A"; }
				// Add into DataGridView
				Players.Rows.Add(button, textBox_Name.Text, textBox_Uniq.Text, textBox_IGN.Text, comboBox_Tier.Text, roles, Duo);
				DataGridViewRow Player = Players.Rows[Players.Rows.Count - 1];
				Player.Cells[0].Value = "X";
				string Tier = comboBox_Tier.Text.Split(' ')[0];
				// Modify colors based on Ranking
				switch (Tier) {
					case "Level":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.LEVELHEX);
						break;
					case "Bronze":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.BRONZEHEX);
						break;
					case "Silver":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.SILVERHEX);
						break;
					case "Gold":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.GOLDHEX);
						break;
					case "Platinum":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.PLATHEX);
						break;
					case "Diamond":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.DIAMONDHEX);
						break;
					case "Master":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.MASTERHEX);
						break;
					case "Challenger":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.CHALLENGERHEX);
						break;
					default:
						break;
				}
				Players.ClearSelection();
			}
		}

		public void EditDialog(ref DataGridView Players) {
			this.Text = "Edit Player";
			button_OK.Text = "Edit";
			// Put it into the Form first.
			textBox_Name.Text = Players.SelectedRows[0].Cells[1].Value.ToString();
			textBox_Uniq.Text = Players.SelectedRows[0].Cells[2].Value.ToString();
			textBox_IGN.Text = Players.SelectedRows[0].Cells[3].Value.ToString();
			comboBox_Tier.Text = Players.SelectedRows[0].Cells[4].Value.ToString();
			string roles_str = Players.SelectedRows[0].Cells[5].Value.ToString();
			char[] roles = roles_str.ToCharArray();
			foreach (char role_char in roles) {
				switch (role_char) {
					case 'T':
						checkBox_Top.Checked = true;
						break;
					case 'J':
						checkBox_Jg.Checked = true;
						break;
					case 'M':
						checkBox_Mid.Checked = true;
						break;
					case 'A':
						checkBox_ADC.Checked = true;
						break;
					case 'S':
						checkBox_Supp.Checked = true;
						break;
					default:
						break;
				}
			}
			textBox_Duo.Text = Players.SelectedRows[0].Cells[6].Value.ToString();
			this.ShowDialog();
			// Copy and paste from above
			if (button_pressed) {
				roles_str = "";
				if (checkBox_Top.Checked) { roles_str += "T"; }
				if (checkBox_Jg.Checked) { roles_str += "J"; }
				if (checkBox_Mid.Checked) { roles_str += "M"; }
				if (checkBox_ADC.Checked) { roles_str += "A"; }
				if (checkBox_Supp.Checked) { roles_str += "S"; }
				string Duo = "";
				if (!string.IsNullOrWhiteSpace(textBox_Duo.Text)) { Duo = textBox_Duo.Text; }
				else { Duo = "N/A"; }
				// Edit into DataGridView
				Players.SelectedRows[0].Cells[1].Value = textBox_Name.Text;
				Players.SelectedRows[0].Cells[2].Value = textBox_Uniq.Text;
				Players.SelectedRows[0].Cells[3].Value = textBox_IGN.Text;
				Players.SelectedRows[0].Cells[4].Value = comboBox_Tier.Text;
				Players.SelectedRows[0].Cells[5].Value = roles_str;
				Players.SelectedRows[0].Cells[6].Value = Duo;
				DataGridViewRow Player = Players.SelectedRows[0];
				Player.Cells[0].Value = "X";
				string Tier = comboBox_Tier.Text.Split(' ')[0];
				// Modify colors based on Ranking
				switch (Tier) {
					case "Level":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.LEVELHEX);
						break;
					case "Bronze":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.BRONZEHEX);
						break;
					case "Silver":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.SILVERHEX);
						break;
					case "Gold":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.GOLDHEX);
						break;
					case "Platinum":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.PLATHEX);
						break;
					case "Diamond":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.DIAMONDHEX);
						break;
					case "Master":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.MASTERHEX);
						break;
					case "Challenger":
						Player.Cells[4].Style.BackColor = ColorTranslator.FromHtml(MainForm.CHALLENGERHEX);
						break;
					default:
						break;
				}
				Players.ClearSelection();
			}
		}

		private void button_OK_Click(object sender, EventArgs e) {
			button_pressed = true;
			this.Close();
		}
	}
}
