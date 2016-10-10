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
	public partial class StatsGen : Form
	{
		public StatsGen() {
			InitializeComponent();
		}

		public void Generate() {

			this.ShowDialog();
			Application.DoEvents();
			Cursor.Current = Cursors.WaitCursor;
			try {
				
			}
			catch (Exception e) {
				this.Close();
				MessageBox.Show("Error: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			Cursor.Current = Cursors.Default;
		}
	}
}
