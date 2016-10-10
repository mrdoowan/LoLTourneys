using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoLBalancing
{
	public partial class MainForm : Form
	{
		public MainForm() {
			InitializeComponent();
		}

		#region Variables and Structs

		// Variables for upgrading
		static private bool upgrading = false;
		private const string version = "1.0.0";

		// Color Codes for Ranks
		public const string levelHex = "#B4A7D6";
		public const string bronzeHex = "#9B5105";
		public const string silverHex = "#C0C0C0";
		public const string goldHex = "#F6B26B";
		public const string platHex = "#5CBFA1";
		public const string diamondHex = "#A4C2F4";
		public const string masterHex = "#E5E5E5";
		public const string challengerHex = "#FFD966";

		// Saveable through Properties.Settings
		static public string points;

		// Balancing Variables
		static private Dictionary<string, int> DefRanktoPts = new Dictionary<string, int>() {
			{ "Level 1-19", 0 },
			{ "Level 20-29", 1 },
			{ "Level 30", 2 },
			{ "Bronze 5", 2 },
			{ "Bronze 4", 2 },
			{ "Bronze 3", 2 },
			{ "Bronze 2", 3 },
			{ "Bronze 1", 3 },
			{ "Silver 5", 4 },
			{ "Silver 4", 4 },
			{ "Silver 3", 5 },
			{ "Silver 2", 5 },
			{ "Silver 1", 6 },
			{ "Gold 5", 6 },
			{ "Gold 4", 7 },
			{ "Gold 3", 8 },
			{ "Gold 2", 9 },
			{ "Gold 1", 10 },
			{ "Platinum 5", 11 },
			{ "Platinum 4", 12 },
			{ "Platinum 3", 13 },
			{ "Platinum 2", 14 },
			{ "Platinum 1", 15 },
			{ "Diamond 5", 16 },
			{ "Diamond 4", 17 },
			{ "Diamond 3", 18 },
			{ "Diamond 2", 19 },
			{ "Diamond 1", 20 },
			{ "Master", 21 },
			{ "Challenger", 23 }
		};

		public class Summoner:IComparable<Summoner>
		{
			public string Name { get; set; }
			public string Uniq { get; set; }
			public string IGN { get; set; }
			public string Rank { get; set; }
			public int Points { get; set; }
			public string Roles { get; set; }
			public int NumRoles { get; set; } // Need this for Duos
			public string Duo { get; set; }
			public int CompareTo(Summoner other) {
				if (Points == other.Points) {
					return NumRoles.CompareTo(other.NumRoles);
				}
				else {
					return Points.CompareTo(other.Points) * -1; // Sort biggest to smallest
				}
			}

			public Summoner(string Name_, string Uniq_, string IGN_, string Rank_, int Points_,
				string Roles_, string Duo_) {
				Name = Name_;
				Uniq = Uniq_;
				IGN = IGN_;
				Rank = Rank_;
				Points = Points_;
				Roles = Roles_;
				NumRoles = Roles_.Length;
				Duo = Duo_;
			}
		}

		public class Team
		{
			public List<Summoner> TeamPlayers;
			public string RolesRemain { get; set; }
			public int TotalPoints { get; set; }

			public Team(string Roles_ = "TJMAS", int TotalPts_ = 0) {
				TeamPlayers = new List<Summoner>();
				RolesRemain = Roles_;
				TotalPoints = TotalPts_;
			}
		}

		// Teams in League
		private List<Team> Teams = new List<Team>();
		private Random Rand = new Random();

		#endregion

		#region Helper Functions

		private void Update_TotPlayers() {
			int numPlayers = dataGridView_Players.Rows.Count;
			label_Total.Text = "Total Players: " + numPlayers;
		}

		// For securing the Trash
		private static void releaseObject(object obj) {
			try {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
				obj = null;
			}
			catch (Exception ex) {
				obj = null;
				throw ex;
			}
			finally {
				GC.Collect();
			}
		}

		// Check update for a new version
		private void Check_Update() {

		}

		// To fill in the background Cell color of a datagridview based on Ranking
		private void FillCellColor(DataGridViewRow Row, int Ind, string Tier) {
			switch (Tier) {
				case "Level":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(levelHex);
					break;
				case "Bronze":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(bronzeHex);
					break;
				case "Silver":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(silverHex);
					break;
				case "Gold":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(goldHex);
					break;
				case "Platinum":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(platHex);
					break;
				case "Diamond":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(diamondHex);
					break;
				case "Master":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(masterHex);
					break;
				case "Challenger":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(challengerHex);
					break;
				default:
					break;
			}
		}

		#endregion

		#region Event Handlers (Opening/Closing MainForm)

		private void MainForm_Load(object sender, EventArgs e) {
			label_Version.Text = "v" + version + " by Steven Duan (sduans@umich.edu)";
			// Load Properties.Settings
			string PtsList = Properties.Settings.Default.PointsList;
			if (!string.IsNullOrWhiteSpace(PtsList)) {
				string[] Pts = PtsList.Split(' ');
				int i = 0;
				foreach (string Rank in DefRanktoPts.Keys) {
					if (i < Pts.Length) { dataGridView_Ranks.Rows.Add(Rank, Pts[i]); }
					i++;
					DataGridViewRow RankRow = dataGridView_Ranks.Rows[dataGridView_Players.Rows.Count - 1];
					string Tier = Rank.Split(' ')[0];
					FillCellColor(RankRow, 0, Tier);
				}
			}
			else {
				// Default
				foreach (string Rank in DefRanktoPts.Keys) {
					string Pts = DefRanktoPts[Rank].ToString();
					dataGridView_Ranks.Rows.Add(Rank, Pts);
					DataGridViewRow RankRow = dataGridView_Ranks.Rows[dataGridView_Ranks.Rows.Count - 1];
					string Tier = Rank.Split(' ')[0];
					FillCellColor(RankRow, 0, Tier);
				}
			}
			textBox_APIKey.Text = Properties.Settings.Default.APIKey;
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
			// Save into Properties.Settings
			string Pts = "";
			for (int i = 1; i < dataGridView_Ranks.Rows.Count; ++i) {
				int Pt = int.Parse(dataGridView_Ranks[1, i].Value.ToString());
				Pts += Pt + " ";
			}
			Pts.TrimEnd(' ');
			Properties.Settings.Default.PointsList = Pts;
			Properties.Settings.Default.APIKey = textBox_APIKey.Text;
			Properties.Settings.Default.Save();
		}

		#endregion

		#region Event Handlers (TabPage: Player Roster)

		// WinForm for Adding a Player
		private void button_AddPlayer_Click(object sender, EventArgs e) {
			AddPlayer Player_Win = new AddPlayer();
			Player_Win.AddDialog(ref dataGridView_Players);
			Update_TotPlayers();
		}

		// OpenDialog Excel
		private void button_LoadPlayers_Click(object sender, EventArgs e) {
			if (dataGridView_Players.Rows.Count > 0) {
				MessageBox.Show("NOTE: You'll lose all data on the Players you currently have on the grid. " +
					"Make sure you save the Players before loading.", "Caution", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
			OpenFileDialog openExcelDialog = new OpenFileDialog();
			openExcelDialog.Filter = "Excel Sheet (*.xlsx)|*.xlsx";
			openExcelDialog.Title = "Load Players";
			openExcelDialog.RestoreDirectory = true;
			if (openExcelDialog.ShowDialog() == DialogResult.OK) {
				Application.DoEvents();
				Cursor.Current = Cursors.WaitCursor;
				// Open Excel
				Excel.Application xlApp = new Excel.Application();
				Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(openExcelDialog.FileName);
				Excel.Worksheet Sheet = xlWorkBook.Worksheets.get_Item(1);
				if (Sheet.Cells[1, 1].value.ToString() != "Name" || Sheet.Cells[1, 2].value.ToString() != "Uniq" ||
					Sheet.Cells[1, 3].value.ToString() != "Summoner" || Sheet.Cells[1, 4].value.ToString() != "Rank" ||
					Sheet.Cells[1, 5].value.ToString() != "Roles" || Sheet.Cells[1, 6].value.ToString() != "Duo") {
					// Validating the correct sheet
					MessageBox.Show("Incorrect Excel Sheet to parse. Make sure the format is correct.", "Wrong Excel", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				}
				dataGridView_Players.Rows.Clear();
				for (int i = 2; Sheet.Cells[i, 1].value != null; ++i) {
					string name = "", uniq = "", summoner = "", rank = "", roles = "", duo = "N/A";
					try { name = Sheet.Cells[i, 1].value.ToString(); } catch { }
					try { uniq = Sheet.Cells[i, 2].value.ToString(); } catch { }
					try { summoner = Sheet.Cells[i, 3].value.ToString(); } catch { }
					try { rank = Sheet.Cells[i, 4].value.ToString(); } catch { }
					try { roles = Sheet.Cells[i, 5].value.ToString(); } catch { }
					try { duo = Sheet.Cells[i, 6].value.ToString(); } catch { }
					DataGridViewButtonColumn button = new DataGridViewButtonColumn();
					dataGridView_Players.Rows.Add(button, name, uniq, summoner, rank, roles, duo);
					DataGridViewRow Player = dataGridView_Players.Rows[dataGridView_Players.Rows.Count - 1];
					Player.Cells[0].Value = "X";
					string Tier = rank.Split(' ')[0];
					FillCellColor(Player, 4, Tier);
				}
				xlWorkBook.Close();
				xlApp.Quit();
				releaseObject(Sheet);
				releaseObject(xlWorkBook);
				releaseObject(xlApp);
				dataGridView_Players.ClearSelection();
				Cursor.Current = Cursors.Default;
			}
			// Update Players
			Update_TotPlayers();
		}

		// SaveDialog Excel
		private void button_SavePlayers_Click(object sender, EventArgs e) {
			SaveFileDialog saveExcelDialog = new SaveFileDialog();
			saveExcelDialog.Filter = "Excel Sheet (*.xlsx)|*.xlsx";
			saveExcelDialog.Title = "Save Players";
			if (saveExcelDialog.ShowDialog() == DialogResult.OK) {
				Application.DoEvents();
				Cursor.Current = Cursors.WaitCursor;
				// Make an Excel Sheet
				Excel.Application xlApp = new Excel.Application();
				Excel.Workbook xlWorkBook;
				Excel.Worksheet xlWorkSheet;
				object mis = System.Reflection.Missing.Value;
				xlWorkBook = xlApp.Workbooks.Add(mis);
				xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
				try {
					xlWorkSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17;
					xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
				}
				catch { }
				xlWorkSheet.Columns["A"].ColumnWidth = 20.00;	// Name
				xlWorkSheet.Columns["B"].ColumnWidth = 10.00;	// Uniq
				xlWorkSheet.Columns["C"].ColumnWidth = 18.00;	// Summoner
				xlWorkSheet.Columns["D"].ColumnWidth = 12.00;	// Rank
				xlWorkSheet.Columns["E"].ColumnWidth = 7.00;	// Roles Pref
				xlWorkSheet.Columns["F"].ColumnWidth = 18.00;   // Duo
				xlWorkSheet.Cells[1, 1] = "Name";
				xlWorkSheet.Cells[1, 2] = "Uniq";
				xlWorkSheet.Cells[1, 3] = "Summoner";
				xlWorkSheet.Cells[1, 4] = "Rank";
				xlWorkSheet.Cells[1, 5] = "Roles";
				xlWorkSheet.Cells[1, 6] = "Duo";
				xlWorkSheet.Rows[1].Font.Bold = true;
				xlWorkSheet.Rows[1].Font.Underline = true;
				xlWorkSheet.get_Range("A1", "F1").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

				int row = 2;
				foreach (DataGridViewRow Player in dataGridView_Players.Rows) {
					xlWorkSheet.Cells[row, 1] = Player.Cells[1].Value.ToString();
					xlWorkSheet.Cells[row, 2] = Player.Cells[2].Value.ToString();
					xlWorkSheet.Cells[row, 3] = Player.Cells[3].Value.ToString();
					xlWorkSheet.Cells[row, 4] = Player.Cells[4].Value.ToString();
					string Tier = Player.Cells[4].Value.ToString().Split(' ')[0];
					switch (Tier) {
						case "Level":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(levelHex));
							break;
						case "Bronze":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(bronzeHex));
							break;
						case "Silver":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(silverHex));
							break;
						case "Gold":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(goldHex));
							break;
						case "Platinum":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(platHex));
							break;
						case "Diamond":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(diamondHex));
							break;
						case "Master":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(masterHex));
							break;
						case "Challenger":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(challengerHex));
							break;
						default:
							break;
					}
					xlWorkSheet.Cells[row, 5] = Player.Cells[5].Value.ToString();
					xlWorkSheet.Cells[row, 6] = Player.Cells[6].Value.ToString();
					row++;
				}

				releaseObject(xlWorkSheet);
				string filename = saveExcelDialog.FileName;
				try {
					xlApp.DisplayAlerts = false;
					xlWorkBook.SaveAs(filename);
					xlApp.Visible = true;
				}
				catch (Exception ex) {
					MessageBox.Show("Can't overwrite file. Please save it as another name." + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				releaseObject(xlWorkBook);
				releaseObject(xlApp);
				Cursor.Current = Cursors.Default;
			}
		}

		// WinForm for Editing the Player
		private void dataGridView_Players_CellDoubleClick(object sender, DataGridViewCellEventArgs e) {
			AddPlayer Player_Win = new AddPlayer();
			Player_Win.EditDialog(ref dataGridView_Players);
		}

		// Removing a Player
		private void dataGridView_Players_CellContentClick(object sender, DataGridViewCellEventArgs e) {
			DataGridView Grid = (DataGridView)sender;
			if (e.RowIndex > 0) {
				DataGridViewRow Player = Grid.Rows[e.RowIndex];
				if (Grid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0) {
					// Button Clicked for that row.
					string message = "Do you want to remove \"" + Player.Cells[1].Value.ToString() + "\"?";
					if (MessageBox.Show(message, "Reminder", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
						Grid.Rows.RemoveAt(Player.Index);
						Grid.Refresh();
						Grid.ClearSelection();
					}
				}
			}
			Update_TotPlayers();
		}

		#endregion

		#region Event Handlers (TabPage: Balancing)

		// Resets the Default in Points
		private void button_ResetPoints_Click(object sender, EventArgs e) {
			int i = 0;
			foreach (int Pts in DefRanktoPts.Values) {
				try { dataGridView_Ranks[1, i].Value = Pts.ToString(); }
				catch { }
				i++;
			}
		}

		#region Helper Functions for Balance Algorithm

		// If there are two Duos, we want to take the Minimum amount of Roles
		private string Combine_UniqueRoles(string Role1, string Role2) {
			if (Role1.Length == 5 || Role2.Length == 5) {
				return "TJMAS";
			}
			string combined = Role1 + Role2;
			return new string(combined.ToCharArray().Distinct().ToArray()); // Removes Duplicate Chars
		}

		// Upon calling the function, it updates that Team's Roles Remaining based on its current Players
		private void Update_TeamRolesRemaining(int Sel) {
			if (Teams[Sel].RolesRemain.Length > 0) {
				string Roles = "";
				foreach (Summoner Player in Teams[Sel].TeamPlayers) {
					Roles += Player.Roles;
				}
				string CombinedRoles = new string(Roles.ToCharArray().Distinct().ToArray());
				foreach (char role in CombinedRoles) {
					int RoleInd = Teams[Sel].RolesRemain.IndexOf(role);
					if (RoleInd != -1) {
						Teams[Sel].RolesRemain = Teams[Sel].RolesRemain.Remove(RoleInd, 1);
					}
				}
			}
		}

		// Adds Player(s) to a Team.
		// Also focus implementation on adding Duos
		// Sel is the index of the Team we're adding the Player
		private void AddPlayerToTeam(Dictionary<string, Summoner> Roster, int Sel, Summoner Added) {
			if (Added.IGN.Contains('@')) {
				// It's a Duo
				int AndSymInd = Added.IGN.IndexOf('@');
				string Duo1_IGN = Added.IGN.Substring(0, AndSymInd);
				string Duo2_IGN = Added.IGN.Substring(AndSymInd + 1);
				Summoner Adding1 = Roster[Duo1_IGN];
				Teams[Sel].TeamPlayers.Add(Adding1);
				Summoner Adding2 = Roster[Duo2_IGN];
				Teams[Sel].TeamPlayers.Add(Adding2);
			}
			else {
				// Solo player
				Teams[Sel].TeamPlayers.Add(Added);
			}
			// After adding Players, Update Team Roles and Points
			Update_TeamRolesRemaining(Sel);
			Teams[Sel].TotalPoints += Added.Points;
		}

		// THIS IS THE BULK OF THE PROGRAM!!!
		// This adds a Player to a Team based on the Team with the Lowest Points
		// A "Random number" is added to the Lowest Points and grabs a Set of Teams within that range.
		// Team added is selected by random, causing constant different outputs each time.
		// Returns True if a Summoner was placed correctly, Returns False otherwise
		private bool Placement_Algorithm(Dictionary<string, Summoner> Roster, Summoner Added) {
			// Check which team has the lowest skill point and set it as the min
			int min = 65536; // Calling this "INFINITY" or maxing out the int
			foreach (Team team in Teams) {
				if (team.TotalPoints < min && team.TeamPlayers.Count < 5) {
					min = team.TotalPoints;
				}
			}
			// Now accumulate a list of teams that we can randomly select to add
			// players and update the teams appropriatedly.
			List<int> TeamInd = new List<int>();
			// First check teams that aren't yet full (or if we're checking a duo,
			// make sure that the team has 2 spots left)
			// List of teams with total_points that are min <= points < min + 2
			for (int i = 0; i < Teams.Count; ++i) {
				if (Teams[i].TeamPlayers.Count < 5) {
					// Now check if this added is 1) Duo AND 2) If team has enough
					if ((Added.Duo == "N/A") || 
						(Added.Duo != "N/A" && (Teams[i].TeamPlayers.Count < 4))) {
						// Add the Index if within Range of "Random Number"
						// Random Number is the numericupdown Control
						if ((Teams[i].TotalPoints >= min) && (Teams[i].TotalPoints <= (min + numericUpDown_RandNum.Value))) {
							TeamInd.Add(i);
						}
					}
				}
			}
			if (TeamInd.Count == 0) {
				// There's no room for this Duo
				return false;
			}
			// Check what roles are remaining for the teams in TeamInd
			List<int> TeamNeedsRole = new List<int>();
			for (int i = 0; i < TeamInd.Count; ++i) {
				// Check every team of their roles remaining and see if Summoner added fits those roles
				if (Teams[TeamInd[i]].RolesRemain.IndexOfAny(Added.Roles.ToCharArray()) != -1) {
					TeamNeedsRole.Add(TeamInd[i]);
				}
			}
			// Now check if TeamNeedsRole is empty. If it's empty, we can just
			// use our normal team_ind because that means none of the teams need
			// a specific required role
			if (TeamNeedsRole.Count > 0) {
				int SelectedInd = Rand.Next(TeamNeedsRole.Count);
				AddPlayerToTeam(Roster, TeamNeedsRole[SelectedInd], Added);
			}
			else {
				int SelectedInd = Rand.Next(TeamInd.Count);
				AddPlayerToTeam(Roster, TeamInd[SelectedInd], Added);
			}
			return true;
		}
		
		// Once Teams are completely balanced, output and save them into an Excel sheet.
		private void Save_TeamsExcel() {
			SaveFileDialog saveExcelDialog = new SaveFileDialog();
			saveExcelDialog.Filter = "Excel Sheet (*.xlsx)|*.xlsx";
			saveExcelDialog.Title = "Save Teams";
			if (saveExcelDialog.ShowDialog() == DialogResult.OK) {
				Application.DoEvents();
				Cursor.Current = Cursors.WaitCursor;
				// Make an Excel Sheet
				Excel.Application xlApp = new Excel.Application();
				Excel.Workbook xlWorkBook;
				Excel.Worksheet xlWorkSheet;
				object mis = System.Reflection.Missing.Value;
				xlWorkBook = xlApp.Workbooks.Add(mis);
				xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
				try {
					xlWorkSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaper11x17;
					xlWorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
				}
				catch { }

				xlWorkSheet.Columns["A"].ColumnWidth = 25.00;   // Name (Uniq)
				xlWorkSheet.Columns["B"].ColumnWidth = 18.00;   // Summoner Name
				xlWorkSheet.Columns["C"].ColumnWidth = 25.00;   // Roles Pref
				xlWorkSheet.Columns["D"].ColumnWidth = 12.00;   // Ranking
				xlWorkSheet.Cells[1, 1] = "Name (Uniq)";
				xlWorkSheet.Cells[1, 2] = "Summoner Name";
				xlWorkSheet.Cells[1, 3] = "Roles Preferred";
				xlWorkSheet.Cells[1, 4] = "Ranking";
				xlWorkSheet.Rows[1].Font.Bold = true;
				xlWorkSheet.Rows[1].Font.Underline = true;
				xlWorkSheet.get_Range("A1", "D1").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

				int row = 3, team = 1;
				foreach (Team TeamSel in Teams) {
					xlWorkSheet.Cells[row, 1] = "Team " + team;
					xlWorkSheet.get_Range("A" + row, "D" + row).Merge();
					foreach (Summoner Player in TeamSel.TeamPlayers) {
						row++;
						xlWorkSheet.Cells[row, 1] = Player.Name + " (" + Player.Uniq + ")";
						xlWorkSheet.Cells[row, 2] = Player.IGN;
						string roles = Player.Roles.Replace("T", "Top, ");
						roles = roles.Replace("J", "Jungle, ");
						roles = roles.Replace("M", "Mid, ");
						roles = roles.Replace("A", "ADC, ");
						roles = roles.Replace("S", "Support, ");
						roles = roles.TrimEnd(',', ' ');
						xlWorkSheet.Cells[row, 3] = roles;
						string ranking = Player.Rank;
						xlWorkSheet.Cells[row, 4] = ranking;
						string Tier = ranking.Split(' ')[0];
						switch (Tier) {
							case "Level":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(levelHex));
								break;
							case "Bronze":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(bronzeHex));
								break;
							case "Silver":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(silverHex));
								break;
							case "Gold":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(goldHex));
								break;
							case "Platinum":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(platHex));
								break;
							case "Diamond":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(diamondHex));
								break;
							case "Master":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(masterHex));
								break;
							case "Challenger":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(challengerHex));
								break;
							default:
								break;
						}
					}
					row++; row++; team++;
				}

				releaseObject(xlWorkSheet);
				string filename = saveExcelDialog.FileName;
				try {
					xlApp.DisplayAlerts = false;
					xlWorkBook.SaveAs(filename);
					xlApp.Visible = true;
				}
				catch (Exception ex) {
					MessageBox.Show("Can't overwrite file. Please save it as another name." + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				releaseObject(xlWorkBook);
				releaseObject(xlApp);
				Cursor.Current = Cursors.Default;
				Cursor.Current = Cursors.Default;
			}
		}

		#endregion

		// Conducts the Balancing Algorithm and Saves once a Team is set
		private void button_Balance_Click(object sender, EventArgs e) {
			int numPlayers = dataGridView_Players.Rows.Count;
			if (numPlayers == 0) {
				MessageBox.Show("There must be more than 0 Players to balance teams.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			if (numPlayers % 5 != 0) {
				MessageBox.Show("You currently have " + numPlayers + " Players\nNumber of players must be divisible by 5.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				return;
			}
			int numTeams = numPlayers / 5;
			if (MessageBox.Show("You will make " + numTeams + " balanced teams. Proceed?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
				for (int Try = 0; Try < 10; ++Try) {
					// This for loop is to see if Teams can be made within 10 tries.
					bool Failed = false;
					// Initialize all Big Variables
					Dictionary<string, int> RankToPts = new Dictionary<string, int>();          // Retrieving past setting
					Dictionary<string, Summoner> Roster = new Dictionary<string, Summoner>();   // Key: IGN, Value: Summoner class
					Dictionary<string, Summoner> Combined = new Dictionary<string, Summoner>(); // Same as above, but when combining Duos
					Teams.Clear();
					for (int i = 0; i < numTeams; ++i) { // Init # Teams
						Teams.Add(new Team());
					}
					// ------------ Get Rank to Points
					foreach (DataGridViewRow RankRow in dataGridView_Ranks.Rows) {
						string Rank = RankRow.Cells[0].Value.ToString();
						int Pts = int.Parse(RankRow.Cells[1].Value.ToString());
						RankToPts.Add(Rank, Pts);
					}
					// ------------ Load Summoners into Roster
					foreach (DataGridViewRow PlayerRow in dataGridView_Players.Rows) {
						string name = PlayerRow.Cells[1].Value.ToString();
						string uniq = PlayerRow.Cells[2].Value.ToString();
						string ign = PlayerRow.Cells[3].Value.ToString();
						string rank = PlayerRow.Cells[4].Value.ToString();
						int pts = RankToPts[rank];
						string roles = PlayerRow.Cells[5].Value.ToString();
						string duo = PlayerRow.Cells[6].Value.ToString();
						Summoner player_Roster = new Summoner(name, uniq, ign, rank, pts, roles, duo);
						Summoner player_Combined = new Summoner(name, uniq, ign, rank, pts, roles, duo);
						Roster.Add(ign, player_Roster);
						Combined.Add(ign, player_Combined);
					}
					// -------------------------------------------------------------------
					// Balancing Algorithm begins here
					// -------------------------------------------------------------------
					// ----------------------
					/* Step 1: Copy all entrants into another unordered_map.
					Combine all duos into a "combined" Summoner. The following are changed:
					- Combine points of the summoner
					- The IGN is both of their summoners, but split by a "@".
					- roles become uniquely combined
					- num_roles = min(num_roles_p1 & num_roles_p2) (For minimum Roles)
					- Ignore uniq and duo
					Once done, ERASE the other duo's information that is also on the Dictionary */
					foreach (string ign in Roster.Keys) {
						string duo = Roster[ign].Duo;
						if (duo != "N/A" && Combined.ContainsKey(ign)) {
							// This summoner has a Duo AND is still in Combined
							if (!Combined.ContainsKey(duo)) {
								MessageBox.Show("Duo does not exist for summoner " + ign, "Failed", MessageBoxButtons.OK, MessageBoxIcon.Stop);
								return;
							}
							// Append name + "@(DUO)"
							// KEEP DUO NAME IN ITS VARIABLE THOUGH, I NEED IT FOR A CHECK LATER ON
							Combined[ign].IGN += "@" + duo;
							Combined[ign].NumRoles = Math.Min(Combined[ign].Roles.Length, Combined[duo].Roles.Length);
							Combined[ign].Roles = Combine_UniqueRoles(Combined[ign].Roles, Combined[duo].Roles);
							Combined[ign].Points += Combined[duo].Points;
							// We can now remove the Duo from Combined
							Combined.Remove(duo);
						}
					}
					if (checkBox_BalByRole.Checked) {
						/* 
						 * Step 2: Next, handle all the entrants that have a num_roles = 1 and
						put those people into teams. The first n summoners placed into the empty
						n teams are randomized. If there are still additional summoners left with
						num_roles = 1 after the first n teams, then assign the next summoners with
						Placement Algorithm 
						*/
						// Add to SingleRoles
						List<Summoner> SingleRoles = new List<Summoner>();
						List<string> RemoveSingleRoles = new List<string>(); // You need this to remove any Duos from Combined
						foreach (string KeyName in Combined.Keys) {
							if (Combined[KeyName].NumRoles == 1) {
								SingleRoles.Add(Combined[KeyName]);
								RemoveSingleRoles.Add(KeyName);
							}
						}
						// Now remove from Combined
						foreach (string RemoveKey in RemoveSingleRoles) {
							Combined.Remove(RemoveKey);
						}
						// Assign first m SingleRole summoners into the n empty teams
						// It also ends when SingleRole.size() < teams.size()
						// Sort each Summoner by a Priority Queue
						SingleRoles.Sort();
						int filled = 0;
						while (filled < Teams.Count && SingleRoles.Count > 0) {
							int TeamInd = Rand.Next(numTeams);  // Generate Random number from 0 to numTeams - 1
							if (Teams[TeamInd].TeamPlayers.Count == 0) {
								AddPlayerToTeam(Roster, TeamInd, SingleRoles[0]);
								SingleRoles.RemoveAt(0);
								filled++;
							}
						}
						// After n teams are filled, if there are SingleRoles still remaining, use Placement Algorithm
						while (SingleRoles.Count > 0) {
							if (!Placement_Algorithm(Roster, SingleRoles[0])) { Failed = true; break; }
							SingleRoles.RemoveAt(0);
						}
						if (Failed) {
							continue; // We failed and need to Try Again
						}
					}
					/* 
					 * Step 3: Place every player (from Combined) into a priority queue (List). 
					The priority queue has to be sorted where the highest skill points 
					is placed at the top. If the highest skill points are equal, the one 
					with less roles is placed at the top. 
					*/
					List<Summoner> Remaining = new List<Summoner>();
					foreach (Summoner Player in Combined.Values) {
						Remaining.Add(Player);
					}
					Remaining.Sort(); // Put into priority queue
					while (Remaining.Count > 0) {
						Summoner Next = Remaining[0];
						if (!Placement_Algorithm(Roster, Next)) { Failed = true; break; }
						Remaining.RemoveAt(0);
					}
					if (Failed) {
						continue; // We failed and need to Try Again
					}
					// Check if each Team has 5 Players
					foreach (Team team in Teams) {
						if (team.TeamPlayers.Count != 5) { Failed = true; break; }
					}
					if (Failed) {
						continue; // We failed and need to Try Again
					}
					// Check if each Team has every Rolefulfilled
					foreach (Team team in Teams) {
						if (team.RolesRemain.Length > 0) { Failed = true; break; }
					}
					if (Failed) {
						continue;
					}
					// ...
					// ...
					// ...
					// If you are at this point, WE MADE A BALANCED TEAM!!! :D :D :D
					/* Step 4: Save the Teams into an Excel Sheet */
					MessageBox.Show("Teams balanced successfully! Save your results!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
					Save_TeamsExcel();
					return;
				}
				MessageBox.Show("Failed to make a Balanced Team. Adjust your settings to be more reasonable or try again.", "Failed", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
			}
		}

		#endregion

		#region Event Handlers (TabPage: Teams & Stats)

		// Loads the total number of Teams. Also sets how many 
		private void button_LoadTeams_Click(object sender, EventArgs e) {

		}

		// Based on the Team selected, display the Team
		private void comboBox_Teams_SelectedIndexChanged(object sender, EventArgs e) {

		}

		// Loads Stats based on the following format.
		private void button_GenStats_Click(object sender, EventArgs e) {

		}

		// Help box for the .txt format
		private void button_HelpStats_Click(object sender, EventArgs e) {

		}

		#endregion
	}
}
