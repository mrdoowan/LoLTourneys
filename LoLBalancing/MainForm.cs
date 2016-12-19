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
using System.Net;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoLBalancing
{
	public partial class MainForm : Form
	{
		public MainForm() {
			InitializeComponent();
		}

		#region Variables and Structs
		// Constants
		public const int NUM_PLAYERS = 5;

		// Variables for upgrading
		static private bool upgrading = false;
		private const string VERSION = "0.Shit.Alpha";

		// Color Codes for Ranks
		public const string LEVELHEX = "#B4A7D6";
		public const string BRONZEHEX = "#9B5105";
		public const string SILVERHEX = "#C0C0C0";
		public const string GOLDHEX = "#F6B26B";
		public const string PLATHEX = "#5CBFA1";
		public const string DIAMONDHEX = "#A4C2F4";
		public const string MASTERHEX = "#E5E5E5";
		public const string CHALLENGERHEX = "#FFD966";

		// Saveable through Properties.Settings


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

        // Txt File strings
        private string matchTxt;
        private string namesTxt;

        private int pageNumber;
        private List<List<string>> IGNs = new List<List<string>>();
        private List<StatsGame> statsGames = new List<StatsGame>();

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

		// Check update for a new VERSION
		private void Check_Update() {

		}

		// To fill in the background Cell color of a datagridview based on Ranking
		private void FillCellColor(DataGridViewRow Row, int Ind, string Tier) {
			switch (Tier) {
				case "Level":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(LEVELHEX);
					break;
				case "Bronze":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(BRONZEHEX);
					break;
				case "Silver":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(SILVERHEX);
					break;
				case "Gold":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(GOLDHEX);
					break;
				case "Platinum":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(PLATHEX);
					break;
				case "Diamond":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(DIAMONDHEX);
					break;
				case "Master":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(MASTERHEX);
					break;
				case "Challenger":
					Row.Cells[Ind].Style.BackColor = ColorTranslator.FromHtml(CHALLENGERHEX);
					break;
				default:
					break;
			}
		}

		#endregion

		#region Event Handlers (Opening/Closing MainForm)

		private void MainForm_Load(object sender, EventArgs e) {
			label_Version.Text = "v" + VERSION + " by Steven Duan (sduans@umich.edu)";
			// Load Properties.Settings
			string PtsList = Properties.Settings.Default.PointsList;
			if (!string.IsNullOrWhiteSpace(PtsList)) {
				string[] Pts = PtsList.Split(' ');
				int i = 0;
				foreach (string Rank in DefRanktoPts.Keys) {
					if (i < Pts.Length) { dataGridView_Ranks.Rows.Add(Rank, Pts[i]); }
					i++;
					DataGridViewRow RankRow = dataGridView_Ranks.Rows[dataGridView_Ranks.Rows.Count - 1];
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
			comboBox_Region.SelectedIndex = 0;
		}

		private void MainForm_FormClosing(object sender, FormClosingEventArgs e) {
			// Save into Properties.Settings
			string Pts = "";
			for (int i = 0; i < dataGridView_Ranks.Rows.Count; ++i) {
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
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(LEVELHEX));
							break;
						case "Bronze":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(BRONZEHEX));
							break;
						case "Silver":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(SILVERHEX));
							break;
						case "Gold":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(GOLDHEX));
							break;
						case "Platinum":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(PLATHEX));
							break;
						case "Diamond":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(DIAMONDHEX));
							break;
						case "Master":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(MASTERHEX));
							break;
						case "Challenger":
							xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(CHALLENGERHEX));
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
			if (MessageBox.Show("Do you want to reset your values back to Default?", "Note", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
				int i = 0;
				foreach (int Pts in DefRanktoPts.Values) {
					try { dataGridView_Ranks[1, i].Value = Pts.ToString(); }
					catch { }
					i++;
				}
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
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(LEVELHEX));
								break;
							case "Bronze":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(BRONZEHEX));
								break;
							case "Silver":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(SILVERHEX));
								break;
							case "Gold":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(GOLDHEX));
								break;
							case "Platinum":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(PLATHEX));
								break;
							case "Diamond":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(DIAMONDHEX));
								break;
							case "Master":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(MASTERHEX));
								break;
							case "Challenger":
								xlWorkSheet.Cells[row, 4].Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(CHALLENGERHEX));
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

		#region Event Handlers (TabPage: Teams)

		// Loads the total number of Teams. Also sets how many 
		private void button_LoadTeams_Click(object sender, EventArgs e) {

		}

        private void button_SaveTeamTxt_Click(object sender, EventArgs e) {

        }

        // Based on the Team selected, display the Team
        private void comboBox_Teams_SelectedIndexChanged(object sender, EventArgs e) {

		}

        #endregion

        #region Event Handlers (TabPage: Stats)
        
        // Loads Stats based on the following format.

        private void button_LoadMatches_Click(object sender, EventArgs e) {
            OpenFileDialog dlgFileOpen = new OpenFileDialog();
            dlgFileOpen.Filter = "Text files (*.txt)|*.txt";
            dlgFileOpen.Title = "Load Match History IDs";
            dlgFileOpen.RestoreDirectory = true;
            if (dlgFileOpen.ShowDialog() == DialogResult.OK) {
                try {
                    StreamReader sr = new StreamReader(dlgFileOpen.FileName);
                    matchTxt = sr.ReadToEnd();
                    label_MatchLoad.Visible = true;
                }
                catch {
                    MessageBox.Show("Error in loading .Txt.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button_LoadNames_Click(object sender, EventArgs e) {
            OpenFileDialog dlgFileOpen = new OpenFileDialog();
            dlgFileOpen.Filter = "Text files (*.txt)|*.txt";
            dlgFileOpen.Title = "Load Summoners";
            dlgFileOpen.RestoreDirectory = true;
            if (dlgFileOpen.ShowDialog() == DialogResult.OK) {
                try {
                    StreamReader sr = new StreamReader(dlgFileOpen.FileName);
                    namesTxt = sr.ReadToEnd();
                    label_NamesLoad.Visible = true;
                }
                catch {
                    MessageBox.Show("Error in loading .Txt.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Helper functions
        private bool parseNames_Txt() {
			IGNs.Clear();	// Clear List
			try {
                // ---------------------------------
                // PARSE NAMES .TXT
                // ---------------------------------
                string[] namesRow = namesTxt.Split('\n');
                int teamNum = 1, numCheck = 0;
                List<string> TeamIGNs = new List<string>();
                for (int i = 0; i < namesRow.Length; ++i) {
                    if (i == 0) {
                        // First Row should be a 1
                        if (int.Parse(namesRow[i]) != 1) {
                            MessageBox.Show("Wrong Format: 1 isn't the beginning of the Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                    }
                    else if (int.TryParse(namesRow[i], out numCheck)) {
                        // Reading Number
                        if (numCheck != teamNum + 1) {
                            MessageBox.Show("Team Numbers are not chronological.\nReload a correct Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        else if (TeamIGNs.Count < NUM_PLAYERS) {
                            MessageBox.Show("There are < 5 people in a Team.\nReload a correct Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }
                        else {
                            // We see a Number, so we add the List of Summoners
                            IGNs.Add(TeamIGNs);
                            TeamIGNs = new List<string>();
                            teamNum++;
                        }
                    }
                    else {
                        // Reading String (Summoner)
                        TeamIGNs.Add(namesRow[i]);
                    }
                }
                // Add the very last team.
                if (TeamIGNs.Count < NUM_PLAYERS) {
                    MessageBox.Show("There are < 5 people in a Team.\nReload a correct Names .txt", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                else {
                    IGNs.Add(TeamIGNs);
                }
            }
            catch (Exception e) {
                MessageBox.Show("Error in parsing Names \nReason: " + 
					e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // Helper function to Load Matches
        private bool parseMatch_Txt() {
			statsGames.Clear(); // Clear List first
			RiotJson json = new RiotJson(comboBox_Region.Text, textBox_APIKey.Text);
			int i = 0; // DEBUGGING
			try {
                // --------------- Retrieve Champion Information
                JToken champJson = json.getChampJson()["data"];
                // --------------- Parsing matchTxt
                string[] matchRow = matchTxt.Split('\n');
                int numTeams = int.Parse(matchRow[0]);
                for (i = 1; i < matchRow.Length; ++i) {
                    string[] Details = matchRow[i].Split(' ');
                    long ID = long.Parse(Details[0]);
                    int BlueTeamNum = int.Parse(Details[1]);
					int RedTeamNum = int.Parse(Details[2]);
					StatsGame parseGame = new StatsGame(ID, RedTeamNum, BlueTeamNum);
					// Instantiate StatsPlayer class in
					JObject match = json.getMatchJson(ID.ToString());
					for (int j = 0; j < NUM_PLAYERS * 2; ++j) {
						JToken summJson = match["participants"];
						string champID = summJson[j]["championId"].ToString();
						string champName = champJson[champID]["name"].ToString();
						string role = summJson[j]["timeline"]["role"].ToString();
						string lane = summJson[j]["timeline"]["lane"].ToString();
                        if (lane == "JUNGLE") { lane = "JNG"; }
                        else if (lane == "MIDDLE") { lane = "MID"; }
						else if (role == "DUO_CARRY") { lane = "ADC"; }
						else if (role == "DUO_SUPPORT") { lane = "SUP"; }
						parseGame.addPlayer(champName, lane);
					}
					statsGames.Add(parseGame);
                }
            }
            catch (Exception e) {
                MessageBox.Show("Error in parsing Matches at i=" + i + "\nReason: " + e.Message, 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // Or otherwise known as the Compile Matches & Names
        // Save a preliminary .txt file
        private void button_GenNames_Click(object sender, EventArgs e) {
            if (string.IsNullOrWhiteSpace(matchTxt) || string.IsNullOrWhiteSpace(namesTxt)) {
                MessageBox.Show("No matches or names loaded.", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
			if (string.IsNullOrWhiteSpace(textBox_APIKey.Text)) {
				MessageBox.Show("API Key not entered.", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
			}
            Application.DoEvents();
            Cursor.Current = Cursors.WaitCursor;
            if (!parseNames_Txt()) { return; }
            if (!parseMatch_Txt()) { return; }
            // Instantiate the first Page
            initialize_GUI_Page(0);
            Cursor.Current = Cursors.Default;
            MessageBox.Show("Finished compilation!", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button_Next_Click(object sender, EventArgs e) {
            // Save Page into statsGames
            save_GUI_Page();
            // Initialize next Page
            initialize_GUI_Page(pageNumber + 1);
        }

        private void button_Prev_Click(object sender, EventArgs e) {
            // Save Page into statsGames
            save_GUI_Page();
            // Initialize next Page
            initialize_GUI_Page(pageNumber - 1);
        }

        // Initialize the Page
        // Sets pageNumber = pageNum
        private void initialize_GUI_Page(int pageNum) {
            try {
                StatsGame page = statsGames[pageNum];
                label_ID.Text = "GAME " + (pageNum + 1) + ": " + page.gameID;
                int blueNum = page.blueTeamNum;
                int redNum = page.redTeamNum;
                label6.Text = "TEAM " + blueNum;
                label7.Text = "TEAM " + redNum;
                // 0
                initialize_GUI_Player(page, 0, blueNum, comboBox_Role0,
                    comboBox_Name0, label_Champ0);
                // 1
                initialize_GUI_Player(page, 1, blueNum, comboBox_Role1,
                    comboBox_Name1, label_Champ1);
                // 2
                initialize_GUI_Player(page, 2, blueNum, comboBox_Role2,
                    comboBox_Name2, label_Champ2);
                // 3
                initialize_GUI_Player(page, 3, blueNum, comboBox_Role3,
                    comboBox_Name3, label_Champ3);
                // 4
                initialize_GUI_Player(page, 4, blueNum, comboBox_Role4,
                    comboBox_Name4, label_Champ4);
                // 5
                initialize_GUI_Player(page, 5, redNum, comboBox_Role5,
                    comboBox_Name5, label_Champ5);
                // 6
                initialize_GUI_Player(page, 6, redNum, comboBox_Role6,
                    comboBox_Name6, label_Champ6);
                // 7
                initialize_GUI_Player(page, 7, redNum, comboBox_Role7,
                    comboBox_Name7, label_Champ7);
                // 8
                initialize_GUI_Player(page, 8, redNum, comboBox_Role8,
                    comboBox_Name8, label_Champ8);
                // 9
                initialize_GUI_Player(page, 9, redNum, comboBox_Role9,
                    comboBox_Name9, label_Champ9);
                // Disable either Last or Next
                if (pageNum == 0) {
                    button_Prev.Enabled = false;
                    button_Next.Enabled = true;
                }
                else if (pageNum == statsGames.Count - 1) {
                    button_Prev.Enabled = true;
                    button_Next.Enabled = false;
                }
                else {
                    button_Prev.Enabled = true;
                    button_Next.Enabled = true;
                }
                // update pageNumber
                pageNumber = pageNum;
            }
            catch (Exception e) {
                MessageBox.Show("Your Match and Names files might be incorrect." + 
                    "\nReason: " + e.Message, "Error", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        // Used to initialize of each Player in Stats
        private void initialize_GUI_Player(StatsGame page, int playerNum,
            int teamNum, ComboBox role, ComboBox names, Label champ) {
            List<string> teamIGNs = IGNs[teamNum - 1];
            role.Text = page.Players[playerNum].role;
            names.Text = "";
            names.Items.Clear();    // Clear the items first
            for (int i = 0; i < teamIGNs.Count; ++i) {
                names.Items.Add(teamIGNs[i]);
            }
            names.Text = page.Players[playerNum].summoner;
            champ.Text = page.Players[playerNum].champ;
        }

        // Saves the exact Player into statsGames
        private void save_Player_statsGames(StatsGame game, int playerNum,
            ComboBox role, Label champ, ComboBox name) {
            game.Players[playerNum].role = role.Text;
            game.Players[playerNum].champ = champ.Text;
            game.Players[playerNum].summoner = name.Text;
        }

        // Saves the entire page into statsGames
        private void save_GUI_Page() {
            StatsGame game = statsGames[pageNumber];
            save_Player_statsGames(game, 0, comboBox_Role0, label_Champ0, comboBox_Name0);
            save_Player_statsGames(game, 1, comboBox_Role1, label_Champ1, comboBox_Name1);
            save_Player_statsGames(game, 2, comboBox_Role2, label_Champ2, comboBox_Name2);
            save_Player_statsGames(game, 3, comboBox_Role3, label_Champ3, comboBox_Name3);
            save_Player_statsGames(game, 4, comboBox_Role4, label_Champ4, comboBox_Name4);
            save_Player_statsGames(game, 5, comboBox_Role5, label_Champ5, comboBox_Name5);
            save_Player_statsGames(game, 6, comboBox_Role6, label_Champ6, comboBox_Name6);
            save_Player_statsGames(game, 7, comboBox_Role7, label_Champ7, comboBox_Name7);
            save_Player_statsGames(game, 8, comboBox_Role8, label_Champ8, comboBox_Name8);
            save_Player_statsGames(game, 9, comboBox_Role9, label_Champ9, comboBox_Name9);
        }

        private void button_GenStats_Click(object sender, EventArgs e) {
            // Check gameStats for validation of every Role
            int gameNum = 1;
            foreach (StatsGame game in statsGames) {
                string[] rolesArr = { "TOP", "JNG", "MID", "ADC", "SUP" };
                // BLUE
                List<string> rolesListBlue = rolesArr.ToList();
                for (int i = 0; i < NUM_PLAYERS; ++i) {
                    rolesListBlue.Remove(game.Players[i].role);
                }
                // RED
                List<string> rolesListRed = rolesArr.ToList();
                for (int i = NUM_PLAYERS; i < NUM_PLAYERS * 2; ++i) {
                    rolesListRed.Remove(game.Players[i].role);
                }
                if (rolesListBlue.Count > 0 || rolesListRed.Count > 0) {
                    MessageBox.Show("Game " + gameNum + " do not have all Roles fulfilled.",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                gameNum++;
            }
            // Check gameStats for validation of every Summoner
            gameNum = 1;
            foreach (StatsGame game in statsGames) {
                // BLUE
                List<string> summTeamBlue = new List<string>();
                for (int i = 0; i < NUM_PLAYERS; ++i) {
                    string summoner = game.Players[i].summoner;
                    if (string.IsNullOrWhiteSpace(summoner)) {
                        MessageBox.Show("Game " + gameNum + " has a blank summoner.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    summTeamBlue.Add(game.Players[i].summoner);
                }
                // RED
                List<string> summTeamRed = new List<string>();
                for (int i = NUM_PLAYERS; i < NUM_PLAYERS * 2; ++i) {
                    string summoner = game.Players[i].summoner;
                    if (string.IsNullOrWhiteSpace(summoner)) {
                        MessageBox.Show("Game " + gameNum + " has a blank summoner.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    summTeamRed.Add(game.Players[i].summoner);
                }
                if (summTeamBlue.Count != summTeamBlue.Distinct().Count() ||
                    summTeamRed.Count != summTeamRed.Distinct().Count()) {
                    MessageBox.Show("Game " + gameNum + " has duplicate summoners.",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                gameNum++;
            }
            // Next check for empty MatchTxt
            if (string.IsNullOrWhiteSpace(matchTxt)) {
                MessageBox.Show("No Match text loaded of any kind.", "Error", 
                    MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }
            // Proceed to make Stats
            label_StatsMsg.Visible = true;
            StatsGen Stats = new StatsGen(comboBox_Region.Text, textBox_APIKey.Text);
            Stats.Generate(statsGames, matchTxt);
            label_StatsMsg.Visible = false;
        }

        private void button_LoadComp_Click(object sender, EventArgs e) {
            OpenFileDialog dlgFileOpen = new OpenFileDialog();
            dlgFileOpen.Filter = "Text files (*.txt)|*.txt";
            dlgFileOpen.Title = "Load Stats Compilation";
            dlgFileOpen.RestoreDirectory = true;
            if (dlgFileOpen.ShowDialog() == DialogResult.OK) {
                try {
                    StreamReader sr_Names = new StreamReader(dlgFileOpen.FileName);
                    string file = sr_Names.ReadToEnd();
                    // Names
                    int endIndex = file.IndexOf("----MATCH_INFO----");
                    namesTxt = file.Substring(0, endIndex);
                    if (!parseNames_Txt()) { return; }
                    // Matches
                    statsGames.Clear();
                    int begIndexMatches = file.IndexOf('\n', endIndex);
                    string matchesString = file.Substring(begIndexMatches + 1);
                    string[] matches = matchesString.Split('\n');
                    StringBuilder sb_Matches = new StringBuilder();
                    sb_Matches.Append(IGNs.Count.ToString() + "\n");
                    for (int i = 0; i < matches.Length; ++i) {
                        // ID BLUE RED
                        sb_Matches.Append(matches[i] + "\n");
                        string[] matchNums = matches[i].Split(' ');
                        long ID = long.Parse(matchNums[0]);
                        int blueNum = int.Parse(matchNums[1]);
                        int redNum = int.Parse(matchNums[2]);
                        StatsGame game = new StatsGame(ID, redNum, blueNum);
                        // Onto the Players
                        for (int j = 0; j < NUM_PLAYERS * 2; ++j) {
                            ++i;
                            string[] playerDets = matches[i].Split(' ');
                            string role = playerDets[0];
                            string champ = playerDets[1].Replace('+', ' ');
                            string summoner = playerDets[2].TrimEnd('\n');
                            game.addPlayer(champ, role, summoner);
                        }
                        statsGames.Add(game);
                    }
                    matchTxt = sb_Matches.ToString().TrimEnd('\n');
                    // Instantiate the first page
                    initialize_GUI_Page(0);
                    // Also notify that Matches and Names are loaded
                    label_MatchLoad.Visible = true;
                    label_NamesLoad.Visible = true;
                }
                catch (Exception ex) {
                    MessageBox.Show("Error in loading Compilation.\nMatch and Names unloaded\nReason: " +
                        ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    matchTxt = "";
                    namesTxt = "";
                    label_MatchLoad.Visible = false;
                    label_NamesLoad.Visible = false;
                }
            }
        }

        private void button_SaveComp_Click(object sender, EventArgs e) {
            if (!string.IsNullOrWhiteSpace(namesTxt) && statsGames.Count > 0) {
                SaveFileDialog saveExcelDialog = new SaveFileDialog();
                saveExcelDialog.Filter = "Text File (*.txt)|*.txt";
                saveExcelDialog.Title = "Save Stats Compilation";
                if (saveExcelDialog.ShowDialog() == DialogResult.OK) {
                    try {
                        StringBuilder sb = new StringBuilder();
                        sb.Append(namesTxt + '\n');
                        sb.Append("----MATCH_INFO----\n");
                        for (int i = 0; i < statsGames.Count; ++i) {
                            StatsGame game = statsGames[i];
                            sb.Append(game.gameID + " " + game.blueTeamNum +
                                " " + game.redTeamNum + '\n');
                            for (int j = 0; j < game.Players.Count; ++j) {
                                sb.Append(game.Players[j].role + " ");
                                string champSpace = game.Players[j].champ.Replace(' ', '+');
                                sb.Append(champSpace + " ");
                                sb.Append(game.Players[j].summoner + '\n');
                            }
                        }
                        string filename = saveExcelDialog.FileName;
                        File.WriteAllText(filename, sb.ToString().TrimEnd('\n'));
                    }
                    catch (Exception ex) {
                        MessageBox.Show("Error in saving Compilation.\nReason: " +
                            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else {
                MessageBox.Show("No record of compilation for the Teams.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion
    }
}
