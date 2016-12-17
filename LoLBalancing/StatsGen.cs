using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoLBalancing
{
	public class StatsGen
	{
		// Default constructor
		public StatsGen() { }

        #region Private Variables / Functions
        private int NumTeams;
        private int TotalGames;
        private string Patch;
        private JToken Champs;

        // Returns Champion Name based on ID
        private string GetChampName(string ID) {
            return Champs[ID]["name"].ToString();
        }

        // Key: TeamNum, Value: List of Names
        private static Dictionary<int, List<string>> TeamNames = new Dictionary<int, List<string>>();
        // Compact List of Teams
        private static List<Team> Teams = new List<Team>();
        // Key: Champ Name, Value: # Times
        private static Dictionary<string, int> Bans = new Dictionary<string, int>();
        private static Dictionary<string, int> Picks = new Dictionary<string, int>();

        // Adds onto Picks and Bans Dict
        private void AddBansPick(ref Dictionary<string, int> Dict, string ChampID) {
            string ChampName = GetChampName(ChampID);
            if (!Dict.ContainsKey(ChampName)) {
                // New Entry
                Dict.Add(ChampName, 1);
            }
            else {
                // Add 1 more
                Dict[ChampName]++;
            }
        }

        // True = Yes
        // False = No
        private string BoolToString(bool yesno) {
            if (yesno) { return "Yes"; }
            else { return "No"; }
        }

        // Calculates the Team Points based on Fantasy LCS Summer 2016
        // (EXCEPTION: Rift Herald is 1 Point)
        private int TeamFantasyPoints(TeamGame Team) {
            int points = 0;
            if (Team.Winner) { points += 2; }
            if (Team.riftHerald) { points += 1; }
            points += (Team.baronKills * 2);
            points += Team.dragonKills;
            if (Team.firstBlood) { points += 2; }
            points += Team.towerKills;
            int Min = Team.matchDuration / 60;
            if (Min < 30 && Team.Winner) { points += 2; }
            return points;
        }

        // Calculates the Player Points based on Fantasy LCS Summer 2016
        private double PlayerFantasyPoints(PlayerGame Player) {
            double points = 0;
            points += Player.Kills * 2;
            points -= Player.Deaths * 0.5;
            points += Player.Assists * 1.5;
            points += Player.CS * 0.01;
            points += Player.Triple * 2;
            points += Player.Quadra * 5;
            points += Player.Penta * 10;
            if (Player.Kills >= 10 || Player.Assists >= 10) { points += 2; }
            return Math.Round(points, 2);
        }
        #endregion

        #region Structs
        private class Team
        {
            public Dictionary<int, TeamGame> TeamGames;             // Key: GameNumber, Value: TeamGame
            public Dictionary<string, List<PlayerGame>> Players;    // Key: IGN, Value: List<PlayerGame>

            public Team() {
                TeamGames = new Dictionary<int, TeamGame>();
                Players = new Dictionary<string, List<PlayerGame>>();
            }
        }

        private class TeamGame
        {
            public int teamKills { get; set; }
            public int teamDeaths { get; set; }
            public int teamGold { get; set; }
            public bool Winner { get; set; }
            public bool firstBlood { get; set; }
            public bool riftHerald { get; set; }
            public int baronKills { get; set; }
            public int dragonKills { get; set; }
            public int towerKills { get; set; }
            public int matchDuration { get; set; } // in seconds

            public TeamGame(int Kills_, int Deaths_, int Gold_, bool Win_, bool FB_, 
                bool Rift_, int Baron_, int Dragon_, int Towers_, int Length_) {
                teamKills = Kills_;
                teamDeaths = Deaths_;
                teamGold = Gold_;
                Winner = Win_;
                firstBlood = FB_;
                riftHerald = Rift_;
                baronKills = Baron_;
                dragonKills = Dragon_;
                towerKills = Towers_;
                matchDuration = Length_;
            }
        }

        private class PlayerGame
        {
            public int GameNumber { get; set; }
            public string champName { get; set; }
            public string Role { get; set; }
            public int matchDuration { get; set; } // in seconds
            public int CSDiff10 { get; set; }
            public int CS { get; set; }
            public int Gold { get; set; }
            public int dmgChamps { get; set; }
            public int dmgTaken { get; set; }
            public int Kills { get; set; }
            public int Deaths { get; set; }
            public int Assists { get; set; }
            public int WardsDes { get; set; }
            public int WardsPla { get; set; }
            public int Double { get; set; }
            public int Triple { get; set; }
            public int Quadra { get; set; }
            public int Penta { get; set; }

            public PlayerGame(int GameNum_, string Champ_, string Role_, int Length_, int CS_10_, int CS_, int Gold_,
                int dmgChamps_, int dmgTaken_, int K_, int D_, int A_, int WardsDes_, int WardsPla_,
                int Double_, int Triple_, int Quadra_, int Penta_) {
                GameNumber = GameNum_;
                champName = Champ_;
                Role = Role_;
                matchDuration = Length_;
                CSDiff10 = CS_10_;
                CS = CS_;
                Gold = Gold_;
                dmgChamps = dmgChamps_;
                dmgTaken = dmgTaken_;
                Kills = K_;
                Deaths = D_;
                Assists = A_;
                WardsDes = WardsDes_;
                WardsPla = WardsPla_;
                Double = Double_;
                Triple = Triple_;
                Quadra = Quadra_;
                Penta = Penta_;
            }
        }
        
        #endregion

        // Input is a .txt file with all the Match History IDs
        public void Generate(string MatchesTxt, string NamesTxt, string APIKey, string region) {
			Application.DoEvents();
			Cursor.Current = Cursors.WaitCursor;
            TeamNames.Clear(); Teams.Clear(); Bans.Clear(); Picks.Clear();
            int requests = 1;
            try {
                // ---------------------------------
                // PARSE NAMES .TXT
                // ---------------------------------
                string[] NamesRow = NamesTxt.Split('\n');
                int TeamNum = 1, NumCheck = 0;
                List<string> IGNs = new List<string>(); // Key: IGN, Value: Role
                for (int i = 0; i < NamesRow.Length; ++i) {
                    if (i == 0) {
                        // First Row should be a 1
                        if (int.Parse(NamesRow[i]) != 1) {
                            MessageBox.Show("Wrong Format: 1 isn't the beginning of the Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else if (int.TryParse(NamesRow[i], out NumCheck)) {
                        // Reading Number
                        if (NumCheck != TeamNum + 1) {
                            MessageBox.Show("Team Numbers are not chronological.\nReload a correct Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else if (IGNs.Count < 5) {
                            MessageBox.Show("There are < 5 people in a Team.\nReload a correct Names .txt", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        else {
                            // We see a Number, so we add the List of Summoners
                            TeamNames.Add(TeamNum, IGNs);
                            IGNs = new List<string>();
                            TeamNum++;
                        }
                    }
                    else {
                        // Reading String (Summoner)
                        IGNs.Add(NamesRow[i]);
                    }
                }
                // Add the very last team.
                if (IGNs.Count < 5) {
                    MessageBox.Show("There are < 5 people in a Team.\nReload a correct Names .txt", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else {
                    TeamNames.Add(TeamNum, IGNs);
                }
                // ---------------------------------
                // PARSE MATCH .TXT
                // ---------------------------------
                string[] MatchRow = MatchesTxt.Split('\n');
				NumTeams = int.Parse(MatchRow[0]);
                // --------------- Initialize Teams with number of Teams
                for (int i = 0; i < NumTeams; ++i) {
                    Teams.Add(new Team());
                }
                TotalGames = MatchRow.Length - 1;
                // --------------- Retrieve Champion Information 
                // (This will also authenticate the API Key)
                string ChampJson = "";
                using (var WC = new WebClient()) {
                    ChampJson = WC.DownloadString("https://global.api.pvp.net/api/lol/static-data/" + region + 
                        "/v1.2/champion?locale=en_US&dataById=true&api_key=" + APIKey);
                }
                Champs = JObject.Parse(ChampJson)["data"];
                // ---------------------------------
                // RETRIEVE MATCH HISTORY
                // ---------------------------------
                for (int i = 1; i < MatchRow.Length; ++i) {
                    // --------------- Parse .Txt file
                    string[] Details = MatchRow[i].Split(' ');
                    string ID = Details[0];
                    int BlueTeamNum = int.Parse(Details[1]);
                    int RedTeamNum = int.Parse(Details[2]);
                    Team BlueTeam = Teams[BlueTeamNum - 1];
                    Team RedTeam = Teams[RedTeamNum - 1];
                    // --------------- Get URL Request of JSON
                    string URL = "https://" + region + ".api.pvp.net/api/lol/" + region +
                        "/v2.2/match/" + ID + "?api_key=" + APIKey;
                    string MatchJson = "";
                    using (var WC = new WebClient()) {
                        // Need to carefully cycle Rate Limits
                        while (true) {
                            try {
                                // Retrieve URL
                                MatchJson = WC.DownloadString(URL);
                                Thread.Sleep(1000);
                                break;
                            }
                            catch (Exception e) {
                                // Expecting 429 Error
                                if (!e.Message.Contains("429")) {
                                    MessageBox.Show("Error: " + e.Message, "Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                Thread.Sleep(1000);
                                continue;
                            }
                        }
                    }
                    JObject MatchParse = JObject.Parse(MatchJson);
                    string[] PatchDetail = MatchParse["matchVersion"].ToString().Split('.');
                    if (requests == 1) { Patch = PatchDetail[0] + "." + PatchDetail[1]; } // Only need Patch from 1
                    // --------------- Team Kills, Team Gold, Match Time, Team Data
                    JToken SummJson = MatchParse["participants"];   // SummJson[0-4] -> Blue, SummJson[5-9] -> Red
                    JToken TeamJson = MatchParse["teams"];          // TeamJson[0] -> Blue, TeamJson[1] -> Red
                    int MatchTime = int.Parse(MatchParse["matchDuration"].ToString());
                    int BlueGameNum = BlueTeam.TeamGames.Count + 1;
                    int RedGameNum = RedTeam.TeamGames.Count + 1;
                    for (int j = 0; j <= 1; ++j) {
                        bool Win = bool.Parse(TeamJson[j]["winner"].ToString());
                        bool FB = bool.Parse(TeamJson[j]["firstBlood"].ToString());
                        int Kills = 0, Deaths = 0, Gold = 0;
                        for (int k = 0; k < 5; ++k) {
                            if (j == 0) {
                                Kills += int.Parse(SummJson[k]["stats"]["kills"].ToString());
                                Deaths += int.Parse(SummJson[k]["stats"]["deaths"].ToString());
                                Gold += int.Parse(SummJson[k]["stats"]["goldEarned"].ToString());
                            }
                            else {
                                Kills += int.Parse(SummJson[k + 5]["stats"]["kills"].ToString());
                                Deaths += int.Parse(SummJson[k + 5]["stats"]["deaths"].ToString());
                                Gold += int.Parse(SummJson[k + 5]["stats"]["goldEarned"].ToString());
                            }
                        }
                        bool RiftHerald = bool.Parse(TeamJson[j]["firstRiftHerald"].ToString());
                        int Dragons = int.Parse(TeamJson[j]["dragonKills"].ToString());
                        int Barons = int.Parse(TeamJson[j]["baronKills"].ToString());
                        int Towers = int.Parse(TeamJson[j]["towerKills"].ToString());
                        // Add to TeamGames
                        TeamGame GameForTeam = new TeamGame(Kills, Deaths, Gold, Win, FB, RiftHerald,
                            Barons, Dragons, Towers, MatchTime);
                        if (j == 0) { BlueTeam.TeamGames.Add(BlueGameNum, GameForTeam); }
                        else { RedTeam.TeamGames.Add(RedGameNum, GameForTeam); }
                    }
                    // --------------- Bans
                    for (int j = 0; j <= 1; ++j) {
                        foreach (JToken ban in TeamJson[j]["bans"]) {
                            AddBansPick(ref Bans, ban["championId"].ToString());
                        }
                    }
                    // --------------- Summoner Data
                    List<PlayerGame> SummonersData = new List<PlayerGame>();
                    // Blue -> j == 0:4, Red -> == 5:9
                    for (int j = 0; j < 10; ++j) {
                        string Champ = GetChampName(SummJson[j]["championId"].ToString());
                        AddBansPick(ref Picks, SummJson[j]["championId"].ToString()); // Add to Picks Dictionary
                        string Role = SummJson[j]["timeline"]["role"].ToString();
                        string Lane = SummJson[j]["timeline"]["lane"].ToString();
                        if (Role == "DUO_CARRY") { Lane = "ADC"; }
                        else if (Role == "DUO_SUPPORT") { Lane = "SUPPORT"; }
                        int CSat10 = (int)(Math.Round(double.Parse(SummJson[j]["timeline"]["creepsPerMinDeltas"]["zeroToTen"].ToString()), 1) * 10);
                        JToken SummStats = SummJson[j]["stats"];
                        int CS = int.Parse(SummStats["minionsKilled"].ToString()) +
                            int.Parse(SummStats["neutralMinionsKilled"].ToString());
                        int Gold = int.Parse(SummStats["goldEarned"].ToString());
                        int DMG_Champs = int.Parse(SummStats["totalDamageDealtToChampions"].ToString());
                        int DMG_Taken = int.Parse(SummStats["totalDamageTaken"].ToString());
                        int Kills = int.Parse(SummStats["kills"].ToString());
                        int Deaths = int.Parse(SummStats["deaths"].ToString());
                        int Assists = int.Parse(SummStats["assists"].ToString());
                        int Wards_Des = int.Parse(SummStats["wardsKilled"].ToString());
                        int Wards_Pla = int.Parse(SummStats["wardsPlaced"].ToString());
                        int Penta = int.Parse(SummStats["pentaKills"].ToString());
                        int Quadra = int.Parse(SummStats["quadraKills"].ToString()) - Penta;
                        int Triple = int.Parse(SummStats["tripleKills"].ToString()) - Quadra - Penta;
                        int Double = int.Parse(SummStats["doubleKills"].ToString()) - Triple - Quadra - Penta;
                        int GameNum = 0;
                        if (j < 5) { GameNum = BlueGameNum; }
                        else { GameNum = RedGameNum; }
                        SummonersData.Add(new PlayerGame(GameNum, Champ, Lane, MatchTime, CSat10, CS, Gold,
                                DMG_Champs, DMG_Taken, Kills, Deaths, Assists, Wards_Des, Wards_Pla, Double,
                                Triple, Quadra, Penta));
                    }
                    // -------------- Check if every role is correct. Prompt a Window if not.
                    
                    // Calculate CSDiff@10 after Role Determination
                    var BlueCS10 = new Dictionary<string, int>();
                    var RedCS10 = new Dictionary<string, int>();
                    int pla = 0;
                    foreach (PlayerGame Player in SummonersData) {
                        if (pla < 5) { BlueCS10.Add(Player.Role, Player.CSDiff10); }
                        else { RedCS10.Add(Player.Role, Player.CSDiff10); }
                        pla++;
                    }
                    string[] Roles = { "TOP", "JUNGLE", "MIDDLE", "ADC", "SUPPORT" };
                    var BlueCSDiff = new Dictionary<string, int>();
                    foreach (string role in Roles) {
                        int CSDiff = BlueCS10[role] - RedCS10[role];
                        BlueCSDiff.Add(role, CSDiff);
                    }
                    for (int j = 0; j < SummonersData.Count; ++j) {
                        string Role = SummonersData[j].Role;
                        if (j < 5) { SummonersData[j].CSDiff10 = BlueCSDiff[Role]; }
                        else { SummonersData[j].CSDiff10 = BlueCSDiff[Role] * -1; }
                    }
                    // --------------- Process Requests
                    requests++;
				}
                // ---------------------------------
                // GENERATE EXCEL SHEET
                // ---------------------------------
                MessageBox.Show("Stats compiled successfully! Please save the Excel file.");
                #region Huge Bulky Code of Making Excel Sheet
                SaveFileDialog saveExcelDialog = new SaveFileDialog();
                saveExcelDialog.Filter = "Excel Sheet (*.xlsx)|*.xlsx";
                saveExcelDialog.Title = "Save Teams";
                if (saveExcelDialog.ShowDialog() == DialogResult.OK) {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkBook;
                    object mis = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(mis);
                    var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
                    // ------- Posting Team and Player Stats
                    for (int i = 0; i < Teams.Count; ++i) {
                        var xlWorkSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[i + 1], Type.Missing, Type.Missing, Type.Missing);
                        xlWorkSheet.Name = "Team " + (i + 1);
                        // Team Title
                        xlWorkSheet.Rows[1].RowWidth = 30.00;
                        xlWorkSheet.Range["A1"].Font.Bold = true;
                        xlWorkSheet.Range["A1"].Font.Size = 25;
                        xlWorkSheet.Range["A1"].Value = "Team " + (i + 1) + " - ";
                        // Team Stats
                        xlWorkSheet.get_Range("A2", "M2").WrapText = true;
                        xlWorkSheet.get_Range("A2", "M2").Font.Bold = true;
                        xlWorkSheet.get_Range("A2", "M2").Font.Underline = true;
                        xlWorkSheet.get_Range("A").ColumnWidth = 12.00;
                        xlWorkSheet.get_Range("B", "D").ColumnWidth = 10.00;
                        xlWorkSheet.get_Range("E").ColumnWidth = 13.00;
                        xlWorkSheet.get_Range("F", "M").ColumnWidth = 8.00;
                        xlWorkSheet.get_Range("A2", "M2").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.Range["A2"].Value = "Team Stats";
                        xlWorkSheet.Range["B2"].Value = "Game #";
                        xlWorkSheet.Range["C2"].Value = "Kills";
                        xlWorkSheet.Range["D2"].Value = "Deaths";
                        xlWorkSheet.Range["E2"].Value = "Match Duration";
                        xlWorkSheet.Range["F2"].Value = "Gold";
                        xlWorkSheet.Range["G2"].Value = "Win";
                        xlWorkSheet.Range["H2"].Value = "First Blood";
                        xlWorkSheet.Range["I2"].Value = "Rift Herald";
                        xlWorkSheet.Range["J2"].Value = "Barons";
                        xlWorkSheet.Range["K2"].Value = "Dragons";
                        xlWorkSheet.Range["L2"].Value = "Towers";
                        xlWorkSheet.Range["M2"].Value = "Fantasy Points";
                        int row = 3;
                        for (int j = 0; j < Teams[i].TeamGames.Count; ++j) {
                            TeamGame Game = Teams[i].TeamGames[j + 1];
                            xlWorkSheet.Range["B" + row].Value = "Game " + (j + 1);
                            xlWorkSheet.Range["C" + row].Value = Game.teamKills;
                            xlWorkSheet.Range["D" + row].Value = Game.teamDeaths;
                            int Min = Game.matchDuration / 60;
                            int Sec = Game.matchDuration % 60;
                            xlWorkSheet.Range["E" + row].Value = Min + " Min " + Sec + " Sec";
                            xlWorkSheet.Range["F" + row].Value = Game.teamGold;
                            xlWorkSheet.Range["G" + row].Value = BoolToString(Game.Winner);
                            xlWorkSheet.Range["H" + row].Value = BoolToString(Game.firstBlood);
                            xlWorkSheet.Range["I" + row].Value = BoolToString(Game.riftHerald);
                            xlWorkSheet.Range["J" + row].Value = Game.baronKills;
                            xlWorkSheet.Range["K" + row].Value = Game.dragonKills;
                            xlWorkSheet.Range["L" + row].Value = Game.towerKills;
                            xlWorkSheet.Range["M" + row].Value = TeamFantasyPoints(Game);
                            row++;
                        }
                        // Team Total
                        xlWorkSheet.Range["B" + row, "M" + row].Font.Bold = true;
                        xlWorkSheet.Range["B" + row].Value = "TOTAL";
                        xlWorkSheet.Range["C" + row].Formula = "=SUM(C3:C" + (row - 1) + ")";
                        xlWorkSheet.Range["D" + row].Formula = "=SUM(D3:D" + (row - 1) + ")";
                        xlWorkSheet.Range["F" + row].Formula = "=SUM(F3:F" + (row - 1) + ")";
                        xlWorkSheet.Range["J" + row].Formula = "=SUM(J3:J" + (row - 1) + ")";
                        xlWorkSheet.Range["K" + row].Formula = "=SUM(K3:K" + (row - 1) + ")";
                        xlWorkSheet.Range["L" + row].Formula = "=SUM(L3:L" + (row - 1) + ")";
                        xlWorkSheet.Range["M" + row].Formula = "=SUM(M3:M" + (row - 1) + ")";
                        int TeamTotalRow = row;
                        row++; row++;
                        // Player Stats
                        xlWorkSheet.get_Range("A" + row, "X" + row).WrapText = true;
                        xlWorkSheet.get_Range("A" + row, "X" + row).Font.Bold = true;
                        xlWorkSheet.get_Range("A" + row, "X" + row).Font.Underline = true;
                        xlWorkSheet.get_Range("N", "O").ColumnWidth = 8.00;
                        xlWorkSheet.get_Range("P", "R").ColumnWidth = 3.00;
                        xlWorkSheet.get_Range("S", "W").ColumnWidth = 5.00;
                        xlWorkSheet.get_Range("X").ColumnWidth = 8.00;
                        xlWorkSheet.Range["A" + row].Value = "Player Stats";
                        xlWorkSheet.Range["B" + row].Value = "Game #";
                        xlWorkSheet.Range["C" + row].Value = "Role";
                        xlWorkSheet.Range["D" + row].Value = "Champion";
                        xlWorkSheet.Range["E" + row].Value = "Match Duration (Min)";
                        xlWorkSheet.Range["F" + row].Value = "KDA";
                        xlWorkSheet.Range["G" + row].Value = "K/P %";
                        xlWorkSheet.Range["H" + row].Value = "D/P %";
                        xlWorkSheet.Range["I" + row].Value = "CSDiff10";
                        xlWorkSheet.Range["J" + row].Value = "CS/Min";
                        xlWorkSheet.Range["K" + row].Value = "DMG Champs / Min";
                        xlWorkSheet.Range["L" + row].Value = "DMG Taken / Min";
                        xlWorkSheet.Range["M" + row].Value = "Gold/Min";
                        xlWorkSheet.Range["N" + row].Value = "Wards Placed / Min";
                        xlWorkSheet.Range["O" + row].Value = "Wards Cleared / Min";
                        xlWorkSheet.Range["P" + row].Value = "K";
                        xlWorkSheet.Range["Q" + row].Value = "D";
                        xlWorkSheet.Range["R" + row].Value = "A";
                        xlWorkSheet.Range["S" + row].Value = "CS";
                        xlWorkSheet.Range["T" + row].Value = "Double";
                        xlWorkSheet.Range["U" + row].Value = "Triple";
                        xlWorkSheet.Range["V" + row].Value = "Quadra";
                        xlWorkSheet.Range["W" + row].Value = "Penta";
                        xlWorkSheet.Range["X" + row].Value = "Fantasy Points";
                        row++;
                        foreach (string IGN in Teams[i].Players.Keys) {
                            List<PlayerGame> Games = Teams[i].Players[IGN];
                            xlWorkSheet.Range["A" + row].Font.Bold = true;
                            xlWorkSheet.Range["A" + row].Value = IGN;
                            int TotDmgChamp = 0, TotDmgTake = 0, TotGold = 0, TotWardDes = 0, TotWardPla = 0;
                            foreach (PlayerGame Game in Games) {
                                int GameNum = Game.GameNumber;
                                TeamGame TeamG = Teams[i].TeamGames[GameNum];
                                xlWorkSheet.Range["B" + row].Value = "Game " + GameNum;
                                xlWorkSheet.Range["C" + row].Value = Game.Role;
                                xlWorkSheet.Range["D" + row].Value = Game.champName;
                                double TimeInMin = Game.matchDuration / 60.0;
                                xlWorkSheet.Range["E" + row].Value = Math.Round(TimeInMin, 2);
                                string KDA;
                                if (Game.Deaths == 0) { KDA = "Perfect"; }
                                else { KDA = Math.Round((double)(Game.Kills + Game.Assists) / Game.Deaths, 2).ToString(); }
                                xlWorkSheet.Range["F" + row].Value = KDA;
                                xlWorkSheet.Range["G" + row].NumberFormat = "###,##%";
                                xlWorkSheet.Range["G" + row].Value = (double)(Game.Kills + Game.Assists) / TeamG.teamKills;
                                xlWorkSheet.Range["H" + row].NumberFormat = "###,##%";
                                xlWorkSheet.Range["H" + row].Value = (double)(Game.Deaths) / TeamG.teamDeaths;
                                xlWorkSheet.Range["I" + row].Value = Game.CSDiff10;
                                xlWorkSheet.Range["J" + row].Value = Math.Round(Game.CS / TimeInMin, 2);
                                xlWorkSheet.Range["K" + row].Value = Math.Round(Game.dmgChamps / TimeInMin, 2);
                                TotDmgChamp += Game.dmgChamps;
                                xlWorkSheet.Range["L" + row].Value = Math.Round(Game.dmgTaken / TimeInMin, 2);
                                TotDmgTake += Game.dmgTaken;
                                xlWorkSheet.Range["M" + row].Value = Math.Round(Game.Gold / TimeInMin, 2);
                                TotGold += Game.Gold;
                                xlWorkSheet.Range["N" + row].Value = Math.Round(Game.WardsPla / TimeInMin, 2);
                                TotWardPla += Game.WardsPla;
                                xlWorkSheet.Range["O" + row].Value = Math.Round(Game.WardsDes / TimeInMin, 2);
                                TotWardDes += Game.WardsDes;
                                xlWorkSheet.Range["P" + row].Value = Game.Kills;
                                xlWorkSheet.Range["Q" + row].Value = Game.Deaths;
                                xlWorkSheet.Range["R" + row].Value = Game.Assists;
                                xlWorkSheet.Range["S" + row].Value = Game.CS;
                                xlWorkSheet.Range["T" + row].Value = Game.Double;
                                xlWorkSheet.Range["U" + row].Value = Game.Triple;
                                xlWorkSheet.Range["V" + row].Value = Game.Quadra;
                                xlWorkSheet.Range["W" + row].Value = Game.Penta;
                                xlWorkSheet.Range["X" + row].Value = PlayerFantasyPoints(Game);
                                row++;
                            }
                            // Player Total
                            int BegRow = row - Games.Count;
                            xlWorkSheet.Range["B" + row, "X" + row].Font.Bold = true;
                            xlWorkSheet.Range["B" + row].Value = "TOTAL";
                            xlWorkSheet.Range["E" + row].Formula = "=SUM(E" + BegRow + ":E" + (row - 1) + ")";
                            xlWorkSheet.Range["F" + row].Formula = "=ROUND((P" + row + "+R" + row + ")/Q" + row + ", 2)";
                            xlWorkSheet.Range["G" + row].NumberFormat = "###,##%";
                            xlWorkSheet.Range["G" + row].Formula = "=ROUND((P" + row + "+R" + row + ")/C" + TeamTotalRow + ", 2)";
                            xlWorkSheet.Range["H" + row].NumberFormat = "###,##%";
                            xlWorkSheet.Range["H" + row].Formula = "=ROUND(Q" + row + "/D" + TeamTotalRow + ", 2)";
                            xlWorkSheet.Range["I" + row].Formula = "=SUM(I" + BegRow + ":I" + (row - 1) + ")/" + Games.Count;
                            xlWorkSheet.Range["J" + row].Formula = "=ROUND(S" + row + "/E" + row + ", 2)";
                            xlWorkSheet.Range["K" + row].Formula = "=ROUND(" + TotDmgChamp + "/E" + row + ", 2)";
                            xlWorkSheet.Range["L" + row].Formula = "=ROUND(" + TotDmgTake + "/E" + row + ", 2)";
                            xlWorkSheet.Range["M" + row].Formula = "=ROUND(" + TotGold + "/E" + row + ", 2)";
                            xlWorkSheet.Range["N" + row].Formula = "=ROUND(" + TotWardPla + "/E" + row + ", 2)";
                            xlWorkSheet.Range["O" + row].Formula = "=ROUND(" + TotWardDes + "/E" + row + ", 2)";
                            xlWorkSheet.Range["P" + row].Formula = "=SUM(P" + BegRow + ":P" + (row - 1) + ")";
                            xlWorkSheet.Range["Q" + row].Formula = "=SUM(Q" + BegRow + ":Q" + (row - 1) + ")";
                            xlWorkSheet.Range["R" + row].Formula = "=SUM(R" + BegRow + ":R" + (row - 1) + ")";
                            xlWorkSheet.Range["S" + row].Formula = "=SUM(S" + BegRow + ":S" + (row - 1) + ")";
                            xlWorkSheet.Range["T" + row].Formula = "=SUM(T" + BegRow + ":T" + (row - 1) + ")";
                            xlWorkSheet.Range["U" + row].Formula = "=SUM(U" + BegRow + ":U" + (row - 1) + ")";
                            xlWorkSheet.Range["V" + row].Formula = "=SUM(V" + BegRow + ":V" + (row - 1) + ")";
                            xlWorkSheet.Range["W" + row].Formula = "=SUM(W" + BegRow + ":W" + (row - 1) + ")";
                            xlWorkSheet.Range["X" + row].Formula = "=SUM(X" + BegRow + ":X" + (row - 1) + ")";
                            row++;
                        }
                        releaseObject(xlWorkSheet);
                    }
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                else { return; }
                #endregion
            }
            catch (Exception e) {
				MessageBox.Show("Error: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
			}
            MessageBox.Show(requests + " requests were made to the Riot API.", "Note");
			Cursor.Current = Cursors.Default;
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
    }
}
