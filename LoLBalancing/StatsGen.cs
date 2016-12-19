using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Net;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace LoLBalancing
{
	public class StatsGen
	{

        #region Private Variables / Functions
        private int numTeams;
        private int totalGames;
        private string Patch;
        private string region;
        private string APIKey;
        private JToken Champs;
        private int NUM_PLAYERS;

        // Colors
        private const string TITLE_BACK = "#FFD966";
        private const string TOTAL_BACK = "#C9DAF8";

        // Default constructor
        public StatsGen(string region_, string APIKey_) {
            region = region_;
            APIKey = APIKey_;
            RiotJson json = new RiotJson(region, APIKey);
            Champs = json.getChampJson()["data"];
            // Initialize champion with the Champion Data
            NUM_PLAYERS = MainForm.NUM_PLAYERS;
        }

        // Returns Champion Name based on ID
        private string GetChampName(string ID) {
            return Champs[ID]["name"].ToString();
        }
        
        // Compact List of Teams
        private static List<Team> Teams = new List<Team>();
        // Key: Champ Name, Value: # Times
        private static Dictionary<string, int> Bans = new Dictionary<string, int>();
        private static Dictionary<string, int> Picks = new Dictionary<string, int>();
        private static Dictionary<string, int> PickWins = new Dictionary<string, int>();

        // Adds onto Bans Dict
        // IDisNum is true if the ID is an actual number
        // False otherwise if it's a name
        private void AddBanDict(string ChampID, bool IDisNum = true) {
            string ChampName = ChampID;
            if (IDisNum) { ChampName = GetChampName(ChampID); }
            if (!Bans.ContainsKey(ChampName)) {
                // New Entry
                Bans.Add(ChampName, 1);
            }
            else {
                // Add 1 more
                Bans[ChampName]++;
            }
        }

        // Adds onto Pick Dict AND Win Dict
        // Follows same protocol as above function
        private void AddPickDict(string ChampID, bool win, bool IDisNum = true) {
            string ChampName = ChampID;
            if (IDisNum) { ChampName = GetChampName(ChampID); }
            if (!Picks.ContainsKey(ChampName)) {
                // New Entry
                Picks.Add(ChampName, 1);
                if (win) { PickWins.Add(ChampName, 1); }
                else { PickWins.Add(ChampName, 0); }
            }
            else {
                // Add 1 More
                Picks[ChampName]++;
                if (win) { PickWins[ChampName]++; }
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
        public void Generate(List<StatsGame> gameList, string matchTxt) {
			Application.DoEvents();
			Cursor.Current = Cursors.WaitCursor;
            Teams.Clear(); Bans.Clear(); Picks.Clear();
            int requests = 1;
            try {
                // ---------------------------------
                // PARSE MATCH .TXT
                // ---------------------------------
                string[] matchRow = matchTxt.Split('\n');
                numTeams = int.Parse(matchRow[0]);
                // --------------- Initialize Teams with number of Teams
                for (int i = 0; i < numTeams; ++i) {
                    Teams.Add(new Team());
                }
                totalGames = matchRow.Length - 1;
                // ---------------------------------
                // RETRIEVE MATCH HISTORY
                // ---------------------------------
                for (int i = 1; i < matchRow.Length; ++i) {
                    // --------------- Parse .Txt file
                    string[] details = matchRow[i].Split(' ');
                    string ID = details[0];
                    int blueTeamNum = int.Parse(details[1]);
                    int redTeamNum = int.Parse(details[2]);
                    Team blueTeam = Teams[blueTeamNum - 1];
                    Team redTeam = Teams[redTeamNum - 1];
                    StatsGame gameDets = gameList[i - 1];
                    // --------------- Get URL Request of JSON
                    RiotJson json = new RiotJson(region, APIKey);
                    JObject matchParse = json.getMatchJson(ID.ToString());
                    string[] patchDetail = matchParse["matchVersion"].ToString().Split('.');
                    if (requests == 1) { Patch = patchDetail[0] + "." + patchDetail[1]; }
                    // Only need Patch from Game 1
                    // --------------- Team Kills, Team Gold, Match Time, Team Data
                    JToken summJson = matchParse["participants"];   // summJson[0-4] -> Blue, summJson[5-9] -> Red
                    JToken teamJson = matchParse["teams"];          // teamJson[0] -> Blue, teamJson[1] -> Red
                    int matchTime = int.Parse(matchParse["matchDuration"].ToString());
                    int blueGameNum = blueTeam.TeamGames.Count + 1;
                    int redGameNum = redTeam.TeamGames.Count + 1;
                    for (int j = 0; j <= 1; ++j) {
                        bool win = bool.Parse(teamJson[j]["winner"].ToString());
                        bool FB = bool.Parse(teamJson[j]["firstBlood"].ToString());
                        int kills = 0, deaths = 0, gold = 0;
                        for (int k = 0; k < NUM_PLAYERS; ++k) {
                            if (j == 0) {
                                kills += int.Parse(summJson[k]["stats"]["kills"].ToString());
                                deaths += int.Parse(summJson[k]["stats"]["deaths"].ToString());
                                gold += int.Parse(summJson[k]["stats"]["goldEarned"].ToString());
                            }
                            else {
                                kills += int.Parse(summJson[k + NUM_PLAYERS]["stats"]["kills"].ToString());
                                deaths += int.Parse(summJson[k + NUM_PLAYERS]["stats"]["deaths"].ToString());
                                gold += int.Parse(summJson[k + NUM_PLAYERS]["stats"]["goldEarned"].ToString());
                            }
                        }
                        bool riftHerald = bool.Parse(teamJson[j]["firstRiftHerald"].ToString());
                        int dragons = int.Parse(teamJson[j]["dragonKills"].ToString());
                        int barons = int.Parse(teamJson[j]["baronKills"].ToString());
                        int towers = int.Parse(teamJson[j]["towerKills"].ToString());
                        // Add to TeamGames
                        TeamGame gameForTeam = new TeamGame(kills, deaths, gold, win, FB, riftHerald,
                            barons, dragons, towers, matchTime);
                        if (j == 0) { blueTeam.TeamGames.Add(blueGameNum, gameForTeam); }
                        else { redTeam.TeamGames.Add(redGameNum, gameForTeam); }
                    }
                    // --------------- Bans
                    for (int j = 0; j <= 1; ++j) {
                        foreach (JToken ban in teamJson[j]["bans"]) {
                            AddBanDict(ban["championId"].ToString());
                        }
                    }
                    // --------------- Summoner Data
                    List<PlayerGame> SummonersData = new List<PlayerGame>();
                    // Blue -> j == 0:4, Red -> == 5:9
                    for (int j = 0; j < NUM_PLAYERS * 2; ++j) {
                        JToken SummStats = summJson[j]["stats"];
                        // player = gameDets.Players[j]
                        string champ = gameDets.Players[j].champ;
                        bool win = bool.Parse(SummStats["winner"].ToString());
                        AddPickDict(champ, win, false); // Add to Picks Dictionary
                        string role = gameDets.Players[j].role;
                        int CSat10 = (int)(Math.Round(double.Parse(summJson[j]["timeline"]["creepsPerMinDeltas"]["zeroToTen"].ToString()), 1) * 10);
                        int CS = int.Parse(SummStats["minionsKilled"].ToString()) +
                            int.Parse(SummStats["neutralMinionsKilled"].ToString());
                        int gold = int.Parse(SummStats["goldEarned"].ToString());
                        int DMG_Champs = int.Parse(SummStats["totalDamageDealtToChampions"].ToString());
                        int DMG_Taken = int.Parse(SummStats["totalDamageTaken"].ToString());
                        int kills = int.Parse(SummStats["kills"].ToString());
                        int deaths = int.Parse(SummStats["deaths"].ToString());
                        int assists = int.Parse(SummStats["assists"].ToString());
                        int wardsDes = int.Parse(SummStats["wardsKilled"].ToString());
                        int wardsPla = int.Parse(SummStats["wardsPlaced"].ToString());
                        int pentaKill = int.Parse(SummStats["pentaKills"].ToString());
                        int quadraKill = int.Parse(SummStats["quadraKills"].ToString()) - pentaKill;
                        int tripleKill = int.Parse(SummStats["tripleKills"].ToString()) - quadraKill - pentaKill;
                        int doubleKill = int.Parse(SummStats["doubleKills"].ToString()) - tripleKill - quadraKill - pentaKill;
                        int gameNum = 0;
                        if (j < NUM_PLAYERS) { gameNum = blueGameNum; }
                        else { gameNum = redGameNum; }
                        SummonersData.Add(new PlayerGame(gameNum, champ, role, matchTime, CSat10, CS, gold,
                                DMG_Champs, DMG_Taken, kills, deaths, assists, wardsDes, wardsPla, doubleKill,
                                tripleKill, quadraKill, pentaKill));
                    }
                    // Calculate CSDiff@10
                    var BlueCS10 = new Dictionary<string, int>();
                    var RedCS10 = new Dictionary<string, int>();
                    int pla = 0;
                    foreach (PlayerGame Player in SummonersData) {
                        if (pla < NUM_PLAYERS) { BlueCS10.Add(Player.Role, Player.CSDiff10); }
                        else { RedCS10.Add(Player.Role, Player.CSDiff10); }
                        pla++;
                    }
                    string[] Roles = { "TOP", "JNG", "MID", "ADC", "SUP" };
                    var BlueCSDiff = new Dictionary<string, int>();
                    foreach (string role in Roles) {
                        int CSDiff = BlueCS10[role] - RedCS10[role];
                        BlueCSDiff.Add(role, CSDiff);
                    }
                    for (int j = 0; j < NUM_PLAYERS * 2; ++j) {
                        string Role = SummonersData[j].Role;
                        if (j < NUM_PLAYERS) { SummonersData[j].CSDiff10 = BlueCSDiff[Role]; }
                        else { SummonersData[j].CSDiff10 = BlueCSDiff[Role] * -1; }
                    }
                    // Add into Players with Summoner Name as Key
                    for (int j = 0; j < NUM_PLAYERS * 2; ++j) {
                        string summoner = gameDets.Players[j].summoner;
                        Team sel_Team = null;
                        if (j < NUM_PLAYERS) { sel_Team = blueTeam; }
                        else { sel_Team = redTeam; }
                        if (!sel_Team.Players.ContainsKey(summoner)) {
                            // Add new entry if Key doesn't exist
                            var addList = new List<PlayerGame>();
                            addList.Add(SummonersData[j]);
                            sel_Team.Players.Add(summoner, addList);
                        }
                        else {
                            // Add into existing
                            sel_Team.Players[summoner].Add(SummonersData[j]);
                        }
                    }
                    // --------------- Process Requests
                    requests++;
                }
            }
            catch (Exception e) {
                MessageBox.Show("Error: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // ---------------------------------
            // GENERATE EXCEL SHEET
            // ---------------------------------
            MessageBox.Show("Stats compiled successfully! Please save the Excel file\n(DO NOT OVERWRITE A FILE)");

            #region Huge Bulky Code of Making Excel Sheet
            var xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            try {
                SaveFileDialog saveExcelDialog = new SaveFileDialog();
                saveExcelDialog.Filter = "Excel Sheet (*.xlsx)|*.xlsx";
                saveExcelDialog.Title = "Save Teams";
                if (saveExcelDialog.ShowDialog() == DialogResult.OK) {
                    object mis = System.Reflection.Missing.Value;
                    xlWorkBook = xlApp.Workbooks.Add(mis);
                    var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
                    // ------- Posting Team and Player Stats
                    for (int i = 0; i < Teams.Count; ++i) {
                        xlWorkSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[i + 1], 
                            Type.Missing, Type.Missing, Type.Missing);
                        xlWorkSheet.Name = "Team " + (i + 1);
                        // Team Title
                        xlWorkSheet.Rows[1].RowHeight = 30.00;
                        xlWorkSheet.Range["A1"].Font.Bold = true;
                        xlWorkSheet.Range["A1"].Font.Size = 25;
                        xlWorkSheet.Range["A1"].Value = "Team " + (i + 1) + " - ";
                        // Team Stats
                        xlWorkSheet.get_Range("A2", "M2").WrapText = true;
                        xlWorkSheet.get_Range("A2", "M2").Font.Bold = true;
                        xlWorkSheet.get_Range("A2", "M2").Font.Underline = true;
                        xlWorkSheet.Columns["A"].ColumnWidth = 18.00;
                        xlWorkSheet.Columns["B"].ColumnWidth = 11.00;
                        xlWorkSheet.Columns["C"].ColumnWidth = 10.00;
                        xlWorkSheet.Columns["D"].ColumnWidth = 10.00;
                        xlWorkSheet.Columns["E"].ColumnWidth = 12.00;
                        xlWorkSheet.Columns["F"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["G"].ColumnWidth = 6.00;
                        xlWorkSheet.Columns["H"].ColumnWidth = 6.00;
                        xlWorkSheet.Columns["I"].ColumnWidth = 8.00;
                        xlWorkSheet.Columns["J"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["K"].ColumnWidth = 8.00;
                        xlWorkSheet.Columns["L"].ColumnWidth = 8.00;
                        xlWorkSheet.Columns["M"].ColumnWidth = 9.00;
                        xlWorkSheet.get_Range("A2", "M2").Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.get_Range("A2", "M2").Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.get_Range("A2", "M2").Interior.Color =
                                ColorTranslator.ToOle(ColorTranslator.FromHtml(TITLE_BACK));
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
                        xlWorkSheet.Range["B" + row, "M" + row].Font.Underline = true;
                        xlWorkSheet.Range["A" + row, "M" + row].Interior.Color =
                                ColorTranslator.ToOle(ColorTranslator.FromHtml(TOTAL_BACK));
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
                        xlWorkSheet.get_Range("A" + row, "X" + row).Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.get_Range("A" + row, "X" + row).Cells.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        xlWorkSheet.get_Range("A" + row, "X" + row).Interior.Color =
                                ColorTranslator.ToOle(ColorTranslator.FromHtml(TITLE_BACK));
                        xlWorkSheet.Columns["N"].ColumnWidth = 9.00;
                        xlWorkSheet.Columns["O"].ColumnWidth = 9.00;
                        xlWorkSheet.Columns["P"].ColumnWidth = 3.00;
                        xlWorkSheet.Columns["Q"].ColumnWidth = 3.00;
                        xlWorkSheet.Columns["R"].ColumnWidth = 3.00;
                        xlWorkSheet.Columns["S"].ColumnWidth = 5.00;
                        xlWorkSheet.Columns["T"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["U"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["V"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["W"].ColumnWidth = 7.00;
                        xlWorkSheet.Columns["X"].ColumnWidth = 8.00;
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
                        xlWorkSheet.Range["K" + row].Value = "DMG Dealt / Min";
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
                                if (Game.Deaths == 0) {
                                    KDA = "Perfect";
                                    xlWorkSheet.Range["F" + row].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                }
                                else {
                                    KDA = Math.Round((double)(Game.Kills + Game.Assists) / Game.Deaths, 2).ToString();
                                }
                                xlWorkSheet.Range["F" + row].Value = KDA;
                                xlWorkSheet.Range["E" + row, "F" + row].NumberFormat = "0.00";
                                xlWorkSheet.Range["G" + row].NumberFormat = "0%";
                                xlWorkSheet.Range["G" + row].Value = (double)(Game.Kills + Game.Assists) / TeamG.teamKills;
                                xlWorkSheet.Range["H" + row].NumberFormat = "0%";
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
                                xlWorkSheet.Range["J" + row, "O" + row].NumberFormat = "0.00";
                                xlWorkSheet.Range["P" + row].Value = Game.Kills;
                                xlWorkSheet.Range["Q" + row].Value = Game.Deaths;
                                xlWorkSheet.Range["R" + row].Value = Game.Assists;
                                xlWorkSheet.Range["S" + row].Value = Game.CS;
                                xlWorkSheet.Range["T" + row].Value = Game.Double;
                                xlWorkSheet.Range["U" + row].Value = Game.Triple;
                                xlWorkSheet.Range["V" + row].Value = Game.Quadra;
                                xlWorkSheet.Range["W" + row].Value = Game.Penta;
                                xlWorkSheet.Range["X" + row].Value = PlayerFantasyPoints(Game);
                                xlWorkSheet.Range["X" + row].NumberFormat = "0.00";
                                row++;
                            }
                            // Player Total
                            int BegRow = row - Games.Count;
                            xlWorkSheet.Range["B" + row, "X" + row].Font.Bold = true;
                            xlWorkSheet.Range["B" + row, "X" + row].Font.Underline = true;
                            xlWorkSheet.Range["A" + row, "X" + row].Interior.Color = 
                                ColorTranslator.ToOle(ColorTranslator.FromHtml(TOTAL_BACK));
                            xlWorkSheet.Range["B" + row].Value = "TOTAL/AVG";
                            xlWorkSheet.Range["E" + row].Formula = "=SUM(E" + BegRow + ":E" + (row - 1) + ")";
                            xlWorkSheet.Range["F" + row].Formula = "=ROUND((P" + row + "+R" + row + ")/Q" + row + ", 2)";
                            xlWorkSheet.Range["E" + row, "F" + row].NumberFormat = "0.00";
                            xlWorkSheet.Range["G" + row].NumberFormat = "0%";
                            xlWorkSheet.Range["G" + row].Formula = "=ROUND((P" + row + "+R" + row + ")/C" + TeamTotalRow + ", 2)";
                            xlWorkSheet.Range["H" + row].NumberFormat = "0%";
                            xlWorkSheet.Range["H" + row].Formula = "=ROUND(Q" + row + "/D" + TeamTotalRow + ", 2)";
                            xlWorkSheet.Range["I" + row].Formula = "=SUM(I" + BegRow + ":I" + (row - 1) + ")/" + Games.Count;
                            xlWorkSheet.Range["J" + row].Formula = "=ROUND(S" + row + "/E" + row + ", 2)";
                            xlWorkSheet.Range["K" + row].Formula = "=ROUND(" + TotDmgChamp + "/E" + row + ", 2)";
                            xlWorkSheet.Range["L" + row].Formula = "=ROUND(" + TotDmgTake + "/E" + row + ", 2)";
                            xlWorkSheet.Range["M" + row].Formula = "=ROUND(" + TotGold + "/E" + row + ", 2)";
                            xlWorkSheet.Range["N" + row].Formula = "=ROUND(" + TotWardPla + "/E" + row + ", 2)";
                            xlWorkSheet.Range["O" + row].Formula = "=ROUND(" + TotWardDes + "/E" + row + ", 2)";
                            xlWorkSheet.Range["I" + row, "O" + row].NumberFormat = "0.00";
                            xlWorkSheet.Range["P" + row].Formula = "=SUM(P" + BegRow + ":P" + (row - 1) + ")";
                            xlWorkSheet.Range["Q" + row].Formula = "=SUM(Q" + BegRow + ":Q" + (row - 1) + ")";
                            xlWorkSheet.Range["R" + row].Formula = "=SUM(R" + BegRow + ":R" + (row - 1) + ")";
                            xlWorkSheet.Range["S" + row].Formula = "=SUM(S" + BegRow + ":S" + (row - 1) + ")";
                            xlWorkSheet.Range["T" + row].Formula = "=SUM(T" + BegRow + ":T" + (row - 1) + ")";
                            xlWorkSheet.Range["U" + row].Formula = "=SUM(U" + BegRow + ":U" + (row - 1) + ")";
                            xlWorkSheet.Range["V" + row].Formula = "=SUM(V" + BegRow + ":V" + (row - 1) + ")";
                            xlWorkSheet.Range["W" + row].Formula = "=SUM(W" + BegRow + ":W" + (row - 1) + ")";
                            xlWorkSheet.Range["X" + row].Formula = "=SUM(X" + BegRow + ":X" + (row - 1) + ")";
                            xlWorkSheet.Range["X" + row].NumberFormat = "0.00";
                            row++;
                        }
                        releaseObject(xlWorkSheet);
                    }
                    // Generate Pick and Ban Stats
                    xlWorkSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[Teams.Count + 1],
                           Type.Missing, Type.Missing, Type.Missing);
                    xlWorkSheet.Name = "Pick Ban Stats";
                    xlWorkSheet.Columns["A"].ColumnWidth = 11.00;
                    xlWorkSheet.Columns["B"].ColumnWidth = 7.00;
                    xlWorkSheet.Columns["C"].ColumnWidth = 7.00;
                    xlWorkSheet.Columns["D"].ColumnWidth = 9.00;
                    xlWorkSheet.Columns["E"].ColumnWidth = 9.00;
                    xlWorkSheet.Columns["F"].ColumnWidth = 7.00;
                    xlWorkSheet.Columns["G"].ColumnWidth = 11.00;
                    xlWorkSheet.Columns["H"].ColumnWidth = 7.00;
                    xlWorkSheet.Columns["I"].ColumnWidth = 9.00;
                    xlWorkSheet.Range["A1"].Font.Bold = true;
                    xlWorkSheet.Range["A1"].Font.Size = 25;
                    xlWorkSheet.Range["A1"].Value = "Champion Pick/Win Rate";
                    xlWorkSheet.Range["G1"].Font.Bold = true;
                    xlWorkSheet.Range["G1"].Font.Size = 25;
                    xlWorkSheet.Range["G1"].Value = "Champion Ban Rate";
                    xlWorkSheet.Range["L1"].Font.Bold = true;
                    xlWorkSheet.Range["L1"].Font.Size = 25;
                    xlWorkSheet.Range["L1"].Value = "Patch " + Patch;
                    xlWorkSheet.Range["A2"].Value = "Games:";
                    xlWorkSheet.Range["B2"].Value = totalGames;
                    xlWorkSheet.Range["A4", "I4"].Font.Bold = true;
                    xlWorkSheet.Range["A4", "I4"].Font.Underline = true;
                    xlWorkSheet.Range["A4", "I4"].Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    xlWorkSheet.Range["A4"].Value = "Champion";
                    xlWorkSheet.Range["B4"].Value = "Wins";
                    xlWorkSheet.Range["C4"].Value = "Games";
                    xlWorkSheet.Range["D4"].Value = "Pick Rate";
                    xlWorkSheet.Range["E4"].Value = "Win Rate";
                    xlWorkSheet.Range["G4"].Value = "Champion";
                    xlWorkSheet.Range["H4"].Value = "Bans";
                    xlWorkSheet.Range["I4"].Value = "Ban Rate";
                    // Pick
                    int row2 = 5;
                    foreach (string champ in Picks.Keys) {
                        int games = Picks[champ];
                        int wins = PickWins[champ];
                        xlWorkSheet.Range["A" + row2].Value = champ;
                        xlWorkSheet.Range["B" + row2].Value = wins;
                        xlWorkSheet.Range["C" + row2].Value = games;
                        xlWorkSheet.Range["D" + row2].Formula = "=C" + row2 + " / B2";
                        xlWorkSheet.Range["E" + row2].Formula = "=B" + row2 + " / C" + row2;
                        xlWorkSheet.Range["D" + row2, "E" + row2].NumberFormat = "0%";
                        row2++;
                    }
                    // Bans
                    row2 = 5;
                    foreach (string champ in Bans.Keys) {
                        int bans = Bans[champ];
                        xlWorkSheet.Range["G" + row2].Value = champ;
                        xlWorkSheet.Range["H" + row2].Value = bans;
                        xlWorkSheet.Range["I" + row2].Formula = "=H" + row2 + " / B2";
                        xlWorkSheet.Range["I" + row2].NumberFormat = "0%";
                        row2++;
                    }
                    releaseObject(xlWorkSheet);
                    // Save as the File
                    string filename = saveExcelDialog.FileName;
                    try {
                        xlApp.DisplayAlerts = false;
                        xlWorkBook.SaveAs(filename);
                    }
                    catch (Exception ex) {
                        MessageBox.Show("Can't overwrite file. Please save it as another name." + 
                            ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    MessageBox.Show("Excel File created!\n" + requests + 
                        " requests were made to the Riot API.", "Note");
                    xlApp.Visible = true;
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                }
                else { return; }
                #endregion
            }
            catch (Exception e) {
				MessageBox.Show("Error: " + e.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                releaseObject(xlApp);
                return;
			}
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
