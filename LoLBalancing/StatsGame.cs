using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoLBalancing {
    public class StatsGame {

        public List<StatsPlayer> Players;
        public long gameID { get; set; }
        public int redTeamNum { get; set; }
        public int blueTeamNum { get; set; }

        public class StatsPlayer {
            public string champ { get; set; }
            public string role { get; set; }
            public string summoner { get; set; }

            // Default Constructor
            public StatsPlayer(string champ_, string role_, string summoner_ = "") {
                champ = champ_;
                role = role_;
                summoner = summoner_;
            }
        }

        // Default Constructor
        public StatsGame(long ID_, int red_, int blue_) {
            Players = new List<StatsPlayer>();
            // We're always going to assume that it contains 2 * NUM_PLAYERS
            gameID = ID_;
            redTeamNum = red_;
            blueTeamNum = blue_;
        }

        // Two varations of the same Function
        public void addPlayer(string champName, string role) {
            Players.Add(new StatsPlayer(champName, role));
        }

        public void addPlayer(string champName, string role, string summoner) {
            Players.Add(new StatsPlayer(champName, role, summoner));
        }
    }
}
