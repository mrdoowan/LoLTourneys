using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using System.Net;
using System.Threading;
using System.Windows.Forms;

namespace LoLBalancing
{
	public class RiotJson
	{
		private string region;
		private string APIKey;

		// Default Constructor
		public RiotJson(string region_, string APIKey_) {
			region = region_;
			APIKey = APIKey_;
		}

		// Based on the Match History ID, retrieve Data for it
		public JObject getMatchJson(string ID) {
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
							return null;
						}
						Thread.Sleep(1000);
						continue;
					}
				}
			}
			return JObject.Parse(MatchJson);
		}

		// Retrieve API information for Champion data
		public JObject getChampJson() {
			string ChampJson = "";
			using (var WC = new WebClient()) {
				ChampJson = WC.DownloadString("https://global.api.pvp.net/api/lol/static-data/" + region +
					"/v1.2/champion?locale=en_US&dataById=true&api_key=" + APIKey);
			}
			return JObject.Parse(ChampJson);
		}
	}
}
