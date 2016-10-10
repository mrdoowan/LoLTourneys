namespace LoLBalancing
{
	partial class StatsGen
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StatsGen));
			this.label_Msg = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// label_Msg
			// 
			this.label_Msg.Location = new System.Drawing.Point(12, 9);
			this.label_Msg.Name = "label_Msg";
			this.label_Msg.Size = new System.Drawing.Size(251, 124);
			this.label_Msg.TabIndex = 0;
			this.label_Msg.Text = "Your stats are currently being generated...";
			this.label_Msg.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// StatsGen
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(275, 142);
			this.Controls.Add(this.label_Msg);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.Name = "StatsGen";
			this.Text = "Stats Generating";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Label label_Msg;
	}
}