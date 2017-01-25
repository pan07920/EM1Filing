namespace EM1Filing
{
    partial class MainEM1Filing
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainEM1Filing));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.eM1FinalFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.eM1TradingFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.eM1FinalFileToolStripMenuItem,
            this.eM1TradingFileToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(956, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // eM1FinalFileToolStripMenuItem
            // 
            this.eM1FinalFileToolStripMenuItem.Name = "eM1FinalFileToolStripMenuItem";
            this.eM1FinalFileToolStripMenuItem.Size = new System.Drawing.Size(91, 20);
            this.eM1FinalFileToolStripMenuItem.Text = "EM1 Final File";
            this.eM1FinalFileToolStripMenuItem.Click += new System.EventHandler(this.eM1FinalFileToolStripMenuItem_Click);
            // 
            // eM1TradingFileToolStripMenuItem
            // 
            this.eM1TradingFileToolStripMenuItem.Name = "eM1TradingFileToolStripMenuItem";
            this.eM1TradingFileToolStripMenuItem.Size = new System.Drawing.Size(106, 20);
            this.eM1TradingFileToolStripMenuItem.Text = "EM1 Trading File";
            this.eM1TradingFileToolStripMenuItem.Click += new System.EventHandler(this.eM1TradingFileToolStripMenuItem_Click);
            // 
            // MainEM1Filing
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(956, 593);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainEM1Filing";
            this.Text = "EM1 Filing Tools";
            this.Load += new System.EventHandler(this.EM1Filing_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem eM1FinalFileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem eM1TradingFileToolStripMenuItem;
    }
}

