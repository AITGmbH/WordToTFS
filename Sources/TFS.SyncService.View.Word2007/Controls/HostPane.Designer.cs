namespace TFS.SyncService.View.Word2007.Controls
{
    partial class HostPane<TControl>
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.WPFContentHost = new System.Windows.Forms.Integration.ElementHost();
            this.SuspendLayout();
            // 
            // WPFContentHost
            // 
            this.WPFContentHost.AutoSize = true;
            this.WPFContentHost.Dock = System.Windows.Forms.DockStyle.Fill;
            this.WPFContentHost.Location = new System.Drawing.Point(0, 0);
            this.WPFContentHost.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.WPFContentHost.Name = "WPFContentHost";
            this.WPFContentHost.Size = new System.Drawing.Size(308, 552);
            this.WPFContentHost.TabIndex = 0;
            this.WPFContentHost.Child = null;
            // 
            // HostPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.WPFContentHost);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "HostPane";
            this.Size = new System.Drawing.Size(308, 552);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Integration.ElementHost WPFContentHost;
    }
}
