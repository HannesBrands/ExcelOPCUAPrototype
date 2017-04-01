namespace ExcelOPCUAPrototype
{
    partial class MyTaskPane
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        private void InitializeComponent()
        {
            this.OpenOPCUAMain = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenOPCUAMain
            // 
            this.OpenOPCUAMain.Location = new System.Drawing.Point(125, 188);
            this.OpenOPCUAMain.Name = "OpenOPCUAMain";
            this.OpenOPCUAMain.Size = new System.Drawing.Size(140, 23);
            this.OpenOPCUAMain.TabIndex = 0;
            this.OpenOPCUAMain.Text = "OpenOPCUAMain";
            this.OpenOPCUAMain.UseVisualStyleBackColor = true;
            this.OpenOPCUAMain.Click += new System.EventHandler(this.OpenOPCUAMain_Click);
            // 
            // MyTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.OpenOPCUAMain);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "MyTaskPane";
            this.Size = new System.Drawing.Size(400, 747);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button OpenOPCUAMain;
    }
}
