using System;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;


namespace ExcelOPCUAPrototype
{
    public partial class MyTaskPane : UserControl , Excel.Tools.ITaskPane
    {
		#region Ctor
        
		public MyTaskPane()
        {
            InitializeComponent();
        }

		#endregion

		#region Properties
		
		private Addin ParentAddin { get; set; }

		#endregion
		
        #region ITaskpane

        public void OnConnection(Excel.Application application, Office._CustomTaskPane definition, object[] customArguments)
        {
			if(customArguments.Length > 0)
				ParentAddin = customArguments[0] as Addin;
        }

        public void OnDisconnection()
        {

        }

        public void OnDockPositionChanged(MsoCTPDockPosition position)
        {
            
        }

        public void OnVisibleStateChanged(bool visible)
        {
			if(null != ParentAddin && null != ParentAddin.RibbonUI)
				ParentAddin.RibbonUI.InvalidateControl("tooglePaneVisibleButton");
        }

        #endregion

        private void OpenOPCUAMain_Click(object sender, EventArgs e)
        {
            new MainFormOPCUA();
        }
    }
}
