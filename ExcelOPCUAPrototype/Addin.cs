using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Windows.Forms;
using NetOffice;
using NetOffice.Tools;
using Excel = NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using NetOffice.ExcelApi.Tools;
using Office = NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using NetOffice.OfficeApi.Tools;
using VBIDE = NetOffice.VBIDEApi;
using NetOffice.VBIDEApi.Enums;

namespace ExcelOPCUAPrototype
{
    // Int- Parameter in "COMAddin" Laden des Automation Addin:
    // 1 = Nicht automatisch laden
    // 2 = Bei Bedarf laden
    // 3 = Beim Start der Office Anwendung automatisch laden
    // 16 = Beim ersten Start automatisch laden, danach bei Bedarf

	[COMAddin("ExcelOPCUAPrototype", "Ein Durchstich um die Datenübertragung von SPS über OPC UA nach Excel zu zeigen", 3), ProgId("ExcelOPCUAPrototype.Addin"), Guid("A9C21BF0-75B6-4E29-B0C3-D077ED21FE5E")]
	[RegistryLocation(RegistrySaveLocation.CurrentUser), CustomUI("ExcelOPCUAPrototype.RibbonUI.xml"), CustomPane(typeof(MyTaskPane), "My TaskPane", true, PaneDockPosition.msoCTPDockPositionRight)]


    public class Addin : Excel.Tools.COMAddin
	{
		public Addin()
		{
			this.OnStartupComplete += new OnStartupCompleteEventHandler(Addin_OnStartupComplete);
			this.OnDisconnection += new OnDisconnectionEventHandler(Addin_OnDisconnection);
		}

		internal Office.IRibbonUI RibbonUI { get; private set; }

		private void Addin_OnStartupComplete(ref Array custom)
		{
			Console.WriteLine("Addin started in Excel Version {0}", Application.Version);
			CreateUserInterface();
		}

		private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
		{
			RemoveUserInterface();
		}

        public void AboutButton_Click(Office.IRibbonControl control)
        {
			MessageBox.Show(String.Format("ExcelOPCUAPrototype Version {0}", this.GetType().Assembly.GetName().Version),
				"About ExcelOPCUAPrototype", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

		public bool TooglePaneVisibleButton_GetPressed(Office.IRibbonControl control)
		{
			return TaskPanes.Count > 0 ? TaskPanes[0].Visible : false;
		}

		public void TooglePaneVisibleButton_Click(Office.IRibbonControl control, bool pressed)
		{
			if(TaskPanes.Count > 0)
				TaskPanes[0].Visible = pressed;
		}

		public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
			RibbonUI = ribbonUI;
        }

		private void CreateUserInterface()
		{
		}

		private void RemoveUserInterface()
		{
		}

		protected override void OnError(ErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "ExcelOPCUAPrototype");
		}

		[RegisterErrorHandler]
		public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, System.Exception exception)
		{
			MessageBox.Show("An error occurend in " + methodKind.ToString(), "ExcelOPCUAPrototype");
		}
    }
}

