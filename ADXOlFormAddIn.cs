using System;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OutlookEvents
{
	/// <summary>
	/// Summary description for ADXOlForm1.
	/// </summary>
	public class ADXOlFormAddIn : AddinExpress.OL.ADXOlForm
	{
		private System.ComponentModel.IContainer components = null;
		private ContextMenuStrip contextMenuStrip1;
		private ToolStripMenuItem cleatToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator1;
		private ToolStripMenuItem selectAllToolStripMenuItem;
		private ToolStripSeparator toolStripSeparator2;
		private ToolStripMenuItem copyToolStripMenuItem;
		private ContextMenuStrip contextMenuStrip2;
		private ToolStripMenuItem expandAllToolStripMenuItem;
		private ToolStripMenuItem collapseAllToolStripMenuItem;
		private ToolStripMenuItem saveAsToolStripMenuItem;
		private SaveFileDialog saveFileDialog1;
		private Hashtable HelpData = new Hashtable();
		private ToolStrip toolStrip1;
		private ToolStripButton toolStripButtonSelectAll;
		private ToolStripButton toolStripButtonClear;
		private ToolStripButton toolStripButtonCopy;
		private ToolStripButton toolStripButtonSaveAs;
		private ToolStripButton toolStripButtonWriteLogToFile;
		private ToolStripButton toolStripButtonStartStopLog;
		private ImageList imageList1;
		private ToolStripSeparator toolStripSeparator3;
		private Panel panel1;
		public TreeView treeView1;
		private Splitter splitter1;
		private Panel panel2;
		private TextBox textBox2;
		public TextBox textBox1;
		private bool saveChangeTreeWiew = false;
		private ToolStripSeparator toolStripSeparator4;
		private ToolStripLabel toolStripLabel1;

		private AddinModule CurrentModule = null;

		public ADXOlFormAddIn()
		{
			// This call is required by the Windows Form Designer.
			InitializeComponent();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if (disposing)
			{
				if (components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}

		#region Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXOlFormAddIn));
			System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("AddinBeginShutdown");
			System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("AddinFinalize");
			System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("AddinInitialize");
			System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("AddinStartupComplete");
			System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("AfterUninstallControls");
			System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("BeforeUninstallControls");
			System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("OfficeColorSchemeChanged");
			System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("OnError");
			System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("OnKeyDown");
			System.Windows.Forms.TreeNode treeNode10 = new System.Windows.Forms.TreeNode("OnRibbonBeforeCreate");
			System.Windows.Forms.TreeNode treeNode11 = new System.Windows.Forms.TreeNode("OnRibbonBeforeLoad");
			System.Windows.Forms.TreeNode treeNode12 = new System.Windows.Forms.TreeNode("OnRibbonLoaded");
			System.Windows.Forms.TreeNode treeNode13 = new System.Windows.Forms.TreeNode("OnSendMessage");
			System.Windows.Forms.TreeNode treeNode14 = new System.Windows.Forms.TreeNode("OnTaskPaneAfterCreate");
			System.Windows.Forms.TreeNode treeNode15 = new System.Windows.Forms.TreeNode("OnTaskPaneAfterShow");
			System.Windows.Forms.TreeNode treeNode16 = new System.Windows.Forms.TreeNode("OnTaskPaneBeforeCreate");
			System.Windows.Forms.TreeNode treeNode17 = new System.Windows.Forms.TreeNode("OnTaskPaneBeforeDestroy");
			System.Windows.Forms.TreeNode treeNode18 = new System.Windows.Forms.TreeNode("OnTaskPaneBeforeShow");
			System.Windows.Forms.TreeNode treeNode19 = new System.Windows.Forms.TreeNode("Add-in Module Events", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4,
            treeNode5,
            treeNode6,
            treeNode7,
            treeNode8,
            treeNode9,
            treeNode10,
            treeNode11,
            treeNode12,
            treeNode13,
            treeNode14,
            treeNode15,
            treeNode16,
            treeNode17,
            treeNode18});
			System.Windows.Forms.TreeNode treeNode20 = new System.Windows.Forms.TreeNode("NewExplorer");
			System.Windows.Forms.TreeNode treeNode21 = new System.Windows.Forms.TreeNode("ExplorerActivate");
			System.Windows.Forms.TreeNode treeNode22 = new System.Windows.Forms.TreeNode("ExplorerAddCommandBars");
			System.Windows.Forms.TreeNode treeNode23 = new System.Windows.Forms.TreeNode("ExplorerBeforeFolderSwitch");
			System.Windows.Forms.TreeNode treeNode24 = new System.Windows.Forms.TreeNode("ExplorerBeforeItemCopy");
			System.Windows.Forms.TreeNode treeNode25 = new System.Windows.Forms.TreeNode("ExplorerBeforeItemCut");
			System.Windows.Forms.TreeNode treeNode26 = new System.Windows.Forms.TreeNode("ExplorerBeforeItemPaste");
			System.Windows.Forms.TreeNode treeNode27 = new System.Windows.Forms.TreeNode("ExplorerBeforeMaximize");
			System.Windows.Forms.TreeNode treeNode28 = new System.Windows.Forms.TreeNode("ExplorerBeforeMinimize");
			System.Windows.Forms.TreeNode treeNode29 = new System.Windows.Forms.TreeNode("ExplorerBeforeMove");
			System.Windows.Forms.TreeNode treeNode30 = new System.Windows.Forms.TreeNode("ExplorerBeforeSize");
			System.Windows.Forms.TreeNode treeNode31 = new System.Windows.Forms.TreeNode("ExplorerBeforeViewSwitch");
			System.Windows.Forms.TreeNode treeNode32 = new System.Windows.Forms.TreeNode("ExplorerClose");
			System.Windows.Forms.TreeNode treeNode33 = new System.Windows.Forms.TreeNode("ExplorerDeactivate");
			System.Windows.Forms.TreeNode treeNode34 = new System.Windows.Forms.TreeNode("ExplorerFolderSwitch");
			System.Windows.Forms.TreeNode treeNode35 = new System.Windows.Forms.TreeNode("ExplorerSelectionChange");
			System.Windows.Forms.TreeNode treeNode36 = new System.Windows.Forms.TreeNode("ExplorerViewSwitch");
			System.Windows.Forms.TreeNode treeNode37 = new System.Windows.Forms.TreeNode("ExplorerAttachmentSelectionChange");
            System.Windows.Forms.TreeNode treeNode201 = new System.Windows.Forms.TreeNode("ExplorerInlineResponse");
            System.Windows.Forms.TreeNode treeNode202 = new System.Windows.Forms.TreeNode("ExplorerInlineResponseClose");
			System.Windows.Forms.TreeNode treeNode38 = new System.Windows.Forms.TreeNode("Explorer(s) Events", new System.Windows.Forms.TreeNode[] {
            treeNode20,
            treeNode21,
            treeNode22,
            treeNode23,
            treeNode24,
            treeNode25,
            treeNode26,
            treeNode27,
            treeNode28,
            treeNode29,
            treeNode30,
            treeNode31,
            treeNode32,
            treeNode33,
            treeNode34,
            treeNode35,
            treeNode36,
            treeNode37,
            treeNode201,
            treeNode202});
			System.Windows.Forms.TreeNode treeNode39 = new System.Windows.Forms.TreeNode("NewInspector");
			System.Windows.Forms.TreeNode treeNode40 = new System.Windows.Forms.TreeNode("InspectorActivate");
			System.Windows.Forms.TreeNode treeNode41 = new System.Windows.Forms.TreeNode("InspectorAddCommandBars");
			System.Windows.Forms.TreeNode treeNode42 = new System.Windows.Forms.TreeNode("InspectorBeforeMaximize");
			System.Windows.Forms.TreeNode treeNode43 = new System.Windows.Forms.TreeNode("InspectorBeforeMinimize");
			System.Windows.Forms.TreeNode treeNode44 = new System.Windows.Forms.TreeNode("InspectorBeforeMove");
			System.Windows.Forms.TreeNode treeNode45 = new System.Windows.Forms.TreeNode("InspectorBeforeSize");
			System.Windows.Forms.TreeNode treeNode46 = new System.Windows.Forms.TreeNode("InspectorClose");
			System.Windows.Forms.TreeNode treeNode47 = new System.Windows.Forms.TreeNode("InspectorDeactivate");
			System.Windows.Forms.TreeNode treeNode48 = new System.Windows.Forms.TreeNode("PageChange");
			System.Windows.Forms.TreeNode treeNode49 = new System.Windows.Forms.TreeNode("InspectorAttachmentSelectionChange");
			System.Windows.Forms.TreeNode treeNode50 = new System.Windows.Forms.TreeNode("Inspector(s) Events", new System.Windows.Forms.TreeNode[] {
            treeNode39,
            treeNode40,
            treeNode41,
            treeNode42,
            treeNode43,
            treeNode44,
            treeNode45,
            treeNode46,
            treeNode47,
            treeNode48,
            treeNode49});
			System.Windows.Forms.TreeNode treeNode51 = new System.Windows.Forms.TreeNode("BeforeReminderShow");
			System.Windows.Forms.TreeNode treeNode52 = new System.Windows.Forms.TreeNode("ReminderAdd");
			System.Windows.Forms.TreeNode treeNode53 = new System.Windows.Forms.TreeNode("ReminderChange");
			System.Windows.Forms.TreeNode treeNode54 = new System.Windows.Forms.TreeNode("ReminderFire");
			System.Windows.Forms.TreeNode treeNode55 = new System.Windows.Forms.TreeNode("ReminderRemove");
			System.Windows.Forms.TreeNode treeNode56 = new System.Windows.Forms.TreeNode("Snooze");
			System.Windows.Forms.TreeNode treeNode57 = new System.Windows.Forms.TreeNode("Reminder Events", new System.Windows.Forms.TreeNode[] {
            treeNode51,
            treeNode52,
            treeNode53,
            treeNode54,
            treeNode55,
            treeNode56});
			System.Windows.Forms.TreeNode treeNode58 = new System.Windows.Forms.TreeNode("AdvancedSearchComplete");
			System.Windows.Forms.TreeNode treeNode59 = new System.Windows.Forms.TreeNode("AdvancedSearchStopped");
			System.Windows.Forms.TreeNode treeNode60 = new System.Windows.Forms.TreeNode("ItemSend");
			System.Windows.Forms.TreeNode treeNode61 = new System.Windows.Forms.TreeNode("NewMail");
			System.Windows.Forms.TreeNode treeNode62 = new System.Windows.Forms.TreeNode("NewMailEx");
			System.Windows.Forms.TreeNode treeNode63 = new System.Windows.Forms.TreeNode("ItemLoad");
			System.Windows.Forms.TreeNode treeNode64 = new System.Windows.Forms.TreeNode("MAPILogonComplete");
			System.Windows.Forms.TreeNode treeNode65 = new System.Windows.Forms.TreeNode("Startup");
			System.Windows.Forms.TreeNode treeNode66 = new System.Windows.Forms.TreeNode("AutoDiscoverComplete");
			System.Windows.Forms.TreeNode treeNode67 = new System.Windows.Forms.TreeNode("Quit");
			System.Windows.Forms.TreeNode treeNode68 = new System.Windows.Forms.TreeNode("OptionPagesAdd");
			System.Windows.Forms.TreeNode treeNode69 = new System.Windows.Forms.TreeNode("BeforeOptionPageAdd");
			System.Windows.Forms.TreeNode treeNode70 = new System.Windows.Forms.TreeNode("Reminder");
			System.Windows.Forms.TreeNode treeNode71 = new System.Windows.Forms.TreeNode("BeforeFolderSharingDialog");
			System.Windows.Forms.TreeNode treeNode72 = new System.Windows.Forms.TreeNode("ContextMenuClose");
			System.Windows.Forms.TreeNode treeNode73 = new System.Windows.Forms.TreeNode("ShortcutContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode74 = new System.Windows.Forms.TreeNode("ViewContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode75 = new System.Windows.Forms.TreeNode("StoreContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode76 = new System.Windows.Forms.TreeNode("AttachmentContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode77 = new System.Windows.Forms.TreeNode("FolderContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode78 = new System.Windows.Forms.TreeNode("ItemContextMenuDisplay");
			System.Windows.Forms.TreeNode treeNode79 = new System.Windows.Forms.TreeNode("Application Events", new System.Windows.Forms.TreeNode[] {
            treeNode58,
            treeNode59,
            treeNode60,
            treeNode61,
            treeNode62,
            treeNode63,
            treeNode64,
            treeNode65,
            treeNode66,
            treeNode67,
            treeNode68,
            treeNode69,
            treeNode70,
            treeNode71,
            treeNode72,
            treeNode73,
            treeNode74,
            treeNode75,
            treeNode76,
            treeNode77,
            treeNode78});
			System.Windows.Forms.TreeNode treeNode80 = new System.Windows.Forms.TreeNode("NamespaceBeforeOptionPageAdd");
			System.Windows.Forms.TreeNode treeNode81 = new System.Windows.Forms.TreeNode("NamespaceOptionPagesAdd");
			System.Windows.Forms.TreeNode treeNode82 = new System.Windows.Forms.TreeNode("Namespace Events", new System.Windows.Forms.TreeNode[] {
            treeNode80,
            treeNode81});
			System.Windows.Forms.TreeNode treeNode83 = new System.Windows.Forms.TreeNode("SyncEnd");
			System.Windows.Forms.TreeNode treeNode84 = new System.Windows.Forms.TreeNode("SyncError");
			System.Windows.Forms.TreeNode treeNode85 = new System.Windows.Forms.TreeNode("SyncProgress");
			System.Windows.Forms.TreeNode treeNode86 = new System.Windows.Forms.TreeNode("SyncStart");
			System.Windows.Forms.TreeNode treeNode87 = new System.Windows.Forms.TreeNode("SyncObject Events", new System.Windows.Forms.TreeNode[] {
            treeNode83,
            treeNode84,
            treeNode85,
            treeNode86});
			System.Windows.Forms.TreeNode treeNode88 = new System.Windows.Forms.TreeNode("OnBeforeFormRegionShow");
			System.Windows.Forms.TreeNode treeNode89 = new System.Windows.Forms.TreeNode("OnGetFormRegionIcon");
			System.Windows.Forms.TreeNode treeNode90 = new System.Windows.Forms.TreeNode("OnGetFormRegionManifest");
			System.Windows.Forms.TreeNode treeNode91 = new System.Windows.Forms.TreeNode("OnGetFormRegionStorage");
			System.Windows.Forms.TreeNode treeNode92 = new System.Windows.Forms.TreeNode("Region Events", new System.Windows.Forms.TreeNode[] {
            treeNode88,
            treeNode89,
            treeNode90,
            treeNode91});
			System.Windows.Forms.TreeNode treeNode93 = new System.Windows.Forms.TreeNode("CommandBarsUpdate");
			System.Windows.Forms.TreeNode treeNode94 = new System.Windows.Forms.TreeNode("Outlook Application Events", new System.Windows.Forms.TreeNode[] {
            treeNode38,
            treeNode50,
            treeNode57,
            treeNode79,
            treeNode82,
            treeNode87,
            treeNode92,
            treeNode93});
			System.Windows.Forms.TreeNode treeNode95 = new System.Windows.Forms.TreeNode("FolderAdd");
			System.Windows.Forms.TreeNode treeNode96 = new System.Windows.Forms.TreeNode("FolderChange");
			System.Windows.Forms.TreeNode treeNode97 = new System.Windows.Forms.TreeNode("FolderRemove");
			System.Windows.Forms.TreeNode treeNode98 = new System.Windows.Forms.TreeNode("Folders Events", new System.Windows.Forms.TreeNode[] {
            treeNode95,
            treeNode96,
            treeNode97});
			System.Windows.Forms.TreeNode treeNode99 = new System.Windows.Forms.TreeNode("ItemAdd");
			System.Windows.Forms.TreeNode treeNode100 = new System.Windows.Forms.TreeNode("ItemChange");
			System.Windows.Forms.TreeNode treeNode101 = new System.Windows.Forms.TreeNode("ItemRemove");
			System.Windows.Forms.TreeNode treeNode102 = new System.Windows.Forms.TreeNode("BeforeFolderMove");
			System.Windows.Forms.TreeNode treeNode103 = new System.Windows.Forms.TreeNode("BeforeItemMove");
			System.Windows.Forms.TreeNode treeNode104 = new System.Windows.Forms.TreeNode("Items Events", new System.Windows.Forms.TreeNode[] {
            treeNode99,
            treeNode100,
            treeNode101,
            treeNode102,
            treeNode103});
			System.Windows.Forms.TreeNode treeNode105 = new System.Windows.Forms.TreeNode("AttachmentAdd");
			System.Windows.Forms.TreeNode treeNode106 = new System.Windows.Forms.TreeNode("AttachmentRead");
			System.Windows.Forms.TreeNode treeNode107 = new System.Windows.Forms.TreeNode("BeforeAttachmentSave");
			System.Windows.Forms.TreeNode treeNode108 = new System.Windows.Forms.TreeNode("BeforeCheckNames");
			System.Windows.Forms.TreeNode treeNode109 = new System.Windows.Forms.TreeNode("Close");
			System.Windows.Forms.TreeNode treeNode110 = new System.Windows.Forms.TreeNode("CustomAction");
			System.Windows.Forms.TreeNode treeNode111 = new System.Windows.Forms.TreeNode("CustomPropertyChange");
			System.Windows.Forms.TreeNode treeNode112 = new System.Windows.Forms.TreeNode("Forward");
			System.Windows.Forms.TreeNode treeNode113 = new System.Windows.Forms.TreeNode("Open");
			System.Windows.Forms.TreeNode treeNode114 = new System.Windows.Forms.TreeNode("PropertyChange");
			System.Windows.Forms.TreeNode treeNode115 = new System.Windows.Forms.TreeNode("Read");
			System.Windows.Forms.TreeNode treeNode116 = new System.Windows.Forms.TreeNode("Reply");
			System.Windows.Forms.TreeNode treeNode117 = new System.Windows.Forms.TreeNode("ReplyAll");
			System.Windows.Forms.TreeNode treeNode118 = new System.Windows.Forms.TreeNode("Send");
			System.Windows.Forms.TreeNode treeNode119 = new System.Windows.Forms.TreeNode("Write");
			System.Windows.Forms.TreeNode treeNode120 = new System.Windows.Forms.TreeNode("BeforeDelete");
			System.Windows.Forms.TreeNode treeNode121 = new System.Windows.Forms.TreeNode("AttachmentRemove");
			System.Windows.Forms.TreeNode treeNode122 = new System.Windows.Forms.TreeNode("BeforeAttachmentAdd");
			System.Windows.Forms.TreeNode treeNode123 = new System.Windows.Forms.TreeNode("BeforeAttachmentPreview");
			System.Windows.Forms.TreeNode treeNode124 = new System.Windows.Forms.TreeNode("BeforeAttachmentRead");
			System.Windows.Forms.TreeNode treeNode125 = new System.Windows.Forms.TreeNode("BeforeAttachmentWriteToTempFile");
			System.Windows.Forms.TreeNode treeNode126 = new System.Windows.Forms.TreeNode("Unload");
			System.Windows.Forms.TreeNode treeNode127 = new System.Windows.Forms.TreeNode("BeforeAutoSave");
			System.Windows.Forms.TreeNode treeNode128 = new System.Windows.Forms.TreeNode("AfterWrite");
			System.Windows.Forms.TreeNode treeNode129 = new System.Windows.Forms.TreeNode("BeforeRead");
            System.Windows.Forms.TreeNode treeNode200 = new System.Windows.Forms.TreeNode("ReadComplete");
			System.Windows.Forms.TreeNode treeNode130 = new System.Windows.Forms.TreeNode("Selected Item Events", new System.Windows.Forms.TreeNode[] {
            treeNode105,
            treeNode106,
            treeNode107,
            treeNode108,
            treeNode109,
            treeNode110,
            treeNode111,
            treeNode112,
            treeNode113,
            treeNode114,
            treeNode115,
            treeNode116,
            treeNode117,
            treeNode118,
            treeNode119,
            treeNode120,
            treeNode121,
            treeNode122,
            treeNode123,
            treeNode124,
            treeNode125,
            treeNode126,
            treeNode127,
            treeNode128,
            treeNode129,
            treeNode200});
			System.Windows.Forms.TreeNode treeNode131 = new System.Windows.Forms.TreeNode("ADXAfterAccessProtectedObject");
			System.Windows.Forms.TreeNode treeNode132 = new System.Windows.Forms.TreeNode("ADXBeforeAccessProtectedObject");
			System.Windows.Forms.TreeNode treeNode133 = new System.Windows.Forms.TreeNode("ADXBeforeFolderSwitch");
			System.Windows.Forms.TreeNode treeNode134 = new System.Windows.Forms.TreeNode("ADXBeforeFolderSwitchEx");
			System.Windows.Forms.TreeNode treeNode135 = new System.Windows.Forms.TreeNode("ADXBeforeFormInstanceCreate");
			System.Windows.Forms.TreeNode treeNode136 = new System.Windows.Forms.TreeNode("ADXFolderSwitch");
			System.Windows.Forms.TreeNode treeNode137 = new System.Windows.Forms.TreeNode("ADXFolderSwitchEx");
			System.Windows.Forms.TreeNode treeNode138 = new System.Windows.Forms.TreeNode("ADXNavigationPaneHide");
			System.Windows.Forms.TreeNode treeNode139 = new System.Windows.Forms.TreeNode("ADXNavigationPaneMinimize");
			System.Windows.Forms.TreeNode treeNode140 = new System.Windows.Forms.TreeNode("ADXNavigationPaneShow");
			System.Windows.Forms.TreeNode treeNode141 = new System.Windows.Forms.TreeNode("ADXNewInspector");
			System.Windows.Forms.TreeNode treeNode142 = new System.Windows.Forms.TreeNode("ADXReadingPaneHide");
			System.Windows.Forms.TreeNode treeNode143 = new System.Windows.Forms.TreeNode("ADXReadingPaneMove");
			System.Windows.Forms.TreeNode treeNode144 = new System.Windows.Forms.TreeNode("ADXReadingPaneShow");
			System.Windows.Forms.TreeNode treeNode145 = new System.Windows.Forms.TreeNode("ADXTodoBarHide");
			System.Windows.Forms.TreeNode treeNode146 = new System.Windows.Forms.TreeNode("ADXTodoBarMinimize");
			System.Windows.Forms.TreeNode treeNode147 = new System.Windows.Forms.TreeNode("ADXTodoBarShow");
			System.Windows.Forms.TreeNode treeNode148 = new System.Windows.Forms.TreeNode("OnError");
			System.Windows.Forms.TreeNode treeNode149 = new System.Windows.Forms.TreeNode("OnInitialize");
			System.Windows.Forms.TreeNode treeNode150 = new System.Windows.Forms.TreeNode("Add-in Express FormsManager Events", new System.Windows.Forms.TreeNode[] {
            treeNode131,
            treeNode132,
            treeNode133,
            treeNode134,
            treeNode135,
            treeNode136,
            treeNode137,
            treeNode138,
            treeNode139,
            treeNode140,
            treeNode141,
            treeNode142,
            treeNode143,
            treeNode144,
            treeNode145,
            treeNode146,
            treeNode147,
            treeNode148,
            treeNode149});
			this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.expandAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.collapseAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
			this.cleatToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.selectAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveAsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.toolStrip1 = new System.Windows.Forms.ToolStrip();
			this.toolStripButtonWriteLogToFile = new System.Windows.Forms.ToolStripButton();
			this.toolStripButtonStartStopLog = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripButtonSelectAll = new System.Windows.Forms.ToolStripButton();
			this.toolStripButtonClear = new System.Windows.Forms.ToolStripButton();
			this.toolStripButtonCopy = new System.Windows.Forms.ToolStripButton();
			this.toolStripButtonSaveAs = new System.Windows.Forms.ToolStripButton();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
			this.imageList1 = new System.Windows.Forms.ImageList(this.components);
			this.panel1 = new System.Windows.Forms.Panel();
			this.treeView1 = new System.Windows.Forms.TreeView();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.panel2 = new System.Windows.Forms.Panel();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.contextMenuStrip2.SuspendLayout();
			this.contextMenuStrip1.SuspendLayout();
			this.toolStrip1.SuspendLayout();
			this.panel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// contextMenuStrip2
			// 
			this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.expandAllToolStripMenuItem,
            this.collapseAllToolStripMenuItem});
			this.contextMenuStrip2.Name = "contextMenuStrip2";
			this.contextMenuStrip2.Size = new System.Drawing.Size(137, 48);
			// 
			// expandAllToolStripMenuItem
			// 
			this.expandAllToolStripMenuItem.Name = "expandAllToolStripMenuItem";
			this.expandAllToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
			this.expandAllToolStripMenuItem.Text = " Expand All";
			this.expandAllToolStripMenuItem.Click += new System.EventHandler(this.expandAllToolStripMenuItem_Click);
			// 
			// collapseAllToolStripMenuItem
			// 
			this.collapseAllToolStripMenuItem.Name = "collapseAllToolStripMenuItem";
			this.collapseAllToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
			this.collapseAllToolStripMenuItem.Text = "Collapse All";
			this.collapseAllToolStripMenuItem.Click += new System.EventHandler(this.collapseAllToolStripMenuItem_Click);
			// 
			// contextMenuStrip1
			// 
			this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cleatToolStripMenuItem,
            this.toolStripSeparator1,
            this.selectAllToolStripMenuItem,
            this.toolStripSeparator2,
            this.copyToolStripMenuItem,
            this.saveAsToolStripMenuItem});
			this.contextMenuStrip1.Name = "contextMenuStrip1";
			this.contextMenuStrip1.Size = new System.Drawing.Size(123, 104);
			// 
			// cleatToolStripMenuItem
			// 
			this.cleatToolStripMenuItem.Name = "cleatToolStripMenuItem";
			this.cleatToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
			this.cleatToolStripMenuItem.Text = "Clear";
			this.cleatToolStripMenuItem.Click += new System.EventHandler(this.cleatToolStripMenuItem_Click);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(119, 6);
			// 
			// selectAllToolStripMenuItem
			// 
			this.selectAllToolStripMenuItem.Name = "selectAllToolStripMenuItem";
			this.selectAllToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
			this.selectAllToolStripMenuItem.Text = "Select All";
			this.selectAllToolStripMenuItem.Click += new System.EventHandler(this.selectAllToolStripMenuItem_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(119, 6);
			// 
			// copyToolStripMenuItem
			// 
			this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
			this.copyToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
			this.copyToolStripMenuItem.Text = "Copy";
			this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
			// 
			// saveAsToolStripMenuItem
			// 
			this.saveAsToolStripMenuItem.Name = "saveAsToolStripMenuItem";
			this.saveAsToolStripMenuItem.Size = new System.Drawing.Size(122, 22);
			this.saveAsToolStripMenuItem.Text = "Save As";
			this.saveAsToolStripMenuItem.Click += new System.EventHandler(this.saveAsToolStripMenuItem_Click);
			// 
			// saveFileDialog1
			// 
			this.saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
			this.saveFileDialog1.InitialDirectory = "c:\\";
			this.saveFileDialog1.RestoreDirectory = true;
			// 
			// toolStrip1
			// 
			this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonWriteLogToFile,
            this.toolStripButtonStartStopLog,
            this.toolStripSeparator3,
            this.toolStripButtonSelectAll,
            this.toolStripButtonClear,
            this.toolStripButtonCopy,
            this.toolStripButtonSaveAs,
            this.toolStripSeparator4,
            this.toolStripLabel1});
			this.toolStrip1.Location = new System.Drawing.Point(0, 0);
			this.toolStrip1.Name = "toolStrip1";
			this.toolStrip1.Size = new System.Drawing.Size(1034, 25);
			this.toolStrip1.TabIndex = 5;
			this.toolStrip1.Text = "toolStrip1";
			// 
			// toolStripButtonWriteLogToFile
			// 
			this.toolStripButtonWriteLogToFile.CheckOnClick = true;
			this.toolStripButtonWriteLogToFile.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonWriteLogToFile.Image")));
			this.toolStripButtonWriteLogToFile.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonWriteLogToFile.Name = "toolStripButtonWriteLogToFile";
			this.toolStripButtonWriteLogToFile.Size = new System.Drawing.Size(108, 22);
			this.toolStripButtonWriteLogToFile.Text = "Write log to file";
			this.toolStripButtonWriteLogToFile.CheckedChanged += new System.EventHandler(this.toolStripButtonWriteLogToFile_CheckedChanged);
			// 
			// toolStripButtonStartStopLog
			// 
			this.toolStripButtonStartStopLog.Checked = true;
			this.toolStripButtonStartStopLog.CheckOnClick = true;
			this.toolStripButtonStartStopLog.CheckState = System.Windows.Forms.CheckState.Checked;
			this.toolStripButtonStartStopLog.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.toolStripButtonStartStopLog.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonStartStopLog.Name = "toolStripButtonStartStopLog";
			this.toolStripButtonStartStopLog.Size = new System.Drawing.Size(81, 22);
			this.toolStripButtonStartStopLog.Text = "Log is started";
			this.toolStripButtonStartStopLog.ToolTipText = "Start/stop log";
			this.toolStripButtonStartStopLog.Click += new System.EventHandler(this.toolStripButtonStartStopLog_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(6, 25);
			// 
			// toolStripButtonSelectAll
			// 
			this.toolStripButtonSelectAll.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
			this.toolStripButtonSelectAll.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonSelectAll.Name = "toolStripButtonSelectAll";
			this.toolStripButtonSelectAll.Size = new System.Drawing.Size(59, 22);
			this.toolStripButtonSelectAll.Text = "Select All";
			this.toolStripButtonSelectAll.TextImageRelation = System.Windows.Forms.TextImageRelation.Overlay;
			this.toolStripButtonSelectAll.Click += new System.EventHandler(this.toolStripButtonSelectAll_Click);
			// 
			// toolStripButtonClear
			// 
			this.toolStripButtonClear.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonClear.Image")));
			this.toolStripButtonClear.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonClear.Name = "toolStripButtonClear";
			this.toolStripButtonClear.Size = new System.Drawing.Size(54, 22);
			this.toolStripButtonClear.Text = "Clear";
			this.toolStripButtonClear.ToolTipText = "Clear log";
			this.toolStripButtonClear.Click += new System.EventHandler(this.toolStripButtonClear_Click);
			// 
			// toolStripButtonCopy
			// 
			this.toolStripButtonCopy.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonCopy.Image")));
			this.toolStripButtonCopy.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonCopy.Name = "toolStripButtonCopy";
			this.toolStripButtonCopy.Size = new System.Drawing.Size(55, 22);
			this.toolStripButtonCopy.Text = "Copy";
			this.toolStripButtonCopy.ToolTipText = "Copy selected/all text";
			this.toolStripButtonCopy.Click += new System.EventHandler(this.toolStripButtonCopy_Click);
			// 
			// toolStripButtonSaveAs
			// 
			this.toolStripButtonSaveAs.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonSaveAs.Image")));
			this.toolStripButtonSaveAs.ImageTransparentColor = System.Drawing.Color.Magenta;
			this.toolStripButtonSaveAs.Name = "toolStripButtonSaveAs";
			this.toolStripButtonSaveAs.Size = new System.Drawing.Size(67, 22);
			this.toolStripButtonSaveAs.Text = "Save As";
			this.toolStripButtonSaveAs.ToolTipText = "Save selected/all text to file";
			this.toolStripButtonSaveAs.Click += new System.EventHandler(this.toolStripButtonSaveAs_Click);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(6, 25);
			// 
			// toolStripLabel1
			// 
			this.toolStripLabel1.IsLink = true;
			this.toolStripLabel1.Name = "toolStripLabel1";
			this.toolStripLabel1.Size = new System.Drawing.Size(150, 22);
			this.toolStripLabel1.Text = "Powered by Add-in Express";
			this.toolStripLabel1.Click += new System.EventHandler(this.toolStripLabel1_Click);
			// 
			// imageList1
			// 
			this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
			this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
			this.imageList1.Images.SetKeyName(0, "delete_x_16_h.bmp");
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.treeView1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
			this.panel1.Location = new System.Drawing.Point(0, 25);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(230, 332);
			this.panel1.TabIndex = 6;
			// 
			// treeView1
			// 
			this.treeView1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.treeView1.CheckBoxes = true;
			this.treeView1.ContextMenuStrip = this.contextMenuStrip2;
			this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.treeView1.Location = new System.Drawing.Point(0, 0);
			this.treeView1.Margin = new System.Windows.Forms.Padding(0);
			this.treeView1.Name = "treeView1";
			treeNode1.Checked = true;
			treeNode1.Name = "Node_AddinBeginShutdown";
			treeNode1.Text = "AddinBeginShutdown";
			treeNode2.Checked = true;
			treeNode2.Name = "Node_AddinFinalize";
			treeNode2.Text = "AddinFinalize";
			treeNode3.Checked = true;
			treeNode3.Name = "Node_AddinInitialize";
			treeNode3.Text = "AddinInitialize";
			treeNode4.Checked = true;
			treeNode4.Name = "Node_AddinStartupComplete";
			treeNode4.Text = "AddinStartupComplete";
			treeNode5.Checked = true;
			treeNode5.Name = "Node_AfterUninstallControls";
			treeNode5.Text = "AfterUninstallControls";
			treeNode6.Checked = true;
			treeNode6.Name = "Node_BeforeUninstallControls";
			treeNode6.Text = "BeforeUninstallControls";
			treeNode7.Checked = true;
			treeNode7.Name = "Node_OfficeColorSchemeChanged";
			treeNode7.Text = "OfficeColorSchemeChanged";
			treeNode8.Checked = true;
			treeNode8.Name = "Node_OnError";
			treeNode8.Text = "OnError";
			treeNode9.Checked = true;
			treeNode9.Name = "Node_OnKeyDown";
			treeNode9.Text = "OnKeyDown";
			treeNode10.Checked = true;
			treeNode10.Name = "Node_OnRibbonBeforeCreate";
			treeNode10.Text = "OnRibbonBeforeCreate";
			treeNode11.Checked = true;
			treeNode11.Name = "Node_OnRibbonBeforeLoad";
			treeNode11.Text = "OnRibbonBeforeLoad";
			treeNode12.Checked = true;
			treeNode12.Name = "Node_OnRibbonLoaded";
			treeNode12.Text = "OnRibbonLoaded";
			treeNode13.Checked = true;
			treeNode13.Name = "Node_OnSendMessage";
			treeNode13.Text = "OnSendMessage";
			treeNode14.Checked = true;
			treeNode14.Name = "Node_OnTaskPaneAfterCreate";
			treeNode14.Text = "OnTaskPaneAfterCreate";
			treeNode15.Checked = true;
			treeNode15.Name = "Node_OnTaskPaneAfterShow";
			treeNode15.Text = "OnTaskPaneAfterShow";
			treeNode16.Checked = true;
			treeNode16.Name = "Node_OnTaskPaneBeforeCreate";
			treeNode16.Text = "OnTaskPaneBeforeCreate";
			treeNode17.Checked = true;
			treeNode17.Name = "Node_OnTaskPaneBeforeDestroy";
			treeNode17.Text = "OnTaskPaneBeforeDestroy";
			treeNode18.Checked = true;
			treeNode18.Name = "Node_OnTaskPaneBeforeShow";
			treeNode18.Text = "OnTaskPaneBeforeShow";
			treeNode19.Checked = true;
			treeNode19.Name = "Node_AddinModule";
			treeNode19.Text = "Add-in Module Events";
			treeNode20.Checked = true;
			treeNode20.Name = "Node_NewExplorer";
			treeNode20.Text = "NewExplorer";
			treeNode21.Checked = true;
			treeNode21.Name = "Node_ExplorerActivate";
			treeNode21.Text = "ExplorerActivate";
			treeNode22.Checked = true;
			treeNode22.Name = "Node_ExplorerAddCommandBars";
			treeNode22.Text = "ExplorerAddCommandBars";
			treeNode23.Checked = true;
			treeNode23.Name = "Node_ExplorerBeforeFolderSwitch";
			treeNode23.Text = "ExplorerBeforeFolderSwitch";
			treeNode24.Checked = true;
			treeNode24.Name = "Node_ExplorerBeforeItemCopy";
			treeNode24.Text = "ExplorerBeforeItemCopy";
			treeNode25.Checked = true;
			treeNode25.Name = "Node_ExplorerBeforeItemCut";
			treeNode25.Text = "ExplorerBeforeItemCut";
			treeNode26.Checked = true;
			treeNode26.Name = "Node_ExplorerBeforeItemPaste";
			treeNode26.Text = "ExplorerBeforeItemPaste";
			treeNode27.Checked = true;
			treeNode27.Name = "Node_ExplorerBeforeMaximize";
			treeNode27.Text = "ExplorerBeforeMaximize";
			treeNode28.Checked = true;
			treeNode28.Name = "Node_ExplorerBeforeMinimize";
			treeNode28.Text = "ExplorerBeforeMinimize";
			treeNode29.Checked = true;
			treeNode29.Name = "Node_ExplorerBeforeMove";
			treeNode29.Text = "ExplorerBeforeMove";
			treeNode30.Checked = true;
			treeNode30.Name = "Node_ExplorerBeforeSize";
			treeNode30.Text = "ExplorerBeforeSize";
			treeNode31.Checked = true;
			treeNode31.Name = "Node_ExplorerBeforeViewSwitch";
			treeNode31.Text = "ExplorerBeforeViewSwitch";
			treeNode32.Checked = true;
			treeNode32.Name = "Node_ExplorerClose";
			treeNode32.Text = "ExplorerClose";
			treeNode33.Checked = true;
			treeNode33.Name = "Node_ExplorerDeactivate";
			treeNode33.Text = "ExplorerDeactivate";
			treeNode34.Checked = true;
			treeNode34.Name = "Node_ExplorerFolderSwitch";
			treeNode34.Text = "ExplorerFolderSwitch";
			treeNode35.Checked = true;
			treeNode35.Name = "Node_ExplorerSelectionChange";
			treeNode35.Text = "ExplorerSelectionChange";
			treeNode36.Checked = true;
			treeNode36.Name = "Node_ExplorerViewSwitch";
			treeNode36.Text = "ExplorerViewSwitch";
			treeNode37.Checked = true;
			treeNode37.Name = "Node_ExplorerAttachmentSelectionChange";
			treeNode37.Text = "ExplorerAttachmentSelectionChange";
            treeNode201.Checked = true;
            treeNode201.Name = "Node_ExplorerInlineResponse";
            treeNode201.Text = "ExplorerInlineResponse";
            treeNode202.Checked = true;
            treeNode202.Name = "Node_ExplorerInlineResponseClose";
            treeNode202.Text = "ExplorerInlineResponseClose";
			treeNode38.Checked = true;
			treeNode38.Name = "Node_ExplorerEvents";
			treeNode38.Text = "Explorer(s) Events";
			treeNode39.Checked = true;
			treeNode39.Name = "Node_NewInspector";
			treeNode39.Text = "NewInspector";
			treeNode40.Checked = true;
			treeNode40.Name = "Node_InspectorActivate";
			treeNode40.Text = "InspectorActivate";
			treeNode41.Checked = true;
			treeNode41.Name = "Node_InspectorAddCommandBars";
			treeNode41.Text = "InspectorAddCommandBars";
			treeNode42.Checked = true;
			treeNode42.Name = "Node_InspectorBeforeMaximize";
			treeNode42.Text = "InspectorBeforeMaximize";
			treeNode43.Checked = true;
			treeNode43.Name = "Node_InspectorBeforeMinimize";
			treeNode43.Text = "InspectorBeforeMinimize";
			treeNode44.Checked = true;
			treeNode44.Name = "Node_InspectorBeforeMove";
			treeNode44.Text = "InspectorBeforeMove";
			treeNode45.Checked = true;
			treeNode45.Name = "Node_InspectorBeforeSize";
			treeNode45.Text = "InspectorBeforeSize";
			treeNode46.Checked = true;
			treeNode46.Name = "Node_InspectorClose";
			treeNode46.Text = "InspectorClose";
			treeNode47.Checked = true;
			treeNode47.Name = "Node_InspectorDeactivate";
			treeNode47.Text = "InspectorDeactivate";
			treeNode48.Checked = true;
			treeNode48.Name = "Node_PageChange";
			treeNode48.Text = "PageChange";
			treeNode49.Checked = true;
			treeNode49.Name = "Node_InspectorAttachmentSelectionChange";
			treeNode49.Text = "InspectorAttachmentSelectionChange";
			treeNode50.Checked = true;
			treeNode50.Name = "Node_InspectorEvents";
			treeNode50.Text = "Inspector(s) Events";
			treeNode51.Checked = true;
			treeNode51.Name = "Node_BeforeReminderShow";
			treeNode51.Text = "BeforeReminderShow";
			treeNode52.Checked = true;
			treeNode52.Name = "Node_ReminderAdd";
			treeNode52.Text = "ReminderAdd";
			treeNode53.Checked = true;
			treeNode53.Name = "Node_ReminderChange";
			treeNode53.Text = "ReminderChange";
			treeNode54.Checked = true;
			treeNode54.Name = "Node_ReminderFire";
			treeNode54.Text = "ReminderFire";
			treeNode55.Checked = true;
			treeNode55.Name = "Node_ReminderRemove";
			treeNode55.Text = "ReminderRemove";
			treeNode56.Checked = true;
			treeNode56.Name = "Node_Snooze";
			treeNode56.Text = "Snooze";
			treeNode57.Checked = true;
			treeNode57.Name = "Node_ReminderEvents";
			treeNode57.Text = "Reminder Events";
			treeNode58.Checked = true;
			treeNode58.Name = "Node_AdvancedSearchComplete";
			treeNode58.Text = "AdvancedSearchComplete";
			treeNode59.Checked = true;
			treeNode59.Name = "Node_AdvancedSearchStopped";
			treeNode59.Text = "AdvancedSearchStopped";
			treeNode60.Checked = true;
			treeNode60.Name = "Node_ItemSend";
			treeNode60.Text = "ItemSend";
			treeNode61.Checked = true;
			treeNode61.Name = "Node_NewMail";
			treeNode61.Text = "NewMail";
			treeNode62.Checked = true;
			treeNode62.Name = "Node_NewMailEx";
			treeNode62.Text = "NewMailEx";
			treeNode63.Checked = true;
			treeNode63.Name = "Node_ItemLoad";
			treeNode63.Text = "ItemLoad";
			treeNode64.Checked = true;
			treeNode64.Name = "Node_MAPILogonComplete";
			treeNode64.Text = "MAPILogonComplete";
			treeNode65.Checked = true;
			treeNode65.Name = "Node_Startup";
			treeNode65.Text = "Startup";
			treeNode66.Checked = true;
			treeNode66.Name = "Node_AutoDiscoverComplete";
			treeNode66.Text = "AutoDiscoverComplete";
			treeNode67.Checked = true;
			treeNode67.Name = "Node_Quit";
			treeNode67.Text = "Quit";
			treeNode68.Checked = true;
			treeNode68.Name = "Node_OptionPagesAdd";
			treeNode68.Text = "OptionPagesAdd";
			treeNode69.Checked = true;
			treeNode69.Name = "Node_BeforeOptionPageAdd";
			treeNode69.Text = "BeforeOptionPageAdd";
			treeNode70.Checked = true;
			treeNode70.Name = "Node_Reminder";
			treeNode70.Text = "Reminder";
			treeNode71.Checked = true;
			treeNode71.Name = "Node_BeforeFolderSharingDialog";
			treeNode71.Text = "BeforeFolderSharingDialog";
			treeNode72.Checked = true;
			treeNode72.Name = "Node_ContextMenuClose";
			treeNode72.Text = "ContextMenuClose";
			treeNode73.Checked = true;
			treeNode73.Name = "Node_ShortcutContextMenuDisplay";
			treeNode73.Text = "ShortcutContextMenuDisplay";
			treeNode74.Checked = true;
			treeNode74.Name = "Node_ViewContextMenuDisplay";
			treeNode74.Text = "ViewContextMenuDisplay";
			treeNode75.Checked = true;
			treeNode75.Name = "Node_StoreContextMenuDisplay";
			treeNode75.Text = "StoreContextMenuDisplay";
			treeNode76.Checked = true;
			treeNode76.Name = "Node_AttachmentContextMenuDisplay";
			treeNode76.Text = "AttachmentContextMenuDisplay";
			treeNode77.Checked = true;
			treeNode77.Name = "Node_FolderContextMenuDisplay";
			treeNode77.Text = "FolderContextMenuDisplay";
			treeNode78.Checked = true;
			treeNode78.Name = "Node_ItemContextMenuDisplay";
			treeNode78.Text = "ItemContextMenuDisplay";
			treeNode79.Checked = true;
			treeNode79.Name = "Node_ApplicationEvents";
			treeNode79.Text = "Application Events";
			treeNode80.Checked = true;
			treeNode80.Name = "Node_NamespaceBeforeOptionPageAdd";
			treeNode80.Text = "NamespaceBeforeOptionPageAdd";
			treeNode81.Checked = true;
			treeNode81.Name = "Node_NamespaceOptionPagesAdd";
			treeNode81.Text = "NamespaceOptionPagesAdd";
			treeNode82.Checked = true;
			treeNode82.Name = "Node_NamespaceEvents";
			treeNode82.Text = "Namespace Events";
			treeNode83.Checked = true;
			treeNode83.Name = "Node_SyncEnd";
			treeNode83.Text = "SyncEnd";
			treeNode84.Checked = true;
			treeNode84.Name = "Node_SyncError";
			treeNode84.Text = "SyncError";
			treeNode85.Checked = true;
			treeNode85.Name = "Node_SyncProgress";
			treeNode85.Text = "SyncProgress";
			treeNode86.Checked = true;
			treeNode86.Name = "Node_SyncStart";
			treeNode86.Text = "SyncStart";
			treeNode87.Checked = true;
			treeNode87.Name = "Node_SyncObjectEvents";
			treeNode87.Text = "SyncObject Events";
			treeNode88.Checked = true;
			treeNode88.Name = "Node_OnBeforeFormRegionShow";
			treeNode88.Text = "OnBeforeFormRegionShow";
			treeNode89.Checked = true;
			treeNode89.Name = "Node_OnGetFormRegionIcon";
			treeNode89.Text = "OnGetFormRegionIcon";
			treeNode90.Checked = true;
			treeNode90.Name = "Node_OnGetFormRegionManifest";
			treeNode90.Text = "OnGetFormRegionManifest";
			treeNode91.Checked = true;
			treeNode91.Name = "Node_OnGetFormRegionStorage";
			treeNode91.Text = "OnGetFormRegionStorage";
			treeNode92.Checked = true;
			treeNode92.Name = "Node_RegionEvents";
			treeNode92.Text = "Region Events";
			treeNode93.Name = "Node_CommandBarsUpdate";
			treeNode93.Text = "CommandBarsUpdate";
			treeNode94.Checked = true;
			treeNode94.Name = "Node_adxOutlookEvents";
			treeNode94.Text = "Outlook Application Events";
			treeNode95.Checked = true;
			treeNode95.Name = "Node_ProcessFolderAdd";
			treeNode95.Text = "FolderAdd";
			treeNode96.Checked = true;
			treeNode96.Name = "Node_ProcessFolderChange";
			treeNode96.Text = "FolderChange";
			treeNode97.Checked = true;
			treeNode97.Name = "Node_ProcessFolderRemove";
			treeNode97.Text = "FolderRemove";
			treeNode98.Checked = true;
			treeNode98.Name = "Node_FoldersEvents";
			treeNode98.Text = "Folders Events";
			treeNode99.Checked = true;
			treeNode99.Name = "Node_ProcessItemAdd";
			treeNode99.Text = "ItemAdd";
			treeNode100.Checked = true;
			treeNode100.Name = "Node_ProcessItemChange";
			treeNode100.Text = "ItemChange";
			treeNode101.Checked = true;
			treeNode101.Name = "Node_ProcessItemRemove";
			treeNode101.Text = "ItemRemove";
			treeNode102.Checked = true;
			treeNode102.Name = "Node_ProcessBeforeFolderMove";
			treeNode102.Text = "BeforeFolderMove";
			treeNode103.Checked = true;
			treeNode103.Name = "Node_ProcessBeforeItemMove";
			treeNode103.Text = "BeforeItemMove";
			treeNode104.Checked = true;
			treeNode104.Name = "Node_ItemsEvents";
			treeNode104.Text = "Items Events";
			treeNode105.Checked = true;
			treeNode105.Name = "Node_ProcessAttachmentAdd";
			treeNode105.Text = "AttachmentAdd";
			treeNode106.Checked = true;
			treeNode106.Name = "Node_ProcessAttachmentRead";
			treeNode106.Text = "AttachmentRead";
			treeNode107.Checked = true;
			treeNode107.Name = "Node_ProcessBeforeAttachmentSave";
			treeNode107.Text = "BeforeAttachmentSave";
			treeNode108.Checked = true;
			treeNode108.Name = "Node_ProcessBeforeCheckNames";
			treeNode108.Text = "BeforeCheckNames";
			treeNode109.Checked = true;
			treeNode109.Name = "Node_ProcessClose";
			treeNode109.Text = "Close";
			treeNode110.Checked = true;
			treeNode110.Name = "Node_ProcessCustomAction";
			treeNode110.Text = "CustomAction";
			treeNode111.Checked = true;
			treeNode111.Name = "Node_ProcessCustomPropertyChange";
			treeNode111.Text = "CustomPropertyChange";
			treeNode112.Checked = true;
			treeNode112.Name = "Node_ProcessForward";
			treeNode112.Text = "Forward";
			treeNode113.Checked = true;
			treeNode113.Name = "Node_ProcessOpen";
			treeNode113.Text = "Open";
			treeNode114.Checked = true;
			treeNode114.Name = "Node_ProcessPropertyChange";
			treeNode114.Text = "PropertyChange";
			treeNode115.Checked = true;
			treeNode115.Name = "Node_ProcessRead";
			treeNode115.Text = "Read";
			treeNode116.Checked = true;
			treeNode116.Name = "Node_ProcessReply";
			treeNode116.Text = "Reply";
			treeNode117.Checked = true;
			treeNode117.Name = "Node_ProcessReplyAll";
			treeNode117.Text = "ReplyAll";
			treeNode118.Checked = true;
			treeNode118.Name = "Node_ProcessSend";
			treeNode118.Text = "Send";
			treeNode119.Checked = true;
			treeNode119.Name = "Node_ProcessWrite";
			treeNode119.Text = "Write";
			treeNode120.Checked = true;
			treeNode120.Name = "Node_ProcessBeforeDelete";
			treeNode120.Text = "BeforeDelete";
			treeNode121.Checked = true;
			treeNode121.Name = "Node_ProcessAttachmentRemove";
			treeNode121.Text = "AttachmentRemove";
			treeNode122.Checked = true;
			treeNode122.Name = "Node_ProcessBeforeAttachmentAdd";
			treeNode122.Text = "BeforeAttachmentAdd";
			treeNode123.Checked = true;
			treeNode123.Name = "Node_ProcessBeforeAttachmentPreview";
			treeNode123.Text = "BeforeAttachmentPreview";
			treeNode124.Checked = true;
			treeNode124.Name = "Node_ProcessBeforeAttachmentRead";
			treeNode124.Text = "BeforeAttachmentRead";
			treeNode125.Checked = true;
			treeNode125.Name = "Node_ProcessBeforeAttachmentWriteToTempFile";
			treeNode125.Text = "BeforeAttachmentWriteToTempFile";
			treeNode126.Checked = true;
			treeNode126.Name = "Node_ProcessUnload";
			treeNode126.Text = "Unload";
			treeNode127.Checked = true;
			treeNode127.Name = "Node_ProcessBeforeAutoSave";
			treeNode127.Text = "BeforeAutoSave";
			treeNode128.Checked = true;
			treeNode128.Name = "Node_ProcessAfterWrite";
			treeNode128.Text = "AfterWrite";
			treeNode129.Checked = true;
			treeNode129.Name = "Node_ProcessBeforeRead";
			treeNode129.Text = "BeforeRead";
            treeNode200.Checked = true;
            treeNode200.Name = "Node_ProcessReadComplete";
            treeNode200.Text = "ReadComplete";
			treeNode130.Checked = true;
			treeNode130.Name = "Node_ItemEvents";
			treeNode130.Text = "Selected Item Events";
			treeNode131.Checked = true;
			treeNode131.Name = "Node_ADXAfterAccessProtectedObject";
			treeNode131.Text = "ADXAfterAccessProtectedObject";
			treeNode132.Checked = true;
			treeNode132.Name = "Node_ADXBeforeAccessProtectedObject";
			treeNode132.Text = "ADXBeforeAccessProtectedObject";
			treeNode133.Checked = true;
			treeNode133.Name = "Node_ADXBeforeFolderSwitch";
			treeNode133.Text = "ADXBeforeFolderSwitch";
			treeNode134.Checked = true;
			treeNode134.Name = "Node_ADXBeforeFolderSwitchEx";
			treeNode134.Text = "ADXBeforeFolderSwitchEx";
			treeNode135.Checked = true;
			treeNode135.Name = "Node_ADXBeforeFormInstanceCreate";
			treeNode135.Text = "ADXBeforeFormInstanceCreate";
			treeNode136.Checked = true;
			treeNode136.Name = "Node_ADXFolderSwitch";
			treeNode136.Text = "ADXFolderSwitch";
			treeNode137.Checked = true;
			treeNode137.Name = "Node_ADXFolderSwitchEx";
			treeNode137.Text = "ADXFolderSwitchEx";
			treeNode138.Checked = true;
			treeNode138.Name = "Node_ADXNavigationPaneHide";
			treeNode138.Text = "ADXNavigationPaneHide";
			treeNode139.Checked = true;
			treeNode139.Name = "Node_ADXNavigationPaneMinimize";
			treeNode139.Text = "ADXNavigationPaneMinimize";
			treeNode140.Checked = true;
			treeNode140.Name = "Node_ADXNavigationPaneShow";
			treeNode140.Text = "ADXNavigationPaneShow";
			treeNode141.Checked = true;
			treeNode141.Name = "Node_ADXNewInspector";
			treeNode141.Text = "ADXNewInspector";
			treeNode142.Checked = true;
			treeNode142.Name = "Node_ADXReadingPaneHide";
			treeNode142.Text = "ADXReadingPaneHide";
			treeNode143.Checked = true;
			treeNode143.Name = "Node_ADXReadingPaneMove";
			treeNode143.Text = "ADXReadingPaneMove";
			treeNode144.Checked = true;
			treeNode144.Name = "Node_ADXReadingPaneShow";
			treeNode144.Text = "ADXReadingPaneShow";
			treeNode145.Checked = true;
			treeNode145.Name = "Node_ADXTodoBarHide";
			treeNode145.Text = "ADXTodoBarHide";
			treeNode146.Checked = true;
			treeNode146.Name = "Node_ADXTodoBarMinimize";
			treeNode146.Text = "ADXTodoBarMinimize";
			treeNode147.Checked = true;
			treeNode147.Name = "Node_ADXTodoBarShow";
			treeNode147.Text = "ADXTodoBarShow";
			treeNode148.Checked = true;
			treeNode148.Name = "Node_OlFormsManagerOnError";
			treeNode148.Text = "OnError";
			treeNode149.Checked = true;
			treeNode149.Name = "Node_OlFormsManagerOnInitialize";
			treeNode149.Text = "OnInitialize";
			treeNode150.Checked = true;
			treeNode150.Name = "Node_adxOlFormsManagerEvents";
			treeNode150.Text = "Add-in Express FormsManager Events";
			this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode19,
            treeNode94,
            treeNode98,
            treeNode104,
            treeNode130,
            treeNode150});
			this.treeView1.RightToLeftLayout = true;
			this.treeView1.Size = new System.Drawing.Size(230, 332);
			this.treeView1.TabIndex = 4;
			this.treeView1.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterCheck);
			this.treeView1.SizeChanged += new System.EventHandler(this.treeView1_SizeChanged);
			this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
			// 
			// splitter1
			// 
			this.splitter1.Location = new System.Drawing.Point(230, 25);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(3, 332);
			this.splitter1.TabIndex = 7;
			this.splitter1.TabStop = false;
			this.splitter1.SizeChanged += new System.EventHandler(this.treeView1_SizeChanged);
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.textBox1);
			this.panel2.Controls.Add(this.textBox2);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panel2.Location = new System.Drawing.Point(233, 25);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(801, 332);
			this.panel2.TabIndex = 8;
			// 
			// textBox1
			// 
			this.textBox1.ContextMenuStrip = this.contextMenuStrip1;
			this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.textBox1.Location = new System.Drawing.Point(0, 0);
			this.textBox1.Multiline = true;
			this.textBox1.Name = "textBox1";
			this.textBox1.ReadOnly = true;
			this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.textBox1.Size = new System.Drawing.Size(801, 296);
			this.textBox1.TabIndex = 7;
			this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
			// 
			// textBox2
			// 
			this.textBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.textBox2.Location = new System.Drawing.Point(0, 296);
			this.textBox2.Multiline = true;
			this.textBox2.Name = "textBox2";
			this.textBox2.ReadOnly = true;
			this.textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.textBox2.Size = new System.Drawing.Size(801, 36);
			this.textBox2.TabIndex = 6;
			// 
			// ADXOlFormAddIn
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1034, 357);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.panel1);
			this.Controls.Add(this.toolStrip1);
			this.Name = "ADXOlFormAddIn";
			this.Text = "Outlook Events";
			this.Load += new System.EventHandler(this.ADXOlFormAddIn_Load);
			this.Shown += new System.EventHandler(this.ADXOlFormAddIn_Shown);
			this.Deactivate += new System.EventHandler(this.ADXOlFormAddIn_Deactivate);
			this.contextMenuStrip2.ResumeLayout(false);
			this.contextMenuStrip1.ResumeLayout(false);
			this.toolStrip1.ResumeLayout(false);
			this.toolStrip1.PerformLayout();
			this.panel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}
		#endregion

		private void SetDataToHelpData()
		{
			HelpData.Add("Node_AddinBeginShutdown", "Occurs before the host application begins its unloading procedures (Outlook 2000 and higher).");
			HelpData.Add("Node_AddinFinalize", "Occurs when the add-in is disconnected from the host application, either programmatically or through the Add-in Manager (Outlook 2000 and higher).");
			HelpData.Add("Node_AddinInitialize", "Occurs when the add-in is connected to the host application (Outlook 2000 and higher).");
			HelpData.Add("Node_AddinStartupComplete", "Occurs when the startup of the host application is complete (after the AddInInitialize event) (Outlook 2000 and higher).");
			HelpData.Add("Node_AfterUninstallControls", "Occurs after Add-in Express has uninstalled controls from the host applications (Outlook 2000 and higher).");
			HelpData.Add("Node_BeforeUninstallControls", "Occurs before Add-in Express uninstalls controls from the host applications (Outlook 2000 and higher).");
			HelpData.Add("Node_OfficeColorSchemeChanged", "Occurs when Office 2007 color scheme is changed (Outlook 2007 and higher).");
			HelpData.Add("Node_OnError", "Occurs when the add-in code generates an exception (Outlook 2000 and higher).");
			HelpData.Add("Node_OnKeyDown", "Occurs when a nonsystem key is pressed (the HandleShortcuts property should be set to True) (Outlook 2000 and higher).");
			HelpData.Add("Node_OnRibbonBeforeCreate", "Occurs once, right before the Ribbon markup file is created (Outlook 2007 and higher).");
			HelpData.Add("Node_OnRibbonBeforeLoad", "Occurs once, right before the Ribbon markup file is loaded (Outlook 2007 and higher).");
			HelpData.Add("Node_OnRibbonLoaded", "Occurs once, after the Ribbon markup file is successfully loaded (Outlook 2007 and higher).");
			HelpData.Add("Node_OnSendMessage", "Occurs after the SendMessage method is called (Outlook 2000 and higher).");
			HelpData.Add("Node_OnTaskPaneAfterCreate", "Occurs after a task pane is added to the host application (Outlook 2007 and higher).");
			HelpData.Add("Node_OnTaskPaneAfterShow", "Occurs after a task pane is shown in the host application (Outlook 2007 and higher).");
			HelpData.Add("Node_OnTaskPaneBeforeCreate", "Occurs before a task pane is added to the host application (Outlook 2007 and higher).");
			HelpData.Add("Node_OnTaskPaneBeforeDestroy", "Occurs before a task pane is destroyed (Outlook 2007 and higher).");
			HelpData.Add("Node_OnTaskPaneBeforeShow", "Occurs before a task pane is shown in the host application (Outlook 2007 and higher).");
			HelpData.Add("Node_AddinModule", "Events of the Add-in Module.");
			HelpData.Add("Node_NewExplorer", "Occurs whenever a new explorer window is opened, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerActivate", "Occurs when an explorer becomes the active window, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerAddCommandBars", "Occurs before Add-in Express adds command bars to the active explorer (Outlook 2003 and higher).");
			HelpData.Add("Node_ExplorerBeforeFolderSwitch", "Occurs before the explorer goes to a new folder, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerBeforeItemCopy", "Occurs when an item is copied. This event can be cancelled after it has started (Outlook 2002 and higher). ");
			HelpData.Add("Node_ExplorerBeforeItemCut", "Occurs when an item is cut from a folder. This method can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeItemPaste", "Occurs when a Microsoft Outlook item is pasted. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeMaximize", "Occurs when an Explorer is maximized by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeMinimize", "Occurs when the active Explorer is minimized by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeMove", "Occurs when the Explorer is moved by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeSize", "Occurs when the user sizes the current Explorer. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_ExplorerBeforeViewSwitch", "Occurs before the explorer changes to a new view, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerClose", "Occurs when an explorer is being closed (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerDeactivate", "Occurs when an explorer stops being the active window, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerFolderSwitch", "Occurs when the explorer goes to a new folder, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerSelectionChange", "Occurs when the user switches to a different item in a folder using the user interface (UI) or programmatically (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerViewSwitch", "Occurs when the view in the explorer changes, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ExplorerEvents", "Events of Outlook Explorer and Explorers collection. ");
			HelpData.Add("Node_NewInspector", "Occurs whenever a new inspector window is opened, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_InspectorActivate", "Occurs when an inspector becomes the active window, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_InspectorAddCommandBars", "Occurs before Add-in Express adds command bars to the active inspector (Outlook 2003 and higher).");
			HelpData.Add("Node_InspectorBeforeMaximize", "Occurs when an Inspector is maximized by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_InspectorBeforeMinimize", "Occurs when the active Inspector is minimized by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_InspectorBeforeMove", "Occurs when the Inspector is moved by the user. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_InspectorBeforeSize", "Occurs when the user sizes the current Inspector. This event can be cancelled after it has started (Outlook 2002 and higher).");
			HelpData.Add("Node_InspectorClose", "Occurs when the inspector associated with a Microsoft Outlook item is being closed (Outlook 2000 and higher).");
			HelpData.Add("Node_InspectorDeactivate", "Occurs when an inspector stops being the active window, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_PageChange", "Occurs when the active form page changes, either programmatically or by user action, on an Inspector object (Outlook 2007 and higher).");
			HelpData.Add("Node_InspectorEvents", "Events of Outlook Inspector and Inspectors collection. ");
			HelpData.Add("Node_BeforeReminderShow", "Occurs before the Reminder dialog box is displayed (Outlook 2002 and higher).");
			HelpData.Add("Node_ReminderAdd", "Occurs after a reminder is added (Outlook 2002 and higher).");
			HelpData.Add("Node_ReminderChange", "Occurs after a reminder has been modified (Outlook 2002 and higher).");
			HelpData.Add("Node_ReminderFire", "Occurs before the reminder is executed (Outlook 2002 and higher).");
			HelpData.Add("Node_ReminderRemove", "Occurs when a Reminder object has been removed from the collection (Outlook 2002 and higher).");
			HelpData.Add("Node_Snooze", "Occurs when a reminder is dismissed using the Snooze button (Outlook 2002 and higher).");
			HelpData.Add("Node_ReminderEvents", "Events of Reminder (Outlook 2002 and higher). ");
			HelpData.Add("Node_AdvancedSearchComplete", "Occurs when the AdvancedSearch method has completed (Outlook 2003 and higher).");
			HelpData.Add("Node_AdvancedSearchStopped", "Occurs when a specified Search object's Stop method has been executed (Outlook 2003 and higher).");
			HelpData.Add("Node_ItemSend", "Occurs whenever an item is sent, either by the user through an Inspector (before the inspector is closed, but after the user clicks the Send button) or when the Send method is used in a program (Outlook 2000 and higher).");
			HelpData.Add("Node_NewMail", "Occurs when one or more new e-mail messages are received in the Inbox (Outlook 2000 and higher).");
			HelpData.Add("Node_NewMailEx", "Occurs when one or more new items are received in the Inbox (Outlook 2003 and higher).");
			HelpData.Add("Node_ItemLoad", "Occurs when an Outlook item is loaded into memory (Outlook 2007 and higher).");
			HelpData.Add("Node_MAPILogonComplete", "Occurs after the user has logged onto the system (Outlook 2002 and higher).");
			HelpData.Add("Node_Startup", "Occurs when Microsoft Outlook is starting, but after all add-in programs have been loaded (Outlook 2000 and higher).");
			HelpData.Add("Node_AutoDiscoverComplete", "Occurs after Outlook has finished accessing the auto-discovery service of an Exchange server and has the related information available in NameSpace.AutoDiscoverXml (Outlook 2007 and higher).");
			HelpData.Add("Node_Quit", "Occurs when Microsoft Outlook begins to close (Outlook 2000 and higher).");
			HelpData.Add("Node_OptionPagesAdd", "Occurs when Outlook updates its option page collection (Outlook 2000 and higher).");
			HelpData.Add("Node_BeforeOptionPageAdd", "Occurs before a new option page is added to the collection of Outlook option pages (Outlook 2000 and higher).");
			HelpData.Add("Node_Reminder", "Occurs immediately before a reminder is displayed (Outlook 2000 and higher).");
			HelpData.Add("Node_BeforeFolderSharingDialog", "Occurs before the Sharing dialog box is displayed for a selected Folder object (Outlook 2007 and higher).");
			HelpData.Add("Node_ApplicationEvents", "Events of Outlook Application.");
			HelpData.Add("Node_NamespaceBeforeOptionPageAdd", "Occurs before a new option page is added to the collection of Outlook option pages (Outlook 2000 and higher).");
			HelpData.Add("Node_NamespaceOptionPagesAdd", "Occurs when Outlook updates the option pages collections of folders (Outlook 2000 and higher).");
			HelpData.Add("Node_NamespaceEvents", "Events of Namespace.");
			HelpData.Add("Node_SyncEnd", "Occurs immediately after Microsoft Outlook finishes synchronizing a user’s folders using the specified Send\\Receive group (Outlook 2000 and higher).");
			HelpData.Add("Node_SyncError", "Occurs when Microsoft Outlook encounters an error while synchronizing a user’s folders using the specified Send\\Receive group (Outlook 2000 and higher).");
			HelpData.Add("Node_SyncProgress", "Occurs periodically while Microsoft Outlook is synchronizing a user’s folders using the specified Send\\Receive group (Outlook 2000 and higher).");
			HelpData.Add("Node_SyncStart", "Occurs when Microsoft Outlook begins synchronizing a user’s folders using the specified Send\\Receive group (Outlook 2000 and higher).");
			HelpData.Add("Node_SyncObjectEvents", "Events of SyncObject.");
			HelpData.Add("Node_ContextMenuClose", "Occurs after a context menu is closed (Outlook 2007 and higher).");
			HelpData.Add("Node_ViewContextMenuDisplay", "Occurs before a context menu is displayed for a view (Outlook 2007 and higher).");
			HelpData.Add("Node_ShortcutContextMenuDisplay", "Occurs before a context menu is displayed for a shortcut (Outlook 2007 and higher).");
			HelpData.Add("Node_AttachmentContextMenuDisplay", "Occurs before a context menu is displayed for a collection of attachments (Outlook 2007 and higher).");
			HelpData.Add("Node_StoreContextMenuDisplay", "Occurs before a context menu is displayed for a store (Outlook 2007 and higher).");
			HelpData.Add("Node_FolderContextMenuDisplay", "Occurs before a context menu is displayed for a folder (Outlook 2007 and higher).");
			HelpData.Add("Node_ItemContextMenuDisplay", "Occurs before a context menu is displayed for a collection of Outlook items (Outlook 2007 and higher).");
			HelpData.Add("Node_ContextMenuEvents", "Events of ContextMenu. ");
			HelpData.Add("Node_OnBeforeFormRegionShow", "Allows an add-in to update the user interface of a form region before it is displayed (Outlook 2007 and higher).");
			HelpData.Add("Node_OnGetFormRegionIcon", "Obtains an icon image that will be displayed for a particular type of icon for the form region (Outlook 2007 and higher).");
			HelpData.Add("Node_OnGetFormRegionManifest", "Obtains the XML manifest for a form region (Outlook 2007 and higher).");
			HelpData.Add("Node_OnGetFormRegionStorage", "Obtains appropriate storage for a form region based on the specified information (Outlook 2007 and higher).");
			HelpData.Add("Node_RegionEvents", "Events of Region");
			HelpData.Add("Node_CommandBarsUpdate", "Occurs when Outlook updates its command bars (Outlook 2000 and higher).");
			HelpData.Add("Node_adxOutlookEvents", "Events of Outlook.");
			HelpData.Add("Node_ProcessFolderAdd", "Occurs when a folder is added to the specified Folders collection (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessFolderChange", "Occurs when a folder in the specified Folders collection is changed (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessFolderRemove", "Occurs when a folder is removed from the specified Folders collection (Outlook 2000 and higher).");
			HelpData.Add("Node_FoldersEvents", "Events of Folders collection.");
			HelpData.Add("Node_ProcessItemAdd", "Occurs when one or more items are added to the specified collection (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessItemChange", "Occurs when an item in the specified collection is changed (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessItemRemove", "Occurs when an item is deleted from the specified collection (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessBeforeFolderMove", "Occurs when a folder is about to be moved or deleted, either as a result of user action or through program code (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeItemMove", "Occurs when an item is about to be moved or deleted from a folder, either as a result of user action or through program code (Outlook 2007 and higher). ");
			HelpData.Add("Node_ItemsEvents", "Events of Items collection and Folder.");
			HelpData.Add("Node_ProcessAttachmentAdd", "Occurs when an attachment has been added to an instance of the parent object (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessAttachmentRead", "Occurs when an attachment in an instance of the parent object has been opened for reading (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessBeforeAttachmentSave", "Occurs just before an attachment is saved (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessBeforeCheckNames", "Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object) (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessClose", "Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessCustomAction", "Occurs when a custom action of an item (which is an instance of the parent object) executes (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessCustomPropertyChange", "Occurs when a custom property of an item (which is an instance of the parent object) is changed (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessForward", "Occurs when the user selects the Forward action for an item, or when the Forward method is called for the item, which is an instance of the parent object (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessOpen", "Occurs when an instance of the parent object is being opened in an Inspector (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessPropertyChange", "Occurs when an explicit built-in property (for example, Subject) of an instance of the parent object is changed (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessRead", "Occurs when an instance of the parent object is opened for editing by the user (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessReply", "Occurs when the user selects the Reply action for an item, or when the Reply method is called for the item, which is an instance of the parent object (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessReplyAll", "Occurs when the user selects the ReplyAll action for an item, or when the ReplyAll method is called for the item, which is an instance of the parent object (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessSend", "Occurs when the user selects the Send action for an item, or when the Send method is called for the item, which is an instance of the parent object (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessWrite", "Occurs when an instance of the parent object is saved, either explicitly (for example, using the Save or SaveAs methods) or implicitly (for example, in response to a prompt when closing the item's inspector) (Outlook 2000 and higher).");
			HelpData.Add("Node_ProcessBeforeDelete", "Occurs before an item (which is an instance of the parent object) is deleted (Outlook 2002 and higher).");
			HelpData.Add("Node_ProcessAttachmentRemove", "Occurs when an attachment has been removed from an instance of the parent object (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeAttachmentAdd", "Occurs before an attachment is added to an instance of the parent object (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeAttachmentPreview", "Occurs before an attachment associated with an instance of the parent object is previewed (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeAttachmentRead", "Occurs before an attachment associated with an instance of the parent object is read from the file system, an attachment stream, or an Attachment object (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeAttachmentWriteToTempFile", "Occurs before an attachment associated with an instance of the parent object is written to a temporary file (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessUnload", "Occurs before an Outlook item is unloaded from memory, either programmatically or by user action (Outlook 2007 and higher).");
			HelpData.Add("Node_ProcessBeforeAutoSave", "Occurs before the item is automatically saved by Outlook (Outlook 2007 and higher).");
			HelpData.Add("Node_ItemEvents", "Events of Item");
			HelpData.Add("Node_ADXAfterAccessProtectedObject", "Occurs immediately after ADXOlFormsManager accesses a protected Outlook object (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXBeforeAccessProtectedObject", "Occurs immediately before ADXOlFormsManager accesses a protected Outlook object (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXBeforeFolderSwitch", "Occurs before an Outlook explorer goes to a new folder, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXBeforeFolderSwitchEx", "Occurs before an Outlook explorer goes to a new folder, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXBeforeFormInstanceCreate", "Occurs before a form instance (of the ADXOlForm type) is created. Allows preventing the creation of the instance (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXFolderSwitch", "Occurs when an Outlook explorer goes to a new folder, either as a result of user action or through program code. Set ShowForm to False to prevent any ADXOlForm showing and prevent ADXFolderSwitch to fire (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXFolderSwitchEx", "Occurs when an Outlook explorer goes to a new folder, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXNavigationPaneHide", "Occurs when the Navigation Pane (2003, 2007), Outlook Bar (2000, 2002) or Folder List (2000, 2002) is hidden (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXNavigationPaneMinimize", "Occurs when the Navigation Pane (2003, 2007) is minimized (Outlook 2003 and higher).");
			HelpData.Add("Node_ADXNavigationPaneShow", "Occurs when the Navigation Pane (2003, 2007), Outlook Bar (2000, 2002) or Folder List (2000, 2002) is shown, but before any forms (of the ADXOlForm type) are shown (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXNewInspector", "Occurs whenever a new inspector window is opened, either as a result of user action or through program code (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXReadingPaneHide", "Occurs when the Reading Pane is hidden (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXReadingPaneMove", "Occurs when the Reading Pane position is changed. Allows determining the layout in which the Reading Pane is shown (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXReadingPaneShow", "Occurs when the Reading Pane is shown, but before any forms (of the ADXOlForm type) are shown. Allows determining the layout in which the Reading Pane is shown (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXTodoBarHide", "Occurs when the Todo Bar is hidden (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXTodoBarMinimize", "Occurs when the Todo Bar is minimized (Outlook 2000 and higher).");
			HelpData.Add("Node_ADXTodoBarShow", "Occurs when the Todo Bar is shown, but before any forms (of the ADXOlForm type) are shown (Outlook 2000 and higher).");
			HelpData.Add("Node_OlFormsManagerOnError", "Occurs when ADXOlFormaManager generates an exception(Outlook 2000 and higher).");
			HelpData.Add("Node_OlFormsManagerOnInitialize", "Occurs before Forms Manager is initialized (Outlook 2000 and higher).");
			HelpData.Add("Node_adxOlFormsManagerEvents", "Events of adxOlFormsManager.");
			HelpData.Add("Node_ProcessAfterWrite", "Occurs after Microsoft Outlook has saved the item (Outlook 2010 and higher).");
			HelpData.Add("Node_ProcessBeforeRead", "Occurs before Microsoft Outlook begins to read the properties for the item (Outlook 2010 and higher).");
			HelpData.Add("Node_InspectorAttachmentSelectionChange", "Occurs when the user selects a different or additional attachment of an item in the active inspector programmatically or by interacting with the user interface (Outlook 2010 and higher). ");
			HelpData.Add("Node_ExplorerAttachmentSelectionChange", "Occurs when the user selects a different or additional attachment in the active explorer programmatically or by interacting with the user interface (Outlook 2010 and higher). ");

            // 2013
            HelpData.Add("Node_ProcessReadComplete", "Occurs when Outlook has completed reading the properties of the item (Outlook 2013).");
            HelpData.Add("Node_ExplorerInlineResponse", "Occurs when the user performs an action that causes an inline response to appear in the Reading Pane (Outlook 2013). ");
            HelpData.Add("Node_ExplorerInlineResponseClose", "Occurs when the user performs an action that causes the active inline response to close in the Reading Pane (Outlook 2013). ");
        }

		private void SetNodeCheckBoxState(TreeNode node)
		{
			if (node.Nodes.Count > 0)
			{
				for (int i = 0; i < node.Nodes.Count; i++)
				{
					node.Nodes[i].Checked = node.Checked;
				}
			}
		}

		private void SetParentNodeState(TreeNode node)
		{
			TreeNode CurrParentNode = node.Parent;
			bool NodeState = false;
			for (int i = 0; i < CurrParentNode.Nodes.Count; i++)
			{
				NodeState = NodeState | CurrParentNode.Nodes[i].Checked;
			}
			node.Parent.Checked = NodeState;
		}

		private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
		{
			this.textBox2.Text = HelpData[e.Node.Name].ToString();
		}

		private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
		{
			if ((e.Node.Nodes.Count > 0) && (saveChangeTreeWiew))
				SetNodeCheckBoxState(e.Node);
			CurrentModule = AddinModule as AddinModule;
			CurrentModule.setTreeView[e.Node.Name] = e.Node.Checked;
		}

		private void ADXOlFormAddIn_Load(object sender, EventArgs e)
		{
			SetDataToHelpData();
			textBox1.BackColor = treeView1.BackColor;
			textBox2.BackColor = treeView1.BackColor;

			RegistryKey key;
			string[] SubKeys = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey).GetSubKeyNames();
			bool newKey = true;
			for (int i = 0; i < SubKeys.Length; i++)
				if (SubKeys[i] == "Forms")
					newKey = false;
			if (newKey)
			{
				key = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey, RegistryKeyPermissionCheck.ReadWriteSubTree).CreateSubKey("Forms", RegistryKeyPermissionCheck.ReadWriteSubTree);
				key.SetValue("TreeViewWidth", this.panel1.Width);
			}
			else
			{
				key = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey).OpenSubKey("Forms", RegistryKeyPermissionCheck.ReadWriteSubTree);
			}
			this.panel1.Width = System.Convert.ToInt32(key.GetValue("TreeViewWidth", 230));

			key.Close();

			CurrentModule = AddinModule as AddinModule;
			if (CurrentModule.setTreeView.Count > 0)
				foreach (TreeNode node in treeView1.Nodes)
				{
					node.Checked = Convert.ToBoolean(CurrentModule.setTreeView[node.Name]);
					if (node.Nodes.Count > 0)
						SetNextLevelTreeView(node);
				}
			saveChangeTreeWiew = true;
		}

		private void SetNextLevelTreeView(TreeNode parentNode)
		{
			CurrentModule = AddinModule as AddinModule;
			foreach (TreeNode node in parentNode.Nodes)
			{
				node.Checked = Convert.ToBoolean(CurrentModule.setTreeView[node.Name]);
				if (node.Nodes.Count > 0)
					SetNextLevelTreeView(node);
			}

		}
		private void ADXOlFormAddIn_Shown(object sender, EventArgs e)
		{
			treeView1.Focus();

		}

		private void cleatToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TextBox1Clear();
		}

		internal void TextBox1Clear()
		{
			this.textBox1.Text = string.Empty;
		}

		private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TextBox1SelectAllText();
		}

		internal void TextBox1SelectAllText()
		{
			this.textBox1.Focus();
			this.textBox1.SelectAll();
		}

		private void copyToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TextBox1Copy();
		}

		internal void TextBox1Copy()
		{
			if (textBox1.SelectedText != "")
				Clipboard.SetDataObject(textBox1.SelectedText);
			else
				Clipboard.SetDataObject(textBox1.Text);
		}

		private void expandAllToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TreeViewExpandAll();
		}

		internal void TreeViewExpandAll()
		{
			treeView1.ExpandAll();
		}

		private void collapseAllToolStripMenuItem_Click(object sender, EventArgs e)
		{
			TreeViewCollapseAll();
		}

		internal void TreeViewCollapseAll()
		{
			treeView1.CollapseAll();
		}

		private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
		{
			SaveLogAs();
		}

		internal void SaveLogAs()
		{
			string s = string.Empty;
			if (textBox1.SelectedText != "")
				s = textBox1.SelectedText;
			else
				s = textBox1.Text;
			if (saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				System.IO.StreamWriter sw = new System.IO.StreamWriter(saveFileDialog1.FileName);
				sw.WriteLine(s);
				sw.Close();
			}
			this.textBox1.Select(this.textBox1.Text.Length, 0);
			this.textBox1.ScrollToCaret();
		}

		private void treeView1_SizeChanged(object sender, EventArgs e)
		{
			if (this.AddinModule == null) return;

			string[] SubKeys = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey).GetSubKeyNames();
			bool newKey = true;
			for (int i = 0; i < SubKeys.Length; i++)
				if (SubKeys[i] == "Forms")
					newKey = false;
			RegistryKey key = null;
			if (newKey)
			{
				key = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey, RegistryKeyPermissionCheck.ReadWriteSubTree).CreateSubKey("Forms", RegistryKeyPermissionCheck.ReadWriteSubTree);
			}
			else
				key = Registry.CurrentUser.OpenSubKey((this.AddinModule as AddinModule).RegistryKey + "\\Forms", RegistryKeyPermissionCheck.ReadWriteSubTree);
			if (saveChangeTreeWiew)
				key.SetValue("TreeViewWidth", this.panel1.Width);
			if (key != null) key.Close();
		}

		private void toolStripButtonSelectAll_Click(object sender, EventArgs e)
		{
			TextBox1SelectAllText();
		}

		private void toolStripButtonClear_Click(object sender, EventArgs e)
		{
			TextBox1Clear();
		}

		private void toolStripButtonCopy_Click(object sender, EventArgs e)
		{
			TextBox1Copy();
		}

		private void toolStripButtonSaveAs_Click(object sender, EventArgs e)
		{
			SaveLogAs();
		}

		private void toolStripButtonWriteLogToFile_CheckedChanged(object sender, EventArgs e)
		{
			CurrentModule = AddinModule as AddinModule;
			CurrentModule.WritteToLogFile(toolStripButtonWriteLogToFile.Checked);
		}

		private void toolStripButtonStartStopLog_Click(object sender, EventArgs e)
		{
			CurrentModule = AddinModule as AddinModule;
			CurrentModule.SetStartStopLog(toolStripButtonStartStopLog.Checked);
			if (toolStripButtonStartStopLog.Checked)
			{
				toolStripButtonStartStopLog.Text = "Log is started";
			}
			else
			{
				toolStripButtonStartStopLog.Text = "Log is stopped";
			}
		}

		internal void SetStateButton()
		{
			CurrentModule = AddinModule as AddinModule;
			if (CurrentModule != null)
			{
				toolStripButtonStartStopLog.Checked = CurrentModule.StartStopLog;
				if (toolStripButtonStartStopLog.Checked)
				{
					toolStripButtonStartStopLog.Text = "Log is started";
				}
				else
				{
					toolStripButtonStartStopLog.Text = "Log is stopped";
				}

				if (CurrentModule.sw == null)
					toolStripButtonWriteLogToFile.Checked = false;
				else
					toolStripButtonWriteLogToFile.Checked = true;
			}
		}

		private void toolStripLabel1_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process p = new System.Diagnostics.Process();
			p.StartInfo.FileName = "http://www.add-in-express.com/";
			p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal;
			p.Start();
		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{
			this.textBox1.Select(this.textBox1.Text.Length, 0);
			this.textBox1.ScrollToCaret();
		}

		private void ADXOlFormAddIn_Deactivate(object sender, EventArgs e)
		{
			CurrentModule = AddinModule as AddinModule;
			CurrentModule.WriteTreeViewState();
		}
	}
}
