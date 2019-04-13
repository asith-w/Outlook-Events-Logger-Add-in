using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Win32;

namespace OutlookEvents
{
	/// <summary>
	///   Add-in Express Add-in Module
	/// </summary>
	[GuidAttribute("9755D90A-F346-415A-84DB-1E0264DDFB80"), ProgId("OutlookEvents.AddinModule")]
	public class AddinModule : AddinExpress.MSO.ADXAddinModule
	{
		public AddinModule()
		{
			InitializeComponent();
		}

		private AddinExpress.OL.ADXOlFormsManager adxOlFormsManager1;
		private AddinExpress.OL.ADXOlFormsCollectionItem adxOlFormsCollectionItem1;
		private AddinExpress.MSO.ADXOutlookAppEvents adxOutlookEvents;

		private ADXOlFormAddIn ResultForm = null;
		private List<string> CurrentEvents = new List<string>();
		private bool isExplorerActivate = true;
		private List<OutlookFoldersEventsClass1> folderEvents = new List<OutlookFoldersEventsClass1>();
		private List<OutlookItemsEventsClass1> itemsEvents = new List<OutlookItemsEventsClass1>();
		private OutlookItemEventsClass1 selectedItemEvents = null;
		internal List<OutlookItemEventsClass1> itemEvents = new List<OutlookItemEventsClass1>();
		public string OutboxFolderEntryID = string.Empty;
		private SaveFileDialog saveFileDialog1;
		internal System.IO.StreamWriter sw = null;
		internal bool StartStopLog = true;
		internal Hashtable setTreeView = new Hashtable();


		#region Component Designer generated code
		/// <summary>
		/// Required by designer
		/// </summary>
		private System.ComponentModel.IContainer components;

		/// <summary>
		/// Required by designer support - do not modify
		/// the following method
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.adxOlFormsManager1 = new AddinExpress.OL.ADXOlFormsManager(this.components);
            this.adxOlFormsCollectionItem1 = new AddinExpress.OL.ADXOlFormsCollectionItem(this.components);
            this.adxOutlookEvents = new AddinExpress.MSO.ADXOutlookAppEvents(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            // 
            // adxOlFormsManager1
            // 
            this.adxOlFormsManager1.Items.Add(this.adxOlFormsCollectionItem1);
            this.adxOlFormsManager1.ADXBeforeFolderSwitch += new AddinExpress.OL.ADXOlFormsManager.BeforeFolderSwitch_EventHandler(this.adxOlFormsManager1_ADXBeforeFolderSwitch);
            this.adxOlFormsManager1.ADXBeforeFolderSwitchEx += new AddinExpress.OL.ADXOlFormsManager.BeforeFolderSwitchEx_EventHandler(this.adxOlFormsManager1_ADXBeforeFolderSwitchEx);
            this.adxOlFormsManager1.ADXFolderSwitch += new AddinExpress.OL.ADXOlFormsManager.FolderSwitch_EventHandler(this.adxOlFormsManager1_ADXFolderSwitch);
            this.adxOlFormsManager1.ADXFolderSwitchEx += new AddinExpress.OL.ADXOlFormsManager.FolderSwitchEx_EventHandler(this.adxOlFormsManager1_ADXFolderSwitchEx);
            this.adxOlFormsManager1.OnInitialize += new AddinExpress.OL.ADXOlFormsManager.OnComponentInitialize_EventHandler(this.adxOlFormsManager1_OnInitialize);
            this.adxOlFormsManager1.ADXNewInspector += new AddinExpress.OL.ADXOlFormsManager.NewInspector_EventHandler(this.adxOlFormsManager1_ADXNewInspector);
            this.adxOlFormsManager1.ADXBeforeFormInstanceCreate += new AddinExpress.OL.ADXOlFormsManager.BeforeFormInstanceCreate_EventHandler(this.adxOlFormsManager1_ADXBeforeFormInstanceCreate);
            this.adxOlFormsManager1.ADXNavigationPaneHide += new AddinExpress.OL.ADXOlFormsManager.NavigationPaneHide_EventHandler(this.adxOlFormsManager1_ADXNavigationPaneHide);
            this.adxOlFormsManager1.ADXNavigationPaneMinimize += new AddinExpress.OL.ADXOlFormsManager.NavigationPaneMinimize_EventHandler(this.adxOlFormsManager1_ADXNavigationPaneMinimize);
            this.adxOlFormsManager1.ADXNavigationPaneShow += new AddinExpress.OL.ADXOlFormsManager.NavigationPaneShow_EventHandler(this.adxOlFormsManager1_ADXNavigationPaneShow);
            this.adxOlFormsManager1.ADXReadingPaneHide += new AddinExpress.OL.ADXOlFormsManager.ReadingPaneHide_EventHandler(this.adxOlFormsManager1_ADXReadingPaneHide);
            this.adxOlFormsManager1.ADXReadingPaneShow += new AddinExpress.OL.ADXOlFormsManager.ReadingPaneShow_EventHandler(this.adxOlFormsManager1_ADXReadingPaneShow);
            this.adxOlFormsManager1.ADXReadingPaneMove += new AddinExpress.OL.ADXOlFormsManager.ReadingPaneMove_EventHandler(this.adxOlFormsManager1_ADXReadingPaneMove);
            this.adxOlFormsManager1.ADXTodoBarHide += new AddinExpress.OL.ADXOlFormsManager.TodoBarHide_EventHandler(this.adxOlFormsManager1_ADXTodoBarHide);
            this.adxOlFormsManager1.ADXTodoBarMinimize += new AddinExpress.OL.ADXOlFormsManager.TodoBarMinimize_EventHandler(this.adxOlFormsManager1_ADXTodoBarMinimize);
            this.adxOlFormsManager1.ADXTodoBarShow += new AddinExpress.OL.ADXOlFormsManager.TodoBarShow_EventHandler(this.adxOlFormsManager1_ADXTodoBarShow);
            this.adxOlFormsManager1.OnError += new AddinExpress.OL.ADXOlFormsManager.Error_EventHandler(this.adxOlFormsManager1_OnError);
            this.adxOlFormsManager1.ADXBeforeAccessProtectedObject += new AddinExpress.OL.ADXOlFormsManager.BeforeAccessProtectedObject_EventHandler(this.adxOlFormsManager1_ADXBeforeAccessProtectedObject);
            this.adxOlFormsManager1.ADXAfterAccessProtectedObject += new AddinExpress.OL.ADXOlFormsManager.AfterAccessProtectedObject_EventHandler(this.adxOlFormsManager1_ADXAfterAccessProtectedObject);
            this.adxOlFormsManager1.SetOwner(this);
            // 
            // adxOlFormsCollectionItem1
            // 
            this.adxOlFormsCollectionItem1.Cached = AddinExpress.OL.ADXOlCachingStrategy.OneInstanceForAllFolders;
            this.adxOlFormsCollectionItem1.ExplorerAllowedDropRegions = AddinExpress.OL.ADXOlExplorerAllowedDropRegions.BottomSubpane;
            this.adxOlFormsCollectionItem1.ExplorerItemTypes = ((AddinExpress.OL.ADXOlExplorerItemTypes)((((((((AddinExpress.OL.ADXOlExplorerItemTypes.olMailItem | AddinExpress.OL.ADXOlExplorerItemTypes.olAppointmentItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olContactItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olTaskItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olJournalItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olNoteItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olPostItem) 
            | AddinExpress.OL.ADXOlExplorerItemTypes.olDistributionListItem)));
            this.adxOlFormsCollectionItem1.ExplorerLayout = AddinExpress.OL.ADXOlExplorerLayout.BottomSubpane;
            this.adxOlFormsCollectionItem1.FormClassName = "OutlookEvents.ADXOlFormAddIn";
            this.adxOlFormsCollectionItem1.UseOfficeThemeForBackground = true;
            // 
            // adxOutlookEvents
            // 
            this.adxOutlookEvents.HandleEvents = new AddinExpress.MSO.ADXOutlookAppEvents.HandledOutlookEvents(true, true);
            this.adxOutlookEvents.ItemSend += new AddinExpress.MSO.ADXOlItemSend_EventHandler(this.adxOutlookEvents_ItemSend);
            this.adxOutlookEvents.NewMail += new System.EventHandler(this.adxOutlookEvents_NewMail);
            this.adxOutlookEvents.Reminder += new AddinExpress.MSO.ADXOlItem_EventHandler(this.adxOutlookEvents_Reminder);
            this.adxOutlookEvents.OptionPagesAdd += new AddinExpress.MSO.ADXOlOptionPages_EventHandler(this.adxOutlookEvents_OptionPagesAdd);
            this.adxOutlookEvents.BeforeOptionPageAdd += new AddinExpress.MSO.ADXOlBeforeOptionPageAdd_EventHandler(this.adxOutlookEvents_BeforeOptionPageAdd);
            this.adxOutlookEvents.Startup += new System.EventHandler(this.adxOutlookEvents_Startup);
            this.adxOutlookEvents.Quit += new System.EventHandler(this.adxOutlookEvents_Quit);
            this.adxOutlookEvents.NewInspector += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_NewInspector);
            this.adxOutlookEvents.InspectorActivate += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorActivate);
            this.adxOutlookEvents.InspectorAddCommandBars += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorAddCommandBars);
            this.adxOutlookEvents.InspectorDeactivate += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorDeactivate);
            this.adxOutlookEvents.InspectorClose += new AddinExpress.MSO.ADXOlInspector_EventHandler(this.adxOutlookEvents_InspectorClose);
            this.adxOutlookEvents.NewExplorer += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_NewExplorer);
            this.adxOutlookEvents.ExplorerActivate += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerActivate);
            this.adxOutlookEvents.ExplorerAddCommandBars += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerAddCommandBars);
            this.adxOutlookEvents.ExplorerFolderSwitch += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerFolderSwitch);
            this.adxOutlookEvents.ExplorerClose += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerClose);
            this.adxOutlookEvents.ExplorerBeforeFolderSwitch += new AddinExpress.MSO.ADXOlExplorerBeforeFolderSwitch_EventHandler(this.adxOutlookEvents_ExplorerBeforeFolderSwitch);
            this.adxOutlookEvents.ExplorerViewSwitch += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerViewSwitch);
            this.adxOutlookEvents.ExplorerBeforeViewSwitch += new AddinExpress.MSO.ADXOlExplorerBeforeViewSwitch_EventHandler(this.adxOutlookEvents_ExplorerBeforeViewSwitch);
            this.adxOutlookEvents.ExplorerDeactivate += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerDeactivate);
            this.adxOutlookEvents.ExplorerSelectionChange += new AddinExpress.MSO.ADXOlExplorer_EventHandler(this.adxOutlookEvents_ExplorerSelectionChange);
            this.adxOutlookEvents.NamespaceOptionPagesAdd += new AddinExpress.MSO.ADXOlNamespaceOptionPages_EventHandler(this.adxOutlookEvents_NamespaceOptionPagesAdd);
            this.adxOutlookEvents.NamespaceBeforeOptionPageAdd += new AddinExpress.MSO.ADXOlNamespaceBeforeOptionPageAdd_EventHandler(this.adxOutlookEvents_NamespaceBeforeOptionPageAdd);
            this.adxOutlookEvents.CommandBarsUpdate += new System.EventHandler(this.adxOutlookEvents_CommandBarsUpdate);
            this.adxOutlookEvents.AdvancedSearchComplete += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxOutlookEvents_AdvancedSearchComplete);
            this.adxOutlookEvents.AdvancedSearchStopped += new AddinExpress.MSO.ADXHostActiveObject_EventHandler(this.adxOutlookEvents_AdvancedSearchStopped);
            this.adxOutlookEvents.MAPILogonComplete += new System.EventHandler(this.adxOutlookEvents_MAPILogonComplete);
            this.adxOutlookEvents.ExplorerBeforeMaximize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeMaximize);
            this.adxOutlookEvents.ExplorerBeforeMinimize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeMinimize);
            this.adxOutlookEvents.ExplorerBeforeMove += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeMove);
            this.adxOutlookEvents.ExplorerBeforeSize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeSize);
            this.adxOutlookEvents.ExplorerBeforeItemCopy += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeItemCopy);
            this.adxOutlookEvents.ExplorerBeforeItemCut += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_ExplorerBeforeItemCut);
            this.adxOutlookEvents.ExplorerBeforeItemPaste += new AddinExpress.MSO.ADXOlExplorerBeforeItemPaste_EventHandler(this.adxOutlookEvents_ExplorerBeforeItemPaste);
            this.adxOutlookEvents.InspectorBeforeMaximize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_InspectorBeforeMaximize);
            this.adxOutlookEvents.InspectorBeforeMinimize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_InspectorBeforeMinimize);
            this.adxOutlookEvents.InspectorBeforeMove += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_InspectorBeforeMove);
            this.adxOutlookEvents.InspectorBeforeSize += new AddinExpress.MSO.ADXHostAction_EventHandler(this.adxOutlookEvents_InspectorBeforeSize);
            this.adxOutlookEvents.NewMailEx += new AddinExpress.MSO.ADXOlNewMailEx_EventHandler(this.adxOutlookEvents_NewMailEx);
            this.adxOutlookEvents.BeforeReminderShow += new AddinExpress.MSO.ADXOlBeforeReminderShow_EventHandler(this.adxOutlookEvents_BeforeReminderShow);
            this.adxOutlookEvents.ReminderAdd += new AddinExpress.MSO.ADXOlReminder_EventHandler(this.adxOutlookEvents_ReminderAdd);
            this.adxOutlookEvents.ReminderChange += new AddinExpress.MSO.ADXOlReminder_EventHandler(this.adxOutlookEvents_ReminderChange);
            this.adxOutlookEvents.ReminderFire += new AddinExpress.MSO.ADXOlReminder_EventHandler(this.adxOutlookEvents_ReminderFire);
            this.adxOutlookEvents.ReminderRemove += new System.EventHandler(this.adxOutlookEvents_ReminderRemove);
            this.adxOutlookEvents.Snooze += new AddinExpress.MSO.ADXOlReminder_EventHandler(this.adxOutlookEvents_Snooze);
            this.adxOutlookEvents.AttachmentContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_AttachmentContextMenuDisplay);
            this.adxOutlookEvents.FolderContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_FolderContextMenuDisplay);
            this.adxOutlookEvents.StoreContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_StoreContextMenuDisplay);
            this.adxOutlookEvents.ShortcutContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_ShortcutContextMenuDisplay);
            this.adxOutlookEvents.ViewContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_ViewContextMenuDisplay);
            this.adxOutlookEvents.ItemContextMenuDisplay += new AddinExpress.MSO.ADXOlContextMenu_EventHandler(this.adxOutlookEvents_ItemContextMenuDisplay);
            this.adxOutlookEvents.ContextMenuClose += new AddinExpress.MSO.ADXOlContextMenuClose_EventHandler(this.adxOutlookEvents_ContextMenuClose);
            this.adxOutlookEvents.ItemLoad += new AddinExpress.MSO.ADXOlItem_EventHandler(this.adxOutlookEvents_ItemLoad);
            this.adxOutlookEvents.BeforeFolderSharingDialog += new AddinExpress.MSO.ADXOlBeforeFolderSharingDialog_EventHandler(this.adxOutlookEvents_BeforeFolderSharingDialog);
            this.adxOutlookEvents.PageChange += new AddinExpress.MSO.ADXOlPageChange_EventHandler(this.adxOutlookEvents_PageChange);
            this.adxOutlookEvents.AutoDiscoverComplete += new System.EventHandler(this.adxOutlookEvents_AutoDiscoverComplete);
            this.adxOutlookEvents.OnGetFormRegionStorage += new AddinExpress.MSO.ADXOlGetFormRegionStorage_EventHandler(this.adxOutlookEvents_OnGetFormRegionStorage);
            this.adxOutlookEvents.OnBeforeFormRegionShow += new AddinExpress.MSO.ADXOlBeforeFormRegionShow_EventHandler(this.adxOutlookEvents_OnBeforeFormRegionShow);
            this.adxOutlookEvents.OnGetFormRegionManifest += new AddinExpress.MSO.ADXOlGetFormRegionManifest_EventHandler(this.adxOutlookEvents_OnGetFormRegionManifest);
            this.adxOutlookEvents.OnGetFormRegionIcon += new AddinExpress.MSO.ADXOlGetFormRegionIcon_EventHandler(this.adxOutlookEvents_OnGetFormRegionIcon);
            this.adxOutlookEvents.SyncError += new AddinExpress.MSO.ADXOlSyncObjectError_EventHandler(this.adxOutlookEvents_SyncError);
            this.adxOutlookEvents.SyncProgress += new AddinExpress.MSO.ADXOlSyncObjectProgress_EventHandler(this.adxOutlookEvents_SyncProgress);
            this.adxOutlookEvents.SyncEnd += new AddinExpress.MSO.ADXOlSyncObjectEnd_EventHandler(this.adxOutlookEvents_SyncEnd);
            this.adxOutlookEvents.SyncStart += new AddinExpress.MSO.ADXOlSyncObjectStart_EventHandler(this.adxOutlookEvents_SyncStart);
            this.adxOutlookEvents.ExplorerAttachmentSelectionChange += new AddinExpress.MSO.ADXOlAttachmentSelectionChange_EventHandler(this.adxOutlookEvents_ExplorerAttachmentSelectionChange);
            this.adxOutlookEvents.ExplorerInlineResponseEx += new AddinExpress.MSO.ADXOlExplorerInlineResponseEx_EventHandler(this.adxOutlookEvents_ExplorerInlineResponseEx);
            this.adxOutlookEvents.ExplorerInlineResponseCloseEx += new AddinExpress.MSO.ADXOlExplorerInlineResponseCloseEx_EventHandler(this.adxOutlookEvents_ExplorerInlineResponseCloseEx);
            this.adxOutlookEvents.InspectorAttachmentSelectionChange += new AddinExpress.MSO.ADXOlAttachmentSelectionChange_EventHandler(this.adxOutlookEvents_InspectorAttachmentSelectionChange);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
            this.saveFileDialog1.InitialDirectory = "c:\\";
            this.saveFileDialog1.RestoreDirectory = true;
            this.saveFileDialog1.Tag = "";
            this.saveFileDialog1.Title = "Select file for save log";
            // 
            // AddinModule
            // 
            this.AddinName = "Add-in Express Events Add-in for Outlook";
            this.Description = "Add-in Express Events Add-in for Outlook";
            this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaOutlook;
            this.AddinInitialize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinInitialize);
            this.AddinStartupComplete += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinStartupComplete);
            this.AddinFinalize += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinFinalize);
            this.AddinBeginShutdown += new AddinExpress.MSO.ADXEvents_EventHandler(this.AddinModule_AddinBeginShutdown);
            this.OnError += new AddinExpress.MSO.ADXError_EventHandler(this.AddinModule_OnError);
            this.OnSendMessage += new AddinExpress.MSO.ADXSendMessage_EventHandler(this.AddinModule_OnSendMessage);
            this.OnKeyDown += new AddinExpress.MSO.ADXKeyDown_EventHandler(this.AddinModule_OnKeyDown);
            this.OnTaskPaneBeforeCreate += new AddinExpress.MSO.ADXTaskPaneBeforeCreate_EventHandler(this.AddinModule_OnTaskPaneBeforeCreate);
            this.OnTaskPaneAfterCreate += new AddinExpress.MSO.ADXTaskPaneAfterCreate_EventHandler(this.AddinModule_OnTaskPaneAfterCreate);
            this.OnTaskPaneBeforeShow += new AddinExpress.MSO.ADXTaskPaneBeforeShow_EventHandler(this.AddinModule_OnTaskPaneBeforeShow);
            this.OnTaskPaneAfterShow += new AddinExpress.MSO.ADXTaskPaneAfterShow_EventHandler(this.AddinModule_OnTaskPaneAfterShow);
            this.OnTaskPaneBeforeDestroy += new AddinExpress.MSO.ADXTaskPaneBeforeDestroy_EventHandler(this.AddinModule_OnTaskPaneBeforeDestroy);
            this.OnRibbonBeforeCreate += new AddinExpress.MSO.ADXRibbonBeforeCreate_EventHandler(this.AddinModule_OnRibbonBeforeCreate);
            this.OnRibbonBeforeLoad += new AddinExpress.MSO.ADXRibbonBeforeLoad_EventHandler(this.AddinModule_OnRibbonBeforeLoad);
            this.OnRibbonLoaded += new AddinExpress.MSO.ADXRibbonLoaded_EventHandler(this.AddinModule_OnRibbonLoaded);
            this.OfficeColorSchemeChanged += new AddinExpress.MSO.OfficeColorSchemeChanged_EventHandler(this.AddinModule_OfficeColorSchemeChanged);

		}


		#endregion

		#region Add-in Express automatic code

		// Required by Add-in Express - do not modify
		// the methods within this region

		public override System.ComponentModel.IContainer GetContainer()
		{
			if (components == null)
				components = new System.ComponentModel.Container();
			return components;
		}

		[ComRegisterFunctionAttribute]
		public static void AddinRegister(Type t)
		{
			AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
		}

		[ComUnregisterFunctionAttribute]
		public static void AddinUnregister(Type t)
		{
			AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
		}

		public override void UninstallControls()
		{
			base.UninstallControls();
		}

		#endregion

		public Outlook._Application OutlookApp
		{
			get
			{
				return (HostApplication as Outlook._Application);
			}
		}

		#region Add-in Code

		internal void SetStartStopLog(bool state)
		{
			StartStopLog = state; ;
		}

		private void ConnectToFolder()
		{
			Outlook._NameSpace ns = OutlookApp.GetNamespace("MAPI");
			if (ns != null)
				try
				{
					Outlook.MAPIFolder inboxFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
					if (inboxFolder != null)
						try
						{
							Outlook.MAPIFolder rootFolder = inboxFolder.Parent as Outlook.MAPIFolder;
							if (rootFolder != null)
								try
								{
									Outlook.MAPIFolder folder1 = ns.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);
									if (folder1 != null)
									{
										OutlookItemsEventsClass1 itemsEventSink = new OutlookItemsEventsClass1(this);
										itemsEventSink.ConnectTo(folder1, true, false);
										itemsEvents.Add(itemsEventSink);
									}

									Outlook.MAPIFolder folder2 = ns.GetFolderFromID(rootFolder.EntryID, rootFolder.StoreID);
									if (folder2 != null)
									{
										OutlookFoldersEventsClass1 foldersEventSink = new OutlookFoldersEventsClass1(this);
										foldersEventSink.ConnectTo(folder2, true, false);
										folderEvents.Add(foldersEventSink);

									}
									ConnectToFolders(ns, rootFolder);
								}
								finally
								{
									Marshal.ReleaseComObject(rootFolder);
								}
						}
						finally
						{
							Marshal.ReleaseComObject(inboxFolder);
						}
					Outlook.MAPIFolder outboxFolder = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox);
					if (outboxFolder != null)
						try
						{
							OutboxFolderEntryID = outboxFolder.EntryID;
						}
						finally
						{
							Marshal.ReleaseComObject(outboxFolder);
						}
				}
				finally
				{
					Marshal.ReleaseComObject(ns);
				}
		}

		private void ConnectToFolders(Outlook._NameSpace ns, Outlook.MAPIFolder rootFolder)
		{
			Outlook.Folders folders = rootFolder.Folders;
			if (folders != null)
				try
				{
					for (int i = 1; i <= folders.Count; i++)
					{
						Outlook.MAPIFolder folder = folders.Item(i);

						if (folder != null)
							try
							{
								Outlook.MAPIFolder folder1 = ns.GetFolderFromID(folder.EntryID, folder.StoreID);
								if (folder1 != null)
								{
									OutlookItemsEventsClass1 itemsEventSink = new OutlookItemsEventsClass1(this);
									itemsEventSink.ConnectTo(folder1, true, false);
									itemsEvents.Add(itemsEventSink);
								}
								Outlook.MAPIFolder folder2 = ns.GetFolderFromID(folder.EntryID, folder.StoreID);
								if (folder2 != null)
								{

									OutlookFoldersEventsClass1 foldersEventSink = new OutlookFoldersEventsClass1(this);
									foldersEventSink.ConnectTo(folder2, true, false);
									folderEvents.Add(foldersEventSink);
								}
								ConnectToFolders(ns, folder);
							}
							finally
							{
								Marshal.ReleaseComObject(folder);
							}
					}
				}
				finally
				{
					Marshal.ReleaseComObject(folders);
				}
		}

		public void DoFolderAdd(Outlook.MAPIFolder folder)
		{
			if (folder != null)
			{
				bool newFolderAdded = true;
				for (int i = folderEvents.Count - 1; i >= 0; i--)
				{
					try
					{
						Outlook.MAPIFolder listFolder = folderEvents[i].FolderObj as Outlook.MAPIFolder;
						if (folder.EntryID == listFolder.EntryID)
						{
							newFolderAdded = false;
							break;
						}
					}
					catch
					{
						folderEvents[i].RemoveConnection();
						folderEvents[i].Dispose();
					}
				}
				if (newFolderAdded)
				{
					Outlook.MAPIFolder newFolder = null;
					Outlook._NameSpace ns = OutlookApp.GetNamespace("MAPI");
					if (ns != null)
						try
						{
							newFolder = ns.GetFolderFromID(folder.EntryID, folder.StoreID);
						}
						finally
						{
							Marshal.ReleaseComObject(ns);
						}
					if (newFolder != null)
					{
						OutlookItemsEventsClass1 itemsEventSink = new OutlookItemsEventsClass1(this);
						itemsEventSink.ConnectTo(newFolder, true, false);
						itemsEvents.Add(itemsEventSink);

						OutlookFoldersEventsClass1 foldersEventSink = new OutlookFoldersEventsClass1(this);
						foldersEventSink.ConnectTo(newFolder, true, false);
						folderEvents.Add(foldersEventSink);
					}
				}
			}
		}

		private bool IsNodeChecked(string NodeName)
		{
			if (NodeName != "")
			{
				bool nodeState = false;
				if (ResultForm != null)
				{
					TreeNode[] tn = ResultForm.treeView1.Nodes.Find(NodeName, true);
					if (tn.Length > 0)
						nodeState = tn[0].Checked;
				}
				return nodeState;
			}
			return true;
		}

		private void WriteLine(string s)
		{
			if (StartStopLog)
				if (ResultForm.textBox1.Text.Length > 0)
				{
					string ss = "= " + DateTime.Now.ToString("T") + s;
					ResultForm.textBox1.Text += Environment.NewLine + ss;
					if (sw != null)
					{
						sw.WriteLine(ss);
					}
				}
				else
				{
					ResultForm.textBox1.Text += "= " + DateTime.Now.ToString("T") + s;
					if (sw != null)
					{
						sw.WriteLine("= " + DateTime.Now.ToString("T") + s);
					}
				}
		}

		public void WriteToLog(string StringRes, string NodeName)
		{
			if (isExplorerActivate)
				if (this.adxOlFormsManager1.Items != null)
				{
					AddinExpress.OL.ADXOlFormsCollectionItem MyFormResultItem = this.adxOlFormsManager1.Items[0];
					if (MyFormResultItem != null)
					{
						//ResultForm = MyFormResultItem.GetCurrentForm(AddinExpress.OL.EmbeddedFormStates.Active) as ADXOlFormAddIn;
                        ResultForm = MyFormResultItem.GetCurrentForm() as ADXOlFormAddIn;
						if (ResultForm != null)
						{
							if (CurrentEvents.Count > 0)
							{
								for (int i = 0; i < CurrentEvents.Count; i++)
								{
									WriteLine(CurrentEvents[i]);
								}
								CurrentEvents.Clear();
								WriteLine("  =  ADXOLForm is created");
							}
							isExplorerActivate = false;
						}
					}
				}
			if (ResultForm != null)
			{
				if (IsNodeChecked(NodeName))
					WriteLine(StringRes);
				ResultForm.SetStateButton();
			}
			else
			{
				if (Convert.ToBoolean(setTreeView[NodeName]))
					CurrentEvents.Add(StringRes);
			}
		}

		public string ItemInfo(object item)
		{
			string s = string.Empty;

			if (item is Outlook.MailItem)
			{
				Outlook.MailItem mail = null;
				mail = item as Outlook.MailItem;
				if (mail != null)
					try
					{
						s += " MailItem with subject '";
						s += mail.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.AppointmentItem)
			{
				Outlook.AppointmentItem appointment = null;
				appointment = item as Outlook.AppointmentItem;
				if (appointment != null)
					try
					{
						s += " AppointmentItem with subject '";
						s += appointment.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.TaskItem)
			{
				Outlook.TaskItem task = null;
				task = item as Outlook.TaskItem;
				if (task != null)
					try
					{
						s += " TaskItem with subject '";
						s += task.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.JournalItem)
			{
				Outlook.JournalItem journal = null;
				journal = item as Outlook.JournalItem;
				if (journal != null)
					try
					{
						s += " JournalItem with subject '";
						s += journal.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.ContactItem)
			{
				Outlook.ContactItem contact = null;
				contact = item as Outlook.ContactItem;
				if (contact != null)
					try
					{
						s += " ContactItem with subject '";
						s += contact.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.PostItem)
			{
				Outlook.PostItem post = null;
				post = item as Outlook.PostItem;
				if (post != null)
					try
					{
						s += " PostItem with subject '";
						s += post.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.NoteItem)
			{
				Outlook.NoteItem note = null;
				note = item as Outlook.NoteItem;
				if (note != null)
					try
					{
						s += " NoteItem with subject '";
						s += note.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}

			if (item is Outlook.DistListItem)
			{
				Outlook.DistListItem distList = null;
				distList = item as Outlook.DistListItem;
				if (distList != null)
					try
					{
						s += " DistListItem with subject '";
						s += distList.Subject;
						s += "'. ";
						return s;
					}
					catch
					{
						return "";
					}
			}
			return s;
		}

		private bool CompareItem(object item, object ItemObj)
		{
			if (itemEvents.Count > 0)
			{
				if ((item is Outlook.MailItem) && (ItemObj is Outlook.MailItem))
					try
					{
						if ((item as Outlook.MailItem).EntryID == (ItemObj as Outlook.MailItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.AppointmentItem) && (ItemObj is Outlook.AppointmentItem))
					try
					{
						if ((item as Outlook.AppointmentItem).EntryID == (ItemObj as Outlook.AppointmentItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.TaskItem) && (ItemObj is Outlook.TaskItem))
					try
					{
						if ((item as Outlook.TaskItem).EntryID == (ItemObj as Outlook.TaskItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.JournalItem) && (ItemObj is Outlook.JournalItem))
					try
					{
						if ((item as Outlook.JournalItem).EntryID == (ItemObj as Outlook.JournalItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.ContactItem) && (ItemObj is Outlook.ContactItem))
					try
					{
						if ((item as Outlook.ContactItem).EntryID == (ItemObj as Outlook.ContactItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.PostItem) && (ItemObj is Outlook.PostItem))
					try
					{
						if ((item as Outlook.PostItem).EntryID == (ItemObj as Outlook.PostItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.NoteItem) && (ItemObj is Outlook.NoteItem))
					try
					{
						if ((item as Outlook.NoteItem).EntryID == (ItemObj as Outlook.NoteItem).EntryID)
							return true;
					}
					catch { return false; }

				if ((item is Outlook.DistListItem) && (ItemObj is Outlook.DistListItem))
					try
					{
						if ((item as Outlook.DistListItem).EntryID == (ItemObj as Outlook.DistListItem).EntryID)
							return true;
					}
					catch { return false; }
			}
			return false;
		}

		private bool isItemEventsConnected(object item)
		{
			for (int i = 0; i < itemEvents.Count; i++)
			{
				if (CompareItem(item, itemEvents[i].ItemObj))
				{
					return true;
				}
			}
			return false;
		}

		private void ConnectToSelectedItem(object explorer)
		{
			Outlook.MAPIFolder currentFolder = null;
			Outlook.Explorer currentExplorer = explorer as Outlook.Explorer;
			if (currentExplorer != null)
				try
				{
					currentFolder = currentExplorer.CurrentFolder;
					if (currentFolder != null)
						if (currentFolder.EntryID != OutboxFolderEntryID)
						{
							Outlook.Selection selection = null;
							try
							{
								selection = currentExplorer.Selection as Outlook.Selection;
								if (selection != null)
									if (selection.Count > 0)
									{
										object item = selection.Item(1);
										if (!isItemEventsConnected(item))
										{
											if (selectedItemEvents == null)
												selectedItemEvents = new OutlookItemEventsClass1(this, true);
											selectedItemEvents.ConnectTo(item, true);
										}
									}
							}
							catch
							{
								// The Explorer has been closed and cannot be used for further operations. Review your code and restart Outlook.
							}
							finally
							{
								if (selection != null)
									Marshal.ReleaseComObject(selection);
							}
						}
				}
				finally
				{
					if (currentFolder != null)
						Marshal.ReleaseComObject(currentFolder);
				}
		}


		internal void WritteToLogFile(bool state)
		{
			if (state)
			{
				if (sw == null)
					if (saveFileDialog1.ShowDialog() == DialogResult.OK)
					{
						sw = new System.IO.StreamWriter(saveFileDialog1.FileName);
						sw.WriteLine(ResultForm.textBox1.Text);
						sw.Flush();
					}
			}
			else
			{
				if (sw != null)
				{
					sw.Close();
					sw = null;
				}
			}
		}

		internal void WriteTreeViewState()
		{
			RegistryKey key = Registry.CurrentUser.OpenSubKey(this.RegistryKey).OpenSubKey("Forms", RegistryKeyPermissionCheck.ReadWriteSubTree);
			if (key != null)
				key = key.CreateSubKey("Nodes", RegistryKeyPermissionCheck.ReadWriteSubTree);

			if (ResultForm != null)
				if (ResultForm.treeView1 != null)
					try
					{
						TreeNodeCollection nodes = ResultForm.treeView1.Nodes;
						foreach (TreeNode currentNode in nodes)
						{
							key.SetValue(currentNode.Name, currentNode.Checked);
							if (currentNode.Nodes.Count > 0)
							{
								WriteNextlevelTreeViewState(key, currentNode);
							}
						}
					}
					catch { }
		}

		private void WriteNextlevelTreeViewState(RegistryKey key, TreeNode node)
		{
			foreach (TreeNode currentNode in node.Nodes)
			{
				key.SetValue(currentNode.Name, currentNode.Checked);
				if (currentNode.Nodes.Count > 0)
				{
					WriteNextlevelTreeViewState(key, currentNode);
				}
			}
		}

		private void WriteSetTreeView()
		{
			RegistryKey key;
			key = Registry.CurrentUser.OpenSubKey(this.RegistryKey + "\\Forms\\Nodes", RegistryKeyPermissionCheck.ReadWriteSubTree);
			if (key != null)
			{
				string[] SubKeys = key.GetValueNames();
				foreach (string s in SubKeys)
					if (s != "Node_CommandBarsUpdate")
						setTreeView.Add(s, key.GetValue(s, true));
					else
						setTreeView.Add(s, key.GetValue(s, false));
			}
		}

		private string sExplorerInfo(object explorer)
		{
			string s = "";
			Outlook.Explorer currentExplorer = explorer as Outlook.Explorer;
			if (currentExplorer != null)
			{
				Outlook.MAPIFolder currFolder = currentExplorer.CurrentFolder as Outlook.MAPIFolder;
				if (currFolder != null)
				{
					s += "Current Folder name is '" + currFolder.Name + "', ";
					Marshal.ReleaseComObject(currFolder);
				}
				s += "Explorer caption is '" + currentExplorer.Caption + "'";
			}
			return s;
		}

		#endregion

		#region AddinModule Events

		private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
		{
			WriteToLog("  =  AddinModule.AddinStartupComplete", "Node_AddinStartupComplete");
		}

		private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
		{
			WriteToLog("  =  AddinModule.AddinBeginShutdown", "Node_AddinBeginShutdown");
			if (selectedItemEvents != null)
				selectedItemEvents.Dispose();
			if (folderEvents != null)
			{
				foreach (OutlookFoldersEventsClass1 folderEventSink in folderEvents)
					folderEventSink.Dispose();
				folderEvents.Clear();
			}
			if (itemsEvents != null)
			{
				foreach (OutlookItemsEventsClass1 itemEventSink in itemsEvents)
					itemEventSink.Dispose();
				itemsEvents.Clear();
			}
			if (itemEvents != null)
			{
				foreach (OutlookItemEventsClass1 itemEventSink in itemEvents)
					itemEventSink.Dispose();
				itemEvents.Clear();
			}
		}

		private void AddinModule_AddinFinalize(object sender, EventArgs e)
		{
			WriteToLog("  =  AddinModule.AddinFinalize", "Node_AddinFinalize");
			if (sw != null)
			{
				sw.Close();
				sw.Dispose();
			}
		}

		private void AddinModule_AddinInitialize(object sender, EventArgs e)
		{
			WriteSetTreeView();
			WriteToLog("  =  AddinModule.AddinInitialize", "Node_AddinInitialize");
			ConnectToFolder();
		}

		private void AddinModule_AfterUninstallControls(object sender, AddinExpress.MSO.ADXOfficeHostApp hostApp, object hostAppObj)
		{
			WriteToLog("  =  AddinModule.AfterUninstallControls", "Node_AfterUninstallControls");
		}

		private void AddinModule_BeforeUninstallControls(object sender, AddinExpress.MSO.ADXOfficeHostApp hostApp, object hostAppObj)
		{
			WriteToLog("  =  AddinModule.BeforeUninstallControls", "Node_BeforeUninstallControls");
		}

		private void AddinModule_OfficeColorSchemeChanged(object sender, AddinExpress.MSO.OfficeColorScheme theme)
		{
			WriteToLog("  =  AddinModule.OfficeColorSchemeChanged", "Node_OfficeColorSchemeChanged");
		}

		private void AddinModule_OnError(AddinExpress.MSO.ADXErrorEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnError, Error is: " + e.ADXError.Message, "Node_OnError");
			e.Handled = true;
		}

		private void AddinModule_OnKeyDown(object sender, AddinExpress.MSO.ADXKeyDownEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnKeyDown", "Node_OnKeyDown");
		}

		private void AddinModule_OnRibbonBeforeCreate(object sender, string ribbonId)
		{
			WriteToLog("  =  AddinModule.OnRibbonBeforeCreate", "Node_OnRibbonBeforeCreate");
		}

		private void AddinModule_OnRibbonBeforeLoad(object sender, AddinExpress.MSO.ADXRibbonBeforeLoadEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnRibbonBeforeLoad", "Node_OnRibbonBeforeLoad");
		}

		private void AddinModule_OnRibbonLoaded(object sender, AddinExpress.MSO.IRibbonUI ribbon)
		{
			WriteToLog("  =  AddinModule.OnRibbonLoaded", "Node_OnRibbonLoaded");
		}

		private void AddinModule_OnSendMessage(object sender, AddinExpress.MSO.ADXSendMessageEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnSendMessage", "Node_OnSendMessagen");
		}

		private void AddinModule_OnTaskPaneAfterCreate(object sender, AddinExpress.MSO.ADXTaskPane.ADXCustomTaskPaneInstance instance, object control)
		{
			WriteToLog("  =  AddinModule.OnTaskPaneAfterCreate", "Node_OnTaskPaneAfterCreate");
		}

		private void AddinModule_OnTaskPaneAfterShow(object sender, AddinExpress.MSO.ADXTaskPane.ADXCustomTaskPaneInstance instance)
		{
			WriteToLog("  =  AddinModule.OnTaskPaneAfterShow", "Node_OnTaskPaneAfterShow");
		}

		private void AddinModule_OnTaskPaneBeforeCreate(object sender, AddinExpress.MSO.ADXTaskPaneCreateEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnTaskPaneBeforeCreate", "Node_OnTaskPaneBeforeCreate");
		}

		private void AddinModule_OnTaskPaneBeforeDestroy(object sender, AddinExpress.MSO.ADXTaskPane.ADXCustomTaskPaneInstance instance)
		{
			WriteToLog("  =  AddinModule.OnTaskPaneBeforeDestroy", "Node_OnTaskPaneBeforeDestroy");
		}

		private void AddinModule_OnTaskPaneBeforeShow(object sender, AddinExpress.MSO.ADXTaskPaneShowEventArgs e)
		{
			WriteToLog("  =  AddinModule.OnTaskPaneBeforeShow", "Node_OnTaskPaneBeforeShow");
		}

		#endregion

		#region Add-in Express Outlook Events

		#region Outlook Explorer(s) Events

		private void adxOutlookEvents_ExplorerBeforeFolderSwitch(object sender, AddinExpress.MSO.ADXOlExplorerBeforeFolderSwitchEventArgs e)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerBeforeFolderSwitch";

			string FolderName = "";
			string OldFolderName = "";

			Outlook.Explorer explorer = e.Explorer as Outlook.Explorer;
			if (explorer != null)
			{
				Outlook.MAPIFolder OldFolder = explorer.CurrentFolder;
				if (OldFolder != null)
				{
					OldFolderName = OldFolder.Name;
					Marshal.ReleaseComObject(OldFolder);
				}
			}

			Outlook.MAPIFolder NewFolder = e.NewFolder as Outlook.MAPIFolder;
			if (NewFolder != null)
			{
				FolderName = NewFolder.Name;
			}
			WriteToLog(s + ", Source Folder name is '" + OldFolderName + "', Destination Folder name is '" + FolderName + "'", "Node_ExplorerBeforeFolderSwitch");
		}

		private void adxOutlookEvents_NewExplorer(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.NewExplorer. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_NewExplorer");
		}

		private void adxOutlookEvents_ExplorerActivate(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerActivate. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_ExplorerActivate");
			isExplorerActivate = true;
			ConnectToSelectedItem(explorer);
		}

		private void adxOutlookEvents_ExplorerAddCommandBars(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerAddCommandBars. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_ExplorerAddCommandBars");
		}

		private void adxOutlookEvents_ExplorerBeforeItemCopy(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeItemCopy", "Node_ExplorerBeforeItemCopy");
		}

		private void adxOutlookEvents_ExplorerBeforeItemCut(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeItemCut", "Node_ExplorerBeforeItemCut");
		}

		private void adxOutlookEvents_ExplorerBeforeItemPaste(object sender, AddinExpress.MSO.ADXOlExplorerBeforeItemPasteEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeItemPaste", "Node_ExplorerBeforeItemPaste");
		}

		private void adxOutlookEvents_ExplorerBeforeMaximize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeMaximize", "Node_ExplorerBeforeMaximize");
		}

		private void adxOutlookEvents_ExplorerBeforeMinimize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeMinimize", "Node_ExplorerBeforeMinimize");
		}

		private void adxOutlookEvents_ExplorerBeforeMove(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeMove", "Node_ExplorerBeforeMove");
		}

		private void adxOutlookEvents_ExplorerBeforeSize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeSize", "Node_ExplorerBeforeSize");
		}

		private void adxOutlookEvents_ExplorerBeforeViewSwitch(object sender, AddinExpress.MSO.ADXOlExplorerBeforeViewSwitchEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerBeforeViewSwitch", "Node_ExplorerBeforeViewSwitch");
		}

		private void adxOutlookEvents_ExplorerClose(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerClose. ";
			Outlook.Explorer currentExplorer = explorer as Outlook.Explorer;
			if (currentExplorer != null)
			{
				s += "Explorer caption is '" + currentExplorer.Caption + "'";
			}
			WriteToLog(s, "Node_ExplorerClose");
		}

		private void adxOutlookEvents_ExplorerDeactivate(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerDeactivate. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_ExplorerDeactivate");
		}

		private void adxOutlookEvents_ExplorerFolderSwitch(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerFolderSwitch. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_ExplorerFolderSwitch");
		}

		private void adxOutlookEvents_ExplorerSelectionChange(object sender, object explorer)
		{
			string s = "  =  ADXOutlookAppEvents.ExplorerSelectionChange. ";
			Outlook.Explorer currentExplorer = explorer as Outlook.Explorer;
			if (currentExplorer != null)
			{
				Outlook.NameSpace ns = currentExplorer.Session as Outlook.NameSpace;
				if (ns != null)
					try
					{
						Outlook.MAPIFolder folder = currentExplorer.CurrentFolder;
						if (folder != null)
							try
							{
								Outlook.MAPIFolder folder2 = ns.GetFolderFromID(folder.EntryID, folder.StoreID);
								if (folder2 != null)
								{
									s += "Current Folder name is '" + folder.Name + "', ";
									bool flagFound = false;
									for (int i = 0; i < this.itemsEvents.Count; i++)
									{
										if ((this.itemsEvents[i].FolderObj as Outlook.MAPIFolder).EntryID == folder.EntryID)
										{

											this.itemsEvents[i].RemoveConnection();
											this.itemsEvents[i].ConnectTo(folder2, true, false);
											flagFound = true;
											break;
										}
									}
									if (!flagFound)
									{
										OutlookItemsEventsClass1 eventSink = new OutlookItemsEventsClass1(this);
										eventSink.ConnectTo(folder2, true, false);
										this.itemsEvents.Add(eventSink);
									}
								}
							}
							finally
							{
								Marshal.ReleaseComObject(folder);
							}
						s += "Explorer caption is '" + currentExplorer.Caption + "'";
					}
					finally
					{
						Marshal.ReleaseComObject(ns);
					}
			}
			WriteToLog(s, "Node_ExplorerSelectionChange");
			ConnectToSelectedItem(explorer);
		}

		private void adxOutlookEvents_ExplorerViewSwitch(object sender, object explorer)
		{
			string s = "  =  DXOutlookAppEvents.ExplorerViewSwitch. ";
			s += sExplorerInfo(explorer);
			WriteToLog(s, "Node_ExplorerViewSwitch");
		}

		private void adxOutlookEvents_ExplorerAttachmentSelectionChange(object sender, object sourceObject)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ExplorerAttachmentSelectionChange", "Node_ExplorerAttachmentSelectionChange");
		}

        private void adxOutlookEvents_ExplorerInlineResponseEx(object sender, object itemObject, object sourceObject)
        {
            WriteToLog("  =  ADXOutlookAppEvents.ExplorerInlineResponse", "Node_ExplorerInlineResponse");
        }

        private void adxOutlookEvents_ExplorerInlineResponseCloseEx(object sender, object sourceObject)
        {
            WriteToLog("  =  ADXOutlookAppEvents.ExplorerInlineResponseClose", "Node_ExplorerInlineResponseClose");
        }

		#endregion

		#region Outlook Inspector(s) Events

		private void adxOutlookEvents_NewInspector(object sender, object inspector, string folderName)
		{
			Outlook._Inspector olInsp = inspector as Outlook._Inspector;
			WriteToLog("  =  ADXOutlookAppEvents.NewInspector. Inspector caption is '" + olInsp.Caption + "', Folder name is '" + folderName + "'", "Node_NewInspector");

			object item = olInsp.CurrentItem;

			if (item != null)
			{
				OutlookItemEventsClass1 itemEventSink = new OutlookItemEventsClass1(this, false);
				itemEventSink.ConnectTo(item, true);
				itemEvents.Add(itemEventSink);
				if (selectedItemEvents != null)
					if (CompareItem(item, selectedItemEvents.ItemObj))
					{
						selectedItemEvents.Dispose();
						selectedItemEvents = null;
					}
			}
		}

		private void adxOutlookEvents_InspectorActivate(object sender, object inspector, string folderName)
		{
			Outlook._Inspector olInsp = inspector as Outlook._Inspector;
			WriteToLog("  =  ADXOutlookAppEvents.InspectorActivate. Inspector caption is '" + olInsp.Caption + "', Folder name is '" + folderName + "'", "Node_InspectorActivate");
		}

		private void adxOutlookEvents_InspectorAddCommandBars(object sender, object inspector, string folderName)
		{
			Outlook._Inspector olInsp = inspector as Outlook._Inspector;
			WriteToLog("  =  ADXOutlookAppEvents.InspectorAddCommandBars. Inspector caption is '" + olInsp.Caption + "', Folder name is '" + folderName + "'", "Node_InspectorAddCommandBars");
		}

		private void adxOutlookEvents_InspectorBeforeMaximize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.InspectorBeforeMaximize", "Node_InspectorBeforeMaximize");
		}

		private void adxOutlookEvents_InspectorBeforeMinimize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.InspectorBeforeMinimize", "Node_InspectorBeforeMinimize");
		}

		private void adxOutlookEvents_InspectorBeforeMove(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.InspectorBeforeMove", "Node_InspectorBeforeMove");
		}

		private void adxOutlookEvents_InspectorBeforeSize(object sender, AddinExpress.MSO.ADXHostActionEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.InspectorBeforeSize", "Node_InspectorBeforeSize");
		}

		private void adxOutlookEvents_InspectorClose(object sender, object inspector, string folderName)
		{
			Outlook._Inspector olInsp = inspector as Outlook._Inspector;
			WriteToLog("  =  ADXOutlookAppEvents.InspectorClose. Inspector caption is '" + olInsp.Caption + "', Folder name is '" + folderName + "'", "Node_InspectorClose");
		}

		private void adxOutlookEvents_InspectorDeactivate(object sender, object inspector, string folderName)
		{
			Outlook._Inspector olInsp = inspector as Outlook._Inspector;
			WriteToLog("  =  ADXOutlookAppEvents.InspectorDeactivate. Inspector caption is '" + olInsp.Caption + "', Folder name is '" + folderName + "'", "Node_InspectorDeactivate");
		}

		private void adxOutlookEvents_PageChange(object sender, AddinExpress.MSO.ADXOlPageChangeEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.PageChange", "Node_PageChange");
		}

		private void adxOutlookEvents_InspectorAttachmentSelectionChange(object sender, object sourceObject)
		{
			WriteToLog("  =  ADXOutlookAppEvents.InspectorAttachmentSelectionChange", "Node_InspectorAttachmentSelectionChange");
		}

		#endregion

		#region Outlook Reminder Events

		private void adxOutlookEvents_BeforeReminderShow(object sender, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.BeforeReminderShow", "Node_BeforeReminderShow");
		}

		private void adxOutlookEvents_ReminderAdd(object sender, object reminderObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ReminderAdd", "Node_ReminderAdd");
		}

		private void adxOutlookEvents_ReminderChange(object sender, object reminderObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ReminderChange", "Node_ReminderChange");
		}

		private void adxOutlookEvents_ReminderFire(object sender, object reminderObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ReminderFire", "Node_ReminderFire");
		}

		private void adxOutlookEvents_ReminderRemove(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ReminderRemove", "Node_ReminderRemove");
		}

		private void adxOutlookEvents_Snooze(object sender, object reminderObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.Snooze", "Node_Snooze");
		}

		#endregion

		#region Outlook Application Events

		private void adxOutlookEvents_ItemSend(object sender, AddinExpress.MSO.ADXOlItemSendEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ItemSend. " + ItemInfo(e.Item), "Node_ItemSend");
		}

		private void adxOutlookEvents_ItemLoad(object sender, object item)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ItemLoad. " + ItemInfo(item), "Node_ItemLoad");
		}

		private void adxOutlookEvents_NewMail(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.NewMail. ", "Node_NewMail");
		}

		private void adxOutlookEvents_NewMailEx(object sender, string entryIDCollection)
		{
			WriteToLog("  =  ADXOutlookAppEvents.NewMailEx", "Node_NewMailEx");
		}

		private void adxOutlookEvents_MAPILogonComplete(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.MAPILogonComplete. ", "Node_MAPILogonComplete");
		}

		private void adxOutlookEvents_Startup(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.Startup", "Node_Startup");
		}

		private void adxOutlookEvents_Quit(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.Quit", "Node_Quit");
		}

		private void adxOutlookEvents_AutoDiscoverComplete(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.AutoDiscoverComplete", "Node_AutoDiscoverComplete");
		}

		private void adxOutlookEvents_BeforeOptionPageAdd(object sender, AddinExpress.MSO.ADXOptionPageAddEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.BeforeOptionPageAdd", "Node_BeforeOptionPageAdd");
		}

		private void adxOutlookEvents_OptionPagesAdd(object sender, object optionPages)
		{
			WriteToLog("  =  ADXOutlookAppEvents.OptionPagesAdd", "Node_OptionPagesAdd");
		}

		private void adxOutlookEvents_Reminder(object sender, object item)
		{
			WriteToLog("  =  ADXOutlookAppEvents.Reminder. " + ItemInfo(item), "Node_Reminder");
		}

		private void adxOutlookEvents_BeforeFolderSharingDialog(object sender, AddinExpress.MSO.ADXOlBeforeFolderSharingDialogEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.BeforeFolderSharingDialog", "Node_BeforeFolderSharingDialog");
		}

		private void adxOutlookEvents_AdvancedSearchComplete(object sender, object hostObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.AdvancedSearchComplete", "Node_AdvancedSearchComplete");
		}

		private void adxOutlookEvents_AdvancedSearchStopped(object sender, object hostObj)
		{
			WriteToLog("  =  ADXOutlookAppEvents.AdvancedSearchStopped", "Node_AdvancedSearchStopped");
		}

		#endregion

		#region Outlook Namespace Events

		private void adxOutlookEvents_NamespaceBeforeOptionPageAdd(object sender, object folder, AddinExpress.MSO.ADXOptionPageAddEventArgs e)
		{
			string s = string.Empty;
			if (folder != null)
				s = "Folder name is '" + (folder as Outlook.MAPIFolder).Name + "'";
			WriteToLog("  =  ADXOutlookAppEvents.NamespaceBeforeOptionPageAdd. " + s, "Node_NamespaceBeforeOptionPageAdd");
		}

		private void adxOutlookEvents_NamespaceOptionPagesAdd(object sender, object optionPages, object folder)
		{
			string s = string.Empty;
			if (folder != null)
				s = "Folder name is '" + (folder as Outlook.MAPIFolder).Name + "'";
			WriteToLog("  =  ADXOutlookAppEvents.NamespaceOptionPagesAdd. " + s, "Node_NamespaceOptionPagesAdd");
		}

		#endregion

		#region Outlook FormsRegion Events

		private void adxOutlookEvents_OnBeforeFormRegionShow(object sender, object formRegion)
		{
			WriteToLog("  =  ADXOutlookAppEvents.OnBeforeFormRegionShow", "Node_OnBeforeFormRegionShow");
		}

		private void adxOutlookEvents_OnGetFormRegionIcon(object sender, AddinExpress.MSO.ADXOlGetFormRegionIconEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.OnGetFormRegionIcon", "Node_OnGetFormRegionIcon");
		}

		private void adxOutlookEvents_OnGetFormRegionManifest(object sender, AddinExpress.MSO.ADXOlGetFormRegionManifestEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.OnGetFormRegionManifest", "Node_OnGetFormRegionManifest");
		}

		private void adxOutlookEvents_OnGetFormRegionStorage(object sender, AddinExpress.MSO.ADXOlGetFormRegionStorageEventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.OnGetFormRegionStorage", "Node_OnGetFormRegionStorage");
		}

		#endregion

		#region Outlook Context Menus Events

		private void adxOutlookEvents_AttachmentContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.AttachmentContextMenuDisplay", "Node_AttachmentContextMenuDisplay");
		}

		private void adxOutlookEvents_ContextMenuClose(object sender, AddinExpress.MSO.ADXOlContextMenu contextMenu)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ContextMenuClose", "Node_ContextMenuClose");
		}

		private void adxOutlookEvents_FolderContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.FolderContextMenuDisplayr", "Node_FolderContextMenuDisplay");
		}

		private void adxOutlookEvents_ItemContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ItemContextMenuDisplay", "Node_ItemContextMenuDisplay");
		}

		private void adxOutlookEvents_ShortcutContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ShortcutContextMenuDisplay", "Node_ShortcutContextMenuDisplay");
		}

		private void adxOutlookEvents_StoreContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.StoreContextMenuDisplay", "Node_StoreContextMenuDisplay");
		}

		private void adxOutlookEvents_ViewContextMenuDisplay(object sender, object commandBar, object target)
		{
			WriteToLog("  =  ADXOutlookAppEvents.ViewContextMenuDisplay", "Node_ViewContextMenuDisplay");
		}

		#endregion

		#region Outlook SyncObject Events

		private void adxOutlookEvents_SyncEnd(object sender, object syncObject)
		{
			WriteToLog("  =  ADXOutlookAppEvents.SyncEnd", "Node_SyncEnd");
		}

		private void adxOutlookEvents_SyncError(object sender, object syncObject, int code, string description)
		{
			WriteToLog("  =  ADXOutlookAppEvents.SyncError. Description is " + description, "Node_SyncError");
		}

		private void adxOutlookEvents_SyncProgress(object sender, object syncObject, AddinExpress.MSO.ADXOlSyncState state, string description, int value, int max)
		{
			WriteToLog("  =  ADXOutlookAppEvents.SyncProgress. Description is " + description, "Node_SyncProgress");
		}

		private void adxOutlookEvents_SyncStart(object sender, object syncObject)
		{
			WriteToLog("  =  ADXOutlookAppEvents.SyncStart", "Node_SyncStart");
		}

		#endregion

		private void adxOutlookEvents_CommandBarsUpdate(object sender, EventArgs e)
		{
			WriteToLog("  =  ADXOutlookAppEvents.CommandBarsUpdate", "Node_CommandBarsUpdate");
		}

		#endregion

		#region Add-in Express Outlook Forms Manager Events

		private void adxOlFormsManager1_ADXAfterAccessProtectedObject(object sender, EventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXAfterAccessProtectedObject", "Node_ADXAfterAccessProtectedObject");
		}

		private void adxOlFormsManager1_ADXBeforeAccessProtectedObject(object sender, EventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXBeforeAccessProtectedObject", "Node_ADXBeforeAccessProtectedObject");
		}

		private void adxOlFormsManager1_ADXBeforeFolderSwitch(object explorerObj, AddinExpress.OL.ADXOlFormsCollectionItem SrcItem, object SrcFolder, AddinExpress.OL.ADXOlFormsCollectionItem DstItem, object DstFolder)
		{
			string s = "  =  ADXOlFormsManager.ADXBeforeFolderSwitch. ";
			if ((SrcFolder as Outlook.MAPIFolder) != null)
			{
				s += "Source Folder name is '" + (SrcFolder as Outlook.MAPIFolder).Name + "'. ";
			}
			if ((DstFolder as Outlook.MAPIFolder) != null)
			{
				s += "Destination Folder name is '" + (DstFolder as Outlook.MAPIFolder).Name + "'.";
			}
			WriteToLog(s, "Node_ADXBeforeFolderSwitch");
		}

		private void adxOlFormsManager1_ADXBeforeFolderSwitchEx(object sender, AddinExpress.OL.BeforeFolderSwitchExEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXBeforeFolderSwitchEx", "Node_ADXBeforeFolderSwitchEx");
		}

		private void adxOlFormsManager1_ADXBeforeFormInstanceCreate(object sender, AddinExpress.OL.BeforeFormInstanceCreateEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXBeforeFormInstanceCreate", "Node_ADXBeforeFormInstanceCreate");
		}

		private void adxOlFormsManager1_ADXFolderSwitch(object sender, AddinExpress.OL.FolderSwitchEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXFolderSwitch", "Node_ADXFolderSwitch");
		}

		private void adxOlFormsManager1_ADXFolderSwitchEx(object sender, AddinExpress.OL.FolderSwitchExEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXFolderSwitchEx", "Node_ADXFolderSwitchEx");
		}

		private void adxOlFormsManager1_ADXNavigationPaneHide(object sender, AddinExpress.OL.NavigationPaneHideEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXNavigationPaneHide", "Node_ADXNavigationPaneHide");
		}

		private void adxOlFormsManager1_ADXNavigationPaneMinimize(object sender, AddinExpress.OL.NavigationPaneMinimizeEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXNavigationPaneMinimize", "Node_ADXNavigationPaneMinimize");
		}

		private void adxOlFormsManager1_ADXNavigationPaneShow(object sender, AddinExpress.OL.NavigationPaneShowEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXNavigationPaneShow", "Node_ADXNavigationPaneShow");
		}

		private void adxOlFormsManager1_ADXNewInspector(object inspectorObj)
		{
			string s = "  =  ADXOlFormsManager.ADXNewInspector";
			Outlook.Inspector inspector = inspectorObj as Outlook.Inspector;
			if (inspector != null)
				s += ",  Inspector Caption  is - " + inspector.Caption;
			WriteToLog(s, "Node_ADXNewInspector");
		}

		private void adxOlFormsManager1_ADXReadingPaneHide(object sender, AddinExpress.OL.ReadingPaneHideEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXReadingPaneHide", "Node_ADXReadingPaneHide");
		}

		private void adxOlFormsManager1_ADXReadingPaneMove(object sender, AddinExpress.OL.ReadingPaneMoveEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXReadingPaneMove", "Node_ADXReadingPaneMove");
		}

		private void adxOlFormsManager1_ADXReadingPaneShow(object sender, AddinExpress.OL.ReadingPaneShowEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXReadingPaneShow", "Node_ADXReadingPaneShow");
		}

		private void adxOlFormsManager1_ADXTodoBarHide(object sender, AddinExpress.OL.TodoBarHideEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXTodoBarHide", "Node_ADXTodoBarHide");
		}

		private void adxOlFormsManager1_ADXTodoBarMinimize(object sender, AddinExpress.OL.TodoBarMinimizeEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXTodoBarMinimize", "Node_ADXTodoBarMinimize");
		}

		private void adxOlFormsManager1_ADXTodoBarShow(object sender, AddinExpress.OL.TodoBarShowEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.ADXTodoBarShow", "Node_ADXTodoBarShow");
		}

		private void adxOlFormsManager1_OnError(object sender, AddinExpress.OL.ErrorEventArgs args)
		{
			WriteToLog("  =  ADXOlFormsManager.OnError, Error is: " + args.Exception.Message, "Node_OlFormsManagerOnError");
			args.Handled = true;
		}

		private void adxOlFormsManager1_OnInitialize()
		{
			WriteToLog("  =  ADXOlFormsManager.OnInitialize", "Node_OlFormsManagerOnInitialize");
		}

		#endregion
	}
}
