using System;

namespace OutlookEvents
{
	/// <summary>
	/// Add-in Express Outlook Folders Events Class
	/// </summary>
	public class OutlookFoldersEventsClass1 : AddinExpress.MSO.ADXOutlookFoldersEvents
	{
		private AddinModule CurrentModule = null;

		public OutlookFoldersEventsClass1(AddinExpress.MSO.ADXAddinModule module)
			: base(module)
		{
			if (CurrentModule == null)
				CurrentModule = module as AddinModule;
		}

		public override void ProcessFolderAdd(object folder)
		{
            string s = "  =  ADXOutlookFoldersEvents.FolderAdd.";
            s += " Folder name is " + (folder as Outlook.MAPIFolder).Name + ".";
            s += " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            CurrentModule.WriteToLog(s, "Node_ProcessFolderAdd");
			CurrentModule.DoFolderAdd(folder as Outlook.MAPIFolder);
		}

		public override void ProcessFolderChange(object folder)
		{
            string s = "  =  ADXOutlookFoldersEvents.FolderChanged.";
            s += " Folder name is " + (folder as Outlook.MAPIFolder).Name + ".";
            s += " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            CurrentModule.WriteToLog(s, "Node_ProcessFolderChange");
		}

		public override void ProcessFolderRemove()
		{
			CurrentModule.WriteToLog("  =  ADXOutlookFoldersEvents.FolderRemove.", "Node_ProcessFolderRemove");
		}
	}
}

