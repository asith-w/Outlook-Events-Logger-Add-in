using System;

namespace OutlookEvents
{
	/// <summary>
	/// Add-in Express Outlook Items Events Class
	/// </summary>
	public class OutlookItemsEventsClass1 : AddinExpress.MSO.ADXOutlookItemsEvents
	{
		private AddinModule CurrentModule = null;

		public OutlookItemsEventsClass1(AddinExpress.MSO.ADXAddinModule module)
			: base(module)
		{
			if (CurrentModule == null)
				CurrentModule = module as AddinModule;
		}

		public override void ProcessItemAdd(object item)
		{
            string s = "  =  ADXOutlookItemsEvents.ItemAdd ";
            if ((FolderObj as Outlook.MAPIFolder).EntryID != CurrentModule.OutboxFolderEntryID)
                s += CurrentModule.ItemInfo(item);
            s += " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            CurrentModule.WriteToLog(s, "Node_ProcessItemAdd");
		}

		public override void ProcessItemChange(object item)
		{

            string s = "  =  ADXOutlookItemsEvents.ItemChange ";
            if ((FolderObj as Outlook.MAPIFolder).EntryID != CurrentModule.OutboxFolderEntryID)
                s += CurrentModule.ItemInfo(item);
            s += " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            CurrentModule.WriteToLog(s, "Node_ProcessItemChange");
		}

		public override void ProcessItemRemove()
		{
            string s = " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            CurrentModule.WriteToLog("  =  ADXOutlookItemsEvents.ItemRemove" + s, "Node_ProcessItemRemove");
		}

		public override void ProcessBeforeFolderMove(object moveTo, AddinExpress.MSO.ADXCancelEventArgs e) 
		{
			string s = "  =  ADXOutlookItemsEvents.BeforeFolderMove. Folder: '";
			if (moveTo != null)
				s += (moveTo as Outlook.MAPIFolder).Name;
			else
				s += "null";
			s += "'. ";
            s += " Parent Folder name is " + (FolderObj as Outlook.MAPIFolder).Name + ".";
			CurrentModule.WriteToLog(s, "Node_ProcessBeforeFolderMove");
		}

		public override void ProcessBeforeItemMove(object item, object moveTo, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			string s = "  =  ADXOutlookItemsEvents.BeforeItemMove. ";
			s += CurrentModule.ItemInfo(item);
            s += " Folder: " + (FolderObj as Outlook.MAPIFolder).Name + ".";
            s += " Destination Folder: '";
			if (moveTo != null)
				s += (moveTo as Outlook.MAPIFolder).Name;
			else
				s += "null";
			s += "'. ";
			CurrentModule.WriteToLog(s, "Node_ProcessBeforeItemMove");
		}
	}
}

