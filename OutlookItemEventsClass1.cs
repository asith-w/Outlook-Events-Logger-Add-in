using System;

namespace OutlookEvents
{
	/// <summary>
	/// Add-in Express Outlook Item Events Class
	/// </summary>
	public class OutlookItemEventsClass1 : AddinExpress.MSO.ADXOutlookItemEvents
	{
		private AddinModule CurrentModule = null;
        private bool isSelectedChanged = false;

		public OutlookItemEventsClass1(AddinExpress.MSO.ADXAddinModule module, bool isSelectedChanged)
			: base(module)
		{
            this.isSelectedChanged = isSelectedChanged;
			if (CurrentModule == null)
				CurrentModule = module as AddinModule;
		}

		public override void ProcessAttachmentAdd(object attachment)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.AttachmentAdd", "Node_ProcessAttachmentAdd");
		}

		public override void ProcessAttachmentRead(object attachment)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.AttachmentRead", "Node_ProcessAttachmentRead");
		}

		public override void ProcessBeforeAttachmentSave(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAttachmentSave", "Node_ProcessBeforeAttachmentSave");
		}

		public override void ProcessBeforeCheckNames(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeCheckNames", "Node_ProcessBeforeCheckNames");
		}

		public override void ProcessClose(AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Close. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessClose");
            if (!isSelectedChanged)
            {
                CurrentModule.itemEvents.Remove(this);
                this.Dispose();
            }
		}

		public override void ProcessCustomAction(object action, object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.CustomAction", "Node_ProcessCustomAction");
		}

		public override void ProcessCustomPropertyChange(string name)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.CustomPropertyChange. CustomProperty name is " + name, "Node_ProcessCustomPropertyChange");
		}

		public override void ProcessForward(object forward, AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Forward. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessForward");
		}

		public override void ProcessOpen(AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Open. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessOpen");
		}

		public override void ProcessPropertyChange(string name)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.PropertyChange. Property name is " + name, "Node_ProcessPropertyChange");
		}

		public override void ProcessRead()
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Read. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessRead");
		}

		public override void ProcessReply(object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Reply. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessReply");
		}

		public override void ProcessReplyAll(object response, AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.ReplyAll. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessReplyAll");
		}

		public override void ProcessSend(AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Send. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessSend");
		}

		public override void ProcessWrite(AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Write. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessWrite");
		}

		public override void ProcessBeforeDelete(object item, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeDelete " + CurrentModule.ItemInfo(item), "Node_ProcessBeforeDelete");
		}

		public override void ProcessAttachmentRemove(object attachment)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.AttachmentRemove", "Node_ProcessAttachmentRemove");
		}

		public override void ProcessBeforeAttachmentAdd(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAttachmentAdd", "Node_ProcessBeforeAttachmentAdd");
		}

		public override void ProcessBeforeAttachmentPreview(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAttachmentPreview", "Node_ProcessBeforeAttachmentPreview");
		}

		public override void ProcessBeforeAttachmentRead(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAttachmentRead", "Node_ProcessBeforeAttachmentRead");
		}

		public override void ProcessBeforeAttachmentWriteToTempFile(object attachment, AddinExpress.MSO.ADXCancelEventArgs e)
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAttachmentWriteToTempFile", "Node_BeforeAttachmentWriteToTempFile");
		}

		public override void ProcessUnload()
		{
			CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.Unload" , "Node_ProcessUnload");
		}

		public override void ProcessBeforeAutoSave(AddinExpress.MSO.ADXCancelEventArgs e)
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeAutoSave. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessBeforeAutoSave");
		}

		public override void ProcessAfterWrite()
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.AfterWrite. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessAfterWrite");
		}

		public override void ProcessBeforeRead()
		{
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.BeforeRead. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessBeforeRead");
		}

        public override void ProcessReadComplete(AddinExpress.MSO.ADXCancelEventArgs e)
        {
            CurrentModule.WriteToLog("  =  ADXOutlookItemEvents.ReadComplete. " + CurrentModule.ItemInfo(this.ItemObj), "Node_ProcessReadComplete");
        }
	}
}

