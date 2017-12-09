using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace MarkdownOutlook
{
    public partial class AppointmentRibbon
    {
        private RenderedMarkdownForm _markdownForm;

        private void AppointmentRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            enableMarkdownMode.Checked = ThisAddIn.CachedMarkdownEnabled;
            _markdownForm = new RenderedMarkdownForm();
        }

        private void enableMarkdownMode_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.CachedMarkdownEnabled = enableMarkdownMode.Checked;
            var context = e.Control.Context;
            var currentAppItem = GetAppointmentItem(context);

            Utility.SetUserProperty(currentAppItem, Constants.EnableMarkdownModeFlag, enableMarkdownMode.Checked.ToString());
        }

        private static AppointmentItem GetAppointmentItem(dynamic context)
        {
            AppointmentItem currentAppItem = null;
            if (context is Inspector)
            {
                currentAppItem = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem as AppointmentItem;
            }
            else if (context is Explorer)
            {
                currentAppItem = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1] as AppointmentItem;
            }

            return currentAppItem;
        }

        private void RenderMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            var context = e.Control.Context;
            AppointmentItem currentAppItem = GetAppointmentItem(context);
            var renderedText = ThisAddIn.RenderMarkdown(currentAppItem.Body);
            _markdownForm.webBrowser.DocumentText = renderedText;
            _markdownForm.ShowDialog();
        }
    }
}
