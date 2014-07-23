using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;

namespace MarkdownOutlook
{
    public partial class MarkdownRibbon
    {
        private RenderedMarkdownForm markdownForm;

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            enableMarkdownMode.Checked = ThisAddIn.CachedMarkdownEnabled;
            markdownForm = new RenderedMarkdownForm();
        }

        private void enableMarkdownMode_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.CachedMarkdownEnabled = enableMarkdownMode.Checked;
            var context = e.Control.Context;
            var currentMailItem = GetMailItem(context);

            Utility.SetUserProperty(currentMailItem, Constants.EnableMarkdownModeFlag, enableMarkdownMode.Checked.ToString());

        }

        private static MailItem GetMailItem(dynamic context)
        {
            MailItem currentMailItem = null;
            if (context is Inspector)
            {
                Inspector inspector = context;
                currentMailItem = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem as MailItem;
            }
            else if (context is Explorer)
            {
                Explorer explorer = context;
                currentMailItem = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1] as MailItem;
            }
            return currentMailItem;
        }

        private void renderMarkdown_Click(object sender, RibbonControlEventArgs e)
        {
            var context = e.Control.Context;
            MailItem currentMailItem = GetMailItem(context);
            var renderedText = ThisAddIn.RenderMarkdown(currentMailItem.Body);
            markdownForm.webBrowser.DocumentText = renderedText;
            markdownForm.ShowDialog();

        }


    }
}
