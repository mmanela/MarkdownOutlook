using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Markdig;
using Markdig.SyntaxHighlighting;
using Microsoft.Office.Tools.Ribbon;

namespace MarkdownOutlook
{
    public partial class ThisAddIn
    {
        private static MarkdownPipeline _pipeline;

        public static bool CachedMarkdownEnabled { get; set; }

        public static string RenderMarkdown(string text)
        {
           return Markdown.ToHtml(text, _pipeline);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            _pipeline = new MarkdownPipelineBuilder()
                .UseAdvancedExtensions()
                .UseSyntaxHighlighting()
                .Build();
            this.Application.ItemSend += Application_ItemSend;
        }

        protected override IRibbonExtension[] CreateRibbonObjects()
        {
            var allRibbons =
                new IRibbonExtension[2];
            allRibbons[0] = new MarkdownRibbon();
            allRibbons[1] = new AppointmentRibbon();
            return allRibbons;
        }

        void Application_ItemSend(object item, ref bool cancel)
        {
            var useMarkdown = false;
            var mailItem = item as MailItem;
            var meetingItem= item as MeetingItem;

            if (mailItem != null)
            {
                useMarkdown = Utility.GetUserProperty<bool>(mailItem, Constants.EnableMarkdownModeFlag);
                if (useMarkdown)
                {
                    mailItem.HTMLBody = RenderMarkdown(mailItem.Body);
                }
            } else if (meetingItem != null)
            {
                useMarkdown = Utility.GetUserProperty<bool>(meetingItem, Constants.EnableMarkdownModeFlag);
                if (useMarkdown)
                {
                    // Use web browser to convert HTML to RTF
                    var webBrowser = new WebBrowser();
                    var htmlText = RenderMarkdown(meetingItem.Body);
                    webBrowser.CreateControl();
                    webBrowser.DocumentText = htmlText;
                    while (webBrowser.DocumentText != htmlText)
                        System.Windows.Forms.Application.DoEvents();
                    if (webBrowser.Document != null)
                    {
                        webBrowser.Document.ExecCommand("SelectAll", false, null);
                        webBrowser.Document.ExecCommand("Copy", false, null);
                    }
                    var i = new RichTextBox();
                    i.Paste();

                    meetingItem.RTFBody = Encoding.ASCII.GetBytes(i.Rtf);
                    webBrowser.Dispose();
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
