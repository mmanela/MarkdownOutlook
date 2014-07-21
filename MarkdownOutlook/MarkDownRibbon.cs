using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace MarkdownOutlook
{
    public partial class MarkdownRibbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            enableMarkdownMode.Checked = ThisAddIn.MarkdownEnabled;
        }

        private void enableMarkdownMode_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.MarkdownEnabled = enableMarkdownMode.Checked;
        }
    }
}
