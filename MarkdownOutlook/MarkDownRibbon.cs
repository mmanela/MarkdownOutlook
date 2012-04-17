using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace MarkdownOutlook
{
    public partial class MarkdownRibbon
    {
        public bool MarkdownEnabled { get; set; }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void enableMarkdownMode_Click(object sender, RibbonControlEventArgs e)
        {
            MarkdownEnabled = enableMarkdownMode.Checked;
        }
    }
}
