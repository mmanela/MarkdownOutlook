using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace MarkdownOutlook
{
    public static class Utility
    {
        public static void SetUserProperty<T>(MailItem currentMailItem, string propName, T propValue)
        {
            if (currentMailItem == null)
            {
                return;
            }

            var markdownModeProperty = currentMailItem.UserProperties.Find(propName);
            if (markdownModeProperty != null)
            {
                markdownModeProperty.Value = propValue.ToString();
            }
            else
            {
                var prop = currentMailItem.UserProperties.Add(Constants.EnableMarkdownModeFlag, OlUserPropertyType.olText);
                prop.Value = propValue.ToString();
            }
        }

        public static T GetUserProperty<T>(MailItem currentMailItem, string propName)
        {
            if (currentMailItem == null)
            {
                return default(T);
            }


            var markdownModeProperty = currentMailItem.UserProperties.Find(propName);
            if (markdownModeProperty != null)
            {
                return Convert.ChangeType(markdownModeProperty.Value, typeof(T));
            }

            return default(T);
        }
    }
}
