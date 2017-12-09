using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;

namespace MarkdownOutlook
{
    public static class Utility
    {
        #region SetUserProperty
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

        public static void SetUserProperty<T>(AppointmentItem currentAppItem, string propName, T propValue)
        {
            if (currentAppItem == null)
            {
                return;
            }

            var markdownModeProperty = currentAppItem.UserProperties.Find(propName);
            if (markdownModeProperty != null)
            {
                markdownModeProperty.Value = propValue.ToString();
            }
            else
            {
                var prop = currentAppItem.UserProperties.Add(Constants.EnableMarkdownModeFlag, OlUserPropertyType.olText);
                prop.Value = propValue.ToString();
            }
        }

        #endregion

        #region GetUserProperty

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

        public static T GetUserProperty<T>(AppointmentItem currentAppItem, string propName)
        {
            if (currentAppItem == null)
            {
                return default(T);
            }

            var markdownModeProperty = currentAppItem.UserProperties.Find(propName);

            if (markdownModeProperty != null)
            {
                return Convert.ChangeType(markdownModeProperty.Value, typeof(T));
            }

            return default(T);
        }

        public static T GetUserProperty<T>(MeetingItem currentMeetingItem, string propName)
        {
            if (currentMeetingItem == null)
            {
                return default(T);
            }

            var markdownModeProperty = currentMeetingItem.UserProperties.Find(propName);

            if (markdownModeProperty != null)
            {
                return Convert.ChangeType(markdownModeProperty.Value, typeof(T));
            }

            return default(T);
        }

        #endregion

    }
}
