using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Reflection;     // to use Missing.Value
using DateLine.Properties;
using System.Windows.Forms;
using System.Windows;
using NLog;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace DateLine
{
    internal sealed class OutlookHelper
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private static readonly OutlookHelper mMeMyself = new OutlookHelper();
        Outlook.Application oApp = null;
        Outlook.NameSpace oNS = null;
        Outlook.MAPIFolder oCalendar = null;

        public static OutlookHelper Instance
        {
            get { return mMeMyself; }
        }

        private OutlookHelper()
        {

        }

        public bool Initialize()
        {
            bool bResult = false;
            try
            {
                // Create the Outlook application.
                oApp = new Outlook.Application();

                // Get the NameSpace and Logon information.
                // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
                oNS = oApp.GetNamespace("MAPI");

                //Log on by using a dialog box to choose the profile.
                //oNS.Logon(Missing.Value, Missing.Value, true, true);

                //Alternate logon method that uses a specific profile.
                // TODO: If you use this logon method, 
                // change the profile name to an appropriate value.
                //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

                // Get the Calendar folder.
                oCalendar = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
                bResult = true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error initializing Outlook Helper.");
                System.Windows.MessageBox.Show("Error during initialization" + ex, "Date Line", MessageBoxButton.OK, MessageBoxImage.Error);
                CleanUp();
            }
            return bResult;
        }

        public void ChangeOutlookFolder()
        {
            try
            {
                var oFolder = oNS?.PickFolder();
                if (oFolder != null)
                {
                    //UpdateCustomFolder(oFolder);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error changing Outlook folder.");
                System.Windows.MessageBox.Show("Error changing Outlook folder." + ex, "Date Line", MessageBoxButton.OK, MessageBoxImage.Error);
                CleanUp();
            }

        }

        public void CleanUp()
        {
            oNS.Logoff();
            oCalendar = null;
            oNS = null;
            oApp = null;
        }

        public Dictionary<DateTime, string> GetTaskStrings(IEnumerable<DateTime> dates)
        {
            _logger.Info("Get Task Strings started");
            Dictionary<DateTime, string> result = new Dictionary<DateTime, string>();
            string task = "";

            foreach (Outlook.AppointmentItem oAppt in oCalendar.Items)
            {
                foreach (var date in dates)
                {
                    if (oAppt.Start.Date == date.Date)
                    {
                        if (!result.ContainsKey(date))
                        {
                            result[date] = "";
                        }
                        result[date] += oAppt.Start.ToString("HH:mm") + "  ";
                        result[date] += oAppt.Subject + "\n";
                        // Show some common properties.
                        //Console.WriteLine("Subject: " + oAppt.Subject);
                        //Console.WriteLine("Organizer: " + oAppt.Organizer);
                        //Console.WriteLine("Start: " + oAppt.Start.ToString());
                        //Console.WriteLine("End: " + oAppt.End.ToString());
                        //Console.WriteLine("Location: " + oAppt.Location);
                        //Console.WriteLine("Recurring: " + oAppt.IsRecurring);
                    }
                }
            }
            _logger.Info("Get Task Strings ended");
            return result;
        }
    }
}
