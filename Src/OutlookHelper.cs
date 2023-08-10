﻿using System;
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

        public string GetTaskString(DateTime date)
        {
            string task = "";

            foreach (Outlook.AppointmentItem oAppt in oCalendar.Items)
            {
                if (oAppt.Start.Date == date.Date)
                {
                    task += oAppt.Start.ToString("HH:mm") + "  ";
                    task += oAppt.Subject + "\n";
                    // Show some common properties.
                    //Console.WriteLine("Subject: " + oAppt.Subject);
                    //Console.WriteLine("Organizer: " + oAppt.Organizer);
                    //Console.WriteLine("Start: " + oAppt.Start.ToString());
                    //Console.WriteLine("End: " + oAppt.End.ToString());
                    //Console.WriteLine("Location: " + oAppt.Location);
                    //Console.WriteLine("Recurring: " + oAppt.IsRecurring);
                }
            }

            return task;
        }

        public void Initialize()
        {
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
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error initializing Outlook Helper.");
                System.Windows.MessageBox.Show("Error during initialization" + ex, "Date Line", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
        }

        internal void ChangeOutlookFolder()
        {
            var oFolder = oNS?.PickFolder();
            if (oFolder != null)
            {
                UpdateCustomFolder(oFolder);
            }
        }

        private void UpdateCustomFolder(Outlook.MAPIFolder oFolder)
        {

        }

        public void CleanUp()
        {
            oNS.Logoff();
            oCalendar = null;
            oNS = null;
            oApp = null;
        }


    }
}
