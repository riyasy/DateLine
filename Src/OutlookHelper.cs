using NLog;
using System;
using System.Collections.Generic;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace DateLine;

internal sealed class OutlookHelper
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private Outlook.Application _oApp;
    private Outlook.NameSpace _oNs;
    private Outlook.MAPIFolder _oCalendar;

    public static OutlookHelper Instance { get; } = new();

    private OutlookHelper()
    {
    }

    public bool Initialize()
    {
        var bResult = false;
        try
        {
            // Create the Outlook application.
            _oApp = new Outlook.Application();

            // Get the NameSpace and Logon information.
            // Outlook.NameSpace oNS = (Outlook.NameSpace)oApp.GetNamespace("mapi");
            _oNs = _oApp.GetNamespace("MAPI");

            //Log on by using a dialog box to choose the profile.
            //oNS.Logon(Missing.Value, Missing.Value, true, true);

            //Alternate logon method that uses a specific profile.
            // TODO: If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            // Get the Calendar folder.
            _oCalendar = _oNs.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            bResult = true;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error initializing Outlook Helper.");
            MessageBox.Show("Error during initialization" + ex, "Date Line", MessageBoxButton.OK,
                MessageBoxImage.Error);
            CleanUp();
        }

        return bResult;
    }

    public void ChangeOutlookFolder()
    {
        try
        {
            var oFolder = _oNs?.PickFolder();
            if (oFolder != null)
            {
                //UpdateCustomFolder(oFolder);
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error changing Outlook folder.");
            MessageBox.Show("Error changing Outlook folder." + ex, "Date Line", MessageBoxButton.OK,
                MessageBoxImage.Error);
            CleanUp();
        }
    }

    public void CleanUp()
    {
        _oNs.Logoff();
        _oCalendar = null;
        _oNs = null;
        _oApp = null;
    }

    public Dictionary<DateTime, string> GetTaskStrings(IEnumerable<DateTime> dates)
    {
        Logger.Info("Get Task Strings started");
        var result = new Dictionary<DateTime, string>();

        foreach (Outlook.AppointmentItem oAppt in _oCalendar.Items)
        foreach (var date in dates)
            if (oAppt.Start.Date == date.Date)
            {
                result.TryAdd(date, "");
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

        Logger.Info("Get Task Strings ended");
        return result;
    }
}