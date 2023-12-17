using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Timers;
using System.Windows;
using HorizontalAlignment = System.Windows.HorizontalAlignment;
using Label = System.Windows.Controls.Label;
using MessageBox = System.Windows.MessageBox;

namespace DateLine;

[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public partial class MainWindow
{
    private const int DAY_WINDOW_SIZE = 15;

    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

    private readonly System.Windows.Forms.NotifyIcon _trayNotify;

    private bool _outLookIntegrationSucceeded;

    private readonly BackgroundWorker _outlookInitializer;

    private readonly List<Label> _dateLabels = [];

    private readonly Style _labelStyle;

    private Timer _refreshTimer = new(new TimeSpan(0, 0, 30));

    private DateTime _today = DateTime.Today;

    public MainWindow()
    {
        InitializeComponent();

        Logger.Debug("Checking to see if there is an instance running.");
        var procName = Process.GetCurrentProcess().ProcessName;
        var processes = Process.GetProcessesByName(procName);

        if (processes.Length > 1)
        {
            Logger.Warn("Instance is already running, exiting.");
            MessageBox.Show("Program is already running", "Date Line", MessageBoxButton.OK,
                MessageBoxImage.Information);
            Close();
            return;
        }

        _labelStyle = Resources["LabelStyle1"] as Style;

        WindowStartupLocation = WindowStartupLocation.Manual;
        Height = SystemParameters.PrimaryScreenHeight;
        Width = 80;
        //this.Opacity = 0.5;           
        Left = SystemParameters.PrimaryScreenWidth - Width;
        Top = 0;
        AddLabels();

        _trayNotify = new System.Windows.Forms.NotifyIcon();
        _trayNotify.Icon = new Icon(GetType(), "application.ico");
        _trayNotify.Text = "Date Line";
        _trayNotify.Visible = true;

        var ctxTrayMenu = new System.Windows.Forms.ContextMenuStrip();
        var mnuChangeFolder = new System.Windows.Forms.ToolStripMenuItem();
        mnuChangeFolder.Text = "Select Outlook Folder";
        mnuChangeFolder.Click += mnuSelectOutlookFolder_Click;

        var mnuExit = new System.Windows.Forms.ToolStripMenuItem();
        mnuExit.Text = "Exit";
        mnuExit.Click += mnuExit_Click;
        //ctxTrayMenu.Items.Add(mnuChangeFolder);
        ctxTrayMenu.Items.Add(mnuExit);
        _trayNotify.ContextMenuStrip = ctxTrayMenu;

        Dispatcher.ShutdownStarted += Dispatcher_ShutdownStarted;

        _outlookInitializer = new BackgroundWorker();
        _outlookInitializer.DoWork += outlookInitializer_DoWork;
        _outlookInitializer.RunWorkerAsync();

        _refreshTimer.Elapsed += _refreshTimer_Elapsed;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        WindowHelper.SetAsDesktopChild(this);
    }

    private static void Dispatcher_ShutdownStarted(object sender, EventArgs e)
    {
        OutlookHelper.Instance.CleanUp();
    }

    private static void mnuSelectOutlookFolder_Click(object sender, EventArgs e)
    {
        OutlookHelper.Instance.ChangeOutlookFolder();
    }

    private void mnuExit_Click(object sender, EventArgs e)
    {
        Close();
        _trayNotify.Visible = false;
    }

    private void AddLabels()
    {
        for (var i = -DAY_WINDOW_SIZE; i < DAY_WINDOW_SIZE; i++)
        {
            var labelDate = new Label
            {
                Width = 70,
                HorizontalContentAlignment = HorizontalAlignment.Right
            };
            var currDate = DateTime.Today.Add(new TimeSpan(i, 0, 0, 0));
            if (currDate == DateTime.Today)
            {
                labelDate.BorderBrush = System.Windows.Media.Brushes.White;
                labelDate.BorderThickness = new Thickness(1.0);
            }

            if (currDate.DayOfWeek is DayOfWeek.Sunday or DayOfWeek.Saturday)
                labelDate.Foreground = System.Windows.Media.Brushes.PaleVioletRed;

            var dayName = currDate.DayOfWeek.ToString();
            dayName = dayName.Substring(0, 2);
            labelDate.Content = string.Concat(currDate.Day.ToString(), " ", dayName);
            labelDate.Style = _labelStyle;
            //labelDate.BorderBrush = System.Windows.Media.Brushes.LightGray;
            //labelDate.BorderThickness = new Thickness(1.0);
            StackPanel.Children.Add(labelDate);

            labelDate.Tag = currDate.Date;
            _dateLabels.Add(labelDate);
        }
    }

    private void outlookInitializer_DoWork(object sender, DoWorkEventArgs e)
    {
        try
        {
            _outLookIntegrationSucceeded = OutlookHelper.Instance.Initialize();
            if (_outLookIntegrationSucceeded)
            {
                RefreshAppointments();
                _refreshTimer.Start();
            }
        }
        catch (Exception err)
        {
            MessageBox.Show(err.Message);
        }
    }

    private void _refreshTimer_Elapsed(object sender, ElapsedEventArgs e)
    {
        if (_today != DateTime.Today)
        {
            _today = DateTime.Today;
            Dispatcher.BeginInvoke(() =>
            {
                StackPanel.Children.Clear();
                _dateLabels.Clear();
                AddLabels();
            });
        }
        RefreshAppointments();
    }

    private void RefreshAppointments()
    {
        var fromDate = DateTime.Today.Subtract(new TimeSpan(DAY_WINDOW_SIZE, 0, 0, 0));
        var toDate = DateTime.Today.Add(new TimeSpan(DAY_WINDOW_SIZE, 0, 0, 0));
        var appointments = OutlookHelper.Instance.GetTaskStrings(fromDate, toDate);

        Dispatcher.BeginInvoke(() =>
        {
            foreach (var lbl in _dateLabels)
            {
                if (!appointments.ContainsKey((DateTime)lbl.Tag)) continue;
                var appointment = appointments[(DateTime)lbl.Tag];
                if (appointment == "") continue;
                var toolTipText = "\n" + ((DateTime)lbl.Tag).ToLongDateString() + "\n\n" + appointment;
                lbl.ToolTip = toolTipText;
                if (!lbl.Content.ToString()!.StartsWith('.'))
                    lbl.Content = "." + lbl.Content;
            }
        });
    }
}