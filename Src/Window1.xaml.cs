using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.ComponentModel;


namespace DateLine
{
    public partial class Window1 : Window
    {

        private System.Windows.Forms.NotifyIcon TrayNotify;
        private System.Windows.Forms.ContextMenuStrip ctxTrayMenu;
        private bool bOutLookIntegrationSucceeded = false;

        private BackgroundWorker mBgWorker = null;

        List<Label> entries = new List<Label>();


        Style labelStyle;
        public Window1()
        {
            mBgWorker = new BackgroundWorker();
            mBgWorker.DoWork += mBgWorker_DoWork;
            mBgWorker.RunWorkerCompleted += mBgWorker_RunWorkerCompleted;

            InitializeComponent();

            mBgWorker.RunWorkerAsync();

            labelStyle = this.Resources["LabelStyle1"] as Style;

            this.WindowStartupLocation = WindowStartupLocation.Manual;
            this.Height = SystemParameters.PrimaryScreenHeight;
            this.Width = 60;
            //this.Opacity = 0.5;           
            this.Left = SystemParameters.PrimaryScreenWidth - this.Width;
            this.Top = 0;
            AddLabels();


            TrayNotify = new System.Windows.Forms.NotifyIcon();
            TrayNotify.Icon = new Icon(this.GetType(), "application.ico");
            TrayNotify.Text = "Date Line";
            TrayNotify.Visible = true;
            ctxTrayMenu = new System.Windows.Forms.ContextMenuStrip();
            System.Windows.Forms.ToolStripMenuItem mnuExit = new System.Windows.Forms.ToolStripMenuItem();
            mnuExit.Text = "Exit";
            mnuExit.Click += new EventHandler(mnuExit_Click);
            ctxTrayMenu.Items.Add(mnuExit);
            TrayNotify.ContextMenuStrip = ctxTrayMenu;
        }

        void mnuExit_Click(object sender, EventArgs e)
        {
            this.Close();
            TrayNotify.Visible = false;
        }

        void mBgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (bOutLookIntegrationSucceeded)
            {
                foreach (Label lbl in entries)
                {
                    string kk = OutlookHelper.Instance.GetTaskString((DateTime)lbl.Tag);
                    if (kk != "")
                    {
                        string toolTipText = "\n" + ((DateTime)lbl.Tag).ToLongDateString() + "\n\n" + kk;
                        lbl.ToolTip = toolTipText;
                        lbl.Content = "." + lbl.Content;
                    }
                }
            }
        }

        void mBgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                OutlookHelper.Instance.Initialize();
                bOutLookIntegrationSucceeded = true;
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        private void AddLabels()
        {
            Label labelDate;
            DateTime currDate;
            TimeSpan time = new TimeSpan(1, 0, 0, 0);
            int yPos = 0;
            string dayName;
            for (int i = -15; i < 15; i++)
            {

                labelDate = new Label();
                labelDate.Width = 50;
                labelDate.HorizontalContentAlignment = HorizontalAlignment.Right;
                currDate = DateTime.Today.Add(new TimeSpan(i, 0, 0, 0));
                if (currDate == DateTime.Today)
                {
                    labelDate.BorderBrush = System.Windows.Media.Brushes.White;
                    labelDate.BorderThickness = new Thickness(1.0);
                }
                if (currDate.DayOfWeek == DayOfWeek.Sunday || currDate.DayOfWeek == DayOfWeek.Saturday)
                {
                    labelDate.Foreground = System.Windows.Media.Brushes.PaleVioletRed;
                }

                dayName = currDate.DayOfWeek.ToString();
                dayName = dayName.Substring(0, 2);
                labelDate.Content = string.Concat(currDate.Day.ToString(), " ", dayName);
                labelDate.Style = labelStyle;
                //labelDate.BorderBrush = System.Windows.Media.Brushes.LightGray;
                //labelDate.BorderThickness = new Thickness(1.0);
                this.stackPanel.Children.Add(labelDate);
                yPos = yPos + 25;

                labelDate.Tag = currDate.Date;
                entries.Add(labelDate);
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            WindowHelper.SetAsDesktopChild(this);
        }

        private void Window_Unloaded(object sender, RoutedEventArgs e)
        {
            OutlookHelper.Instance.CleanUp();
        }
    }
}
