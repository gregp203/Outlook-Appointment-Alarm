using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static AppointmentsApp.Form1;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace AppointmentsApp {

    
    public partial class Form1 : Form {
        //get outlook application
        Outlook.Application outlookApp = new Outlook.Application();

        //create a list for appointments
        List<Appointment> appointments = new List<Appointment>();
        BindingSource appointmentsBS = new BindingSource();

        //timeout cells from being left selected to prevent bindings from reseting(update of data)
        int editTimeout = 5;

        public Form1()
        {
            InitializeComponent();

            this.Icon = AppointmentsApp.Properties.Resources.alarm;

            //set the bindingsource datasource to be the appointments list so changes to the list reflect in the data grid view
            appointmentsBS.DataSource = appointments;             

            //set data grid view source and columns widths
            apptDGV.DataSource = appointmentsBS;
            apptDGV.Columns[0].Visible = false; //entityID invisible            
            apptDGV.Columns[1].Width = 30;  // enable alarm check box column
            apptDGV.Columns[1].HeaderText = "On";
            apptDGV.Columns[2].ReadOnly = true;
            apptDGV.Columns[2].Width = 500;
            apptDGV.Columns[2].HeaderText = "Subject";
            apptDGV.Columns[3].Width = 60; // start time column
            apptDGV.Columns[3].ReadOnly = true;
            apptDGV.Columns[4].Width = 60; // lead time column
            apptDGV.Columns[5].Width = 60; // alarm time column
            apptDGV.Columns[5].ReadOnly = true;
            apptDGV.Columns[6].Width = 60; // time left column
            apptDGV.Columns[6].ReadOnly = true;
            apptDGV.ClearSelection();

            //event for every second populate the appointment list and check is alarm is ready
            Timer timer = new Timer();
            timer.Interval = 1000;   // milliseconds
            timer.Tick += TimerTick;  // set handler
            timer.Start();
        }

        //event handler for every second
        private void TimerTick(object sender, EventArgs e)  //run this logic each timer tick
        {
            PopulateApptsListAndAlarm(outlookApp,ref appointments);

            //if cell is selected to edit dont reset bindings
            if (apptDGV.CurrentCell==null || !apptDGV.CurrentCell.Selected)
            {
                appointmentsBS.ResetBindings(false);
                //reset of bindings seems to select the first cell so clear selection
                apptDGV.ClearSelection();
            }
            else editTimeout--;

            //if editTimeout is zero reset bindings
            if(editTimeout < 1)
            {
                editTimeout = 5;
                appointmentsBS.ResetBindings(false);
                //reset of bindings seems to select the first cell so clear selection
                apptDGV.ClearSelection();
            }

        }

        static private void PopulateApptsListAndAlarm(Outlook.Application outlookApp, ref List<Appointment> appointments)
        {
            //get calandar folder
            Outlook.Folder calFolder =
                outlookApp.Session.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderCalendar)
                as Outlook.Folder;
            
            //set time range of appointment to put inthe list
            DateTime start = DateTime.Now;
            DateTime end = DateTime.Today.AddDays(1);

            //Get the appointments filtered for the range
            Outlook.Items rangeAppts = GetAppointmentsInRange(calFolder, start, end);
            
            if (rangeAppts != null)
            {
                //loop trough the appointments from outlook 
                foreach (Outlook.AppointmentItem outlookAppt in rangeAppts)
                {
                    //find if outlook appointment is not already in the list and add it
                    int i = appointments.FindIndex(x => x.itemId == outlookAppt.EntryID.ToString());
                    if (i == -1)
                    {
                        //create apppointment object to add in to the list
                        Appointment addAppt = new Appointment();
                        addAppt.itemId = outlookAppt.EntryID.ToString();
                        addAppt.alarmEnable = true;
                        addAppt.apptDesc = outlookAppt.Subject;
                        addAppt.startTime = outlookAppt.Start.TimeOfDay;
                        addAppt.leadMinutes = 5;
                        addAppt.alarmTime = outlookAppt.Start.AddMinutes(-addAppt.leadMinutes).TimeOfDay;
                        addAppt.timeLeft = addAppt.alarmTime.Subtract(DateTime.Now.TimeOfDay);
                        appointments.Add(addAppt);
                    }
                    //if  outlook appointment is on the list update properties and check if ready for alarm
                    else
                    {
                        appointments[i].apptDesc = outlookAppt.Subject;
                        appointments[i].startTime = outlookAppt.Start.TimeOfDay;
                        appointments[i].alarmTime = outlookAppt.Start.AddMinutes(-appointments[i].leadMinutes).TimeOfDay;
                        appointments[i].timeLeft = appointments[i].alarmTime.Subtract(DateTime.Now.TimeOfDay);
                        
                        //check if appointment is ready to alarm 
                        if (appointments[i].alarmEnable && appointments[i].alarmTime < DateTime.Now.TimeOfDay)
                        {
                            appointments[i].alarmEnable = false;
                            AlarmForm alarm = new AlarmForm(outlookAppt);
                            alarm.Show();                            
                        }
                    }
                }

                //loop through appointments to remove if not in outlook apointments
                for (int i = 0; i < appointments.Count; i++)
                {
                    //I couldnt figurea out a way to find an outlookAppt in rangeAppts by EntryID so loop through all rangeAppts to see if one is found
                    bool found = false;
                    foreach (Outlook.AppointmentItem outlookAppt in rangeAppts)
                    {
                        if (outlookAppt.EntryID.ToString() == appointments[i].itemId)
                        {
                            found = true;
                        }                        
                    }
                    if (!found) { appointments.RemoveAt(i); }    
                }
            }

            //sort list by start time
            appointments.Sort(delegate (Appointment x, Appointment y)
            {
                return x.startTime.CompareTo(y.startTime);
            });
        }

        //taken from https://learn.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-search-and-obtain-appointments-in-a-time-range
        static private Outlook.Items GetAppointmentsInRange(
            Outlook.Folder folder, DateTime startTime, DateTime endTime)
        {            
            string filter = "[Start] >= '"
                + startTime.ToString("g")
                + "' AND [End] <= '"
                + endTime.ToString("g") + "'";            
            try
            {
                Outlook.Items calItems = folder.Items;
                calItems.IncludeRecurrences = true;
                calItems.Sort("[Start]", Type.Missing);
                Outlook.Items restrictItems = calItems.Restrict(filter);
                if (restrictItems.Count > 0)
                {
                    return restrictItems;
                }
                else
                {
                    return null;
                }
            }
            catch { return null; }
        }

        private void apptDGV_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            apptDGV.ClearSelection();

        }

        private void apptDGV_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (apptDGV.IsCurrentCellDirty && apptDGV.CurrentCell is DataGridViewCheckBoxCell)
            {
                apptDGV.EndEdit();
                apptDGV.ClearSelection();
            }
        }
    }

    //struct for appointments
    public class Appointment {
        public string itemId { get; set; }
        public bool alarmEnable { get; set; }
        public string apptDesc { get; set; }
        public TimeSpan startTime { get; set; }
        public int leadMinutes { get; set; }
        public TimeSpan alarmTime { get; set; }
        public TimeSpan timeLeft { get; set; }
    }
}
