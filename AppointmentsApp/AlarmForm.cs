using AppointmentsApp.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Media;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace AppointmentsApp {
    public partial class AlarmForm : Form {

        SoundPlayer simpleSound;
        Outlook.AppointmentItem _outlookApptItem;

        //if clicked the appointment is displayed. we dont want it to display again when the form is closed 
        bool wasClicked = false;

        //to always be on top suff
        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        static readonly IntPtr HWND_TOP = new IntPtr(0);
        static readonly IntPtr HWND_BOTTOM = new IntPtr(1);
        const UInt32 SWP_NOSIZE = 0x0001;
        const UInt32 SWP_NOMOVE = 0x0002;
        const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        
        public AlarmForm(Outlook.AppointmentItem outlookApptItem)
        {
            InitializeComponent();
            this.Icon = AppointmentsApp.Properties.Resources.alarm;
            _outlookApptItem = outlookApptItem;
            simpleSound = new SoundPlayer(@"c:\windows\Media\Alarm10.wav");
            simpleSound.PlayLooping();
        }

        private void AlarmForm_Load(object sender, EventArgs e)
        {
            //always on top stuff
            this.TopMost = true;
            this.BringToFront();
            this.TopLevel = true;
            this.Focus();
            SetWindowPos(this.Handle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS);
        }           

        private void AlarmForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //stop sound and open the appointment if not already opened
            simpleSound.Stop();
            //if already clicked the appointment is displayed. we dont want it to display again when the form is closed
            if (!wasClicked) _outlookApptItem.Display(false);
        }

        private void AlarmForm_Shown(object sender, EventArgs e)
        {
            //set lable2 to subject of appointment
            label2.Text = _outlookApptItem.Subject;

            //timer to blink label1
            Timer timer = new Timer();
            timer.Interval = 1000;   // milliseconds
            timer.Tick += TimerTick;  // set handler
            timer.Start();
        }

        //event handler for every second
        private void TimerTick(object sender, EventArgs e)  //run this logic each timer tick
        {
            //toggle background color
            label1.BackColor = label1.BackColor == Color.White ? Color.Black : Color.White;
            label1.ForeColor = label1.ForeColor == Color.Black ? Color.White : Color.Black;
            this.TopMost = true;
            this.BringToFront();
            this.TopLevel = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {
            //if clicked the appointment is displayed. we dont want it to display again when the form is closed
            wasClicked = true;
            _outlookApptItem.Display(false);
        }

        private void label2_Click(object sender, EventArgs e)
        {
            //if clicked the appointment is displayed. we dont want it to display again when the form is closed
            wasClicked = true;
            _outlookApptItem.Display(false);
        }
    }
}
