using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using System.Windows;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace dsoFramerTestUse
{
    public partial class Form1 : Form
    {
        private bool fInCall_KB = false;
        private bool fInCall_MS = false;
        private Timer tmrDelayedCall_KB;
        private Timer tmrDelayedCall_MS;
        private const string cst_wordClassName = "_WwG";    //Word Class Name
        enum KeyType
        {
            ENTER,
            CTL_A,
            NONE
        }
        KeyType keyType = KeyType.NONE;

        //define a custom menu, Here we create a menu of type ContextMenu instead of
        //ContextMenuStrip, the later occurred some problem after show().
        ContextMenu ctxMenu;

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool ScreenToClient(IntPtr hWnd, ref POINT point);

        public Form1()
        {
            InitializeComponent();

            WindowState = FormWindowState.Maximized;
            
            //We create a custom menu filling with temp contents.
            ctxMenu = new ContextMenu();
            ctxMenu.MenuItems.Add("www.Medalsoft.com");
            ctxMenu.MenuItems.Add("author");
            ctxMenu.MenuItems.Add("Jerryh");

            Program.kh.KeyDown += new KeyEventHandler(kh_KeyDown);  //subscribe key down event
            Program.mh.MouseRBtnUp += new MouseEventHandler(mh_RBtnUp); //subscribe mouse right button down event.
        }

        void kh_KeyDown(object sender, KeyEventArgs e)
        {
            //Key Down event Handler, our form is active

            if (!fInCall_KB && (Form.ActiveForm == this))
            {
                //a sample while users press [Ctl+A] key, do something...
                if (e.KeyCode == Keys.A && (int)Control.ModifierKeys == (int)Keys.Control)
                {
                    string ctlClassName = GetFocusedControlClassName();

                    if (ctlClassName.Equals(cst_wordClassName))
                    {
                        //The hook procedure should process a message in less time than the data 
                        //entry specified in the LowLevelHooksTimeout value in the following registry key: 
                        //HKEY_CURRENT_USER\Control Panel\Desktop
                        //The value is in milliseconds. If the hook procedure does not return during this
                        //interval, the system will pass the message to the next hook.
                        keyType = KeyType.CTL_A;

                        //Here we create a timer to delay showing of the dialog!
                        e.Handled = true;
                        fInCall_KB = true;
                        tmrDelayedCall_KB = new Timer();
                        tmrDelayedCall_KB.Interval = 1;
                        tmrDelayedCall_KB.Tick += new
                        EventHandler(tmrDelayedCall_KB_Tick);
                        tmrDelayedCall_KB.Start();
                    }
                }
                else if (e.KeyCode == Keys.Enter)
                {
                    string ctlClassName = GetFocusedControlClassName();

                    if (ctlClassName.Equals(cst_wordClassName))
                    {
                        keyType = KeyType.ENTER;
                        e.Handled = true;
                        fInCall_KB = true;
                        tmrDelayedCall_KB = new Timer();
                        tmrDelayedCall_KB.Interval = 1;
                        tmrDelayedCall_KB.Tick += new
                        EventHandler(tmrDelayedCall_KB_Tick);
                        tmrDelayedCall_KB.Start();
                    }
                }
            }
        }

        void tmrDelayedCall_KB_Tick(object sender, EventArgs e)
        {
            tmrDelayedCall_KB.Stop();
            switch (keyType)
            {
                case KeyType.CTL_A:
                    MessageBox.Show("MedalSoft HotKey CTL+A");
                    break;
                case KeyType.ENTER:
                    MessageBox.Show("MedalSoft HotKey ENTER");
                    break;
                default:
                    break;
            }

            fInCall_KB = false;
        }

        void mh_RBtnUp(object sender, MouseEventArgs e)
        {
            if (!fInCall_MS && (Form.ActiveForm == this))
            {
                string ctlClassName = GetFocusedControlClassName();
                if (ctlClassName.Equals(cst_wordClassName))
                {
                    //Here we create a timer to delay showing of the dialog!
                    fInCall_MS = true;
                    tmrDelayedCall_MS = new Timer();
                    tmrDelayedCall_MS.Interval = 1;
                    tmrDelayedCall_MS.Tick += new
                    EventHandler(tmrDelayedCall_MS_Tick);
                    tmrDelayedCall_MS.Start();
                }
            }
        }

        void tmrDelayedCall_MS_Tick(object sender, EventArgs e)
        {
            tmrDelayedCall_MS.Stop();

            //Here, we should confirm the mouse pointer falled into word client area.
            //if yes, we pop up the custom menu!
            if (CheckMousePointInFocusClientRect() == true)
            {
                //we need exchange the mouse pos from screen to client.
                ctxMenu.Show(this, this.PointToClient(Cursor.Position));
            }

            fInCall_MS = false;
        }

        //add this function to check whether the mouse cursor fall into focused client rect.
        private bool CheckMousePointInFocusClientRect()
        {
            IntPtr focusedHandle = GetFocus();
            StringBuilder strClassName = new StringBuilder(100);
            Point mousePt = Control.MousePosition;
            RECT rect;

            if (focusedHandle != IntPtr.Zero)
            {
                GetClassName(focusedHandle, strClassName, strClassName.Capacity);

                if (GetWindowRect(focusedHandle, out rect) == true)
                {
                    System.Drawing.Rectangle rt = new System.Drawing.Rectangle(rect.X, rect.Y, rect.WIDTH, rect.HEIGHT);
                    POINT stmousePt = new POINT(mousePt.X, mousePt.Y);

                    if (rt.Contains(new Point(stmousePt.x, stmousePt.y)) == true)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        //add this func to get class name of the focused control.("word")
        private string GetFocusedControlClassName()
        {
            // To get hold of the focused control: 
            IntPtr focusedHandle = GetFocus();
            StringBuilder strClassName = new StringBuilder(100);

            if (focusedHandle != IntPtr.Zero)
            {
                int ret = GetClassName(focusedHandle, strClassName, strClassName.Capacity);
                if (ret != 0)   //successful
                {
                    return strClassName.ToString();
                }
                else    //error occurred.
                {
                    return "";
                }
            }

            return "";
        }

        private void axFramerControl1_OnDocumentOpened(object sender, AxDSOFramer._DFramerCtlEvents_OnDocumentOpenedEvent e)
        {
            Document document = e.document as Document;
            document.ActiveWindow.DisplayRulers = false;
            document.ActiveWindow.DisplayRightRuler = false;
            document.ActiveWindow.DisplayVerticalRuler = false;
            document.ActiveWindow.DisplayScreenTips = false;
            document.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            document.ActiveWindow.View.ShowBookmarks = true;
            document.ActiveWindow.ActivePane.DisplayRulers = false;
            document.ActiveWindow.View.DisplaySmartTags = false;

            //disable word commandbars 
            try
            {
                //test debug.write
                /*Debug.Listeners.Add(new TextWriterTraceListener(Console.Out));
                Debug.AutoFlush = true;
                Debug.Indent();
                Debug.WriteLine(Name.ToString());
                Debug.Unindent();*/

                Document wordDoc = document as Document;

                //Disable Text popup menu（while right click on the word client area）.
                //Index is 78, in word 2003.
                //Name=Text Index=78 Type=msoBarTypePopup
                //The reason disable this first is when word started, we right click word
                //client area at once, the text popup menu perhaps still enable at this moment.
                wordDoc.CommandBars[78].Enabled = false;
                wordDoc.CommandBars[78].Visible = false;

                //this course cost a few time, disable all tool bars, menu bars.
                for (int i = 1; i <= wordDoc.CommandBars.Count; i++)
                {
                    wordDoc.CommandBars[i].Enabled = false;
                    wordDoc.CommandBars[i].Visible = false;
                }
            }
            catch (Exception)
            {
            }
        }

        private void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.InitialDirectory = System.Environment.CurrentDirectory;
            fd.Filter = "Microsoft Word Files|*.docx|All Files|*.*";
            fd.RestoreDirectory = true;
            fd.FilterIndex = 1;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                this.axFramerControl1.Open(fd.FileName);
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.axFramerControl1.Close();
        }

        private void SetFocus_Click(object sender, EventArgs e)
        {
            this.axFramerControl1.Focus();
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            Size ctlSize = new Size(this.ClientRectangle.Size.Width - 20,
                this.ClientRectangle.Size.Height - 50);
            this.axFramerControl1.Size = ctlSize;
        }
    }
}
