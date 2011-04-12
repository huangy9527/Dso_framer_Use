using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;

namespace dsoFramerTestUse
{
    [StructLayout(LayoutKind.Sequential)]
    public class POINT
    {
        public int x;
        public int y;

        public POINT(int x, int y)
        {
            this.x = x;
            this.y = y;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MouseHookStruct
    {
        public POINT pt;
        public int hWnd;
        public int wHitTestCode;
        public int dwExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        int left;
        int top;
        int right;
        int bottom;

        public int X
        {
            get { return left; }
        }
        public int Y
        {
            get { return top; }
        }
        public int WIDTH
        {
            get { return right-left; }
        }
        public int HEIGHT
        {
            get { return bottom - top; }
        }
    }

    class MouseHook : IDisposable
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
        LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
        IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        //hook enables you to monitor mouse input events about to 
        //be posted in a thread input queue
        private const int WH_MOUSE_LL = 14;
        private const int WM_RBUTTONUP = 0x0205;
        private IntPtr _hookID = IntPtr.Zero;

        public event MouseEventHandler MouseRBtnUp;

        private LowLevelMouseProc llMouseProc;

        public MouseHook()
        {
            //Install Hook, return handle to the hook procedure
            llMouseProc = new LowLevelMouseProc(HookCallback);
            _hookID = SetHook(llMouseProc);
        }

#region IDisposable Members
        public void Dispose()
        {
            //Uninstall hook;
            UnhookWindowsHookEx(_hookID);
        }
#endregion

        private IntPtr SetHook(LowLevelMouseProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                //set hook
                return SetWindowsHookEx(WH_MOUSE_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        //function pointer
        private delegate IntPtr LowLevelMouseProc(
        int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(
        int nCode, IntPtr wParam, IntPtr lParam)
        {
            //keydown occurred
            if (nCode >= 0 && wParam == (IntPtr)WM_RBUTTONUP)
            {
                MouseHookStruct MHookStruct = 
                    (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));
                
                if (MouseRBtnUp != null)
                {
                    //MouseButtonEventArgs args = new MouseButtonEventArgs()
                    MouseEventArgs args = new MouseEventArgs(
                        MouseButtons.Right, 1, MHookStruct.pt.x, MHookStruct.pt.y, 0);
                    MouseRBtnUp(this, args);    //Raise Event.
                    //return new IntPtr(1);   // indicate processed
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
    }
}
