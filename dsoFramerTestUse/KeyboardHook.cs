using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;

namespace dsoFramerTestUse
{
    //Using SetWindowsHookEx create gloal hook, raise event while 
    //our winform activing
    class KeyboardHook : IDisposable
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
        LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
        IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        //hook enables you to monitor keyboard input events about to 
        //be posted in a thread input queue
        private const int WH_KEYBOARD_LL = 13;
        //Posted to the window with the keyboard focus when a 
        //nonsystem key is pressed
        private const int WM_KEYDOWN = 0x0100;
        private IntPtr _hookID = IntPtr.Zero;

        public event KeyEventHandler KeyDown;

        private LowLevelKeyboardProc llKBProc;

        public KeyboardHook()
        {
            //Install Hook, return handle to the hook procedure
            llKBProc = new LowLevelKeyboardProc(HookCallback);
            _hookID = SetHook(llKBProc);
        }

#region IDisposable Members
        public void Dispose()
        {
            //Uninstall hook;
            UnhookWindowsHookEx(_hookID);
        }
#endregion

        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                //set hook
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        //function pointer
        private delegate IntPtr LowLevelKeyboardProc(
        int nCode, IntPtr wParam, IntPtr lParam);

        private IntPtr HookCallback(
        int nCode, IntPtr wParam, IntPtr lParam)
        {
            //keydown occurred
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;
                
                if (KeyDown != null)
                {
                    KeyEventArgs args = new KeyEventArgs(key);
                    KeyDown(this, args);    //Raise Event.
                    if (args.Handled)
                    {
                       return new IntPtr(1);   // indicate processed
                    }
                }
            }
            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }
    }
}
