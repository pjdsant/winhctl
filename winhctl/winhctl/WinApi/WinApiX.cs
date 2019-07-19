using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace LikeWater.WinHCtl.WinApi
{
    public class WinApiX
    {
        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder buf, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, [Out] StringBuilder lParam);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr zeroOnly, string lpWindowName);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr i);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);

        [DllImport("user32.dll", EntryPoint = "SendMessage", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageClick(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false, EntryPoint = "SendMessage")]
        private static extern IntPtr SendRefMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

        private delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);


        private static bool EnumChildWindowsCallback(IntPtr handle, IntPtr pointer)
        {
            var gcHandle = GCHandle.FromIntPtr(pointer);
            var list = gcHandle.Target as List<IntPtr>;

            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }

            list.Add(handle);

            return true;
        }

        private static IEnumerable<IntPtr> GetChildWindows(IntPtr parent)
        {
            var result = new List<IntPtr>();
            var listHandle = GCHandle.Alloc(result);

            try
            {
                EnumChildWindows(parent, EnumChildWindowsCallback, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }

            return result;
        }

        private static string GetTextX(IntPtr handle)
        {
            const uint WM_GETTEXTLENGTH = 0x000E;
            const uint WM_GETTEXT = 0x000D;

            var sbi = new StringBuilder(100);
            var length = (int)SendMessage(handle, WM_GETTEXTLENGTH, IntPtr.Zero, null);
            var sb = new StringBuilder(length + 1);
            GetClassName(handle, sbi, sbi.Capacity);
            SendMessage(handle, WM_GETTEXT, (IntPtr)sb.Capacity, sb);

            return sb.ToString();
        }
        public static string GetText(string windowTitle, int index)
        {
            var sb = new StringBuilder();
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                var childWindowText = GetTextX(childWindows.ToArray()[index]);
                sb.Append(childWindowText);

                return sb.ToString();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        public static void SendText(string windowTitle, int index, string message)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const uint WM_SETTEXT = 0x000C;
                var hdchild = childWindows.ToArray()[index];

                SendMessage(childWindows.ToArray()[index], WM_SETTEXT, 0, message);
            }

            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        public static void SendClick(string windowTitle, int idx)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int BM_CLICK = 0x00F5;
                SendMessageClick(childWindows.ToArray()[idx], BM_CLICK, new IntPtr(0), new IntPtr(0));
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        public static string GetComboItem(string windowTitle, int index, int item)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int CB_GETLBTEXT = 0x0148;
                StringBuilder ssb = new StringBuilder(256, 256);

                SendRefMessage(childWindows.ToArray()[index], CB_GETLBTEXT, item, ssb);
                return ssb.ToString();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public static void SetComboItem(string windowTitle, int index, int item)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int CB_SETCURSEL = 0x014E;
                SendMessage(childWindows.ToArray()[index], CB_SETCURSEL, item, "0");
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }



    }
}
