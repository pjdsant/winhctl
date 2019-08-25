using RegMan.Resource;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace LikeWater.WinHCtl.WinApi
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("LikeWater.WinHandlerControl")]
    [ComVisible(true)]

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

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false, EntryPoint = "SendMessage")]
        static extern IntPtr SendMessageList(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint",CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg,IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", EntryPoint = "mouse_event")]
        public static extern void mouse_event(MouseEventFlag dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public enum MouseEventFlag : uint
        {
            Move = 0x0001,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            XDown = 0x0080,
            XUp = 0x0100,
            Wheel = 0x0800,
            VirtualDesk = 0x4000,
            Absolute = 0x8000
        }

        public enum EventType
        {
            Down,
            Up,
            Click
        }


        public enum KeyEventFlag : int
        {
            Down = 0x0000,
            Up = 0x0002,
        }

        private void virtual_MouseMove(Point p)
        {
            SetCursorPos(p.X, p.Y);
        }

        private void virtual_MouseEvent(MouseButtons button, EventType type = EventType.Click, bool doubleClick = false)

        {

            MouseEventFlag flagUp = new MouseEventFlag();

            MouseEventFlag flagDown = new MouseEventFlag();

            switch (button)
            {
                case MouseButtons.Left:
                    flagUp = MouseEventFlag.LeftUp;
                    flagDown = MouseEventFlag.LeftDown;
                    break;
                case MouseButtons.Middle:
                    flagUp = MouseEventFlag.MiddleUp;
                    flagDown = MouseEventFlag.MiddleDown;
                    break;
                case MouseButtons.Right:
                    flagUp = MouseEventFlag.RightUp;
                    flagDown = MouseEventFlag.RightDown;
                    break;
                default://defaul left button
                    flagUp = MouseEventFlag.LeftUp;
                    flagDown = MouseEventFlag.LeftDown;
                    break;
            }

            if (type == EventType.Click)
            {
                int count = doubleClick == false ? 1 : 2;//check
                for (int i = 0; i < count; i++)
                {
                    mouse_event(flagUp, 0, 0, 0, 0);
                    mouse_event(flagDown, 0, 0, 0, 0);
                }
            }

            else
            {
                if (type == EventType.Down)
                    mouse_event(flagDown, 0, 0, 0, 0);
                else if (type == EventType.Up)
                    mouse_event(flagUp, 0, 0, 0, 0);
            }
        }

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
        public string GetText(string windowTitle, int index)
        {
            var sb = new StringBuilder();
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                var childWindowText = GetTextX(childWindows.ToArray()[index]);
                sb.Append(childWindowText);

                RegistryManager rm = new RegistryManager();
                rm.WriteRegistryEvents("GetText");

                return sb.ToString();
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        public void SendText(string windowTitle, int index, string message)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const uint WM_SETTEXT = 0x000C;
                var hdchild = childWindows.ToArray()[index];

                SendMessage(childWindows.ToArray()[index], WM_SETTEXT, 0, message);

                RegistryManager rm = new RegistryManager();
                rm.WriteRegistryEvents("SendText");
            }

            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        public void SendClick(string windowTitle, int index)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int BM_CLICK = 0x00F5;
                SendMessageClick(childWindows.ToArray()[index], BM_CLICK, new IntPtr(0), new IntPtr(0));

                RegistryManager rm = new RegistryManager();
                rm.ReadRegistryEvents("SendClick");
                rm.WriteRegistryEvents("SendClick");
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }


        public string GetComboItem(string windowTitle, int index, int item)
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

        public void SetComboItem(string windowTitle, int index, int item)
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

        private Button button1 = new Button();
        public void DoubleClickListBox(string windowTitle, int index)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                IntPtr hWnd = childWindows.ToArray()[index];
                RECT rct;
                int width = 0;
                int height = 0;
                int x = 0;
                int y = 0;
                Point point;
                bool coordinatesFound = false;

                rct = new RECT();
                GetWindowRect(hWnd, ref rct);

                width = rct.Right - rct.Left;
                height = rct.Bottom - rct.Top;
                point = new Point();
                x = (width / 2);
                y = (height / 2);

                coordinatesFound = ClientToScreen(hWnd, ref point);

                if (coordinatesFound == true)
                {
                    SetCursorPos(point.X + x, point.Y + y);
                }

                virtual_MouseMove(point);
                virtual_MouseEvent(MouseButtons.Left);
                virtual_MouseEvent(MouseButtons.Left, type: EventType.Down);
                virtual_MouseEvent(MouseButtons.Left, type: EventType.Up);
                Thread.Sleep(200);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        
    }
}
