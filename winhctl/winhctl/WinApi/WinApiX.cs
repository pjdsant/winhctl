using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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
        private static extern IntPtr SendMessageList(IntPtr hWnd, uint Msg, UIntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint", CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref Point lpPoint);

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        private static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", EntryPoint = "mouse_event")]
        private static extern void mouse_event(MouseEventFlag dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private enum MouseEventFlag : uint
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

        private enum EventType
        {
            Down,
            Up,
            Click
        }


        private enum KeyEventFlag : int
        {
            Down = 0x0000,
            Up = 0x0002,
        }

        private void Virtual_MouseMove(Point p)
        {
            try
            {
                SetCursorPos(p.X, p.Y);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void Virtual_MouseEvent(MouseButtons button, EventType type = EventType.Click, bool doubleClick = false)

        {
            try
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
            catch (Exception)
            {
                throw;
            }
        }

        private delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        private static bool EnumChildWindowsCallback(IntPtr handle, IntPtr pointer)
        {
            var gcHandle = GCHandle.FromIntPtr(pointer);
            var list = gcHandle.Target as List<IntPtr>;
            try
            {
                if (list == null)
                {
                    throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
                }

                list.Add(handle);

                return true;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private static IEnumerable<IntPtr> GetChildWindows(IntPtr parent)
        {
            var result = new List<IntPtr>();
            var listHandle = GCHandle.Alloc(result);

            try
            {
                EnumChildWindows(parent, EnumChildWindowsCallback, GCHandle.ToIntPtr(listHandle));
            }
            catch (Exception)
            {
                throw;
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
            try
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
            catch (Exception)
            {
                throw;
            }
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
                return sb.ToString();
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void SendText(string windowTitle, int index, string message)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const uint WM_SETTEXT = 0x000C;
                var hdchild = childWindows.ToArray()[index];

                SendMessage(childWindows.ToArray()[index], WM_SETTEXT, 0, message);
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void SendClick(string windowTitle, int index)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int BM_CLICK = 0x00F5;
                SendMessageClick(childWindows.ToArray()[index], BM_CLICK, new IntPtr(0), new IntPtr(0));
            }
            catch (Exception)
            {
                throw;
            }
        }


        private string GetComboItem(string windowTitle, int index, int item)
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
            catch (Exception)
            {
                throw;
            }
        }

        private void SetComboItem(string windowTitle, int index, int item)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int CB_SETCURSEL = 0x014E;
                SendMessage(childWindows.ToArray()[index], CB_SETCURSEL, item, "0");
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void SendDoubleClickListBox(string windowTitle, int index)
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

                Virtual_MouseMove(point);
                Virtual_MouseEvent(MouseButtons.Left);
                Virtual_MouseEvent(MouseButtons.Left, type: EventType.Down);
                Virtual_MouseEvent(MouseButtons.Left, type: EventType.Up);
                Thread.Sleep(200);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public string GetPhonesGEO(string campanha)
        {
            try
            {
                var phones = "";
                var phonesMainTitle = "";
                int countGeoFoneTotal = 0;
                var mainTitle = "";
                var childTitle = "";
                int idxPhone1;
                int idxPhone2;
                int idxPhone3;
                int idxPhone4;
                int idxPhone4A;
                int idxPhone4L;
                var phone1 = "";
                var phone2 = "";
                var phone3 = "";
                var phone4 = "";
                var phone4A = "";

                mainTitle = "EDGE Client - Telemarketing";
                childTitle = "EDGE Client - ";

                if (!ScreenActive(mainTitle) && !ScreenActive(childTitle))
                {
                    return "";
                }

                if (campanha == "LIME")
                {
                    idxPhone1 = 85;
                    idxPhone2 = 81;
                    idxPhone3 = 77;
                       
                    phone1 = Regex.Replace(GetText(mainTitle, idxPhone1), "[^0-9]", "");
                    phone2 = Regex.Replace(GetText(mainTitle, idxPhone2), "[^0-9]", "");
                    phone3 = Regex.Replace(GetText(mainTitle, idxPhone3), "[^0-9]", "");
                }
                else
                {
                    if (campanha == "BSP_VMM")
                    {
                        if (ScreenActive(childTitle))
                        {
                            if (CountHandlesOnScreen(childTitle) <= 172)
                            {
                                idxPhone1 = 153;
                                idxPhone2 = 149;
                                idxPhone3 = 145;
                                idxPhone4 = 15;
                                idxPhone4A = 16;
                            }
                            else
                            {
                                idxPhone1 = 169;
                                idxPhone2 = 165;
                                idxPhone3 = 161;
                                idxPhone4 = 31;
                                idxPhone4A = 32;
                            }
                            phone1 = Regex.Replace(GetText(childTitle, idxPhone1), "[^0-9]", "");
                            phone2 = Regex.Replace(GetText(childTitle, idxPhone2), "[^0-9]", "");
                            phone3 = Regex.Replace(GetText(childTitle, idxPhone3), "[^0-9]", "");
                        }
                        else
                        {
                            idxPhone1 = 85;
                            idxPhone2 = 81;
                            idxPhone3 = 77;
                            idxPhone4 = 15;
                            idxPhone4A = 16;
                            idxPhone4L = 59;
                            phone1 = Regex.Replace(GetText(mainTitle, idxPhone1), "[^0-9]", "");
                            phone2 = Regex.Replace(GetText(mainTitle, idxPhone2), "[^0-9]", "");
                            phone3 = Regex.Replace(GetText(mainTitle, idxPhone3), "[^0-9]", "");

                            Thread.Sleep(50);
                            SetFocusScreen(mainTitle);
                            Thread.Sleep(50);
                            SendDoubleClickListBox(mainTitle, idxPhone4L);
                            Thread.Sleep(100);
                            SetFocusScreen(mainTitle);
                            Thread.Sleep(50);
                            SendDoubleClickListBox(mainTitle, idxPhone4L);
                            Thread.Sleep(50);


                        }

                    }
                    else
                    {
                        if (ScreenActive(childTitle))
                        {
                            if (CountHandlesOnScreen(childTitle) <= 173)
                            {
                                idxPhone1 = 154;
                                idxPhone2 = 150;
                                idxPhone3 = 146;
                                idxPhone4 = 29;
                                idxPhone4A = 30;
                            }
                            else
                            {
                                idxPhone1 = 170;
                                idxPhone2 = 166;
                                idxPhone3 = 162;
                                idxPhone4 = 45;
                                idxPhone4A = 46;
                            }
                            phone1 = Regex.Replace(GetText(childTitle, idxPhone1), "[^0-9]", "");
                            phone2 = Regex.Replace(GetText(childTitle, idxPhone2), "[^0-9]", "");
                            phone3 = Regex.Replace(GetText(childTitle, idxPhone3), "[^0-9]", "");
                        }
                        else
                        {
                            idxPhone1 = 85;
                            idxPhone2 = 81;
                            idxPhone3 = 77;
                            idxPhone4 = 29;
                            idxPhone4A = 30;
                            idxPhone4L = 59;
                            phone1 = Regex.Replace(GetText(mainTitle, idxPhone1), "[^0-9]", "");
                            phone2 = Regex.Replace(GetText(mainTitle, idxPhone2), "[^0-9]", "");
                            phone3 = Regex.Replace(GetText(mainTitle, idxPhone3), "[^0-9]", "");

                            Thread.Sleep(50);
                            SetFocusScreen(mainTitle);
                            Thread.Sleep(50);
                            SendDoubleClickListBox(mainTitle, idxPhone4L);
                            Thread.Sleep(100);
                            SetFocusScreen(mainTitle);
                            Thread.Sleep(50);
                            SendDoubleClickListBox(mainTitle, idxPhone4L);
                            Thread.Sleep(50);

                        }

                    }

                    phone4A = Regex.Replace(GetText(childTitle, idxPhone4A), "[^0-9]", "");
                    phone4 = phone4A + Regex.Replace(GetText(childTitle, idxPhone4), "[^0-9]", "");

                }

                phone1 = phone1.Trim();

                if (phone1 == "" || phone1 == null || phone1 == ",")
                {

                }
                else
                {
                    countGeoFoneTotal++;
                    if (phone1.Substring(0, 1) == "0") { phone1 = phone1.Substring(1, phone1.Length - 1); }
                    phonesMainTitle = phone1 + ",";
                }

                if (phone2 == "" || phone2 == null || phone2 == ",")
                {

                }
                else
                {
                    countGeoFoneTotal++;
                    if (phone2.Substring(0, 1) == "0") { phone2 = phone2.Substring(1, phone2.Length - 1); }
                    phonesMainTitle = phonesMainTitle + phone2 + ",";
                }

                if (phone3 == "" || phone3 == null || phone3 == ",")
                {

                }
                else
                {
                    countGeoFoneTotal++;
                    if (phone3.Substring(0, 1) == "0") { phone3 = phone3.Substring(1, phone3.Length - 1); }
                    phonesMainTitle = phonesMainTitle + phone3 + ",";
                }

                if (phone4 == "" || phone4 == null || phone4 == ",")
                {
                    if (phone4A == "" || phone4A == null)
                    {

                    }
                    phones = phonesMainTitle;
                }
                else
                {
                    countGeoFoneTotal++;
                    countGeoFoneTotal++;

                    if (phone4.Substring(0, 1) == "0") { phone4 = phone4.Substring(1, phone4.Length - 1); }

                    if (phonesMainTitle.Contains(phone4))
                    {
                        phones = phonesMainTitle;
                    }
                    else
                    {
                        phones = phonesMainTitle + phone4;
                    }
                }


                List<string> phoneList = phones.ToLower().Split(',').Distinct().ToList();
                phones = string.Join(",", phoneList);


                return phones;
            }
            catch (Exception)
            {
                throw;
            }

        }

        bool ScreenActive(string windowTitle)
        {
            bool result = false;

            IntPtr retVal = FindWindowByCaption(IntPtr.Zero, windowTitle);

            if (retVal != null && retVal != new IntPtr(0))
            {
                result = true;
            }

            return result;
        }

        public string GetTextFields(string campanha)
        {
            var sb = new StringBuilder();

            int idxCampanha;
            int idxLista;

            try
            {
                string windowTitle = "";
                string childTitle = "";

                windowTitle = "EDGE Client - Telemarketing";
                childTitle = "EDGE Client - ";

                if (!ScreenActive(windowTitle) && (!ScreenActive(childTitle)))
                {
                    return "";
                }

                if (campanha == "BSP_VMM")
                {
                    if (ScreenActive(childTitle))
                    {
                        if (CountHandlesOnScreen(childTitle) <= 172)
                        {
                            idxCampanha = 114;
                            idxLista = 106;
                        }
                        else
                        {
                            idxCampanha = 130;
                            idxLista = 122;
                        }
                    }
                    else
                    {
                        idxCampanha = 46;
                        idxLista = 38;
                    }
                }
                else
                {
                    if (ScreenActive(childTitle))
                    {
                        if (CountHandlesOnScreen(childTitle) <= 173)
                        {
                            idxCampanha = 115;
                            idxLista = 107;
                        }
                        else
                        {
                            idxCampanha = 131;
                            idxLista = 123;
                        }
                    }
                    else
                    {
                        idxCampanha = 46;
                        idxLista = 38;
                    }
                }

                var windowHWnd = FindWindowByCaption(IntPtr.Zero, childTitle);
                var childWindows = GetChildWindows(windowHWnd);
                var textCampanha = GetTextX(childWindows.ToArray()[idxCampanha]);
                sb.Append(textCampanha);
                sb.Append(",");
                var textLista = GetTextX(childWindows.ToArray()[idxLista]);
                sb.Append(textLista);
                sb.Append(",");

                return sb.ToString();

            }
            catch (Exception)
            {
                throw;
            }
        }

        int CountHandlesOnScreen(string windowTitle)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                return childWindows.ToArray().Length;
            }
            catch (Exception)
            {
                throw;
            }
        }

        void SetFocusScreen(string windowTitle)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                SetForegroundWindow(windowHWnd);
            }
            catch (Exception)
            {

                throw;
            }
        }


        string GetListItem(string windowTitle, int index, int item)
        {
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                const int LB_GETTEXT = 0x0189;
                StringBuilder ssb = new StringBuilder(256, 256);

                SendRefMessage(childWindows.ToArray()[index], LB_GETTEXT, item, ssb);
                return ssb.ToString();
            }
            catch (Exception)
            {
                throw;
            }
        }
        string GetOperation(string windowTitle,int idxChildHwd, int idxChildItem)
        {
            try
            {
                return GetListItem(windowTitle, idxChildHwd, idxChildItem);
            }
            catch (Exception)
            {
                throw;
            }

        }

        public string GetVersion()
        {
            //AssemblyName oAssemblyName = Assembly.GetExecutingAssembly().GetName();
            //return oAssemblyName.Version.ToString();
            try
            {
                return Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
            catch (Exception)
            {
                throw;
            }
           
        }

    }
}