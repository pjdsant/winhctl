using RegMan.Resource;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: SuppressIldasmAttribute()]

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

        [DllImport("user32.dll", EntryPoint = "WindowFromPoint",CharSet = CharSet.Auto, ExactSpelling = true)]
        private static extern IntPtr WindowFromPoint(Point point);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg,IntPtr wParam, IntPtr lParam);

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
            RegistryManager rm = new RegistryManager();
            try
            {
                SetCursorPos(p.X, p.Y);
            }
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1000]WinApiX.Virtual_MouseMove --> " + ex.Message.ToString());
                throw new Exception(ex.Message); ;
            }
           
        }

        private void Virtual_MouseEvent(MouseButtons button, EventType type = EventType.Click, bool doubleClick = false)

        {
            RegistryManager rm = new RegistryManager();
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
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1001]WinApiX.Virtual_MouseEvent --> " + ex.Message.ToString());
                throw new Exception(ex.Message);
            }
        }

        private delegate bool EnumWindowProc(IntPtr hWnd, IntPtr parameter);

        private static bool EnumChildWindowsCallback(IntPtr handle, IntPtr pointer)
        {
            var gcHandle = GCHandle.FromIntPtr(pointer);
            var list = gcHandle.Target as List<IntPtr>;
            RegistryManager rm = new RegistryManager();
            try
            {
                if (list == null)
                {
                    throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
                }

                list.Add(handle);

                return true;
            }
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1002]WinApiX.EnumChildWindowsCallback --> " +
                "Params: handler: " + handle.ToString() +
                "pointer: " + pointer.ToString() +
                "Message Exception: " + ex.Message.ToString());
                throw new Exception(ex.Message);
            }
        }

        private static IEnumerable<IntPtr> GetChildWindows(IntPtr parent)
        {
            var result = new List<IntPtr>();
            var listHandle = GCHandle.Alloc(result);
            RegistryManager rm = new RegistryManager();

            try
            {
                EnumChildWindows(parent, EnumChildWindowsCallback, GCHandle.ToIntPtr(listHandle));
            }
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1003]WinApiX.GetChildWindows --> " +
                "Params: parent: " + parent.ToString() +
                "Message Exception: " + ex.Message.ToString());
                throw new Exception(ex.Message);
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
            RegistryManager rm = new RegistryManager();

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
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1004]WinApiX.GetTextX --> " +
                "Params: handler: " + handle.ToString() +
                "Message Exception: " + ex.Message.ToString());
                throw new Exception(ex.Message);
            }
        }
        public string GetText(string windowTitle, int index)
        {
            var sb = new StringBuilder();
            RegistryManager rm = new RegistryManager();
            try
            {
                var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                var childWindows = GetChildWindows(windowHWnd);
                var childWindowText = GetTextX(childWindows.ToArray()[index]);
                sb.Append(childWindowText);
                return sb.ToString();
            }
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1005]WinApiX.GetText --> " +
                    "Params: windowTitle: " + windowTitle +
                    " Index: " + index.ToString() +
                    "Message Exception: " + ex.Message.ToString());
                throw new Exception(ex.Message);
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

                RegistryManager rm = new RegistryManager();
                rm.WriteRegistryEvents("SendText", 1);
            }

            catch (Exception e)
            {
                throw new Exception(e.Message);
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

                RegistryManager rm = new RegistryManager();
                rm.ReadRegistryEvents("SendClick");
                rm.WriteRegistryEvents("SendClick", 1);
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
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
            catch (Exception e)
            {
                throw new Exception(e.Message);
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
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }

        private void SendDoubleClickListBox(string windowTitle, int index)
        {
            RegistryManager rm = new RegistryManager();

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
            catch (Exception ex)
            {
                rm.SetAlarm("[Erro-1006]WinApiX.SendDoubleClickListBox --> " +
                    "Params: windowTitle: " + windowTitle +
                    " Index: " + index.ToString() + 
                    "Message Exception: " + ex.Message.ToString() );
                throw new Exception(ex.Message);
            }
        }

        public string GetPhonesGEO()
        {
            RegistryManager registryManager = new RegistryManager();
            string alarmMessage = "";

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


                mainTitle = registryManager.ReadRegistryValue("GEOMain_Title");
                childTitle = registryManager.ReadRegistryValue("GEOChild_Title");

                if (!ScreenActive(mainTitle) && !ScreenActive(childTitle))
                {
                    return "";
                }

                if (ScreenActive(childTitle)) //12,11,10,3,4
                {
                    idxPhone1 = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexPhone1")); 
                    idxPhone2 = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexPhone2"));
                    idxPhone3 = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexPhone3"));
                    idxPhone4 = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexPhone4"));
                    idxPhone4A = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexPhone4A"));
                    phone1 = Regex.Replace(GetText(childTitle, idxPhone1), "[^0-9]", "");
                    phone2 = Regex.Replace(GetText(childTitle, idxPhone2), "[^0-9]", "");
                    phone3 = Regex.Replace(GetText(childTitle, idxPhone3), "[^0-9]", "");

                }
                else
                {
                    idxPhone1 = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone1"));
                    idxPhone2 = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone2"));
                    idxPhone3 = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone3"));
                    idxPhone4 = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone4"));
                    idxPhone4A = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone4A"));
                    idxPhone4L = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexPhone4L"));

                    phone1 = Regex.Replace(GetText(mainTitle, idxPhone1), "[^0-9]", "");
                    phone2 = Regex.Replace(GetText(mainTitle, idxPhone2), "[^0-9]", "");
                    phone3= Regex.Replace(GetText(mainTitle, idxPhone3), "[^0-9]", "");

                    Thread.Sleep(200);
                    SendDoubleClickListBox(mainTitle, idxPhone4L);
                    Thread.Sleep(200);
                }

                phone4A = Regex.Replace(GetText(childTitle, idxPhone4A), "[^0-9]", "");
                phone4 = phone4A + Regex.Replace(GetText(childTitle, idxPhone4), "[^0-9]", "");

                phone1 = phone1.Trim();

                if (phone1 == "" || phone1 == null || phone1 == ",")
                {
                    alarmMessage = "[Erro-2001]GetPhonesGEO - Phone1 null or Empty - ";
                    
                }
                else
                {
                    countGeoFoneTotal++;
                    registryManager.WriteRegistryEvents("GeoFoneComCount", 1);
                    if (phone1.Substring(0, 1) == "0") { phone1 = phone1.Substring(1, phone1.Length - 1); }
                    phonesMainTitle = phone1 + ",";
                }

                if (phone2 == "" || phone2 == null || phone2 == ",")
                {
                    alarmMessage += alarmMessage +  "[Erro-2002]GetPhonesGEO - Phone2 null or Empty - ";
                    
                }
                else
                {
                    countGeoFoneTotal++;
                    registryManager.WriteRegistryEvents("GeoFoneResCount", 1);
                    if (phone2.Substring(0, 1) == "0") { phone2 = phone2.Substring(1, phone2.Length - 1); }
                    phonesMainTitle = phonesMainTitle + phone2 + ",";
                }

                if (phone3 == "" || phone3 == null || phone3 == ",")
                {
                    alarmMessage += alarmMessage + "[Erro-2003]GetPhonesGEO - Phone3 null or Empty - ";
                }
                else
                {
                    countGeoFoneTotal++;
                    registryManager.WriteRegistryEvents("GeoFoneCelCount", 1);
                    if (phone3.Substring(0, 1) == "0") { phone3 = phone3.Substring(1, phone3.Length - 1); }
                    phonesMainTitle = phonesMainTitle + phone3 + ",";
                }

                if (phone4 == "" || phone4 == null || phone4 == ",")
                {
                    if (phone4A == "" || phone4A == null)
                    {
                        alarmMessage += alarmMessage + "[Erro-2005]GetPhonesGEO - Phone4A null or Empty - ";
                    }
                   
                    alarmMessage += alarmMessage +  "[Erro-2004]GetPhonesGEO - Phone4 null or Empty - ";

                    phones = phonesMainTitle;
                }
                else
                {
                    countGeoFoneTotal++;
                    countGeoFoneTotal++;
                    registryManager.WriteRegistryEvents("GeoFoneSegTelaCount", 2);

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

                if (countGeoFoneTotal > 0)
                {
                    registryManager.WriteRegistryEvents("GeoTotalFoneCount", countGeoFoneTotal);
                }
                List<string> phoneList = phones.ToLower().Split(',').Distinct().ToList();
                phones = string.Join(",", phoneList);

                if (!String.IsNullOrWhiteSpace(alarmMessage))
                {
                    registryManager.SetAlarm(alarmMessage);
;                }

                return phones;
            }
            catch (Exception ex)
            {
                registryManager.SetAlarm(alarmMessage + " -->" +  ex.Message.ToString());

                throw new Exception(ex.Message);
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

        public string GetTextFields()
        {
            var sb = new StringBuilder();
            RegistryManager registryManager = new RegistryManager();
            int idxCampanha;
            int idxLista;
           

            try
            {
                string windowTitle = "";
                string childTitle = "";

                windowTitle = registryManager.ReadRegistryValue("GEOMain_Title");
                childTitle = registryManager.ReadRegistryValue("GEOChild_Title");

                if (!ScreenActive(windowTitle) && (!ScreenActive(childTitle)))
                {
                    return "";
                }

                if (ScreenActive(childTitle))
                {
                    idxCampanha = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexCampanha"));
                    idxLista = Int32.Parse(registryManager.ReadRegistryValue("GEOChild_IndexLista"));
                    var windowHWnd = FindWindowByCaption(IntPtr.Zero, childTitle);
                    var childWindows = GetChildWindows(windowHWnd);
                    var textCampanha = GetTextX(childWindows.ToArray()[idxCampanha]);
                    sb.Append(textCampanha);
                    sb.Append(",");
                    var textLista = GetTextX(childWindows.ToArray()[idxLista]);
                    sb.Append(textLista);
                    sb.Append(",");
                }
                else
                {
                    idxCampanha = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexCampanha"));
                    idxLista = Int32.Parse(registryManager.ReadRegistryValue("GEOMain_IndexLista"));
                    var windowHWnd = FindWindowByCaption(IntPtr.Zero, windowTitle);
                    var childWindows = GetChildWindows(windowHWnd);
                    var textCampanha = GetTextX(childWindows.ToArray()[idxCampanha]);
                    sb.Append(textCampanha);
                    sb.Append(",");
                    var textLista = GetTextX(childWindows.ToArray()[idxLista]);
                    sb.Append(textLista);
                    sb.Append(",");
                }

                return sb.ToString();

            }
            catch (Exception ex)
            {
                registryManager.SetAlarm("[Erro-1007]WinApiX.GetText() --> " +
                    "Message Exception: " + ex.Message.ToString());
                throw new Exception(ex.Message);
            }

        }
    }
}
