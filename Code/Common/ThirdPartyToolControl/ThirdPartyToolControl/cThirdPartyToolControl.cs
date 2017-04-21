using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace ThirdPartyToolControl
{
    public class cThirdPartyToolControl
    {
        #region -- CONST --

        // WM
        const int WM_SETFOCUS       = 0x0007;
        const int WM_KILLFOCUS      = 0x0008;
        const int WM_CLOSE          = 0x0010;
        const int WM_SETTEXT        = 0x000C;
        const int WM_KEYDOWN        = 0x0100;
        const int WM_KEYUP          = 0x0101;
        const int WM_CHAR           = 0x0102;
        const int WM_SYSKEYDOWN     = 0x0104;
        const int WM_SYSKEYUP       = 0x0105;
        const int WM_COMMAND        = 0x0111;
        const int WM_SYSCOMMAND     = 0x0112;
        const int WM_RBUTTONDOWN    = 0x0204;
        const int WM_RBUTTONUP      = 0x0205;
        
        // GW
        const int GW_HWNDNEXT       = 0x0002;
        const int GW_HWNDPREV       = 0x0003;

        // virtual key
        const int VK_BACK           = 0x0008;
        const int VK_TAB            = 0x0009;
        const int VK_CLEAR          = 0x000C;
        const int VK_RETURN         = 0x000D;
        const int VK_SHIFT          = 0x0010;
        const int VK_CONTROL        = 0x0011;
        const int VK_ESCAPE         = 0x001B;
        const int VK_SNAPSHOT       = 0x002A;
        const int VK_UP             = 0x0026;
        const int VK_DOWN           = 0x0028;
        const int VK_F1             = 0x0070;
        const int VK_F2             = 0x0071;
        const int VK_F3             = 0x0072;
        const int VK_F4             = 0x0073;
        const int VK_F5             = 0x0074;
        const int VK_F6             = 0x0075;
        const int VK_F7             = 0x0076;
        const int VK_F8             = 0x0077;
        const int VK_F9             = 0x0078;
        const int VK_F10            = 0x0079;
        const int VK_F11            = 0x007A;
        const int VK_F12            = 0x007B;
        

        // Button Control Messages
        const int BM_GETCHECK       = 0x00F0;
        const int BM_SETCHECK       = 0x00F1;
        const int BM_GETSTATE       = 0x00F2;
        const int BM_SETSTATE       = 0x00F3;
        const int BM_SETSTYLE       = 0x00F4;
        const int BM_CLICK          = 0x00F5;
        const int BM_GETIMAGE       = 0x00F6;
        const int BM_SETIMAGE       = 0x00F7;
        const int BM_SETDONTCLICK   = 0x00F8;

        const int BST_UNCHECKED     = 0x0000;
        const int BST_CHECKED       = 0x0001;
        const int BST_INDETERMINATE = 0x0002;
        const int BST_PUSHED        = 0x0004;
        const int BST_FOCUS         = 0x0008;

        // others
        const int SC_CLOSE          = 0xF060;
        const int MK_RBUTTON        = 0x0002;
        const int CB_SETCURSEL      = 0x014E;
        const int LB_GETCOUNT       = 0x018B;
        const int LB_GETTEXT        = 0x0189;
        const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
        const uint KEYEVENTF_KEYUP = 0x0002;
        
        #endregion

        #region Windows -- user32 API --

        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern int SendMessage(int hWnd, int Msg, int wParam, string lParam);

        [DllImport("user32.dll")]
        private static extern int PostMessage(int hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        private static extern int EnumChildWindows(int hWndParent, CallBack lpfn, int lParam);

        [DllImport("user32.dll")]
        private static extern int GetWindowText(int hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern int FindWindow(string strclassName, string strWindowName);

        [DllImport("user32.dll")]
        private static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern int GetDlgItem(int hDlg, int nIDDlgItem);

        [DllImport("user32.dll", EntryPoint = "CheckRadioButton")]
        private static extern bool CheckRadioButton(IntPtr hwnd, int firstID, int lastID, int checkedID);

        [DllImport("user32.dll")]
        private static extern int SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);

        [DllImport("user32.dll")]
        private static extern int GetWindowRect(int hwnd, ref RECT lpRect);

        [DllImport("user32")]
        private static extern int SetFocus(int hwnd);

        [DllImport("user32")]
        private static extern int GetWindow(int hwnd, int Msg);

        //[DllImport("user32.dll", SetLastError = true)]
        //public static extern bool RegisterHotKey(IntPtr hWnd, int id, int iModifiers, Keys vk); 

        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(int hWnd);

        // load ini files
        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern bool WritePrivateProfileString(string sectionName, string keyName, string keyValue, string filePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetPrivateProfileString(string sectionName, string keyName, string defaultReturnString, StringBuilder returnString, int returnStringLength, string filePath);

        public delegate bool CallBack(int hwnd, int lParam);

        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }
        #endregion


        //// Member variable ////
        public int V_WM_SETFOCUS { get { return WM_SETFOCUS; } }
        public int V_WM_KILLFOCUS { get { return WM_KILLFOCUS; } }
        public int V_WM_CLOSE { get { return WM_CLOSE; } }
        public int V_WM_SETTEXT { get { return WM_SETTEXT; } }
        public int V_WM_KEYDOWN { get { return WM_KEYDOWN; } }
        public int V_WM_KEYUP { get { return WM_KEYUP; } }
        public int V_WM_CHAR { get { return WM_CHAR; } }
        public int V_WM_SYSKEYDOWN { get { return WM_SYSKEYDOWN; } }
        public int V_WM_SYSKEYUP { get { return WM_SYSKEYUP; } }
        public int V_WM_COMMAND { get { return WM_COMMAND; } }
        public int V_WM_SYSCOMMAND { get { return WM_SYSCOMMAND; } }
        public int V_WM_RBUTTONDOWN { get { return WM_RBUTTONDOWN; } }
        public int V_WM_RBUTTONUP { get { return WM_RBUTTONUP; } }

        public int V_GW_HWNDNEXT { get { return GW_HWNDNEXT; } }
        public int V_GW_HWNDPREV { get { return GW_HWNDPREV; } }

        public int V_VK_BACK { get { return VK_BACK; } }
        public int V_VK_TAB { get { return VK_TAB; } }
        public int V_VK_CLEAR { get { return VK_CLEAR; } }
        public int V_VK_RETURN { get { return VK_RETURN; } }
        public int V_VK_SHIFT { get { return VK_SHIFT; } }
        public int V_VK_CONTROL { get { return VK_CONTROL; } }
        public int V_VK_ESCAPE { get { return VK_ESCAPE; } }
        public int V_VK_SNAPSHOT { get { return VK_SNAPSHOT; } }
        public int V_VK_UP { get { return VK_UP; } }
        public int V_VK_DOWN { get { return VK_DOWN; } }
        public int V_VK_F1 { get { return VK_F1; } }
        public int V_VK_F2 { get { return VK_F2; } }
        public int V_VK_F3 { get { return VK_F3; } }
        public int V_VK_F4 { get { return VK_F4; } }
        public int V_VK_F5 { get { return VK_F5; } }
        public int V_VK_F6 { get { return VK_F6; } }
        public int V_VK_F7 { get { return VK_F7; } }
        public int V_VK_F8 { get { return VK_F8; } }
        public int V_VK_F9 { get { return VK_F9; } }
        public int V_VK_F10 { get { return VK_F10; } }
        public int V_VK_F11 { get { return VK_F11; } }
        public int V_VK_F12 { get { return VK_F12; } }

        public int V_BM_GETCHECK { get { return BM_GETCHECK; } }
        public int V_BM_SETCHECK { get { return BM_SETCHECK; } }
        public int V_BM_GETSTATE { get { return BM_GETSTATE; } }
        public int V_BM_SETSTATE { get { return BM_SETSTATE; } }
        public int V_BM_SETSTYLE { get { return BM_SETSTYLE; } }
        public int V_BM_CLICK { get { return BM_CLICK; } }
        public int V_BM_GETIMAGE { get { return BM_GETIMAGE; } }
        public int V_BM_SETIMAGE { get { return BM_SETIMAGE; } }
        public int V_BM_SETDONTCLICK { get { return BM_SETDONTCLICK; } }

        public int V_BST_UNCHECKED { get { return BST_UNCHECKED; } }
        public int V_BST_CHECKED { get { return BST_CHECKED; } }
        public int V_BST_INDETERMINATE { get { return BST_INDETERMINATE; } }
        public int V_BST_PUSHED { get { return BST_PUSHED; } }
        public int V_BST_FOCUS { get { return BST_FOCUS; } }

        public int V_SC_CLOSE { get { return SC_CLOSE; } }
        public int V_MK_RBUTTON { get { return MK_RBUTTON; } }
        public int V_CB_SETCURSEL { get { return CB_SETCURSEL; } }
        public int V_LB_GETCOUNT { get { return LB_GETCOUNT; } }
        public int V_LB_GETTEXT { get { return LB_GETTEXT; } }
        public uint V_KEYEVENTF_EXTENDEDKEY { get { return KEYEVENTF_EXTENDEDKEY; } }
        public uint V_KEYEVENTF_KEYUP { get { return KEYEVENTF_KEYUP; } }

        //// Member function ////
        public int F_SendMessage(int hWnd, int Msg, int wParam, int lParam)
        {
            return SendMessage(hWnd, Msg, wParam, lParam);
        }

        public int F_SendMessage(int hWnd, int Msg, int wParam, string lParam)
        {
            return SendMessage(hWnd, Msg, wParam, lParam);
        }

        public int F_PostMessage(int hWnd, int Msg, int wParam, int lParam)
        {
            return PostMessage(hWnd, Msg, wParam, lParam);
        }

        public int F_EnumChildWindows(int hWndParent, CallBack lpfn, int lParam)
        {
            return EnumChildWindows(hWndParent, lpfn, lParam);
        }

        public int F_GetWindowText(int hWnd, StringBuilder lpString, int nMaxCount)
        {
            return GetWindowText(hWnd, lpString, nMaxCount);
        }

        public int F_FindWindow(string strclassName, string strWindowName)
        {
            return FindWindow(strclassName, strWindowName);
        }

        public int F_FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow)
        {
            return FindWindowEx(hwndParent, hwndChildAfter, lpszClass, lpszWindow);
        }

        public int F_GetDlgItem(int hDlg, int nIDDlgItem)
        {
            return GetDlgItem(hDlg, nIDDlgItem);
        }

        public int F_SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags)
        {
            return SetWindowPos(hwnd, hWndInsertAfter, x, y, cx, cy, wFlags);
        }

        public int F_SetFocus(int hwnd)
        {
            return SetFocus(hwnd);
        }

        public bool F_SetForegroundWindow(int hWnd)
        {
            return SetForegroundWindow(hWnd);
        }

        public int F_GetWindow(int hwnd, int Msg)
        {
            return GetWindow(hwnd, Msg);
        }

        public bool F_WritePrivateProfileString(string sectionName, string keyName, string keyValue, string filePath)
        {
            return WritePrivateProfileString(sectionName, keyName, keyValue, filePath);
        }

        public int F_GetPrivateProfileString(string sectionName, string keyName, string defaultReturnString, StringBuilder returnString, int returnStringLength, string filePath)
        {
            return GetPrivateProfileString(sectionName, keyName, defaultReturnString, returnString, returnStringLength, filePath);
        }

        public void F_KeybdEvent(int bVk, int bScan, uint dwFlags, int dwExtraInfo)
        {
            keybd_event((byte)bVk, (byte)bScan, dwFlags, dwExtraInfo);
        }
        
    }
}
