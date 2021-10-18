using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using VB = Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    internal partial class FileListBoxForm : System.Windows.Forms.Form
    {
        private struct OPENFILENAME
        {
            public int lStructSize;
            public int hwndOwner;
            public int hInstance;
            public string lpstrFilter;
            public string lpstrCustomFilter;
            public int nMaxCustFilter;
            public int nFilterIndex;
            public string lpstrFile;
            public int nMaxFile;
            public string lpstrFileTitle;
            public int nMaxFileTitle;
            public string lpstrInitialDir;
            public string lpstrTitle;
            public int Flags;
            public short nFileOffset;
            public short nFileExtension;
            public string lpstrDefExt;
            public int lCustData;
            public int lpfnHook;
            public string IpTemplateName;
        }

        OPENFILENAME ofn;
        private const short OFN_READONLY = 0x1;
        private const short OFN_OVERWRITEPROMPT = 0x2;
        private const short OFN_HIDEREADONLY = 0x4;
        private const short OFN_NOCHANGEDIR = 0x8;
        private const short OFN_SHOWHELP = 0x10;
        private const short OFN_ENABLEHOOK = 0x20;
        private const short OFN_ENABLETEMPLATE = 0x40;
        private const short OFN_ENABLETEMPLATEHANDLE = 0x80;
        private const short OFN_NOVALIDATE = 0x100;

        private const short OFN_ALLOWMULTISELECT = 0x200;
        private const short OFN_EXTENSIONDIFFERENT = 0x400;
        private const short OFN_PATHMUSTEXIST = 0x800;
        private const short OFN_FILEMUSTEXIST = 0x1000;
        private const short OFN_CREATEPROMPT = 0x2000;
        private const short OFN_SHAREAWARE = 0x400;
        private const int OFN_NOREADONLYRETURN = 0x8000;
        private const int OFN_NOTESTFILECREATE = 0x10000;
        private const int OFN_NONENETWORKBUTTON = 0x20000;
        private const int OFN_NOLONGNAMES = 0x40000;
        private const int OFN_EXPLORER = 0x80000;
        private const int OFN_NODEREFERENCELINKS = 0x100000;
        private const int OFN_LONGNAMES = 0x20000;

        private const short OFN_SHAREFALLTHROUGH = 2;
        private const short OFN_SHARENOWARN = 1;
        private const short OFN_SHAREWARN = 0;

        private const short conHwndTopmost = -1;
        private const short conHwndNoTopmost = -2;
        private const short conSwpNoActivate = 0x10;

        private const short conSwpShowWindow = 0x40;
        [DllImport("comdlg32.dll", EntryPoint = "GetOpenFileNameA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int GetOpenFileName(ref OPENFILENAME pOpenfilename);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        private static extern int DestroyWindow(int hwnd);

        public string OpenDialog(System.Windows.Forms.Form Form1, string Filter_String, string Title, string InitDir)
        {
            string returnValue = null;

            int A = 0;

            ofn.lStructSize = Strings.Len(ofn);
            ofn.hwndOwner = Form1.Handle.ToInt32();
            //ofn.hInstance = Microsoft.VisualBasic.Compatibility.VB6.Support.GetHInstance().ToInt32();
            if (Strings.Right(Filter_String, 1) != "|")
            {
                Filter_String = Filter_String + "|";
            }

            Filter_String = Filter_String.Replace("|", Strings.Chr(0).ToString());

            ofn.lpstrFilter = Filter_String;
            ofn.lpstrFile = Strings.Space(254);
            ofn.nMaxFile = 255;
            ofn.lpstrFileTitle = Strings.Space(254);
            ofn.nMaxFileTitle = 255;
            ofn.lpstrInitialDir = InitDir;
            ofn.lpstrTitle = Title;
            ofn.Flags = OFN_HIDEREADONLY | OFN_FILEMUSTEXIST;
            A = GetOpenFileName(ref ofn);

            if ((A != 0))
            {
                returnValue = Strings.Trim(ofn.lpstrFile);
            }
            else
            {
                returnValue = "";
            }
            return returnValue;
        }

        private void FileListBoxForm_Load(System.Object eventSender, System.EventArgs eventArgs)
        {
            SetWindowPos(Handle.ToInt32(), conHwndTopmost, 0, 0, 0, 0, conSwpNoActivate | conSwpShowWindow);
        }
    }
}
