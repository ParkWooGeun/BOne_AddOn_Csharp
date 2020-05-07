using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace PSH_BOne_AddOn
{
    static class MessageAPIs
    {

        //// Part of the MSG structure - receives the location of the mouse
        public struct POINTAPI
        {
            public int x;
            public int y;
        }

        //// The message structure
        public struct Msg
        {
            public int hwnd;
            public int message;
            public int wParam;
            public int lParam;
            public int Time;
            public POINTAPI pt;
        }
        [DllImport("user32", EntryPoint = "GetMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]

        //// Retrieves messages sent to the calling thread's message queue
        //UPGRADE_WARNING: Msg 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern Boolean GetMessage(ref Msg lpMsg, int hwnd, int wMsgFilterMin, int wMsgFilterMax);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]

        //// Translates virtual-key messages into character messages
        //UPGRADE_WARNING: Msg 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern int TranslateMessage(ref Msg lpMsg);
        [DllImport("user32", EntryPoint = "DispatchMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]

        //// Forwards the message on to the window represented by the
        //// hWnd member of the Msg structure
        //UPGRADE_WARNING: Msg 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern int DispatchMessage(ref Msg lpMsg);

        //UPGRADE_NOTE: Msg이(가) Msg_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        public static Msg Msg_Renamed;
        [DllImport("kernel32", EntryPoint = "FindFirstFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]


        //UPGRADE_WARNING: WIN32_FIND_DATA 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern int FindFirstFile(string lpFileName, ref WIN32_FIND_DATA lpFindFileData);
        [DllImport("kernel32", EntryPoint = "FindNextFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        //UPGRADE_WARNING: WIN32_FIND_DATA 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern int FindNextFile(int hFindFile, ref WIN32_FIND_DATA lpFindFileData);

        ////파일명의 최대 길이
        public const short MAX_PATH = 260;

        public struct FILETIME
        {
            ////리턴받을 파일에 관한 일부 세부정보 구조체
            public int dwLowDateTime;
            public int dwHighDateTime;
        }

        public struct WIN32_FIND_DATA
        {
            ////검색된 파일 또는 하위디렉토리의 정보를 받을 구조체
            public int dwFileAttributes;
            public FILETIME ftCreationTime;
            public FILETIME ftLastAccessTime;
            public FILETIME ftLastWriteTime;
            public int nFileSizeHigh;
            public int nFileSizeLow;
            public int dwReserved0;
            public int dwReserved1;
            //UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
            [VBFixedString(MessageAPIs.MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = MessageAPIs.MAX_PATH)]
            public char[] cFileName;
            //UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
            [VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
            public char[] cAlternate;
        }
        [DllImport("shell32.dll", EntryPoint = "SHFileOperationA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]


        //파일 조작하는 함수
        //UPGRADE_WARNING: SHFILEOPSTRUCT 구조체에는 이 Declare 문에 인수로 전달할 마샬링 특성이 있어야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
        public static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);

        //파일 조작에 관련된 정보를 정의하는 사용자정의 데이터형
        public struct SHFILEOPSTRUCT
        {
            public int hwnd;
            public int wfunc;
            public string pfrom;
            public string pto;
            public int fFlags;
            public int fAnyOperationsAborted;
            public int hNamemappings;
            public string lpszProgressTitle;
        }
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]


        public static extern void Sleep(int dwMilliseconds);
    }
}
