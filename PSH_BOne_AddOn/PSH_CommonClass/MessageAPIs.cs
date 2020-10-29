using System;
using System.Runtime.InteropServices;

namespace PSH_BOne_AddOn
{
    static class MessageAPIs
    {
        public static Msg structMsg;

        //Part of the MSG structure - receives the location of the mouse
        public struct POINTAPI
        {
            public int x;
            public int y;
        }

        //The message structure
        public struct Msg
        {
            public int hwnd;
            public int message;
            public int wParam;
            public int lParam;
            public int Time;
            public POINTAPI pt;
        }

        //Retrieves messages sent to the calling thread's message queue
        [DllImport("user32", EntryPoint = "GetMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern Boolean GetMessage(ref Msg lpMsg, int hwnd, int wMsgFilterMin, int wMsgFilterMax);

        //Translates virtual-key messages into character messages
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int TranslateMessage(ref Msg lpMsg);

        //Forwards the message on to the window represented by the
        //hWnd member of the Msg structure
        [DllImport("user32", EntryPoint = "DispatchMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int DispatchMessage(ref Msg lpMsg);

        ////파일명의 최대 길이
        //public const short MAX_PATH = 260;

        //public struct FILETIME
        //{
        //    //리턴받을 파일에 관한 일부 세부정보 구조체
        //    public int dwLowDateTime;
        //    public int dwHighDateTime;
        //}

        //public struct WIN32_FIND_DATA
        //{
        //    //검색된 파일 또는 하위디렉토리의 정보를 받을 구조체
        //    public int dwFileAttributes;
        //    public FILETIME ftCreationTime;
        //    public FILETIME ftLastAccessTime;
        //    public FILETIME ftLastWriteTime;
        //    public int nFileSizeHigh;
        //    public int nFileSizeLow;
        //    public int dwReserved0;
        //    public int dwReserved1;
        //    [VBFixedString(MessageAPIs.MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = MessageAPIs.MAX_PATH)]
        //    public char[] cFileName;
        //    [VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 14)]
        //    public char[] cAlternate;
        //}
        //[DllImport("kernel32", EntryPoint = "FindFirstFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        //public static extern int FindFirstFile(string lpFileName, ref WIN32_FIND_DATA lpFindFileData);
        //[DllImport("kernel32", EntryPoint = "FindNextFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        //public static extern int FindNextFile(int hFindFile, ref WIN32_FIND_DATA lpFindFileData);

        //[DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        //public static extern void Sleep(int dwMilliseconds);

        ////파일 조작에 관련된 정보를 정의하는 사용자정의 데이터형
        //public struct SHFILEOPSTRUCT
        //{
        //    public int hwnd;
        //    public int wfunc;
        //    public string pfrom;
        //    public string pto;
        //    public int fFlags;
        //    public int fAnyOperationsAborted;
        //    public int hNamemappings;
        //    public string lpszProgressTitle;
        //}
        ////파일 조작하는 함수
        //[DllImport("shell32.dll", EntryPoint = "SHFileOperationA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        //public static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);
    }
}
