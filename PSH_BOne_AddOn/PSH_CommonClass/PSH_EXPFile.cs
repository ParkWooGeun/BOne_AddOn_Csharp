using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace PSH_BOne_AddOn.ExportFile
{
    class EXPFile
    {
        [DllImport("ExportCustomFile.dll")]
        public static extern int NTS_GetFileSize([MarshalAs(UnmanagedType.LPStr)] string szIn, [MarshalAs(UnmanagedType.LPStr)] string szPassword, [MarshalAs(UnmanagedType.LPStr)] string szName, int bAnsi);

        [DllImport("ExportCustomFile.dll")]
        public static extern int NTS_GetFileBuf([MarshalAs(UnmanagedType.LPStr)] string szIn, [MarshalAs(UnmanagedType.LPStr)] string szPassword, [MarshalAs(UnmanagedType.LPStr)] string szName, [In, Out] byte[] pcBuffer, int bAnsi);
    }
}
