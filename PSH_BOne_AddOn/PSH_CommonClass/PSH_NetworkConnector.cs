using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace PSH_BOne_AddOn

{
    public class NetworkConnector
    {

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]

        public struct NETRESOURCE
        {
            public uint dwScope;
            public uint dwType;
            public uint dwDisplayType;
            public uint dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }


        [DllImport("mpr.dll", CharSet = CharSet.Auto)]

        public static extern int WNetUseConnection(
                    IntPtr hwndOwner,
                    [MarshalAs(UnmanagedType.Struct)] ref NETRESOURCE lpNetResource,
                    string lpPassword,
                    string lpUserID,
                    uint dwFlags,
                    StringBuilder lpAccessName,
                    ref int lpBufferSize,
                    out uint lpResult);

       

        

        public int TryConnectNetwork(string remotePath, string userID, string pwd, string strLocalName)
        {
            int capacity = 64;
            uint resultFlags = 0;
            uint flags = 0;

            StringBuilder sb = new StringBuilder(capacity);
            NETRESOURCE ns = new NETRESOURCE();
            ns.dwType = 1; // 공유 디스크
            ns.lpLocalName = strLocalName;  // 로컬 드라이브 지정하지 않음
            ns.lpRemoteName = remotePath;
            ns.lpProvider = null;

            
            int result = WNetUseConnection(IntPtr.Zero, ref ns, pwd, userID, flags, sb, ref capacity, out resultFlags);
            return result;
        }


        [DllImport("mpr.dll", EntryPoint = "WNetCancelConnection2", CharSet = CharSet.Auto)]

        public static extern int WNetCancelConnection2(String lpName, int dwFlags, int fForce);

        //public void DisconnectNetwork()
        //{
        //    WNetCancelConnection2A(NETRESOURCE.lpRemoteName, 1, 0);
        //}
        public void DisconnectNetwork(string strRemoteConnectString)
        {
            WNetCancelConnection2(strRemoteConnectString, 1, 1);
        }

    }
}