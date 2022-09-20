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
        public NETRESOURCE NetResource = new NETRESOURCE();
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

        [DllImport("mpr.dll", EntryPoint = "WNetCancelConnection2", CharSet = CharSet.Auto)]

        public static extern int WNetCancelConnection2A(String lpName, int dwFlags, int fForce);

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

        public int TryConnectNetwork(string remotePath, string userID, string pwd)
        {
            int capacity = 64;
            uint resultFlags = 0;
            uint flags = 0;
            StringBuilder sb = new StringBuilder(capacity);
            NetResource.dwType = 1; // 공유 디스크
            NetResource.lpLocalName = null;  // 로컬 드라이브 지정하지 않음
            NetResource.lpRemoteName = "\\\\191.1.1.220\\b1_shr";
            NetResource.lpProvider = null;

            
            int result = WNetUseConnection(IntPtr.Zero, ref NetResource, pwd, userID, flags, sb, ref capacity, out resultFlags);
            return result;
        }

        public void DisconnectNetwork()
        {
            WNetCancelConnection2A(NetResource.lpRemoteName, 1, 0);
        }
    }
}