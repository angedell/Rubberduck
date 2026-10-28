using System;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;

namespace Rubberduck.UI.CodeExplorer
{
    /// <summary>Whether it is fine to download something optional: there is a connection and Windows does not count it as metered, roaming or near a data limit.</summary>
    public static class NetworkCost
    {
        private const uint Unrestricted = 0x1;

        public static bool CanDownloadOptionalData()
        {
            try
            {
                return NetworkInterface.GetIsNetworkAvailable() && IsUnrestricted();
            }
            catch (Exception exception) when (exception is NetworkInformationException || exception is COMException || exception is InvalidCastException)
            {
                //When the cost cannot be told, nothing is downloaded.
                return false;
            }
        }

        private static bool IsUnrestricted()
        {
            var manager = (INetworkCostManager)new NetworkListManagerClass();
            try
            {
                //Exactly "unrestricted": any other flag (fixed or variable cost, over or approaching the limit, congested, roaming) or an unknown cost says no.
                return manager.GetCost(out var cost, IntPtr.Zero) == 0 && cost == Unrestricted;
            }
            finally
            {
                Marshal.ReleaseComObject(manager);
            }
        }

        [ComImport, Guid("DCB00C01-570F-4A9B-8D69-199FDBA5723B")]
        private class NetworkListManagerClass
        {
        }

        [ComImport, Guid("DCB00008-570F-4A9B-8D69-199FDBA5723B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface INetworkCostManager
        {
            [PreserveSig]
            int GetCost(out uint cost, IntPtr destinationAddress);

            [PreserveSig]
            int GetDataPlanStatus(IntPtr status, IntPtr destinationAddress);

            [PreserveSig]
            int SetDestinationAddresses(uint length, IntPtr addresses, [MarshalAs(UnmanagedType.VariantBool)] bool append);
        }
    }
}
