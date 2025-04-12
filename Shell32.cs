using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

/// <summary>
/// A class for executing Shell32 related operations.
/// </summary>
public static class Shell32
{
    // Define constants for SHFileOperation
    const int FO_DELETE = 0x0003;
    const int FOF_ALLOWUNDO = 0x0040;
    const int FOF_NOCONFIRMATION = 0x0010;

    // Define the SHFILEOPSTRUCT structure
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct SHFILEOPSTRUCT
    {
        public IntPtr hwnd;
        public int wFunc;
        public string pFrom;
        public string? pTo;
        public ushort fFlags;
        public bool fAnyOperationsAborted;
        public IntPtr hNameMappings;
        public string? lpszProgressTitle;
    }

    // P/Invocations
    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern int SHFileOperation(ref SHFILEOPSTRUCT FileOp);

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    private static extern uint ExtractIconEx(string lpszFile, int nIconIndex, IntPtr[] phiconLarge, IntPtr[] phiconSmall, uint nIcons);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern bool DestroyIcon(IntPtr handle);

    /// <summary>
    /// A Shell32 subclass for executing functions relating to the user's recycle bin.
    /// </summary>
    public static class RecycleBin
    {
        /// <summary>
        /// Asks the user to delete the specified path and returns true if the operation succeeds.
        /// </summary>
        /// <param name="path">The file that should be deleted.</param>
        /// <returns>A boolean indicating if the operation was a success.</returns>
        public static bool RecycleFile(string path)
        {
            string pathToDelete = @"" + path;
            
            SHFILEOPSTRUCT operation = new SHFILEOPSTRUCT();
            operation.wFunc = FO_DELETE;
            operation.pFrom = pathToDelete + '\0';
            operation.pTo = null;
            operation.fFlags = (ushort)(FOF_ALLOWUNDO);
            operation.hwnd = IntPtr.Zero;
            operation.fAnyOperationsAborted = false;
            operation.hNameMappings = IntPtr.Zero;
            operation.lpszProgressTitle = null;

            int result = SHFileOperation(ref operation);

            if (result == 0)
            {
                return true;
            }

            return false;
        }
    }

    /// <summary>
    /// A Shell32 subclass for retrieving Windows icons.
    /// </summary>
    public static class WindowsIcon
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hFile, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr LoadImage(IntPtr hinst, IntPtr lpszName, uint uType, int cxDesired, int cyDesired, uint fuLoad);

        const uint LOAD_LIBRARY_AS_DATAFILE = 0x00000002;
        const uint IMAGE_ICON = 1;
        const uint LR_DEFAULTSIZE = 0x00000040;
        const uint LR_LOADFROMFILE = 0x00000010;
        const uint LR_SHARED = 0x00008000;

        /// <summary>
        /// Gets and returns an icon by it's resource ID. Returns null if the icon doesn't exist.
        /// </summary>
        /// <param name="directory">The directory which contains the icon. Can be either "imageres.dll" or "shell32.dll".</param>
        /// <param name="resourceId">The icon's resource ID.</param>
        /// <param name="size">The size of the icon.</param>
        /// <returns>An Icon instance containing the specified icon.</returns>
        public static Icon GetIconByResourceId(string directory, int resourceId, int size = 32)
        {
            IntPtr hModule = LoadLibraryEx(directory, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE);
            if (hModule != IntPtr.Zero)
            {
                IntPtr hIcon = LoadImage(hModule, (IntPtr)resourceId, IMAGE_ICON, size, size, LR_DEFAULTSIZE | LR_SHARED);
                if (hIcon != IntPtr.Zero)
                {
                    return Icon.FromHandle(hIcon);
                }
            }
            return null;
        }
    }
}