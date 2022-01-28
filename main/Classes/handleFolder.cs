namespace TaskbarGroups.Classes
{
    using System;
    using System.Drawing;
    using System.Runtime.InteropServices;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct Shfileinfo
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    internal static class handleFolder
    {
        public const uint ShgfiIcon = 0x000000100;
        public const uint ShgfiLargeicon = 0x000000000;
        public const uint FileAttributeDirectory = 0x00000010;

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, out Shfileinfo psfi,
            uint cbFileInfo, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DestroyIcon(IntPtr hIcon);

        public static Icon GetFolderIcon(string path)
        {
            // Need to add size check, although errors generated at present!
            var flags = ShgfiIcon | ShgfiLargeicon;

            // Get the folder icon
            var shfi = new Shfileinfo();

            var res = SHGetFileInfo(path, FileAttributeDirectory, out shfi, (uint) Marshal.SizeOf(shfi), flags);

            if (res == IntPtr.Zero)
                throw Marshal.GetExceptionForHR(Marshal.GetHRForLastWin32Error());

            // Load the icon from an HICON handle
            Icon.FromHandle(shfi.hIcon);

            // Now clone the icon, so that it can be successfully stored in an ImageList
            var icon = (Icon) Icon.FromHandle(shfi.hIcon).Clone();

            DestroyIcon(shfi.hIcon); // Cleanup

            return icon;
        }
    }
}