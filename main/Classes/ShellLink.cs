namespace TaskbarGroups.Classes
{
    using System;
    using System.Runtime.InteropServices;
    using System.Text;

    internal static class ShellLink
    {
        public static void InstallShortcut(string exePath, string appId, string desc, string workingDirectory,
            string iconLocation, string saveLocation, string arguments)
        {
            // Use passed parameters as to construct the shortcut
            var newShortcut = (IShellLinkW) new CShellLink();
            newShortcut.SetPath(exePath);
            newShortcut.SetDescription(desc);
            newShortcut.SetWorkingDirectory(workingDirectory);
            newShortcut.SetArguments(arguments);
            newShortcut.SetIconLocation(iconLocation, 0);

            // Set the classID of the shortcut that is created
            // Crucial to avoid program stacking
            var newShortcutProperties = (IPropertyStore) newShortcut;

            var varAppId = new PropVariantHelper();
            varAppId.SetValue(appId);
            newShortcutProperties?.SetValue(Propertykey.AppUserModel_ID, varAppId.Propvariant);

            // Save the shortcut as per passed save location
            var newShortcutSave = (IPersistFile) newShortcut;
            newShortcutSave.Save(saveLocation, true);
        }

        internal class PropVariantHelper
        {
            private Propvariant variant;
            public Propvariant Propvariant => variant;

            public void SetValue(string val)
            {
                NativeMethods.PropVariantClear(ref variant);
                variant.vt = (ushort) VarEnum.VT_LPWSTR;
                variant.unionmember = Marshal.StringToCoTaskMemUni(val);
            }

            private static class NativeMethods
            {
                [DllImport("Ole32.dll", PreserveSig = false)]
                internal static extern void PropVariantClear(ref Propvariant pvar);
            }
        }

        #region COM APIs

        [ComImport]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IShellLinkW
        {
            void GetPath([Out] [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, IntPtr pfd,
                uint fFlags);

            void GetIDList(out IntPtr ppidl);

            void SetIDList(IntPtr pidl);

            void GetDescription([Out] [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxName);

            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);

            void GetWorkingDirectory([Out] [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);

            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);

            void GetArguments([Out] [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);

            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);

            void GetHotKey(out short wHotKey);

            void SetHotKey(short wHotKey);

            void GetShowCmd(out uint iShowCmd);

            void SetShowCmd(uint iShowCmd);

            void GetIconLocation([Out] [MarshalAs(UnmanagedType.LPWStr)] out StringBuilder pszIconPath, int cchIconPath,
                out int iIcon);

            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);

            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);

            void Resolve(IntPtr hwnd, uint fFlags);

            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        [ComImport]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IPersistFile
        {
            void GetCurFile([Out] [MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile);

            void IsDirty();

            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, [MarshalAs(UnmanagedType.U4)] long dwMode);

            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);

            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct Propvariant
        {
            [FieldOffset(0)] public ushort vt;

            [FieldOffset(8)] public IntPtr unionmember;

            [FieldOffset(8)] public ulong forceStructToLargeEnoughSize;
        }

        [ComImport]
        [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStore
        {
            void GetCount([Out] out uint propertyCount);

            void GetAt([In] uint propertyIndex, [Out] [MarshalAs(UnmanagedType.Struct)] out Propertykey key);

            void GetValue([In] [MarshalAs(UnmanagedType.Struct)] ref Propertykey key,
                [Out] [MarshalAs(UnmanagedType.Struct)] out Propvariant pv);

            void SetValue([In] [MarshalAs(UnmanagedType.Struct)] ref Propertykey key,
                [In] [MarshalAs(UnmanagedType.Struct)] ref Propvariant pv);

            void Commit();
        }

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        [ClassInterface(ClassInterfaceType.None)]
        internal class CShellLink
        {
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct Propertykey
        {
            public Guid fmtid;
            public uint pid;

            public Propertykey(Guid guid, uint id)
            {
                fmtid = guid;
                pid = id;
            }

            public static readonly Propertykey AppUserModel_ID =
                new Propertykey(new Guid("9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3"), 5);
        }

        #endregion COM APIs
    }
}