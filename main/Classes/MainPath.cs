namespace TaskbarGroups.Classes
{
    // Function that is accessed by all forms to get the starting absolute path of the .exe
    // Added as to not keep generating the path in each form
    internal static class MainPath
    {
        public static string path;
        public static string exeString;
    }
}