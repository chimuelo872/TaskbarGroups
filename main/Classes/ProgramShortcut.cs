namespace TaskbarGroups.Classes
{
    public class ProgramShortcut
    {
        public string Arguments = "";
        public string WorkingDirectory = MainPath.exeString;

        public string FilePath { get; set; }
        public bool IsWindowsApp { get; set; }

        public string Name { get; set; } = "";
    }
}