using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        try 
        {
            // 确保控制台编码设置为UTF-8
            Console.OutputEncoding = Encoding.UTF8;
            // 可能需要明确设置输入编码
            Console.InputEncoding = Encoding.UTF8;
        }
        catch (Exception ex)
        {
            // 如果无法设置编码，继续执行而不终止程序
            Console.WriteLine($"警告: 无法设置控制台编码: {ex.Message}");
        }


        try
        {
            // 解析命令行参数
            if (args.Length < 2)
            {
                Console.WriteLine("错误: 请提供目标路径和源路径参数");
                Console.WriteLine("用法: scripts.exe <targetPath> <sourcePath>");
                return;
            }

            string targetPath = args[0];
            string sourcePath = args[1];

            Console.WriteLine($"目标路径: {targetPath}");
            Console.WriteLine($"源路径: {sourcePath}");

            // 验证路径
            if (!Directory.Exists(targetPath))
            {
                Console.WriteLine($"目标路径不存在，尝试创建: {targetPath}");
                try
                {
                    Directory.CreateDirectory(targetPath);
                    Console.WriteLine("目标路径创建成功");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"错误: 无法创建目标路径: {ex.Message}");
                    return;
                }
            }

            if (!Directory.Exists(sourcePath))
            {
                Console.WriteLine($"错误: 源路径不存在: {sourcePath}");
                return;
            }

            // 收集源路径下的所有文件夹
            Console.WriteLine("开始收集源路径下的文件夹...");
            var directories = GetAllDirectories(sourcePath);
            Console.WriteLine($"找到 {directories.Count} 个文件夹");

            // 使用线程安全的计数器
            int created = 0, skipped = 0, failed = 0;
            var options = new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount };

            // 并行处理文件夹
            Parallel.ForEach(directories, options, directory =>
            {
                try
                {
                    var folderName = Path.GetFileName(directory);
                    if (!HasSevenDigitPrefix(folderName))
                    {
                        Console.WriteLine($"跳过: 文件夹名称不符合要求 - {directory}");
                        Interlocked.Increment(ref skipped);
                        return;
                    }

                    string shortcutName = Path.GetFileName(directory);
                    string shortcutPath = Path.Combine(targetPath, $"{shortcutName}.lnk");

                    if (HasShortcutForTarget(targetPath, directory))
                    {
                        Console.WriteLine($"跳过: 已存在指向该目录的快捷方式 - {directory}");
                        Interlocked.Increment(ref skipped);
                        return;
                    }

                    if (File.Exists(shortcutPath))
                    {
                        Console.WriteLine($"跳过: 快捷方式已存在 - {shortcutPath}");
                        Interlocked.Increment(ref skipped);
                        return;
                    }

                    if (CreateShortcut(shortcutPath, directory, $"Shortcut to {shortcutName}"))
                    {
                        Console.WriteLine($"创建: {shortcutPath} -> {directory}");
                        Interlocked.Increment(ref created);
                    }
                    else
                    {
                        Console.WriteLine($"错误: 无法创建 {directory} 的快捷方式");
                        Interlocked.Increment(ref failed);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"错误: 无法创建 {directory} 的快捷方式: {ex.Message}");
                    Interlocked.Increment(ref failed);
                }
            });

            // 输出统计信息
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("操作完成!");
            Console.WriteLine($"已创建: {created} 个快捷方式");
            Console.WriteLine($"已跳过: {skipped} 个快捷方式");
            Console.WriteLine($"失败: {failed} 个快捷方式");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"发生未处理异常: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static List<string> GetAllDirectories(string sourcePath)
    {
        var result = new List<string>();
        try
        {

            // 递归获取所有子目录
            foreach (var directory in Directory.GetDirectories(sourcePath, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    result.Add(directory);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"警告: 无法访问目录 {directory}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"警告: 列出目录时出错: {ex.Message}");
        }
        return result;
    }

    internal static bool CreateShortcut(string shortcutPath, string targetPath, string description)
    {
        try
        {
            IShellLinkW link = (IShellLinkW)new ShellLink();
            link.SetPath(targetPath);

            if (!string.IsNullOrEmpty(description))
            {
                link.SetDescription(description);
            }

            string workingDirectory = Path.GetDirectoryName(targetPath) ?? targetPath;
            link.SetWorkingDirectory(workingDirectory);

            IPersistFile file = (IPersistFile)link;
            file.Save(shortcutPath, true);

            return File.Exists(shortcutPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建快捷方式时出错: {ex.Message}");
            return false;
        }
    }

    internal static bool HasSevenDigitPrefix(string name)
    {
        if (string.IsNullOrEmpty(name) || name.Length < 7)
        {
            return false;
        }

        for (int i = 0; i < 7; i++)
        {
            if (!char.IsDigit(name[i]))
            {
                return false;
            }
        }

        return true;
    }

    internal static bool HasShortcutForTarget(string shortcutsRoot, string targetPath)
    {
        if (!Directory.Exists(shortcutsRoot))
        {
            return false;
        }

        string normalizedTarget = NormalizePath(targetPath);

        foreach (var shortcut in Directory.EnumerateFiles(shortcutsRoot, "*.lnk", SearchOption.TopDirectoryOnly))
        {
            try
            {
                string? existingTarget = GetShortcutTarget(shortcut);
                if (existingTarget is null)
                {
                    continue;
                }

                if (string.Equals(NormalizePath(existingTarget), normalizedTarget, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"警告: 检查快捷方式 {shortcut} 时出错: {ex.Message}");
            }
        }

        return false;
    }

    internal static string? GetShortcutTarget(string shortcutPath)
    {
        if (!File.Exists(shortcutPath))
        {
            return null;
        }

        IShellLinkW link = (IShellLinkW)new ShellLink();
        IPersistFile file = (IPersistFile)link;
        file.Load(shortcutPath, 0);

        StringBuilder buffer = new StringBuilder(1024);
        link.GetPath(buffer, buffer.Capacity, out var data, 0);
        string target = buffer.ToString();

        return string.IsNullOrWhiteSpace(target) ? null : target;
    }

    internal static string NormalizePath(string path)
    {
        string fullPath = Path.GetFullPath(path);
        return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WIN32_FIND_DATAW
    {
        public uint dwFileAttributes;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
        public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
        public uint nFileSizeHigh;
        public uint nFileSizeLow;
        public uint dwReserved0;
        public uint dwReserved1;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string cFileName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 14)]
        public string cAlternate;
    }

    [ComImport]
    [Guid("00021401-0000-0000-C000-000000000046")]
    private class ShellLink
    {
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000214F9-0000-0000-C000-000000000046")]
    private interface IShellLinkW
    {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, out WIN32_FIND_DATAW pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("0000010B-0000-0000-C000-000000000046")]
    private interface IPersistFile
    {
        void GetClassID(out Guid pClassID);
        void IsDirty();
        void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
        void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, [MarshalAs(UnmanagedType.Bool)] bool fRemember);
        void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
        void GetCurFile([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFileName);
    }
}
