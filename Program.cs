using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Concurrent;

class Program
{
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
                    if (!char.IsDigit(Path.GetFileName(directory)[0]))
                    {
                        Console.WriteLine($"跳过: 文件夹名称不符合要求 - {directory}");
                        Interlocked.Increment(ref skipped);
                        return;
                    }

                    string relativePath = Path.GetRelativePath(sourcePath, directory);
                    string shortcutName = Path.GetFileName(directory);
                    string shortcutPath = Path.Combine(targetPath, $"{shortcutName}.lnk");

                    if (File.Exists(shortcutPath))
                    {
                        Console.WriteLine($"跳过: 快捷方式已存在 - {shortcutPath}");
                        Interlocked.Increment(ref skipped);
                        return;
                    }

                    if (CreateShortcutWithPowerShell(shortcutPath, directory, $"Shortcut to {shortcutName}"))
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
            // 添加顶级目录
            result.Add(sourcePath);

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

    static bool CreateShortcutWithPowerShell(string shortcutPath, string targetPath, string description)
    {
        try
        {
            // 使用PowerShell创建快捷方式
            string psCommand = $@"$WshShell = New-Object -ComObject WScript.Shell; " +
                              $@"$Shortcut = $WshShell.CreateShortcut('{shortcutPath.Replace("'", "''")}'); " +
                              $@"$Shortcut.TargetPath = '{targetPath.Replace("'", "''")}'; " +
                              $@"$Shortcut.Description = '{description.Replace("'", "''")}'; " +
                              $@"$Shortcut.Save()";

            // 创建PowerShell进程
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "powershell.exe";
                process.StartInfo.Arguments = $"-Command \"{psCommand}\"";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.StandardOutputEncoding = Encoding.UTF8;
                process.StartInfo.StandardErrorEncoding = Encoding.UTF8;

                process.Start();
                process.WaitForExit();

                // 检查是否成功
                return process.ExitCode == 0 && File.Exists(shortcutPath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"创建快捷方式时出错: {ex.Message}");
            return false;
        }
    }
}
