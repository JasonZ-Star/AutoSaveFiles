using Newtonsoft.Json;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Channels;
using GetOfficeHook;
namespace AutoSaveFiles;

class Program
{
    static void Main()
    {
        var fileNameAndPathDic = new Dictionary<string, string>();
        var lastBackupFileNameDic = new Dictionary<string, string>();
        var backupFileNameAndPath = new Dictionary<string, string>();
        if (ConfigOperator.IsFirstTimeRun())
        {
            ConfigOperator.Initialize(ref fileNameAndPathDic, ref lastBackupFileNameDic, ref backupFileNameAndPath);
        }
        else
        {
            ConfigOperator.LoadFiles(ref fileNameAndPathDic, ref lastBackupFileNameDic, ref backupFileNameAndPath);
            if (!ConfigOperator.IsFirstTimeRun())
            {
                Console.WriteLine("是否需要添加新的备份文件？按Y添加");
                Thread.Sleep(1000);
                Console.WriteLine("5秒无操作后默认不添加");
                Thread.Sleep(1000);
                bool addNewBackup = false;
                for (int i = 5; i > 0; i--)
                {
                    Thread.Sleep(1000);
                    if (Console.KeyAvailable)
                    {
                        ConsoleKeyInfo key = Console.ReadKey(true);
                        // 检查用户是否按下了 'Y' 或 'y' 键
                        if (key.KeyChar == 'Y' || key.KeyChar == 'y')
                        {
                            Console.WriteLine($"你按下了{Char.ToUpper(key.KeyChar)}");
                            addNewBackup = true;
                            break;
                        }

                        if (key.KeyChar == 'N' || key.KeyChar == 'n')
                        {
                            Console.SetCursorPosition(0, Console.CursorTop);
                            break;
                        }
                    }

                }

                if (addNewBackup)
                {
                    ConfigOperator.LoadFiles
                        (ref fileNameAndPathDic, ref lastBackupFileNameDic, ref backupFileNameAndPath);
                }
                else
                {
                    Console.WriteLine("不添加新的文件");
                }

                // 暂停1秒
                Thread.Sleep(1000);
            }

            while (true)
            {
                foreach (var files in fileNameAndPathDic)
                {
                    var file = new MyFile(files.Value);
                    Console.WriteLine($"{files.Key}: {lastBackupFileNameDic[files.Key]}");
                    Console.WriteLine($"{lastBackupFileNameDic[files.Key]}" +
                                      $"--({File.GetLastWriteTime(files.Value)}): " +
                                      "-------" + $"{lastBackupFileNameDic[files.Key]}" +
                                      $"--({File.GetLastWriteTimeUtc(file.GetBackupFileNameAndPath(files.Value))})");
                }

                try
                {
                    Console.WriteLine($"{DateTime.Now}程序正常运行");
                    foreach (var file in fileNameAndPathDic)
                    {
                        var loadedFile = new MyFile(file.Value);
                        Console.WriteLine("------------------------------");
                        Console.WriteLine($"File-load: {file.Key}");
                        Console.WriteLine(File.GetLastWriteTime(fileNameAndPathDic[loadedFile.FileName]));
                        Console.WriteLine(
                            File.GetLastWriteTime(backupFileNameAndPath[lastBackupFileNameDic[loadedFile.FileName]]));
                        Console.WriteLine("------------------------------");
                        if (loadedFile.IsFileModified(file.Value))
                        {
                            Console.WriteLine("-------------------------Modified----------------");
                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("配置文件错误，请重新配置");
                }

                Thread.Sleep(30000);
            }
        }
    }


}