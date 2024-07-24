using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace AutoSaveFiles
{
    internal class ConfigOperator
    {
        public static bool IsFirstTimeRun()
        {
            if (!File.Exists("config.json"))
            {
                File.Create("config.json").Close();
                return true;
            }
            return false;
        }

        private static void LoadConfig(ref Dictionary<string, string>? fileNameAndPathDic,
        ref Dictionary<string, string>? lastBackupFileNameDic,
        ref Dictionary<string, string> backupFileNameAndPath)
        {
            Console.WriteLine("检测到配置文件，正在读取中");
            Thread.Sleep(1000);
            var config = File.ReadAllText("config.json");
            if (String.IsNullOrEmpty(config))
            {
                Console.WriteLine("--------------------");
                Console.WriteLine("配置文件为空，请重新配置");
                Console.WriteLine("--------------------");
                Initialize(ref fileNameAndPathDic, ref lastBackupFileNameDic, ref backupFileNameAndPath);
            }
            else
            {
                try
                {
                    var allConfig =
                        JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(config);

                    if (allConfig != null && allConfig.TryGetValue("FilenameAndPath", out var value))
                    {
                        fileNameAndPathDic = value;
                        Console.WriteLine("---需备份文件数据加载成功---");
                        Thread.Sleep(1000);
                    }

                    if (allConfig != null && allConfig.TryGetValue("LastBackupFile", out var value1) &&
                        allConfig.TryGetValue("BackupFileNameAndPath", out var value2))
                    {
                        lastBackupFileNameDic = value1;
                        backupFileNameAndPath = value2;
                        Console.WriteLine("---已备份文件数据加载成功---");
                        Thread.Sleep(1000);
                    }

                    Console.WriteLine("已添加的文件如下");
                    foreach (var file in allConfig?["FilenameAndPath"]!)
                    {
                        Console.WriteLine(file.Key);
                        Thread.Sleep(200);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    Console.WriteLine("配置文件错误，请重新配置");
                    throw;
                }
            }
        }

        public static void Initialize(ref Dictionary<string, string>? fileNameAndPathDic,
            ref Dictionary<string, string>? lastBackupFileNameDic,
            ref Dictionary<string, string>? backupFileAndPathDic)
        {
            Console.WriteLine("--------------------------");
            Console.WriteLine("第一次运行，请跟随提示操作");
            Console.WriteLine("--------------------------");
            Thread.Sleep(1000);
            MyFile.CreateDirectory();
            Console.WriteLine("配置目录已创建");
            ConfigOperator.LoadFiles(ref fileNameAndPathDic, ref lastBackupFileNameDic, ref backupFileAndPathDic);

            // File.WriteAllText("./source/source.txt", sourcePath);
        }
    public static void LoadFiles
            (ref Dictionary<string, string>? fileNameAndPathDic,
                ref Dictionary<string, string>? lastBackupFileNameDic,
                ref Dictionary<string, string> backupFileNameAndPath)
        {
            string sourcePath = null;
            if (IsFirstTimeRun())
            {
                Console.WriteLine("请输入需要备份的文件路径（包含文件名称），多个文件路径请用空格隔开，输入E结束路径配置");
                while (sourcePath != "E")
                {
                    sourcePath = Console.ReadLine();
                    var currentAddedFiles = new Dictionary<string, string>();
                    // 读取文件路径
                    if (sourcePath != "E")
                    {
                        var pathList = new List<string>(sourcePath?
                            .Split(' ', ';', '，', '；') ?? Array.Empty<string>());

                        foreach (var path in pathList)
                        {
                            if (lastBackupFileNameDic != null && fileNameAndPathDic != null && File.Exists(path) && !fileNameAndPathDic.ContainsKey(Path.GetFileName(path))
                                && !lastBackupFileNameDic.ContainsKey(Path.GetFileName(path)))
                            {
                                var loadedFile = new MyFile(path);
                                fileNameAndPathDic.Add(Path.GetFileName(path), path);
                                currentAddedFiles.Add(Path.GetFileName(path), path);
                                loadedFile.CreativeBackupDirectory().BackupFile(lastBackupFileNameDic, fileNameAndPathDic, backupFileNameAndPath, path);
                            }
                            else if (fileNameAndPathDic != null && fileNameAndPathDic.ContainsKey(Path.GetFileName(path)))
                            {
                                Console.WriteLine($"\"{path}\"文件已被添加，请勿重复添加");
                            }
                            else if (!File.Exists(path))
                            {
                                Console.WriteLine($"--{path}---所指向的文件不存在，请重新输入");
                            }
                        }
                    }
                    if (currentAddedFiles.Count != 0)
                    {
                        Console.WriteLine("你添加了");
                        foreach (var file in currentAddedFiles)
                        {
                            Console.WriteLine(file.Value);
                        }

                        if (fileNameAndPathDic != null)
                        {
                            foreach (var files in fileNameAndPathDic)
                            {
                                Console.WriteLine($"------fileNameAndPathDic：{files.Key} {files.Value}");
                            }
                        }
                        WriteSourceConfig(fileNameAndPathDic);
                        WriteLastBackupConfig(lastBackupFileNameDic);
                        WriteBackupFileConfig(backupFileNameAndPath);
                    }
                }
            
            }
        }

        public static void WriteConfig(string key, string value)
        {
            string json = JsonConvert.SerializeObject
                (new Dictionary<string, string>() { [key] = value });
            File.WriteAllText("config.json", json);
        }
        public static void WriteLastBackupConfig(Dictionary<string, string>? dataDic)
        {
            // 读取和解析原有的 JSON 数据
            Dictionary<string, Dictionary<string, string>> config =
                JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>
                    (File.ReadAllText("config.json")) ?? new Dictionary<string, Dictionary<string, string>>();
            // 在数据结构中添加新的数据
            if (config.TryGetValue("LastBackupFile", out var value))
            {
                foreach (var data in value)
                {
                    if (!config["LastBackupFile"].ContainsKey(data.Key))
                    {
                        config["LastBackupFile"].Add(data.Key, data.Value);
                    }
                }
            }
            else
            {
                if (dataDic != null) config.TryAdd("LastBackupFile", dataDic);
            }
            // 将整个数据结构序列化回 JSON 格式
            string json = JsonConvert.SerializeObject(config);
            // 写入文件
            File.WriteAllText("config.json", json);
        }
        public static void WriteSourceConfig(Dictionary<string, string>? dataDic)
        {
            // 读取和解析原有的 JSON 数据
            Dictionary<string, Dictionary<string, string>> config =
                JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>
                    (File.ReadAllText("config.json")) ?? new Dictionary<string, Dictionary<string, string>>();
            // 在数据结构中添加新的数据
            if (config.TryGetValue("FilenameAndPath", out var value))
            {
                foreach (var data in value)
                {
                    if (config["FilenameAndPath"].ContainsKey(data.Key))
                    {
                        config["FilenameAndPath"].Add(data.Key, data.Value);
                    }
                }
            }
            else
            {
                if (dataDic != null) config.TryAdd("FilenameAndPath", dataDic);
            }
            // 将整个数据结构序列化回 JSON 格式
            string json = JsonConvert.SerializeObject(config);
            // 写入文件
            File.WriteAllText("config.json", json);
        }

        public static void WriteBackupFileConfig(Dictionary<string, string>? dataDic)
        {

            // 读取和解析原有的 JSON 数据
            Dictionary<string, Dictionary<string, string>> config =
                JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>
                    (File.ReadAllText("config.json")) ?? new Dictionary<string, Dictionary<string, string>>();
            // 在数据结构中添加新的数据
            if (config.TryGetValue("BackupFileNameAndPath", out var value))
            {
                foreach (var data in value)
                {
                    if (config["BackupFileNameAndPath"].ContainsKey(data.Key))
                    {
                        config["BackupFileNameAndPath"].Add(data.Key, data.Value);
                    }
                }
            }
            else
            {
                if (dataDic != null) config.TryAdd("BackupFileNameAndPath", dataDic);
            }
            // 将整个数据结构序列化回 JSON 格式
            string json = JsonConvert.SerializeObject(config);
            // 写入文件
            File.WriteAllText("config.json", json);
        }
    }
}
