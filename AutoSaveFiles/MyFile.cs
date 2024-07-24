using System.Globalization;

namespace AutoSaveFiles;

public class MyFile(string path)
{
    private string _lastModifiedTime = GetFileLastModifiedTime(path).ToString(CultureInfo.InvariantCulture);

    public string FileName { get; } = Path.GetFileName(path);

    public string LastModifiedTime
    {
        get => _lastModifiedTime;
        set => _lastModifiedTime = GetFileLastModifiedTime(value).ToString(CultureInfo.InvariantCulture);
    }

    // 创建配置文件
    public static void CreateDirectory()
    {
        if (!Directory.Exists("./data"))
        {
            Directory.CreateDirectory("./data");
            File.Create("./data/fileList.json").Close();
            ConfigOperator.WriteConfig("dataPath","./data");
        }
        if (!Directory.Exists("./source"))
        {
            Directory.CreateDirectory("./source");
            ConfigOperator.WriteConfig("sourcePath","./source");
        }
    }
    // 创建备份文件夹
    public MyFile CreativeBackupDirectory()
    {
        if (!Directory.Exists($"{Path.GetDirectoryName(path)}\\backup"))
        {
            Directory.CreateDirectory($"{Path.GetDirectoryName(path)}\\backup");
        }
        return this;
    }
    public bool IsFileExist(string path)
    {
        return File.Exists(path);
    }
    // 获得备份文件路径及文件名
    public string GetBackupFileNameAndPath(string sourcePath)
    {
        return Path.GetDirectoryName(sourcePath) + @"\backup\" + Path.GetFileNameWithoutExtension(sourcePath) + "-Ver." + DateTime.Now.ToString("yyyy-MM-dd~HH-mm") + Path.GetExtension(sourcePath);
    }
    // 获取文件最后修改时间
    public static string GetFileLastModifiedTime(string path)
    {
        return new DateTimeOffset(File.GetLastWriteTime(path)).ToUnixTimeSeconds().ToString();
    }

    // 检查文件是否被修改
    public bool IsFileModified(string path)
    {
        if (IsFileExist(GetBackupFileNameAndPath(path)) == false || LastModifiedTime ==
            GetFileLastModifiedTime(GetBackupFileNameAndPath(path)))
        {
            return false;
        }
        return true;
    }

    // 备份文件
    public void BackupFile(Dictionary<string, string>? lastBackupFileNameDictionary,
        Dictionary<string, string>? fileNameAndPathDic,
        Dictionary<string, string>? backupFileNameAndPathDic, 
        string sourcePath)
    {
        string fileName = Path.GetFileName(sourcePath);
        // 如果文件被修改或者备份文件不存在，则备份文件
        if ((lastBackupFileNameDictionary != null && backupFileNameAndPathDic != null
            && !lastBackupFileNameDictionary.ContainsKey(fileName)) 
            ||  (IsFileModified(sourcePath) ))
        {
            var backupFilePath = GetBackupFileNameAndPath(sourcePath);
            var backupFileName = Path.GetFileName(backupFilePath);
            File.Copy(sourcePath, backupFilePath, overwrite:true);
            if (backupFileNameAndPathDic != null) backupFileNameAndPathDic[backupFileName] = backupFilePath;
            if (lastBackupFileNameDictionary != null) lastBackupFileNameDictionary[FileName] = backupFileName;
        }
    }

}