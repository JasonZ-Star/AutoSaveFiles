using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace GetOfficeHook
{
    internal class OfficeWindowHook
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr GetClassName(IntPtr hWnd);
        // 导入user32.dll中的GetWindowText函数，该函数用于获取窗口的标题文本
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        private static extern int GetWindowTextLength(IntPtr hWnd);

        // 导入user32.dll中的FindWindow函数，该函数用于查找窗口句柄
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // Word窗口的类名，用于FindWindow函数识别Word的窗口
        private const string WordWindowClass = "OpusApp";

        private const int EVERYTHING_OK = 0;
        private const int EVERYTHING_ERROR_MEMORY = 1;
        private const int EVERYTHING_ERROR_IPC = 2;
        private const int EVERYTHING_ERROR_REGISTERCLASSEX = 3;
        private const int EVERYTHING_ERROR_CREATEWINDOW = 4;
        private const int EVERYTHING_ERROR_CREATETHREAD = 5;
        private const int EVERYTHING_ERROR_INVALIDINDEX = 6;
        private const int EVERYTHING_ERROR_INVALIDCALL = 7;

        private const int EVERYTHING_REQUEST_FILE_NAME = 0x00000001;
        private const int EVERYTHING_REQUEST_PATH = 0x00000002;
        private const int EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME = 0x00000004;
        private const int EVERYTHING_REQUEST_EXTENSION = 0x00000008;
        private const int EVERYTHING_REQUEST_SIZE = 0x00000010;
        private const int EVERYTHING_REQUEST_DATE_CREATED = 0x00000020;
        private const int EVERYTHING_REQUEST_DATE_MODIFIED = 0x00000040;
        private const int EVERYTHING_REQUEST_DATE_ACCESSED = 0x00000080;
        private const int EVERYTHING_REQUEST_ATTRIBUTES = 0x00000100;
        private const int EVERYTHING_REQUEST_FILE_LIST_FILE_NAME = 0x00000200;
        private const int EVERYTHING_REQUEST_RUN_COUNT = 0x00000400;
        private const int EVERYTHING_REQUEST_DATE_RUN = 0x00000800;
        private const int EVERYTHING_REQUEST_DATE_RECENTLY_CHANGED = 0x00001000;
        private const int EVERYTHING_REQUEST_HIGHLIGHTED_FILE_NAME = 0x00002000;
        private const int EVERYTHING_REQUEST_HIGHLIGHTED_PATH = 0x00004000;
        private const int EVERYTHING_REQUEST_HIGHLIGHTED_FULL_PATH_AND_FILE_NAME = 0x00008000;

        private const int EVERYTHING_SORT_NAME_ASCENDING = 1;
        private const int EVERYTHING_SORT_NAME_DESCENDING = 2;
        private const int EVERYTHING_SORT_PATH_ASCENDING = 3;
        private const int EVERYTHING_SORT_PATH_DESCENDING = 4;
        private const int EVERYTHING_SORT_SIZE_ASCENDING = 5;
        private const int EVERYTHING_SORT_SIZE_DESCENDING = 6;
        private const int EVERYTHING_SORT_EXTENSION_ASCENDING = 7;
        private const int EVERYTHING_SORT_EXTENSION_DESCENDING = 8;
        private const int EVERYTHING_SORT_TYPE_NAME_ASCENDING = 9;
        private const int EVERYTHING_SORT_TYPE_NAME_DESCENDING = 10;
        private const int EVERYTHING_SORT_DATE_CREATED_ASCENDING = 11;
        private const int EVERYTHING_SORT_DATE_CREATED_DESCENDING = 12;
        private const int EVERYTHING_SORT_DATE_MODIFIED_ASCENDING = 13;
        private const int EVERYTHING_SORT_DATE_MODIFIED_DESCENDING = 14;
        private const int EVERYTHING_SORT_ATTRIBUTES_ASCENDING = 15;
        private const int EVERYTHING_SORT_ATTRIBUTES_DESCENDING = 16;
        private const int EVERYTHING_SORT_FILE_LIST_FILENAME_ASCENDING = 17;
        private const int EVERYTHING_SORT_FILE_LIST_FILENAME_DESCENDING = 18;
        private const int EVERYTHING_SORT_RUN_COUNT_ASCENDING = 19;
        private const int EVERYTHING_SORT_RUN_COUNT_DESCENDING = 20;
        private const int EVERYTHING_SORT_DATE_RECENTLY_CHANGED_ASCENDING = 21;
        private const int EVERYTHING_SORT_DATE_RECENTLY_CHANGED_DESCENDING = 22;
        private const int EVERYTHING_SORT_DATE_ACCESSED_ASCENDING = 23;
        private const int EVERYTHING_SORT_DATE_ACCESSED_DESCENDING = 24;
        private const int EVERYTHING_SORT_DATE_RUN_ASCENDING = 25;
        private const int EVERYTHING_SORT_DATE_RUN_DESCENDING = 26;

        private const int EVERYTHING_TARGET_MACHINE_X86 = 1;
        private const int EVERYTHING_TARGET_MACHINE_X64 = 2;
        private const int EVERYTHING_TARGET_MACHINE_ARM = 3;



        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern UInt32 Everything_SetSearchW(string lpSearchString);
        public static extern UInt32 Everything_SetSearch(string lpSearchString);

        [DllImport("Everything64.dll")]
        private static extern void Everything_SetMatchPath(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMatchCase(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMatchWholeWord(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetRegex(bool bEnable);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetMax(UInt32 dwMax);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetOffset(UInt32 dwOffset);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetMatchPath();

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetMatchCase();

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetMatchWholeWord();

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetRegex();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetMax();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetOffset();

        [DllImport("Everything64.dll")]
        public static extern IntPtr Everything_GetSearchW();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetLastError();

        [DllImport("Everything64.dll")]
        public static extern bool Everything_QueryW(bool bWait);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SortResultsByPath();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetNumFileResults();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetNumFolderResults();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetNumResults();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetTotFileResults();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetTotFolderResults();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetTotResults();

        [DllImport("Everything64.dll")]
        public static extern bool Everything_IsVolumeResult(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_IsFolderResult(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_IsFileResult(UInt32 nIndex);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern void Everything_GetResultFullPathName(UInt32 nIndex, StringBuilder lpString,
            UInt32 nMaxCount);

        [DllImport("Everything64.dll")]
        public static extern void Everything_Reset();

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultFileName(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetSort(UInt32 dwSortType);

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetSort();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetResultListSort();

        [DllImport("Everything64.dll")]
        public static extern void Everything_SetRequestFlags(UInt32 dwRequestFlags);

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetRequestFlags();

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetResultListRequestFlags();

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultExtension(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultSize(UInt32 nIndex, out long lpFileSize);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultDateCreated(UInt32 nIndex, out long lpFileTime);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultDateModified(UInt32 nIndex, out long lpFileTime);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultDateAccessed(UInt32 nIndex, out long lpFileTime);

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetResultAttributes(UInt32 nIndex);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultFileListFileName(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetResultRunCount(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultDateRun(UInt32 nIndex, out long lpFileTime);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_GetResultDateRecentlyChanged(UInt32 nIndex, out long lpFileTime);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultHighlightedFileName(UInt32 nIndex);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultHighlightedPath(UInt32 nIndex);

        [DllImport("Everything64.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr Everything_GetResultHighlightedFullPathAndFileName(UInt32 nIndex);

        [DllImport("Everything64.dll")]
        public static extern UInt32 Everything_GetRunCountFromFileName(string lpFileName);
        [DllImport("Everything64.dll")]
        public static extern StringBuilder Everything_GetResultPath(UInt32 index);

        [DllImport("Everything64.dll")]
        public static extern bool Everything_SetRunCountFromFileName(string lpFileName, UInt32 dwRunCount);
        public const int MaxPath = 260;
        public static List<string> OfficeList = new List<string> { "WINWORD", "EXCEL", "POWERPNT" };

        private static string ExtractDocumentTitle(string windowTitle)
        {
            string pattern = @"^(.+?)( - Word$| - Excel$| - PowerPoint$)";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);
            Match match = regex.Match(windowTitle);
            if (match.Success)
            {
                // 返回捕获的文件名部分
                return match.Groups[1].Value;
            }
            return null;
        }

        internal static List<string> GetOfficeWindowsTitles(string officeProcessName)
        {
            List<string> documentTitles = new List<string>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                StringBuilder windowText = new StringBuilder(256);
                if (GetWindowText(hWnd, windowText, 256) > 0)
                {

                    // 获取窗口的进程ID
                    GetWindowThreadProcessId(hWnd, out uint processId);

                    // 获取进程名称
                    Process process = Process.GetProcessById((int)processId);
                    if (process.ProcessName == officeProcessName)
                    {
                        // 从窗口标题中提取文件名
                        string title = windowText.ToString();
                        string documentTitle = ExtractDocumentTitle(title);
                        if (!string.IsNullOrEmpty(documentTitle))
                        {
                            documentTitles.Add(documentTitle);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            return documentTitles;
        }

        internal static List<string> SearchFile(string fileName)
        {
            List<string> fileSearchResult = [];
            // set the search
            Everything_SetSearchW($"\"{fileName}\" !*.lnk");

            // use our own custom scrollbar... 			
            // Everything_SetMax(listBox1.ClientRectangle.Height / listBox1.ItemHeight);
            // Everything_SetOffset(VerticalScrollBarPosition...);

            // request name and size
            Everything_SetRequestFlags(EVERYTHING_REQUEST_FILE_NAME
                                       | EVERYTHING_REQUEST_PATH
                                       | EVERYTHING_REQUEST_DATE_MODIFIED
                                       | EVERYTHING_REQUEST_SIZE
                                       | EVERYTHING_REQUEST_FULL_PATH_AND_FILE_NAME);

            Everything_SetSort(13);

            // execute the query
            Everything_QueryW(true);

            // sort by path
            // Everything_SortResultsByPath();


            // total number of results
            // string text = $"\n{Everything_GetNumResults()} Results" ;
            // Console.WriteLine(text);

            for (UInt32 i = 0; i < Everything_GetNumResults(); i++)
            {
                Everything_GetResultDateModified(i, out long dateModified);
                Everything_GetResultSize(i, out long size);
                StringBuilder fullPath = new StringBuilder(MaxPath);
                Everything_GetResultFullPathName(i, fullPath, (uint)fullPath.Capacity);
                fileSearchResult.Add(fullPath.ToString());
                // Print all information				
                // Console.WriteLine(
                //     $"第{i+1}个 Size: {size} , Date Modified: {DateTime.FromFileTime(dateModified).Year} / " +
                //     $" {DateTime.FromFileTime(dateModified).Month} / {DateTime.FromFileTime(dateModified).Day} " +
                //     $"{DateTime.FromFileTime(dateModified).Hour}:{DateTime.FromFileTime(dateModified).Minute:D2}, " +
                //     $"Filename: {Marshal.PtrToStringUni(Everything_GetResultFileName(i))}, " +
                //     $"FilePath: {fullPath}");
            }
            return fileSearchResult;
        }
        private static bool IsUserAdministrator()
        {
            bool isAdmin;
            try
            {
                WindowsIdentity user = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(user);
                isAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch (UnauthorizedAccessException)
            {
                isAdmin = false;
            }
            catch (Exception)
            {
                isAdmin = false;
            }

            return isAdmin;
        }

        // 以管理员权限运行程序
        private static void RunAsAdministrator()
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo();
            processStartInfo.UseShellExecute = true;
            processStartInfo.WorkingDirectory = Environment.CurrentDirectory;
            processStartInfo.FileName = Environment.ProcessPath;
            processStartInfo.Verb = "runas";
            try
            {
                Process.Start(processStartInfo);
            }
            catch (Exception)
            {
                Console.WriteLine("无法以管理员权限运行程序");
            }

            // 提示用户按任意键退出程序
            Console.WriteLine("按任意键退出...");
            // 等待用户按键
            Console.ReadKey();
        }
    }


}
