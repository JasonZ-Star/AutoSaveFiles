using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using static System.Net.Mime.MediaTypeNames;


namespace GetOfficeHook
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            while (true)
            {
                List<string> allWindowTiltes = new List<string>();
                List<string> officeTitles = new List<string>();
                foreach (var title in OfficeWindowHook.OfficeList)
                {
                    List<string> windowsTitles = OfficeWindowHook.GetOfficeWindowsTitles(title);
                    foreach (var windowsTitle in windowsTitles)
                    {
                        Console.WriteLine($"windowsTitle： {windowsTitle}");
                        allWindowTiltes.Add(windowsTitle);
                    }
                }
                string targetString = "  -  兼容性模式";
                
                List<string> fileSearchResult = [];
                if (fileSearchResult.Count > 0)
                {
                    Console.WriteLine("Results are as follows:");
                    foreach (var title in allWindowTiltes)
                    {
                        Console.WriteLine(title);
                        string resultString = title;
                        if (!title.EndsWith('x') && !title.EndsWith('X'))
                        {
                            resultString = resultString.Replace(targetString, "", StringComparison.OrdinalIgnoreCase);
                        }
                        List<string> fileSearchResultPart = OfficeWindowHook.SearchFile(resultString);
                        foreach (var fileNameAndPath in fileSearchResultPart)
                        {
                            fileSearchResult.Add(fileNameAndPath);
                        }
                    }
                    foreach (var file in fileSearchResult)
                    {
                        Console.WriteLine("{0}\n", file);
                    }
                }
                else
                {
                    Console.WriteLine("There is no any office software running.");
                }
                

                Thread.Sleep(5000);
                Console.Clear();
            }
        }


        
    }
}

