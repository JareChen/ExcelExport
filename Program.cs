using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace ExcleExport
{
    public class Program
    {
        /// <summary>
        /// 配置文件
        /// </summary>
        static string configPath = "config.txt";

        /// <summary>
        /// 默认输入目录
        /// </summary>
        static string inputPath = "Input";

        /// <summary>
        /// 默认输出目录
        /// </summary>
        static string outputPath = "Output";
        
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            InitInputPathAndOutputPath();
            // 导出自定义格式的配置，修改 new 出的对象。
            OutputFormat outputFormat = new GameFrameworkOutputFormat();
            outputFormat.Output(inputPath, outputPath);
        }

        /// <summary>
        /// 初始化配置输入和输出目录，默认为应用程序统计目录的 Input 和 Output 文件夹
        /// </summary>
        static void InitInputPathAndOutputPath()
        {
            if (File.Exists(configPath))
            {
                StreamReader sr = File.OpenText(configPath);
                string line = sr.ReadLine();
                string[] inputStrArray = line.Split(':');
                inputPath = inputStrArray[1];
                line = sr.ReadLine();
                string[] outputStrArray = line.Split(':');
                outputPath = outputStrArray[1];
                line = sr.ReadLine();
            }
            if (!Directory.Exists(inputPath))
            {
                Directory.CreateDirectory(inputPath);
            }
            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }
            Directory.CreateDirectory(outputPath);
        }
    }
}
