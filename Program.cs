using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace ExcleExport
{
    public class Program
    {
        static string configPath = "config.txt";
        static string fileName = "abc.xlsx";
        static string intputPath = "Input";
        static string outputPath = "Output";
        
        static void Main(string[] args)
        {
            InitInputPathAndOutputPath();

            string filePath = Path.Combine(intputPath, fileName);
            FileInfo fi = new FileInfo(filePath);
            ExcelPackage package = new ExcelPackage(fi);
            ExcelWorksheets sheets = package.Workbook.Worksheets;
            foreach (var workSheetItem in sheets)
            {
                OutputFile(workSheetItem);
            }
        }

        /// <summary>
        /// 导出 Excel Sheet
        /// </summary>
        /// <param name="workSheetItem">单个worksheet</param>
        static void OutputFile(ExcelWorksheet workSheetItem)
        {
            int rowCount = workSheetItem.Dimension.End.Row;
            int columnCount = workSheetItem.Dimension.End.Column;
            StringBuilder outputTxt = new StringBuilder();
            StringBuilder outputCS = new StringBuilder();

            StringBuilder csClassFormat = new StringBuilder();
            csClassFormat.AppendLine("public class {0} : IDataRow");
            csClassFormat.AppendLine("{{");
            csClassFormat.AppendLine("{1}");
            csClassFormat.AppendLine("}}");
            
            StringBuilder csFieldFormat = new StringBuilder();
            csFieldFormat.AppendLine("\t/// <summary>");
            csFieldFormat.AppendLine("\t/// {0}");
            csFieldFormat.AppendLine("\t/// <summary>");
            csFieldFormat.AppendLine("\tpublic {1} {2}");
            csFieldFormat.AppendLine("\t{{");
            csFieldFormat.AppendLine("\t\tget;");
            csFieldFormat.AppendLine("\t\tprivate set;");
            csFieldFormat.AppendLine("\t}}");

            
            ExcelRange classInfo = workSheetItem.Cells[4, columnCount];
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= columnCount; j++)
                {
                    var cellValue = workSheetItem.GetValue(i, j);
                    string value = cellValue == null ? string.Empty : cellValue.ToString(); 
                    outputTxt.Append(value + "\t");

                    if (i == 4 && j > 1)
                    {
                        string expositionStr = classInfo[2, j].GetValue<string>();
                        string typeStr = classInfo[3, j].GetValue<string>();
                        string fieldNameStr = classInfo[4, j].GetValue<string>();
                        outputCS.AppendLine(string.Format(csFieldFormat.ToString(), expositionStr, typeStr, fieldNameStr));
                    }
                }
                outputTxt.AppendLine();
            }

            StringBuilder parseDataRowMethodFormat = new StringBuilder();
            parseDataRowMethodFormat.AppendLine("\tpublic void ParseDataRow(string dataRowText)");
            parseDataRowMethodFormat.AppendLine("\t{");
            parseDataRowMethodFormat.AppendLine("\t\tstring[] text = DataTableExtension.SplitDataRow(dataRowText);");
            parseDataRowMethodFormat.AppendLine("\t\tint index = 1;");
            for (int j = 2; j <= columnCount; j++)
            {
                string typeStr = classInfo[3, j].GetValue<string>();
                string fieldNameStr = classInfo[4, j].GetValue<string>();
                if (typeStr == "int")
                {
                    parseDataRowMethodFormat.AppendLine(string.Format("\t\t{0} = int.Parse(text[++index]);", fieldNameStr));
                }
                else
                {
                    parseDataRowMethodFormat.AppendLine(string.Format("\t\t{0} = text[++index];", fieldNameStr));
                }
            }
            parseDataRowMethodFormat.Append("\t}");

            string txtNameFormat = "{0}Config";
            string txtFileName = string.Format(txtNameFormat, workSheetItem.Name);
            OutoutFile(txtFileName + ".txt", outputTxt.ToString());

            string csNameFormat = "{0}Config";
            outputCS.Append(parseDataRowMethodFormat.ToString());
            string csFileName = string.Format(csNameFormat, workSheetItem.Name);
            string csFileStr = string.Format(csClassFormat.ToString(), csFileName, outputCS);
            OutoutFile(csFileName + ".cs", csFileStr.ToString());
        }
        static void OutoutFile(string fileName, string outputStr)
        {
            string outputFilePath = Path.Combine(outputPath, fileName);
            if (!File.Exists(outputFilePath))
            {
                using (StreamWriter sw = File.CreateText(outputFilePath))
                {
                    sw.Write(outputStr.ToString());
                }
            }
            else
            {
                File.WriteAllText(outputFilePath, outputStr.ToString());
            }
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
                intputPath = inputStrArray[1];
                line = sr.ReadLine();
                string[] outputStrArray = line.Split(':');
                outputPath = outputStrArray[1];
                line = sr.ReadLine();
                string[] fileNameStrArray = line.Split(':');
                fileName = fileNameStrArray[1];
            }

            if (!Directory.Exists(intputPath))
            {
                Directory.CreateDirectory(intputPath);
            }
            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }
        }
    }
}
