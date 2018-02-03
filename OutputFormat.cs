using System.IO;
using System.Text;
using OfficeOpenXml;

public abstract class OutputFormat
{
    /// <summary>
    /// 输出目录
    /// </summary>
    protected string outputPath;

    /// <summary>
    /// 文件名，不包含后缀
    /// </summary>
    protected string fileNameWithoutExt;
    protected abstract void Generate(ExcelWorksheet workSheetItem);

    /// <summary>
    /// 输出文件
    /// </summary>
    /// <param name="fileName">文件名</param>
    /// <param name="outputStr">输出目录</param>
    protected void OutputFile(string fileName, string outputStr)
    {
        string outputFilePath = Path.Combine(outputPath, fileName);
        if (!File.Exists(outputFilePath))
        {
            using (StreamWriter sw = File.CreateText(outputFilePath))
            {
                sw.Write(outputStr);
            }
        }
        else
        {
            File.WriteAllText(outputFilePath, outputStr);
        }
    }

    /// <summary>
    /// 生成文件，对外暴露的方法
    /// </summary>
    /// <param name="inputPath">输入目录</param>
    /// <param name="outputPath">输出目录</param>
    public void Output(string inputPath, string outputPath)
    {
        this.outputPath = outputPath;
        string[] filePaths = Directory.GetFiles(inputPath, "*.xlsx");
        foreach (var itemPath in filePaths)
        {
            FileInfo fi = new FileInfo(itemPath);
            ExcelPackage package = new ExcelPackage(fi);
            ExcelWorksheets sheets = package.Workbook.Worksheets;
            foreach (var workSheetItem in sheets)
            {
                Generate(workSheetItem);
            }
        }
    }
}
