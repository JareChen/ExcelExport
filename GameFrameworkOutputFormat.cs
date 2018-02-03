using System.IO;
using System.Text;
using OfficeOpenXml;

public class GameFrameworkOutputFormat : OutputFormat
{
    /// <summary>
    /// 输出的配置文件
    /// </summary>
    private string outputCSStr;

    /// <summary>
    /// 输出的配置脚本
    /// </summary>
    private string outputTxtStr;

    /// <summary>
    /// 生成文件的格式，包含配置文件和脚本文件
    /// </summary>
    /// <param name="workSheetItem"></param>
    protected override void Generate(ExcelWorksheet workSheetItem)
    {
        fileNameWithoutExt = workSheetItem.Name + "Config";
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
        outputCS.Append(parseDataRowMethodFormat.ToString());
        outputTxtStr = outputTxt.ToString(); 
        outputCSStr = string.Format(csClassFormat.ToString(), fileNameWithoutExt, outputCS);

        string txtFileName = fileNameWithoutExt + ".txt";
        OutputFile(txtFileName, outputTxtStr);

        string csFileName = fileNameWithoutExt + ".cs";
        OutputFile(csFileName, outputCSStr);
    }
}
