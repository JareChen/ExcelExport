public class classConfig : IDataRow
{
	/// <summary>
	/// id
	/// <summary>
	public int Id
	{
		get;
		private set;
	}

	/// <summary>
	/// 名称
	/// <summary>
	public string Name
	{
		get;
		private set;
	}

	/// <summary>
	/// 级别
	/// <summary>
	public int Level
	{
		get;
		private set;
	}

	public void ParseDataRow(string dataRowText)
	{
		string[] text = DataTableExtension.SplitDataRow(dataRowText);
		int index = 1;
		Id = int.Parse(text[++index]);
		Name = text[++index];
		Level = int.Parse(text[++index]);
	}
}
