public class persionConfig : IDataRow
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
	/// 姓名
	/// <summary>
	public string Name
	{
		get;
		private set;
	}

	/// <summary>
	/// 性别
	/// <summary>
	public int Sex
	{
		get;
		private set;
	}

	/// <summary>
	/// 年龄
	/// <summary>
	public int Age
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
		Sex = int.Parse(text[++index]);
		Age = int.Parse(text[++index]);
	}
}
