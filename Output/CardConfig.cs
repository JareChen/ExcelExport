public class CardConfig : Config
{
	/// <summary>
	/// 资源名
	/// <summary>
	public string ResName
	{
		get;
		private set;
	}

	/// <summary>
	/// 资源名(大)
	/// <summary>
	public string ResNameLarge
	{
		get;
		private set;
	}

	/// <summary>
	/// 颜色
	/// <summary>
	public int Color
	{
		get;
		private set;
	}

	/// <summary>
	/// 类型(数字牌|功能牌|万能牌|牌背)
	/// <summary>
	public int Type
	{
		get;
		private set;
	}

	/// <summary>
	/// 牌堆
	/// <summary>
	public int Modulus
	{
		get;
		private set;
	}

	public override void ParseDataRow(string dataRowText)
	{
		string[] text = dataRowText.Split('\t');
		int index = 0;
		Id = text[index++];
		ResName = text[index++];
		ResNameLarge = text[index++];
		Color = int.Parse(text[index++]);
		Type = int.Parse(text[index++]);
		Modulus = int.Parse(text[index++]);
	}
}
