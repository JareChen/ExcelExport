public class CardConfig : Config
{
	/// <summary>
	/// 牌数
	/// <summary>
	public int CardNumber;

	/// <summary>
	/// 资源名
	/// <summary>
	public string ResName;

	/// <summary>
	/// 资源名(大)
	/// <summary>
	public string ResNameLarge;

	/// <summary>
	/// 颜色
	/// <summary>
	public int Color;

	/// <summary>
	/// 类型(数字牌|功能牌|万能牌|牌背)
	/// <summary>
	public string Type;

	/// <summary>
	/// 牌堆
	/// <summary>
	public int Modulus;

	public override void ParseDataRow(string dataRowText)
	{
		string[] text = dataRowText.Split('\t');
		int index = 0;
		Id = text[index++];
		CardNumber = int.Parse(text[index++]);
		ResName = text[index++];
		ResNameLarge = text[index++];
		Color = int.Parse(text[index++]);
		Type = text[index++];
		Modulus = int.Parse(text[index++]);
	}
}
