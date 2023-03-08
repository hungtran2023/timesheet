using System;
using System.Collections.Specialized;
using System.Runtime.InteropServices;

namespace MSDN
{
	[Serializable]
	public class mySession 
	{

		private HybridDictionary dic = new HybridDictionary();

		public mySession()
		{
		}

		public string this [string name]
		{
			get
			{
				return (string)dic[name.ToLower()];
			}
			set
			{
				dic[name.ToLower()] = value;
			}
		}
	}
}
