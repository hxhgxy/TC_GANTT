using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Worksheet.Utility
{
	public static class Utility
	{

		internal static readonly Regex CellReferenceRegex = new Regex(
			"^\\s*(?:(?<abs_col>\\$)?(?<col>[A-Z]+)(?<abs_row>\\$)?(?<row>[0-9]+))\\s*$",
			RegexOptions.Compiled | RegexOptions.Singleline);
		internal static readonly Regex RangeReferenceRegex = new Regex(
			"^\\s*(?:(?<abs_from_col>\\$)?(?<from_col>[A-Z]+)(?<abs_from_row>\\$)?(?<from_row>[0-9]+):(?<abs_to_col>\\$)?(?<to_col>[A-Z]+)(?<abs_to_row>\\$)?(?<to_row>[0-9]+))\\s*$",
			RegexOptions.Compiled | RegexOptions.Singleline);
		internal static readonly Regex SingleAbsoulteRangeRegex = new Regex(
			"^\\s*(" +
				"((?<abs_from_row>\\$)?(?<from_row>[0-9]+):(?<abs_to_row>\\$)?(?<to_row>[0-9]+))|" +
				"((?<abs_from_col>\\$)?(?<from_col>[A-Z]+):(?<abs_to_col>\\$)?(?<to_col>[A-Z]+)))\\s*$",
			RegexOptions.Compiled | RegexOptions.Singleline);
		internal static readonly Regex NameRegex = new Regex("^\\s*[A-Za-z0-9_$]+\\s*$",
			RegexOptions.Compiled | RegexOptions.Singleline);
		private const string AlphaChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		private const int AlphaCharsLength = 26;
		public static string GetAlphaChar(long a)
		{
			char[] v = new char[10];
			int i = 9;
			while (a >= AlphaCharsLength)
			{
				v[i] = AlphaChars[((int)(a % AlphaCharsLength))];
				a = a / AlphaCharsLength - 1;
				i--;
			}
			v[i] = AlphaChars[((int)(a % AlphaCharsLength))];
			return new string(v, i, 10 - i);
		}

		public static int GetNumberOfChar(string address)
		{
			if (string.IsNullOrEmpty(address) || address.Length < 1 || address.Any(c => c < 'A' || c > 'Z'))
			{
				throw new ArgumentException("cannot convert into number of index from empty address", "id");
			}
			else
			{
				int idx = address[0] - AlphaChars[0] + 1;
				for (int i = 1; i < address.Length; i++)
				{
					idx *= AlphaCharsLength;
					idx += address[i] - AlphaChars[0] + 1;
				}
				return idx - 1;
			}
		}

		public static string ToAddress(int row, int col)
		{
			char columnLetter = (char)('A' + col);
			return $"{columnLetter}{row + 1}";
		}

		public static CellPosition FromAddress(string address)
		{
			int col = address[0] - 'A';
			int row = int.Parse(address.Substring(1)) - 1;
			return new CellPosition(row, col);
		}
	}
}
