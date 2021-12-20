using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.Models {
	public enum MaterialType {

		Unknown,

		SolidBirch,

		EconomyBirch,

		HybridBirch,

		Walnut,

		UnFinishedWalnut,

		WhiteOak,

		UnFinishedWhiteOak,

		Plywood1_2,

		Plywood1_4,

		Plywood3_8,

		WhiteMela1_2,
		
		WhiteMela1_4,
		
		BlackMela1_2,

		BlackMela1_4

	}

	public class MaterialFunctions {

		public static string TypeToString(MaterialType matType) {
			switch (matType) {
				case MaterialType.SolidBirch:
					return "solid_birch";
				case MaterialType.EconomyBirch:
					return "economy_birch";
				case MaterialType.HybridBirch:
					return "hybrid_birch";
				case MaterialType.Walnut:
					return "walnut";
				case MaterialType.UnFinishedWalnut:
					return "walnut_unfinished";
				case MaterialType.WhiteOak:
					return "white_oak";
				case MaterialType.UnFinishedWhiteOak:
					return "white_oak_unfinished";
				case MaterialType.Plywood1_2:
					return "plywood_1_2";
				case MaterialType.Plywood1_4:
					return "plywood_1_4";
				case MaterialType.Plywood3_8:
					return "plywood_3_8";
				case MaterialType.WhiteMela1_2:
					return "whitemela_1_2";
				case MaterialType.WhiteMela1_4:
					return "whitemela_1_4";
				case MaterialType.BlackMela1_2:
					return "blackmela_1_2";
				case MaterialType.BlackMela1_4:
					return "blackmela_1_4";
				default:
					return "Unknown";
			}
		}

		public static MaterialType StringToType(string matType) {
			switch (matType) {
				case "solid_birch":
					return MaterialType.SolidBirch;
				case "economy_birch":
					return MaterialType.EconomyBirch;
				case "hybrid_birch":
					return MaterialType.HybridBirch;
				case "walnut":
					return MaterialType.Walnut;
				case "walnut_unfinished":
					return MaterialType.UnFinishedWalnut;
				case "white_oak":
					return MaterialType.WhiteOak;
				case "white_oak_unfinished":
					return MaterialType.UnFinishedWhiteOak;
				case "plywood_1_2":
					return MaterialType.Plywood1_2;
				case "plywood_1_4":
					return MaterialType.Plywood1_4;
				case "plywood_3_8":
					return MaterialType.Plywood3_8;
				case "whitemela_1_2":
					return MaterialType.WhiteMela1_2;
				case "whitemela_1_4":
					return MaterialType.WhiteMela1_4;
				case "blackmela_1_2":
					return MaterialType.BlackMela1_2;
				case "blackmela_1_4":
					return MaterialType.BlackMela1_4;
				default:
					return MaterialType.Unknown;
			}
		}


	}

}
