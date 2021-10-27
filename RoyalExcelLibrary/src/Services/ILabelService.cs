using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Services {

	public enum LabelFieldType {
		Unknown,
		Text,
		Address
	}

	public class LabelField {
		public LabelFieldType Type { get; set; }
		public object Value { get; set; }
	}

	public interface ILabelService {
		void PrintLabels();
	}

}