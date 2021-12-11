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

		Label CreateLabel();

		void AddLabel(Label label, int qty);

	}

	public class Label : IReadOnlyCollection<object> {
		public Dictionary<string, LabelField> LabelFields { get; set; }
		public int Count => LabelFields.Count;

		public object this[string fieldName] {
			get => LabelFields[fieldName].Value;
			set => LabelFields[fieldName].Value = value;
		}

		public IEnumerator<object> GetEnumerator() {
			return LabelFields.Values.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return LabelFields.Keys.GetEnumerator();
		}

	}

}