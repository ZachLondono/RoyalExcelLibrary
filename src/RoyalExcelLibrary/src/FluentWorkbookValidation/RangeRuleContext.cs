using ClosedXML.Excel;

namespace RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation {
    public class RangeRuleContext : AbstractRuleContext {

        private readonly IXLRange _target;
        private readonly string _rangeName;
        internal RangeRuleContext(IXLRange target, string rangeName, WkbkValidator validator) : base(validator) {
            _target = target;
            _rangeName = rangeName;
        }

        public RangeRuleContext NotEmpty() {

            bool passed;

            try {
                var val = _target.FirstCell().GetStringValue();
                if (string.IsNullOrEmpty(val)) passed = false;
                else passed = true;
            } catch {
                passed = false;
            }

            AddRule(new Rule {
                Passed = passed,
                DefaultMessage = $"Range '{_rangeName}' is empty"
            });

            return this;

        }

        public RangeRuleContext ContainsDouble() {

            bool passed;

            try {
                var val = _target.FirstCell().GetDoubleValue();
                passed = true;
            } catch {
                passed = false;
            }

            AddRule(new Rule {
                Passed = passed,
                DefaultMessage = $"Range '{_rangeName}' does not contain a double"
            });

            return this;

        }

    }

}
