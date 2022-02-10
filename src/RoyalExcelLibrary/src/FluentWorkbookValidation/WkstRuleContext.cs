using ClosedXML.Excel;

namespace RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation {
    public class WkstRuleContext : AbstractRuleContext {

        private readonly IXLWorksheet _target;
        internal WkstRuleContext(IXLWorksheet target, WkbkValidator validator) : base(validator) {
            _target = target;
        }

        public WkstRuleContext HasRange(string rangeName) {

            bool passed;

            try {
                var range = _target.Range(rangeName);
                if (range is null) passed = false;
                else passed = true;
            } catch {
                passed = false;
            }

            AddRule(new Rule {
                Passed = passed,
                DefaultMessage = $"Worksheet does not have range '{rangeName}'"
            });

            return this;

        }

        public RangeRuleContext ForRange(string rangeName) {
            var context = new RangeRuleContext(_target.Range(rangeName), rangeName, validator);
            validator.AddContext(context);
            return context;
        }

    }

}
