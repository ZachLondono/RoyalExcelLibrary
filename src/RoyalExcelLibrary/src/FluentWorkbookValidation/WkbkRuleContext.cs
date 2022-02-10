using ClosedXML.Excel;

namespace RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation {
    public class WkbkRuleContext : AbstractRuleContext {

        private readonly XLWorkbook _target;
        internal WkbkRuleContext(XLWorkbook target, WkbkValidator validator) : base(validator) {
            _target = target;
        }

        public WkbkRuleContext HasSheet(string sheetName) {

            AddRule(new Rule {
                Passed = _target.Worksheets.Contains(sheetName),
                DefaultMessage = $"Worksheet does not exist '{sheetName}'"
            });

            return this;

        }

        public WkstRuleContext ForSheet(string sheetName) {
            var context =  new WkstRuleContext(_target.Worksheet(sheetName), validator);
            validator.AddContext(context);
            return context;
        }

        public RangeRuleContext ForRange(string rangeName) {
            var context = new RangeRuleContext(_target.Range(rangeName), rangeName, validator);
            validator.AddContext(context);
            return context;
        }

    }

}
