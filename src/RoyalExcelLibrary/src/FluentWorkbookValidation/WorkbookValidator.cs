using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation {

    public class WkbkValidator {

        private readonly XLWorkbook _target;
        private readonly List<AbstractRuleContext> _ruleContexts;

        public WkbkValidator(XLWorkbook target) {
            _target = target;
            _ruleContexts = new List<AbstractRuleContext>();
        }

        public WkbkRuleContext WkbkRule() {
            var context = new WkbkRuleContext(_target, this);
            _ruleContexts.Add(context);
            return context;
        }

        internal void AddContext(AbstractRuleContext context) {
            _ruleContexts.Add(context);
        }

        public void Validate() {

            foreach(var context in _ruleContexts) {
                context.Validate();
            }

        }

    }

    public class Rule {
        public bool Passed { get; set; }
        public string DefaultMessage { get; set; }
    }

    public class AbstractRuleContext {

        protected readonly WkbkValidator validator;

        private string _message = null;
        private readonly List<Rule> rules = new List<Rule>();

        public AbstractRuleContext(WkbkValidator validator) {
            this.validator = validator;
        }

        protected void AddRule(Rule rule) {
            rules.Add(rule);
        }

        public void WithMessage(string message) {
            _message = message;
        }

        internal void Validate() {

            foreach (var rule in rules) {

                if (rule.Passed) continue;

                if (_message is null) throw new Exception(rule.DefaultMessage);
                throw new Exception(_message);

            }
        }

    }

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

    }

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
