using System;
using System.Collections.Generic;

namespace RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation {
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

}
