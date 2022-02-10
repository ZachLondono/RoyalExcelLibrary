using ClosedXML.Excel;
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

}
