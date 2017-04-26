using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Markup;

namespace test
{
    class EnvironmentVarExtension : MarkupExtension
    {
        private string _variableName;

        public EnvironmentVarExtension(string variableName)
        {
            _variableName = variableName;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return Environment.GetEnvironmentVariable(VariableName);
        }

        [ConstructorArgument("variableName")]
        public string VariableName
        {
            get { return _variableName; }
            set { _variableName = value; }
        }
    }
}
