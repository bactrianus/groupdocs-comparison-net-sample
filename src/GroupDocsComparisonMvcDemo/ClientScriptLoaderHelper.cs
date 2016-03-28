using System.Text;
using System.Web;
using GroupDocs.Web.UI.Comparison;

namespace GroupDocsComparisonMvcDemo
{
    /// <summary>
    /// Client Script Loader Helper
    /// </summary>
    public class ClientScriptLoaderHelper : GroupDocs.Web.UI.Comparison.Helpers.ClientScriptLoaderHelper, IHtmlString
    {
        private ClientHelper helper;

        internal ClientScriptLoaderHelper(ComparisonWidgetSettings settings, ClientHelper helper) : base(settings) 
        {
            //Set ClientHelper
            this.helper = helper;
        }
        
        public override string ToString()
        {
            var result = new StringBuilder();

            //Add Comparison scripts
            result.AppendLine(base.ToString());
            result.Append(helper.GenerateClientCode());

            return result.ToString();
        }

        string IHtmlString.ToHtmlString()
        {
            return ToString();
        }
    }
}
