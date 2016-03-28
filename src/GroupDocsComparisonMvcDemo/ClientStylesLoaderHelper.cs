using System.Text;
using System.Web;
using GroupDocs.Web.UI.Comparison;

namespace GroupDocsComparisonMvcDemo
{
    /// <summary>
    /// Client Styles Loader Helper
    /// </summary>
    public class ClientStylesLoaderHelper : GroupDocs.Web.UI.Comparison.Helpers.ClientStylesLoaderHelper, IHtmlString
    {
        private readonly ComparisonWidgetSettings _settings;

        internal ClientStylesLoaderHelper(ComparisonWidgetSettings settings):base(settings)
        {
            //Set comparison settings
            _settings = settings;
        }

        /// <summary>
        /// Set true if need to use widget in ASP.NET WebForms project
        /// </summary>
        /// <param name="useHttpHandlers">True by default</param>
        /// <returns></returns>
        public override string ToString()
        {
            var result = new StringBuilder();
            //Add comparison css scripts
            result.Append(base.ToString());
            return result.ToString();
        }

        string IHtmlString.ToHtmlString()
        {
            return ToString();
        }
    }
}
