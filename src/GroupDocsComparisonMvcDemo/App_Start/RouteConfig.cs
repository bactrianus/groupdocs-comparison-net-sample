using System.Web.Mvc;
using System.Web.Routing;
using Groupdocs.Web.UI;
using GroupDocs.Web.UI.Comparison;

namespace GroupDocsComparisonMvcDemo
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes, ComparisonWidgetSettings settings)
        {
            //Add route "{resource}.axd/{*pathInfo}" to ignore
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");
            //Init Viewer routes
            Viewer.InitRoutes();
            //Init Comparison routes
            RouteConfigurator.Configure(settings);
            
            //Add default route for Home controller
            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}