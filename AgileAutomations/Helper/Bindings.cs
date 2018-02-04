using Ninject.Modules;
using AgileAutomations.Helper;
using AgileAutomations.Interface;
using AgileAutomations.ViewModel;

namespace AgileAutomations.Helper
{
    public class Bindings : NinjectModule
    {
        public override void Load()
        {
            Bind<IMainWindowViewModel>().To<MainWindowViewModel>();
            Bind<IExcelParser>().To<ExcelParser>();
            Bind<IFileBrowser>().To<FileBrowser>();
            Bind<IHtmlHelper>().To<HtmlHelper>();
        }
    }
}