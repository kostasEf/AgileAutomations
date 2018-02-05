using AgileAutomations.Interface;
using AgileAutomations.ViewModel;
using Ninject.Modules;

namespace AgileAutomations.Helper
{
    public class Bindings : NinjectModule
    {
        public override void Load()
        {
            Bind<IMainWindowViewModel>().To<MainWindowViewModel>();
            Bind<IExcelHelper>().To<ExcelHelper>();
            Bind<IFileBrowser>().To<FileBrowser>();
            Bind<IHtmlHelper>().To<HtmlHelper>();
        }
    }
}