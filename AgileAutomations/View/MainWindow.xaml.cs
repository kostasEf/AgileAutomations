using System.Reflection;
using AgileAutomations.Interface;
using Ninject;

namespace AgileAutomations.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            var kernel = new StandardKernel();
            kernel.Load(Assembly.GetExecutingAssembly());

            var mainWindowViewModel = kernel.Get<IMainWindowViewModel>();
            DataContext = mainWindowViewModel;

            InitializeComponent();
        }
    }
}
    