using System.Reflection;
using System.Windows;
using AgileAutomations.Helper;
using AgileAutomations.Interface;
using AgileAutomations.ViewModel;
using Ninject;

namespace AgileAutomations.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IMainWindowViewModel mainWindowViewModel;

        public MainWindow()
        {
            var kernel = new StandardKernel();
            kernel.Load(Assembly.GetExecutingAssembly());

            mainWindowViewModel = kernel.Get<IMainWindowViewModel>();
            DataContext = mainWindowViewModel;

            InitializeComponent();
        }
    }
}
    