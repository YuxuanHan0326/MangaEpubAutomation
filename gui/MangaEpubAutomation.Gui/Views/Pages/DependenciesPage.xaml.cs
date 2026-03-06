using System.Windows;
using System.Windows.Controls;

namespace MangaEpubAutomation.Gui.Views.Pages;

public partial class DependenciesPage : Page
{
    public DependenciesPage()
    {
        InitializeComponent();
        EnsureDataContext();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        EnsureDataContext();
    }

    private void EnsureDataContext()
    {
        if (DataContext is null && System.Windows.Application.Current.MainWindow is MainWindow mainWindow)
        {
            DataContext = mainWindow.ViewModel;
        }
    }
}
