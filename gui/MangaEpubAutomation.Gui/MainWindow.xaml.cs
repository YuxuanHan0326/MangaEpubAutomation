using System.Windows;
using MangaEpubAutomation.Gui.ViewModels;
using MangaEpubAutomation.Gui.Views.Pages;
using Wpf.Ui.Appearance;
using Wpf.Ui.Controls;

namespace MangaEpubAutomation.Gui;

public partial class MainWindow : Window
{
    private bool _isFollowingSystemTheme;

    public MainViewModel ViewModel { get; }

    public MainWindow()
    {
        InitializeComponent();
        ViewModel = new MainViewModel();
        DataContext = ViewModel;

        SetThemeFollowSystem();
        Loaded += OnLoaded;
        Closed += OnClosed;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        RootNavigation.Navigate(typeof(RunPage));
    }

    private void OnClosed(object? sender, EventArgs e)
    {
        if (_isFollowingSystemTheme)
        {
            SystemThemeWatcher.UnWatch(this);
            _isFollowingSystemTheme = false;
        }
    }

    private void OnLightThemeClick(object sender, RoutedEventArgs e) => SetTheme(ApplicationTheme.Light);

    private void OnDarkThemeClick(object sender, RoutedEventArgs e) => SetTheme(ApplicationTheme.Dark);

    private void OnFollowSystemThemeClick(object sender, RoutedEventArgs e) => SetThemeFollowSystem();

    private void SetThemeFollowSystem()
    {
        if (_isFollowingSystemTheme)
        {
            SystemThemeWatcher.UnWatch(this);
        }

        SystemThemeWatcher.Watch(this, WindowBackdropType.None, updateAccents: true);
        _isFollowingSystemTheme = true;
        RootNavigation.UpdateLayout();
        UpdateLayout();
    }

    private void SetTheme(ApplicationTheme theme)
    {
        if (_isFollowingSystemTheme)
        {
            SystemThemeWatcher.UnWatch(this);
            _isFollowingSystemTheme = false;
        }

        ApplicationThemeManager.Apply(theme, WindowBackdropType.None, updateAccent: true);
        RootNavigation.UpdateLayout();
        UpdateLayout();
    }
}
