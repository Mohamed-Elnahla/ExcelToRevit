using System.Diagnostics;
using System.Reflection;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;

namespace Elnahla.ExcelToRevit.Views;

public partial class AboutView
{
    public AboutView()
    {
        InitializeComponent();
        UpdateContent();
        EventManager.RegisterClassHandler(typeof(Hyperlink), Hyperlink.RequestNavigateEvent,
            new RequestNavigateEventHandler(Hyperlink_OnClick));
    }

    private void UpdateContent()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var info = FileVersionInfo.GetVersionInfo(assembly.Location);
        VersionRun.Text = info.ProductVersion;
    }

    private void Window_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs eventArgs)
    {
        if (eventArgs.Source is Run) return;
        Close();
    }

    private static void Hyperlink_OnClick(object sender, RequestNavigateEventArgs eventArgs)
    {
        Process.Start(eventArgs.Uri.AbsoluteUri);
        eventArgs.Handled = true;
    }
}