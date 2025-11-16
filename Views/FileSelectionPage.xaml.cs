using ElectronicSpreadsheet.Services;

namespace ElectronicSpreadsheet.Views;

public partial class FileSelectionPage : ContentPage
{
    private TaskCompletionSource<GoogleDriveFile?> _tcs;

    public FileSelectionPage(List<GoogleDriveFile> files)
    {
        InitializeComponent();
        FilesCollection.ItemsSource = files;
        _tcs = new TaskCompletionSource<GoogleDriveFile?>();
    }

    public Task<GoogleDriveFile?> GetSelectedFileAsync()
    {
        return _tcs.Task;
    }

    private void OnFileSelected(object sender, SelectionChangedEventArgs e)
    {
        if (e.CurrentSelection.Count > 0)
        {
            var selectedFile = e.CurrentSelection[0] as GoogleDriveFile;
            _tcs.TrySetResult(selectedFile);
            Navigation.PopModalAsync();
        }
    }

    private void OnFileTapped(object sender, EventArgs e)
    {
        if (sender is Border border && border.BindingContext is GoogleDriveFile file)
        {
            _tcs.TrySetResult(file);
            Navigation.PopModalAsync();
        }
    }

    private void OnCancelClicked(object sender, EventArgs e)
    {
        _tcs.TrySetResult(null);
        Navigation.PopModalAsync();
    }
}
