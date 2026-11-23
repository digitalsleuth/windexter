using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Windexter.Core;

namespace Windexter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private static readonly DispatcherTimer? elapsedTimer = new();
        private static readonly Stopwatch? stopWatch = new();
        private static readonly string displayName = "Windexter";
        private static readonly string githubBinaryRepo = "https://github.com/digitalsleuth/windexter";
        private static readonly Version? appVersion = new(Assembly.GetExecutingAssembly().GetName().Version!.ToString(3));
        private string? detectedGatherFile;
        private string? detectedPropMapFile;
        private string? detectedSecStoreFile;
        private int detectionVersion;

        public MainWindow()
        {
            InitializeComponent();
            DBPath.TextChanged += (sender, e) => DetectCompanionFiles();
            DataContext = this;
            this.Title = $"{displayName} v{appVersion}";
            elapsedTimer!.Tick += new EventHandler(ElapsedTime!);
            elapsedTimer.Interval = TimeSpan.FromSeconds(1);
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.AboutClick, (sender, e) => { AboutClick(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.AboutClick, new KeyGesture(Key.A, ModifierKeys.Control | ModifierKeys.Shift)));
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.DBPicker, (sender, e) => { DBPicker(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.DBPicker, new KeyGesture(Key.D, ModifierKeys.Control)));
            CommandBindings.Add(new CommandBinding(KeyboardShortcuts.OutputPicker, (sender, e) => { OutputPicker(sender, e); }, (sender, e) => { e.CanExecute = true; }));
            InputBindings.Add(new KeyBinding(KeyboardShortcuts.OutputPicker, new KeyGesture(Key.O, ModifierKeys.Control)));
        }

        public static class KeyboardShortcuts
        {
            static KeyboardShortcuts()
            {
                AboutClick = new RoutedCommand("AboutClick", typeof(MainWindow));
                DBPicker = new RoutedCommand("DBPicker", typeof(MainWindow));
                OutputPicker = new RoutedCommand("OutputPicker", typeof(MainWindow));
            }
            public static RoutedCommand AboutClick { get; private set; }
            public static RoutedCommand DBPicker { get; private set; }
            public static RoutedCommand OutputPicker { get; private set; }

        }

        private void ElapsedTime(object source, EventArgs e)
        {
            TimeSpan timeSpan = stopWatch!.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            TimerLabel.Content = elapsedTime;
        }

        private void OutputPicker(object sender, RoutedEventArgs e)
        {
            string initialDirectory = DBPath.Text != "" ? DBPath.Text : Environment.SpecialFolder.Desktop.ToString();
            string selectedPath = "";
            FolderBrowserDialog folderDlg = new()
            {
                Description = "Select the output directory for your processing results",
                ShowNewFolderButton = true,
                UseDescriptionForTitle = true,
                RootFolder = Environment.SpecialFolder.Desktop,
                InitialDirectory = initialDirectory
            };

            DialogResult result = folderDlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                selectedPath = folderDlg.SelectedPath;
            }
            if (selectedPath != "")
            {
                OutputPath.Text = selectedPath;
                OutputPath.ToolTip = selectedPath;
            }
        }

        private void DBPicker(object sender, RoutedEventArgs e)
        {
            try
            {
                Microsoft.Win32.OpenFileDialog openFile = new()
                {
                    Title = "Select a Windows.db or Windows.edb file",
                    Filter = "Windows Search Index Files|*.db;*.edb;"
                };
                if (openFile.ShowDialog() == true)
                {
                    DBPath.Text = openFile.FileName;
                    DBPath.ToolTip = openFile.FileName;
                    OutputPath.Text = Path.GetDirectoryName(openFile.FileName);
                    OutputPath.ToolTip = OutputPath.Text;
                }
            }
            catch (Exception ex)
            {
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Unable to load Database:\n\n{ex.Message}", "Database Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
                GoButton.IsEnabled = true;
            }
        }

        private async void DetectCompanionFiles()
        {
            int version = ++detectionVersion;
            string path = DBPath.Text.Trim().Trim('"');
            (string? gather, string? propMap, string? secStore) found = await Task.Run(() => (SearchIndex.FindGatherFile(path), SearchIndex.FindPropMapFile(path), SearchIndex.FindSecStoreFile(path)));
            if (version == detectionVersion)
            {
                detectedGatherFile = found.gather;
                detectedPropMapFile = found.propMap;
                detectedSecStoreFile = found.secStore;
                ApplyCompanionState();
            }
        }

        private void ApplyCompanionState()
        {
            GatherCheck.IsEnabled = detectedGatherFile is not null;
            GatherCheck.ToolTip = detectedGatherFile is null ? "No Windows-gather.db was found next to the selected Windows.db." : $"Also processes {detectedGatherFile}";
            if (!GatherCheck.IsEnabled)
            {
                GatherCheck.IsChecked = false;
            }
            PropMapCheck.IsEnabled = detectedPropMapFile is not null;
            PropMapCheck.ToolTip = detectedPropMapFile is null ? "No PropMap.db was found in or below the folder of the selected Windows.db." : $"Also processes {detectedPropMapFile}";
            if (!PropMapCheck.IsEnabled)
            {
                PropMapCheck.IsChecked = false;
            }
            SecStoreCheck.IsEnabled = detectedSecStoreFile is not null;
            SecStoreCheck.ToolTip = detectedSecStoreFile is null ? "No SecStore.db was found in or below the folder of the selected Windows.db." : $"Also processes {detectedSecStoreFile}";
            if (!SecStoreCheck.IsEnabled)
            {
                SecStoreCheck.IsChecked = false;
            }
        }

        private void SelectDeselectAll(object sender, RoutedEventArgs e)
        {
            bool isSelectAll = SelectDeselectButton.Content.ToString() == "Select All";

            foreach (object child in LogicalTreeHelper.GetChildren(MainGrid))
            {
                if (child is System.Windows.Controls.CheckBox { IsEnabled: true } cb)
                {
                    cb.IsChecked = isSelectAll;
                }
            }

            SelectDeselectButton.Content = isSelectAll ? "Deselect All" : "Select All";
        }

        private void ResetUI()
        {
            stopWatch?.Stop();
            elapsedTimer?.Stop();
            GoButton.IsEnabled = true;
            ProgressBarControl.IsIndeterminate = false;
            ProgressBarControl.Value = 100;
            ProgressBarControl.Visibility = Visibility.Collapsed;
            StatusBox.Visibility = Visibility.Collapsed;
            StatusBox.Text = "";
            StatusBox.Foreground = System.Windows.Media.Brushes.Black;
            TimerLabel.Visibility = Visibility.Collapsed;
            URLCheck.IsEnabled = true;
            URLCheck.IsChecked = false;
            GPSCheck.IsEnabled = true;
            GPSCheck.IsChecked = false;
            SummaryCheck.IsEnabled = true;
            SummaryCheck.IsChecked = false;
            CompInfoCheck.IsEnabled = true;
            CompInfoCheck.IsChecked = false;
            ActivityCheck.IsEnabled = true;
            ActivityCheck.IsChecked = false;
            TimelineCheck.IsEnabled = true;
            TimelineCheck.IsChecked = false;
            GatherCheck.IsChecked = false;
            PropMapCheck.IsChecked = false;
            SecStoreCheck.IsChecked = false;
            ApplyCompanionState();
            XlsxCheck.IsEnabled = true;
            SqliteCheck.IsEnabled = true;
            DBPath.IsEnabled = true;
            DBPathPicker.IsEnabled = true;
            OutputPath.IsEnabled = true;
            OutputPathPicker.IsEnabled = true;
            SelectDeselectButton.Content = "Select All";
            SelectDeselectButton.IsEnabled = true;
            SearchIndex.Reset();
        }

        private void SetupUI()
        {
            stopWatch?.Reset();
            stopWatch?.Start();
            elapsedTimer?.Start();
            GoButton.IsEnabled = false;
            ProgressBarControl.Visibility = Visibility.Visible;
            StatusBox.Visibility = Visibility.Visible;
            TimerLabel.Visibility = Visibility.Visible;
            ProgressBarControl.IsIndeterminate = true;
            URLCheck.IsEnabled = false;
            GPSCheck.IsEnabled = false;
            SummaryCheck.IsEnabled = false;
            CompInfoCheck.IsEnabled = false;
            ActivityCheck.IsEnabled = false;
            TimelineCheck.IsEnabled = false;
            GatherCheck.IsEnabled = false;
            PropMapCheck.IsEnabled = false;
            SecStoreCheck.IsEnabled = false;
            XlsxCheck.IsEnabled = false;
            SqliteCheck.IsEnabled = false;
            SelectDeselectButton.IsEnabled = false;
            DBPath.IsEnabled = false;
            DBPathPicker.IsEnabled = false;
            OutputPath.IsEnabled = false;
            OutputPathPicker.IsEnabled = false;
            SearchIndex.Reset();
        }

        private async void GoButtonClick(object sender, RoutedEventArgs e)
        {
            string now = DateTime.Now.ToUniversalTime().ToString("yyyyMMdd-HHmmss");
            try
            {
                if (DBPath.Text == "")
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show("No database was selected.\n\nPlease select a valid database path and try again.", "Database Path is required.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                if (OutputPath.Text == "")
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show("No output path was selected.\n\nPlease select a valid output path and try again.", "Output Path is required.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                bool xlsx = XlsxCheck.IsChecked == true;
                bool sqlite = SqliteCheck.IsChecked == true;
                if (!xlsx && !sqlite)
                {
                    App.Current.MainWindow.Activate();
                    System.Windows.MessageBox.Show("No output format was selected.\n\nPlease select Spreadsheet, SQLite or both and try again.", "Output format is required.", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    return;
                }
                var filters = new SearchIndex.FilterOptions(
                    URLCheck.IsChecked == true,
                    GPSCheck.IsChecked == true,
                    SummaryCheck.IsChecked == true,
                    CompInfoCheck.IsChecked == true,
                    ActivityCheck.IsChecked == true,
                    TimelineCheck.IsChecked == true);
                string? gatherFile = GatherCheck.IsChecked == true ? detectedGatherFile : null;
                string? propMapFile = PropMapCheck.IsChecked == true ? detectedPropMapFile : null;
                string? secStoreFile = SecStoreCheck.IsChecked == true ? detectedSecStoreFile : null;
                SetupUI();
                SearchIndex.DbFile = DBPath.Text.Trim().Trim('"');
                await Task.Yield();
                await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render);
                StatusBox.Text = "Extracting data ...";
                await Task.Run(SearchIndex.ExtractData);
                if (gatherFile is not null)
                {
                    await Task.Run(() => SearchIndex.ExtractGather(gatherFile));
                }
                if (propMapFile is not null)
                {
                    await Task.Run(() => SearchIndex.ExtractPropMap(propMapFile));
                }
                if (secStoreFile is not null)
                {
                    await Task.Run(() => SearchIndex.ExtractSecStore(secStoreFile));
                }
                StatusBox.Text = "Parsing data ...";
                await SearchIndex.ParseDataAsync();
                List<string> outputFiles = [];
                if (xlsx)
                {
                    if (filters.Any)
                    {
                        await SearchIndex.ParseSubData(filters, status => StatusBox.Text = status);
                    }
                    StatusBox.Text = "Exporting to XLSX...";
                    outputFiles.Add(Path.Combine(OutputPath.Text, SearchIndex.OutputFileName(now)));
                    await SearchIndex.ExportToExcel(outputFiles[^1]);
                }
                if (sqlite)
                {
                    StatusBox.Text = "Exporting to SQLite...";
                    outputFiles.Add(Path.Combine(OutputPath.Text, SearchIndex.OutputFileName(now, "sqlite")));
                    await SearchIndex.ExportToSqlite(outputFiles[^1], $"{displayName} v{appVersion}");
                }
                StatusBox.Text = "Extraction and parsing completed!";
                ProgressBarControl.Value = 100;
                ProgressBarControl.IsIndeterminate = false;
                StatusBox.Foreground = System.Windows.Media.Brushes.White;
                stopWatch?.Stop();
                elapsedTimer?.Stop();
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Data successfully exported to:\n\n{string.Join("\n", outputFiles)}", "Export Successful", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                stopWatch?.Stop();
                elapsedTimer?.Stop();
                App.Current.MainWindow.Activate();
                System.Windows.MessageBox.Show($"Error with data extraction:\n\n{ex.Message}", "Data Extraction Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                ResetUI();
            }
        }

        private void AboutClick(object sender, RoutedEventArgs e)
        {
            App.Current.MainWindow.Activate();
            MessageBoxResult result = System.Windows.MessageBox.Show(
                $"{displayName} v{appVersion}\n" +
                "A simple Windows Search Index parser for the SQLite and ESE DB versions.\n\n" +
                $"Author: Corey Forman (digitalsleuth)\n" +
                $"Source: {githubBinaryRepo}\n\n" +
                $"Would you like to visit the repo on GitHub?",
                $"{displayName} v{appVersion}", MessageBoxButton.YesNo, MessageBoxImage.Information);
            if (result == MessageBoxResult.Yes)
            {
                Process.Start(new ProcessStartInfo($"{githubBinaryRepo}") { UseShellExecute = true });
            }
        }
    }
}
