using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MiniExcelLibs;
using ParquetViewer;
using ParquetViewer.Analytics;
using ParquetViewer.Engine;
using ParquetViewer.Engine.Exceptions;
using ParquetViewer.Exceptions;
using ParquetViewer.Helpers;
using ParquetViewer.Services;
using Resources = ParquetViewer.Resources;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace ParquetViewer.Main.Local.ViewModels
{
    public sealed partial class MainViewModel : ObservableObject
    {
        private const int DefaultOffset = 0;
        private const int DefaultRowCount = 1000;

        private readonly IFieldSelectionService _fieldSelectionService;
        private readonly IMetadataService _metadataService;
        private IParquetEngine? _openParquetEngine;

        [ObservableProperty] private string _title = "Parquet Viewer";
        [ObservableProperty] private string _searchFilter = string.Empty;
        [ObservableProperty] private int _currentOffset = DefaultOffset;
        [ObservableProperty] private int _currentMaxRowCount = DefaultRowCount;
        [ObservableProperty] private string _recordCountText = "0";
        [ObservableProperty] private string _totalRowCountText = "0";
        [ObservableProperty] private string _actualShownRecordCount = "0";
        [ObservableProperty] private string _performanceTooltip = string.Empty;
        [ObservableProperty] private bool _isLoading;

        private CancellationTokenSource? _loadCts;

        [NotifyCanExecuteChangedFor(nameof(ChangeFieldsCommand))]
        [NotifyCanExecuteChangedFor(nameof(RunQueryCommand))]
        [NotifyCanExecuteChangedFor(nameof(SaveAsCommand))]
        [NotifyCanExecuteChangedFor(nameof(ShowMetadataCommand))]
        [NotifyCanExecuteChangedFor(nameof(GetSqlCreateTableScriptCommand))]
        [ObservableProperty] private bool _isAnyFileOpen;

        [NotifyCanExecuteChangedFor(nameof(LoadAllRowsCommand))]
        [ObservableProperty] private bool _hasMoreRows;

        private string? _openFileOrFolderPath;
        public string? OpenFileOrFolderPath
        {
            get => _openFileOrFolderPath;
            private set
            {
                _openFileOrFolderPath = value;
                _openParquetEngine?.DisposeSafely();
                _openParquetEngine = null;
                _selectedFields = null;
                MainDataSource?.Dispose();
                MainDataSource = null;
                HasMoreRows = false;
                RecordCountText = "0";
                TotalRowCountText = "0";
                ActualShownRecordCount = "0";
                IsAnyFileOpen = false;
                CurrentOffset = DefaultOffset;
                SearchFilter = string.Empty;

                if (string.IsNullOrWhiteSpace(value))
                    Title = "Parquet Viewer";
                else
                    Title = string.Format(
                        File.Exists(value)
                            ? Resources.Strings.MainWindowOpenFileTitleFormat
                            : Resources.Strings.MainWindowOpenFolderTitleFormat,
                        value);
            }
        }

        private DataTable? _mainDataSource;
        public DataTable? MainDataSource
        {
            get => _mainDataSource;
            private set
            {
                if (_mainDataSource == value) return;
                _mainDataSource?.Dispose();
                _mainDataSource = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(DataView));
                if (_mainDataSource is not null)
                    HasMoreRows = _mainDataSource.Rows.Count < (_openParquetEngine?.RecordCount ?? 0);
            }
        }

        public DataView? DataView => _mainDataSource?.DefaultView;

        public bool IsDarkMode => AppSettings.DarkMode;
        public bool AlwaysLoadAllRecords => AppSettings.AlwaysLoadAllRecords;
        public bool AnalyticsConsent => AppSettings.AnalyticsDataGatheringConsent;
        public bool IsDefaultDateFormat => AppSettings.DateTimeDisplayFormat == DateFormat.Default;
        public bool IsIso8601DateFormat => AppSettings.DateTimeDisplayFormat == DateFormat.ISO8601;
        public bool IsCustomDateFormat => AppSettings.DateTimeDisplayFormat == DateFormat.Custom;
        public bool IsEnglish => AppSettings.UserSelectedCulture is null || AppSettings.UserSelectedCulture.TwoLetterISOLanguageName == "en";
        public bool IsTurkish => AppSettings.UserSelectedCulture?.TwoLetterISOLanguageName == "tr";
        public bool IsKorean => AppSettings.UserSelectedCulture?.TwoLetterISOLanguageName == "ko";

        private List<string>? _selectedFields;

        public MainViewModel(IFieldSelectionService fieldSelectionService, IMetadataService metadataService)
        {
            _fieldSelectionService = fieldSelectionService;
            _metadataService = metadataService;
        }

        partial void OnSearchFilterChanged(string value)
        {
            // live filter handled by RunQueryCommand
        }

        partial void OnCurrentOffsetChanged(int value) => _ = LoadDataAsync();
        partial void OnCurrentMaxRowCountChanged(int value) => _ = LoadDataAsync();

        [RelayCommand]
        private async Task OpenFileAsync()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*",
                Multiselect = false
            };
            if (dialog.ShowDialog() != true) return;

            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.FileOpen);
            await OpenNewFileOrFolderAsync(dialog.FileName);
        }

        [RelayCommand]
        private async Task OpenFolderAsync()
        {
            var dialog = new OpenFolderDialog();
            if (dialog.ShowDialog() != true) return;

            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.FolderOpen);
            await OpenNewFileOrFolderAsync(dialog.FolderName);
        }

        [RelayCommand]
        private void NewFile()
        {
            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.FileNew);
            OpenFileOrFolderPath = null;
        }

        [RelayCommand(CanExecute = nameof(IsAnyFileOpen))]
        private async Task ChangeFieldsAsync()
        {
            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.ChangeFields);
            var fields = await OpenFieldSelectionAsync(forceOpen: true);
            if (fields is not null)
                await SetSelectedFieldsAsync(fields);
        }

        [RelayCommand(CanExecute = nameof(IsAnyFileOpen))]
        private void RunQuery()
        {
            if (MainDataSource is null) return;

            var queryText = SearchFilter ?? string.Empty;
            queryText = Regex.Replace(queryText, "^WHERE ", string.Empty, RegexOptions.IgnoreCase).Trim();

            if (string.IsNullOrWhiteSpace(queryText) || MainDataSource.DefaultView.RowFilter == queryText)
                return;

            var stopwatch = Stopwatch.StartNew();
            var queryEvent = new ExecuteQueryEvent
            {
                RecordCountTotal = MainDataSource.Rows.Count,
                ColumnCount = MainDataSource.Columns.Count
            };

            try
            {
                MainDataSource.DefaultView.RowFilter = queryText;
                queryEvent.IsValid = true;
                queryEvent.RecordCountFiltered = MainDataSource.DefaultView.Count;
            }
            catch (Exception ex)
            {
                MainDataSource.DefaultView.RowFilter = null;
                throw new InvalidQueryException(ex);
            }
            finally
            {
                queryEvent.RunTimeMS = stopwatch.ElapsedMilliseconds;
                _ = queryEvent.Record();
                ActualShownRecordCount = MainDataSource.DefaultView.Count.ToString();
            }
        }

        [RelayCommand]
        private void ClearFilter()
        {
            if (string.IsNullOrEmpty(MainDataSource?.DefaultView.RowFilter)) return;
            MainDataSource.DefaultView.RowFilter = null;
            SearchFilter = string.Empty;
            ActualShownRecordCount = (MainDataSource?.DefaultView.Count ?? 0).ToString();
        }

        private bool CanLoadAllRows() => HasMoreRows;

        [RelayCommand(CanExecute = nameof(CanLoadAllRows))]
        private async Task LoadAllRowsAsync()
        {
            if (_openParquetEngine is null) return;
            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.LoadAllRows);
            CurrentMaxRowCount = (int)_openParquetEngine.RecordCount;
        }

        [RelayCommand(CanExecute = nameof(IsAnyFileOpen))]
        private async Task SaveAsAsync()
        {
            if (MainDataSource?.DefaultView.Count <= 0) return;

            var dialog = new SaveFileDialog
            {
                Title = Resources.Strings.RecordsToBeExportedTitleFormat.Format(MainDataSource!.DefaultView.Count),
                Filter = "CSV file (*.csv)|*.csv|JSON file (*.json)|*.json|Excel '93 file (*.xls)|*.xls|Excel '07 file (*.xlsx)|*.xlsx"
            };

            if (_openParquetEngine?.Metadata.SchemaTree?.Children.All(s => s.IsPrimitive) == true
                && _openParquetEngine is Engine.ParquetNET.ParquetEngine)
            {
                dialog.Filter += "|Parquet file (*.parquet)|*.parquet";
            }

            if (dialog.ShowDialog() != true) return;

            var filePath = dialog.FileName;
            var extension = Path.GetExtension(filePath);
            var fileType = UtilityMethods.ExtensionToFileType(extension)
                ?? throw new ArgumentOutOfRangeException(extension);

            var stopwatch = Stopwatch.StartNew();

            try
            {
                IsLoading = true;
                await ExportResultsAsync(MainDataSource, fileType, _openParquetEngine, filePath,
                    CancellationToken.None, null, OpenFileOrFolderPath);

                var fileSizeInBytes = new FileInfo(filePath).Length;
                FileExportEvent.FireAndForget(fileType, fileSizeInBytes, MainDataSource.DefaultView.Count,
                    MainDataSource.Columns.Count, stopwatch.ElapsedMilliseconds);

                MessageBox.Show(
                    Resources.Strings.ExportSuccessfulMessageFormat.Format(Math.Round((fileSizeInBytes / 1024.0) / 1024.0, 2)),
                    Resources.Strings.ExportSuccessfulTitle,
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (IOException ex)
            {
                TryDeleteFile(filePath);
                ShowError(ex.Message, Resources.Errors.ExportFailedErrorTitle);
            }
            finally
            {
                IsLoading = false;
            }
        }

        [RelayCommand]
        private void ToggleDarkMode()
        {
            AppSettings.DarkMode = !AppSettings.DarkMode;
            OnPropertyChanged(nameof(IsDarkMode));
        }

        [RelayCommand]
        private void ToggleAlwaysLoadAllRecords()
        {
            AppSettings.AlwaysLoadAllRecords = !AppSettings.AlwaysLoadAllRecords;
            OnPropertyChanged(nameof(AlwaysLoadAllRecords));
        }

        [RelayCommand]
        private void SetDateFormat(string tag)
        {
            if (int.TryParse(tag, out var val))
            {
                AppSettings.DateTimeDisplayFormat = val.ToEnum(DateFormat.Default);
                OnPropertyChanged(nameof(IsDefaultDateFormat));
                OnPropertyChanged(nameof(IsIso8601DateFormat));
                OnPropertyChanged(nameof(IsCustomDateFormat));
            }
        }

        [RelayCommand]
        private void SetLanguage(string? cultureName)
        {
            // empty/null means English (default) — stored as null in AppSettings
            System.Globalization.CultureInfo? newCulture = null;
            if (!string.IsNullOrWhiteSpace(cultureName))
            {
                if (!UtilityMethods.TryParseCultureInfo(cultureName, out newCulture)) return;
            }

            if (newCulture?.Name == AppSettings.UserSelectedCulture?.Name) return;

            if (MessageBox.Show(
                Resources.Strings.LanguageChangeConfirmationMessage,
                Resources.Strings.LanguageChangeConfirmationTitle,
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

            AppSettings.UserSelectedCulture = newCulture;
            UtilityMethods.RestartApplication();
        }

        [RelayCommand(CanExecute = nameof(IsAnyFileOpen))]
        private void ShowMetadata()
        {
            if (_openParquetEngine is null) return;
            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.MetadataViewer);
            _metadataService.Show(_openParquetEngine);
        }

        [RelayCommand(CanExecute = nameof(IsAnyFileOpen))]
        private void GetSqlCreateTableScript()
        {
            if (MainDataSource?.Columns.Count > 0)
            {
                var path = OpenFileOrFolderPath?.TrimEnd('/');
                var tableName = Path.GetFileNameWithoutExtension(path) ?? "MY_TABLE";

                var dataset = new DataSet();
                MainDataSource.TableName = tableName;
                dataset.Tables.Add(MainDataSource);

                var adapter = new CustomScriptBasedSchemaAdapter();
                var sql = adapter.GetSchemaScript(dataset, false);

                dataset.Tables.Remove(MainDataSource);

                MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.SQLCreateTable);
                Clipboard.SetText(sql);
                MessageBox.Show(Resources.Strings.CreateTableScriptCopiedToClipboardMessage,
                    "ParquetViewer", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(Resources.Strings.CreateTableScriptFailedWithNoFieldsMessage,
                    "ParquetViewer", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void OpenUserGuide()
        {
            MenuBarClickEvent.FireAndForget(MenuBarClickEvent.ActionId.UserGuide);
            Process.Start(new ProcessStartInfo(Constants.WikiURL) { UseShellExecute = true });
        }

        [RelayCommand]
        private void ToggleAnalyticsConsent()
        {
            AppSettings.AnalyticsDataGatheringConsent = !AppSettings.AnalyticsDataGatheringConsent;
            OnPropertyChanged(nameof(AnalyticsConsent));
        }

        [RelayCommand]
        private void CancelLoad()
        {
            _loadCts?.Cancel();
        }

        [RelayCommand]
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        [RelayCommand]
        private void ShowAbout()
        {
            MessageBox.Show(
                $"Parquet Viewer {Env.AssemblyVersion}",
                "About Parquet Viewer",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // ── internal helpers ──────────────────────────────────────────────

        public async Task OpenNewFileOrFolderAsync(string path)
        {
            OpenFileOrFolderPath = path;

            var fields = await OpenFieldSelectionAsync(forceOpen: false);
            if (_openParquetEngine is null) return;

            if (AppSettings.AlwaysLoadAllRecords)
                CurrentMaxRowCount = (int)_openParquetEngine.RecordCount;
            else
                CurrentMaxRowCount = DefaultRowCount;

            if (fields is not null)
            {
                await SetSelectedFieldsAsync(fields);
                AppSettings.OpenedFileCount++;
            }
        }

        private async Task<List<string>?> OpenFieldSelectionAsync(bool forceOpen)
        {
            if (string.IsNullOrWhiteSpace(OpenFileOrFolderPath)) return null;

            if (_openParquetEngine is null)
            {
                try
                {
                    _openParquetEngine = await Engine.ParquetNET.ParquetEngine.OpenFileOrFolderAsync(
                        OpenFileOrFolderPath, default);
                }
                catch (Exception ex)
                {
                    if (_openParquetEngine is null) OpenFileOrFolderPath = null;
                    HandleEngineException(ex);
                    return null;
                }
            }

            List<string>? fields = null;
            try { fields = _openParquetEngine.Fields; }
            catch (ArgumentException ex) when (ex.Message.StartsWith("at least one field is required")) { }
            catch (Exception ex)
            {
                throw new Parquet.ParquetException(Resources.Errors.ParquetSchemaReadErrorMessage, ex);
            }

            if (fields?.Count > 0)
            {
                if (AppSettings.AlwaysSelectAllFields && !forceOpen) return fields;

                return await _fieldSelectionService.ShowAsync(fields, _selectedFields ?? new List<string>());
            }

            ShowError(Resources.Errors.NoFieldsFoundErrorMessage, Resources.Errors.NoFieldsFoundErrorTitle);
            return null;
        }

        private async Task SetSelectedFieldsAsync(List<string> fields)
        {
            var dupes = fields.GroupBy(f => f.ToUpperInvariant()).Where(g => g.Count() > 1).SelectMany(g => g).ToList();
            if (dupes.Count > 0)
            {
                fields = fields.Where(f => !dupes.Any(d => d.Equals(f, StringComparison.InvariantCultureIgnoreCase))).ToList();
                MessageBox.Show(
                    $"The following duplicate fields could not be loaded: {string.Join(',', dupes)}.\n\nCase sensitive field names are not currently supported.",
                    "Duplicate fields detected", MessageBoxButton.OK, MessageBoxImage.Warning);
            }

            _selectedFields = fields;
            if (_selectedFields.Count > 0)
                await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            if (_openParquetEngine is null || _selectedFields is null || _selectedFields.Count == 0) return;
            if (!File.Exists(OpenFileOrFolderPath) && !Directory.Exists(OpenFileOrFolderPath))
            {
                ShowError(Resources.Errors.OpenFileNoLongerExistsErrorMessageFormat.Format(OpenFileOrFolderPath + Environment.NewLine));
                return;
            }

            var stopwatch = Stopwatch.StartNew();
            var loadTime = TimeSpan.Zero;
            var indexTime = TimeSpan.Zero;
            var wasSuccessful = false;

            _loadCts?.Cancel();
            _loadCts?.Dispose();
            _loadCts = new CancellationTokenSource();
            var cts = _loadCts;

            try
            {
                IsLoading = true;

                var intermediate = await Task.Run(async () =>
                    await _openParquetEngine.ReadRowsAsync(_selectedFields, CurrentOffset, CurrentMaxRowCount,
                        cts.Token, null));

                loadTime = stopwatch.Elapsed;

                var finalResult = await Task.Run(() => intermediate.Invoke(false));
                indexTime = stopwatch.Elapsed - loadTime;

                RecordCountText = string.Format(Resources.Strings.LoadedRecordCountRangeFormat,
                    CurrentOffset, CurrentOffset + finalResult.Rows.Count);
                TotalRowCountText = _openParquetEngine.RecordCount.ToString();
                ActualShownRecordCount = finalResult.Rows.Count.ToString();

                IsAnyFileOpen = true;
                MainDataSource = finalResult;
                wasSuccessful = true;
            }
            catch (AllFilesSkippedException ex) { HandleAllFilesSkippedException(ex); }
            catch (SomeFilesSkippedException ex) { HandleSomeFilesSkippedException(ex); }
            catch (FileReadException ex) { HandleFileReadException(ex); }
            catch (MultipleSchemasFoundException ex) { HandleMultipleSchemasFoundException(ex); }
            catch (MalformedFieldException ex) { ShowError(Resources.Errors.MalformedFieldErrorMessageFormat.Format(ex.Message)); }
            catch (DecimalOverflowException ex) { HandleDecimalOverflowException(ex); }
            catch (OperationCanceledException) { }
            finally
            {
                stopwatch.Stop();
                IsLoading = false;
                var renderTime = stopwatch.Elapsed - loadTime - indexTime;
                PerformanceTooltip = $"Total: {stopwatch.Elapsed:mm\\:ss\\.ff}\n  Load: {loadTime:mm\\:ss\\.ff}\n  Index: {indexTime:mm\\:ss\\.ff}\n  Render: {renderTime:mm\\:ss\\.ff}";

                if (wasSuccessful && _openParquetEngine is not null)
                {
                    var engineType = _openParquetEngine is Engine.ParquetNET.ParquetEngine
                        ? FileOpenEvent.ParquetEngineTypeId.ParquetNET
                        : FileOpenEvent.ParquetEngineTypeId.DuckDB;

                    FileOpenEvent.FireAndForget(
                        Directory.Exists(OpenFileOrFolderPath),
                        _openParquetEngine.NumberOfPartitions,
                        _openParquetEngine.RecordCount,
                        _openParquetEngine.Metadata.RowGroups.Count,
                        _openParquetEngine.Fields.Count,
                        MainDataSource!.Columns.Cast<DataColumn>().Select(c => c.DataType.Name).Distinct().Order().ToArray(),
                        CurrentOffset,
                        CurrentMaxRowCount,
                        MainDataSource!.Columns.Count,
                        (long)stopwatch.Elapsed.TotalMilliseconds,
                        (long)loadTime.TotalMilliseconds,
                        (long)indexTime.TotalMilliseconds,
                        (long)renderTime.TotalMilliseconds,
                        engineType);
                }
            }
        }

        // ── exception helpers ─────────────────────────────────────────────

        private void HandleEngineException(Exception ex)
        {
            if (ex is AllFilesSkippedException afse) HandleAllFilesSkippedException(afse);
            else if (ex is SomeFilesSkippedException sfse) HandleSomeFilesSkippedException(sfse);
            else if (ex is FileReadException fre) HandleFileReadException(fre);
            else if (ex is MultipleSchemasFoundException msfe) HandleMultipleSchemasFoundException(msfe);
            else if (ex is FileNotFoundException fnfe) ShowError(fnfe.Message);
            else if (ex is not OperationCanceledException) throw ex;
        }

        private static void HandleAllFilesSkippedException(AllFilesSkippedException ex)
        {
            var sb = new StringBuilder(Resources.Errors.NoValidParquetFilesFoundErrorMessage);
            foreach (var f in ex.SkippedFiles) sb.AppendLine($"-{f.FileName}");
            ShowError(sb.ToString());
        }

        private static void HandleSomeFilesSkippedException(SomeFilesSkippedException ex)
        {
            var sb = new StringBuilder(Resources.Errors.SomeInvalidParquetFilesFoundErrorMessage);
            foreach (var f in ex.SkippedFiles) sb.AppendLine($"-{f.FileName}");
            ShowError(sb.ToString());
        }

        private static void HandleFileReadException(FileReadException ex)
            => ShowError(Resources.Errors.UnexpectedFileReadErrorMessageFormat.Format(ex));

        private static void HandleMultipleSchemasFoundException(MultipleSchemasFoundException ex)
        {
            var sb = new StringBuilder(Resources.Errors.MultipleSchemasDetectedErrorMessage);
            sb.AppendLine(" ");
            var index = 1;
            foreach (var schema in ex.Schemas)
            {
                sb.AppendLine(Resources.Errors.MultipleSchemasDetectedEntriesErrorMessageFormat.Format(index++));
                foreach (var field in schema.Take(5)) sb.AppendLine($"  {field}");
                if (index > 10) { sb.AppendLine("..."); break; }
            }
            ShowError(sb.ToString());
        }

        private static void HandleDecimalOverflowException(DecimalOverflowException ex)
            => ShowError(
                (ex.HasDetailedInfo
                    ? Resources.Errors.DecimalValueTooLargeErrorMessageFormat
                    : Resources.Errors.DecimalValueUnknownSizeTooLargeErrorMessageFormat)
                .Format(ex.FieldName, ex.Precision, ex.Scale,
                    DecimalOverflowException.MAX_DECIMAL_PRECISION,
                    DecimalOverflowException.MAX_DECIMAL_SCALE),
                Resources.Errors.DecimalValueTooLargeErrorTitle);

        private static void ShowError(string message, string? title = null)
            => MessageBox.Show(message, title ?? Resources.Errors.GenericErrorMessage,
                MessageBoxButton.OK, MessageBoxImage.Error);

        private static void TryDeleteFile(string? path)
        {
            try { if (!string.IsNullOrWhiteSpace(path)) File.Delete(path); }
            catch { }
        }

        // ── export helpers (ported from MainForm.Helpers.cs) ──────────────

        private static Task ExportResultsAsync(DataTable dataTable, FileType fileType, IParquetEngine? engine,
            string filePath, CancellationToken ct, IProgress<int>? progress, string? sourcePath)
        {
            return fileType switch
            {
                FileType.CSV => WriteCsvAsync(dataTable, filePath, ct, progress),
                FileType.XLS => WriteExcel93Async(dataTable, filePath, ct, progress),
                FileType.XLSX => WriteExcel2007Async(dataTable, filePath, Path.GetFileNameWithoutExtension(sourcePath) ?? "Sheet1", ct, progress),
                FileType.JSON => WriteJsonAsync(dataTable, filePath, ct, progress),
                FileType.PARQUET => WriteParquetAsync(engine!, dataTable, filePath, ct, progress),
                _ => throw new Exception(string.Format(Resources.Errors.UnsupportedExportTypeFormat, fileType))
            };
        }

        private static async Task WriteExcel2007Async(DataTable dt, string path, string sheetName, CancellationToken ct, IProgress<int>? progress)
        {
            const int MaxSheetNameLength = 31;
            sheetName = Regex.Replace(sheetName, "[^a-zA-Z0-9 _\\-()]", string.Empty).Left(MaxSheetNameLength);
            using var fs = new FileStream(path, FileMode.OpenOrCreate);
            await fs.SaveAsAsync(dt, printHeader: true, sheetName, ExcelType.XLSX, configuration: null, progress, ct);
        }

        private static Task WriteCsvAsync(DataTable dt, string path, CancellationToken ct, IProgress<int>? progress)
            => Task.Run(() =>
            {
                using var writer = new StreamWriter(path, false, Encoding.UTF8);
                var row = new StringBuilder();
                bool first = true;
                foreach (DataColumn col in dt.Columns)
                {
                    if (!first) row.Append(','); else first = false;
                    row.Append(col.ColumnName.Replace("\r", "").Replace("\n", "").Replace(",", ""));
                }
                writer.WriteLine(row.ToString());

                var dateFmt = AppSettings.DateTimeDisplayFormat.GetDateFormat();
                var dateonlyFmt = AppSettings.DateTimeDisplayFormat.GetDateOnlyFormat();
                var timeonlyFmt = AppSettings.DateTimeDisplayFormat.GetTimeOnlyFormat();

                foreach (DataRowView rowView in dt.DefaultView)
                {
                    row.Clear(); first = true;
                    foreach (var value in rowView.Row.ItemArray)
                    {
                        if (ct.IsCancellationRequested) break;
                        if (!first) row.Append(','); else first = false;
                        row.Append(value switch
                        {
                            DateTime dt2 => UtilityMethods.CleanCSVValue(dt2.ToString(dateFmt)),
                            DateOnly d => UtilityMethods.CleanCSVValue(d.ToString(dateonlyFmt)),
                            TimeOnly t => UtilityMethods.CleanCSVValue(t.ToString(timeonlyFmt)),
                            _ => UtilityMethods.CleanCSVValue(value!.ToString()!)
                        });
                        progress?.Report(1);
                    }
                    writer.WriteLine(row.ToString());
                }
            }, ct);

        private static Task WriteExcel93Async(DataTable dt, string path, CancellationToken ct, IProgress<int>? progress)
            => Task.Run(() =>
            {
                var dateFmt = AppSettings.DateTimeDisplayFormat.GetDateFormat();
                var dateonlyFmt = AppSettings.DateTimeDisplayFormat.GetDateOnlyFormat();
                var timeonlyFmt = AppSettings.DateTimeDisplayFormat.GetTimeOnlyFormat();
                using var fs = new FileStream(path, FileMode.OpenOrCreate);
                var writer = new ExcelWriter(fs);
                writer.BeginWrite();
                for (int i = 0; i < dt.Columns.Count; i++)
                    writer.WriteCell(0, i, dt.Columns[i].ColumnName);
                for (int i = 0; i < dt.DefaultView.Count; i++)
                {
                    if (ct.IsCancellationRequested) break;
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        var val = dt.DefaultView[i][j];
                        if (val == DBNull.Value) writer.WriteCell(i + 1, j);
                        else if (val is int or uint or sbyte or byte) writer.WriteCell(i + 1, j, Convert.ToInt32(val));
                        else if (val is double or decimal or float or long) writer.WriteCell(i + 1, j, Convert.ToDouble(val));
                        else if (val is DateTime dt2) writer.WriteCell(i + 1, j, dt2.ToString(dateFmt));
                        else if (val is DateOnly d) writer.WriteCell(i + 1, j, d.ToString(dateonlyFmt));
                        else if (val is TimeOnly t) writer.WriteCell(i + 1, j, t.ToString(timeonlyFmt));
                        else
                        {
                            var s = val.ToString()!;
                            if (s.Length > 255) throw new XlsCellLengthException(255);
                            writer.WriteCell(i + 1, j, s);
                        }
                        progress?.Report(1);
                    }
                }
                writer.EndWrite();
            }, ct);

        private static Task WriteJsonAsync(DataTable dt, string path, CancellationToken ct, IProgress<int>? progress)
            => Task.Run(() =>
            {
                using var fs = new FileStream(path, FileMode.OpenOrCreate);
                using var jsonWriter = new Engine.Utf8JsonWriterWithRunningLength(fs);
                jsonWriter.WriteStartArray();
                foreach (DataRowView row in dt.DefaultView)
                {
                    jsonWriter.WriteStartObject();
                    for (int i = 0; i < row.Row.ItemArray.Length; i++)
                    {
                        if (ct.IsCancellationRequested) break;
                        jsonWriter.WritePropertyName(dt.Columns[i].ColumnName);
                        Engine.Helpers.WriteValue(jsonWriter, row.Row.ItemArray[i]!, false);
                        progress?.Report(1);
                    }
                    jsonWriter.WriteEndObject();
                }
                jsonWriter.WriteEndArray();
            }, ct);

        private static Task WriteParquetAsync(IParquetEngine engine, DataTable dt, string path,
            CancellationToken ct, IProgress<int>? progress)
            => Task.Run(async () =>
            {
                var meta = new Dictionary<string, string>
                {
                    ["ParquetViewer"] = $"{{\"CreatedWith\":\"ParquetViewer\",\"Version\":\"{Env.AssemblyVersion}\"}}"
                };
                await engine.WriteDataToParquetFileAsync(dt, path, ct, progress, meta);
            }, ct);
    }
}
