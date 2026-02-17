using System.ComponentModel;
using System.IO.Ports;
using System.Text.Json;
using System.Text;

namespace GlueConfigApp;

public sealed class MainForm : Form
{
    private const int MaxLinesPerGun = 32;
    private const double MinPatternMm = 0.0;
    private const double MaxPatternMm = 100000.0;

    private readonly ComboBox _portCombo = new() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Button _refreshPortsButton = new() { Text = "Refresh Ports" };
    private readonly Button _connectButton = new() { Text = "Connect" };
    private readonly Label _statusLabel = new() { Text = "Disconnected", AutoSize = true };
    private readonly Label _programStateLabel = new() { AutoSize = true, Font = new Font("Segoe UI", 9F, FontStyle.Bold) };

    private readonly NumericUpDown _pulsesPerMm = new() { DecimalPlaces = 4, Minimum = 0.0001M, Maximum = 100000M, Value = 10M, Increment = 0.1M };
    private readonly NumericUpDown _maxMsPerMm = new() { DecimalPlaces = 0, Minimum = 1, Maximum = 60000, Value = 100 };
    private readonly NumericUpDown _photocellOffset = new() { DecimalPlaces = 2, Minimum = -5000, Maximum = 5000, Value = 0, Increment = 1 };
    private readonly NumericUpDown _debounceMs = new() { DecimalPlaces = 0, Minimum = 0, Maximum = 1000, Value = 20 };

    private readonly Button _sendAllButton = new() { Text = "Send All (Config + Patterns)" };

    private readonly Button _activateButton = new() { Text = "Activate" };
    private readonly Button _deactivateButton = new() { Text = "Deactivate" };

    private readonly Button _testOpenGun1Button = new() { Text = "Test 1" };
    private readonly Button _testCloseGun1Button = new() { Text = "Close 1" };
    private readonly Button _testOpenGun2Button = new() { Text = "Test 2" };
    private readonly Button _testCloseGun2Button = new() { Text = "Close 2" };
    private readonly Button _testOpenBothButton = new() { Text = "Test Both" };
    private readonly Button _testCloseBothButton = new() { Text = "Close Both" };

    private readonly NumericUpDown _calibPaperLength = new() { DecimalPlaces = 2, Minimum = 1, Maximum = 10000, Value = 297, Increment = 1 };
    private readonly Button _calibArmButton = new() { Text = "Arm Calibration" };

    private readonly Button _saveProgramButton = new() { Text = "Save Program" };
    private readonly Button _loadProgramButton = new() { Text = "Load Program" };

    private readonly TextBox _logBox = new()
    {
        Multiline = true,
        ScrollBars = ScrollBars.Vertical,
        ReadOnly = true,
        Dock = DockStyle.Fill,
        Font = new Font("Consolas", 9F)
    };

    private readonly DataGridView _gun1Grid = new();
    private readonly DataGridView _gun2Grid = new();

    private readonly Button _gun1AddLineButton = new() { Text = "Add Line" };
    private readonly Button _gun1RemoveLineButton = new() { Text = "Remove Selected" };
    private readonly Button _gun2AddLineButton = new() { Text = "Add Line" };
    private readonly Button _gun2RemoveLineButton = new() { Text = "Remove Selected" };

    private readonly PatternPreviewPanel _gun1Preview = new() { PatternColor = Color.OrangeRed };
    private readonly PatternPreviewPanel _gun2Preview = new() { PatternColor = Color.SteelBlue };

    private readonly BindingList<PatternLine> _gun1Lines = new();
    private readonly BindingList<PatternLine> _gun2Lines = new();

    private SerialPort? _serial;
    private readonly StringBuilder _serialRxBuffer = new();
    private readonly System.Windows.Forms.Timer _serialPollTimer = new() { Interval = 100 };
    private bool _hasUnsavedChanges;
    private bool _hasUnsentChanges;
    private bool _suspendDirtyTracking;

    private static readonly string AppStateDirectory = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "GlueConfigApp");
    private static readonly string AppStateFilePath = Path.Combine(AppStateDirectory, "last_state.json");

    public MainForm()
    {
        Text = "Hot Glue Gun Configuration";
        Width = 1200;
        Height = 760;
        StartPosition = FormStartPosition.CenterScreen;

        BuildLayout();
        WireEvents();
        ConfigurePatternGrid(_gun1Grid, _gun1Lines);
        ConfigurePatternGrid(_gun2Grid, _gun2Lines);

        _gun1Lines.ListChanged += (_, _) => RefreshPreviews();
        _gun2Lines.ListChanged += (_, _) => RefreshPreviews();

        _pulsesPerMm.ValueChanged += (_, _) => MarkConfigChanged();
        _maxMsPerMm.ValueChanged += (_, _) => MarkConfigChanged();
        _photocellOffset.ValueChanged += (_, _) => MarkConfigChanged();
        _debounceMs.ValueChanged += (_, _) => MarkConfigChanged();
        _calibPaperLength.ValueChanged += (_, _) => MarkConfigChanged();

        _gun1Lines.ListChanged += (_, _) => MarkPatternChanged();
        _gun2Lines.ListChanged += (_, _) => MarkPatternChanged();

        RefreshSerialPorts();
        RefreshPreviews();
        LoadLastAppState();
        UpdateProgramStateUi();
    }

    private void PollSerial()
    {
        try
        {
            if (_serial?.IsOpen != true)
            {
                return;
            }

            var text = _serial.ReadExisting();
            if (!string.IsNullOrEmpty(text))
            {
                AppendSerialText(text);
            }
        }
        catch
        {
            // Ignore transient serial read errors.
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        SaveLastAppState();
        DisconnectSerial();
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(8)
        };

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

        var connectionPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            Padding = new Padding(0, 0, 0, 8)
        };

        connectionPanel.Controls.Add(new Label { Text = "COM Port:", AutoSize = true, Margin = new Padding(0, 7, 6, 0) });
        connectionPanel.Controls.Add(_portCombo);
        connectionPanel.Controls.Add(_refreshPortsButton);
        connectionPanel.Controls.Add(_connectButton);
        connectionPanel.Controls.Add(new Label { Text = "Status:", AutoSize = true, Margin = new Padding(20, 7, 6, 0) });
        connectionPanel.Controls.Add(_statusLabel);
        connectionPanel.Controls.Add(new Label { Text = "Program:", AutoSize = true, Margin = new Padding(20, 7, 6, 0) });
        connectionPanel.Controls.Add(_programStateLabel);

        var middle = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1
        };
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
        middle.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 62));

        middle.Controls.Add(BuildConfigGroup(), 0, 0);
        middle.Controls.Add(BuildPatternsGroup(), 1, 0);

        var logGroup = new GroupBox { Text = "Serial Log", Dock = DockStyle.Fill };
        logGroup.Controls.Add(_logBox);

        root.Controls.Add(connectionPanel, 0, 0);
        root.Controls.Add(middle, 0, 1);
        root.Controls.Add(logGroup, 0, 2);

        Controls.Add(root);
    }

    private Control BuildConfigGroup()
    {
        var group = new GroupBox { Text = "Configuration & Commands", Dock = DockStyle.Fill };

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 20,
            Padding = new Padding(10)
        };

        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));

        panel.Controls.Add(new Label { Text = "Pulses per mm", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
        panel.Controls.Add(_pulsesPerMm, 1, 0);

        panel.Controls.Add(new Label { Text = "Max ms per mm", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
        panel.Controls.Add(_maxMsPerMm, 1, 1);

        panel.Controls.Add(new Label { Text = "Photocell offset (mm)", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 2);
        panel.Controls.Add(_photocellOffset, 1, 2);

        panel.Controls.Add(new Label { Text = "Debounce (ms)", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 3);
        panel.Controls.Add(_debounceMs, 1, 3);

        panel.Controls.Add(_sendAllButton, 0, 5);
        panel.SetColumnSpan(_sendAllButton, 2);

        panel.Controls.Add(_activateButton, 0, 8);
        panel.Controls.Add(_deactivateButton, 1, 8);

        panel.Controls.Add(_testOpenGun1Button, 0, 10);
        panel.Controls.Add(_testCloseGun1Button, 1, 10);
        panel.Controls.Add(_testOpenGun2Button, 0, 11);
        panel.Controls.Add(_testCloseGun2Button, 1, 11);
        panel.Controls.Add(_testOpenBothButton, 0, 12);
        panel.Controls.Add(_testCloseBothButton, 1, 12);

        panel.Controls.Add(new Label { Text = "Calib paper length (mm)", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 14);
        panel.Controls.Add(_calibPaperLength, 1, 14);
        panel.Controls.Add(_calibArmButton, 0, 15);

        panel.Controls.Add(_saveProgramButton, 0, 17);
        panel.Controls.Add(_loadProgramButton, 1, 17);

        group.Controls.Add(panel);
        return group;
    }

    private Control BuildPatternsGroup()
    {
        var group = new GroupBox { Text = "Pattern Editor", Dock = DockStyle.Fill };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(10)
        };

        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

        layout.Controls.Add(BuildGunPatternPanel("Gun 1", _gun1Grid, _gun1Preview, _gun1AddLineButton, _gun1RemoveLineButton), 0, 0);
        layout.Controls.Add(BuildGunPatternPanel("Gun 2", _gun2Grid, _gun2Preview, _gun2AddLineButton, _gun2RemoveLineButton), 0, 1);

        group.Controls.Add(layout);
        return group;
    }

    private static Control BuildGunPatternPanel(string title, DataGridView grid, PatternPreviewPanel preview, Button addButton, Button removeButton)
    {
        var group = new GroupBox { Text = title, Dock = DockStyle.Fill };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3
        };

        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

        grid.Dock = DockStyle.Fill;
        preview.Dock = DockStyle.Fill;
        preview.MinimumSize = new Size(100, 70);

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        buttons.Controls.Add(addButton);
        buttons.Controls.Add(removeButton);

        layout.Controls.Add(grid, 0, 0);
        layout.Controls.Add(buttons, 0, 1);
        layout.Controls.Add(preview, 0, 2);

        group.Controls.Add(layout);
        return group;
    }

    private static void ConfigurePatternGrid(DataGridView grid, BindingList<PatternLine> source)
    {
        grid.AutoGenerateColumns = false;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = false;
        grid.RowHeadersVisible = false;

        var startCol = new DataGridViewTextBoxColumn
        {
            HeaderText = "Start (mm)",
            DataPropertyName = nameof(PatternLine.StartMm),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        };
        var endCol = new DataGridViewTextBoxColumn
        {
            HeaderText = "End (mm)",
            DataPropertyName = nameof(PatternLine.EndMm),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        };

        grid.Columns.Add(startCol);
        grid.Columns.Add(endCol);
        grid.DataSource = source;
    }

    private void WireEvents()
    {
        _refreshPortsButton.Click += (_, _) => RefreshSerialPorts();
        _connectButton.Click += (_, _) => ToggleConnection();

        _sendAllButton.Click += (_, _) => { SendSetConfig(); SendPatterns(); };

        _activateButton.Click += (_, _) => SendJson(new { cmd = "set_active", active = true });
        _deactivateButton.Click += (_, _) => SendJson(new { cmd = "set_active", active = false });

        _testOpenGun1Button.Click += (_, _) => SendJson(new { cmd = "test_open", gun = 1, timeout_ms = 2000 });
        _testCloseGun1Button.Click += (_, _) => SendJson(new { cmd = "test_close", gun = 1 });
        _testOpenGun2Button.Click += (_, _) => SendJson(new { cmd = "test_open", gun = 2, timeout_ms = 2000 });
        _testCloseGun2Button.Click += (_, _) => SendJson(new { cmd = "test_close", gun = 2 });
        _testOpenBothButton.Click += (_, _) => SendJson(new { cmd = "test_open", gun = "both", timeout_ms = 2000 });
        _testCloseBothButton.Click += (_, _) => SendJson(new { cmd = "test_close", gun = "both" });
        _calibArmButton.Click += (_, _) => SendCalibrationArm();

        _saveProgramButton.Click += (_, _) => SaveProgramToFile();
        _loadProgramButton.Click += (_, _) => LoadProgramFromFile();

        _gun1AddLineButton.Click += (_, _) => TryAddLine(_gun1Lines, "Gun 1");
        _gun2AddLineButton.Click += (_, _) => TryAddLine(_gun2Lines, "Gun 2");

        _gun1RemoveLineButton.Click += (_, _) => RemoveSelectedLine(_gun1Grid, _gun1Lines);
        _gun2RemoveLineButton.Click += (_, _) => RemoveSelectedLine(_gun2Grid, _gun2Lines);

        _gun1Grid.CellEndEdit += (_, _) => { NormalizePattern(_gun1Lines); RefreshPreviews(); };
        _gun2Grid.CellEndEdit += (_, _) => { NormalizePattern(_gun2Lines); RefreshPreviews(); };
        _gun1Grid.CellValueChanged += (_, _) => RefreshPreviews();
        _gun2Grid.CellValueChanged += (_, _) => RefreshPreviews();
        _gun1Grid.DataError += (_, e) => e.ThrowException = false;
        _gun2Grid.DataError += (_, e) => e.ThrowException = false;

        _serialPollTimer.Tick += (_, _) => PollSerial();
    }

    private void RefreshSerialPorts()
    {
        var current = _portCombo.SelectedItem?.ToString();
        var ports = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();

        _portCombo.Items.Clear();
        _portCombo.Items.AddRange(ports);

        if (!string.IsNullOrWhiteSpace(current) && ports.Contains(current, StringComparer.OrdinalIgnoreCase))
        {
            _portCombo.SelectedItem = current;
        }
        else if (_portCombo.Items.Count > 0)
        {
            _portCombo.SelectedIndex = 0;
        }
    }

    private void ToggleConnection()
    {
        if (_serial?.IsOpen == true)
        {
            DisconnectSerial();
            return;
        }

        if (_portCombo.SelectedItem is not string port || string.IsNullOrWhiteSpace(port))
        {
            MessageBox.Show("Please select a COM port.", "Missing Port", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            _serial = new SerialPort(port, 115200)
            {
                NewLine = "\n",
                ReadTimeout = 1000,
                WriteTimeout = 1000,
                DtrEnable = true,
                RtsEnable = true,
                Handshake = Handshake.None
            };
            _serial.DataReceived += SerialDataReceived;
            _serial.Open();
            _serial.DiscardInBuffer();
            _serial.DiscardOutBuffer();
            _serialPollTimer.Start();

            _connectButton.Text = "Disconnect";
            _statusLabel.Text = $"Connected ({port})";
            AppendLog($"Connected to {port}; waiting for device messages...");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Could not open serial port: {ex.Message}", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            DisconnectSerial();
        }
    }

    private void DisconnectSerial()
    {
        if (_serial is null)
        {
            _connectButton.Text = "Connect";
            _statusLabel.Text = "Disconnected";
            return;
        }

        try
        {
            _serialPollTimer.Stop();
            _serial.DataReceived -= SerialDataReceived;
            if (_serial.IsOpen)
            {
                _serial.Close();
            }
            _serial.Dispose();
        }
        catch
        {
            // Ignore cleanup errors.
        }
        finally
        {
            _serial = null;
            _connectButton.Text = "Connect";
            _statusLabel.Text = "Disconnected";
            AppendLog("Disconnected");
        }
    }

    private void SerialDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        try
        {
            var text = _serial?.ReadExisting();
            if (string.IsNullOrEmpty(text))
            {
                return;
            }
            BeginInvoke(() => AppendSerialText(text));
        }
        catch
        {
            // Ignore transient serial read errors.
        }
    }

    private void AppendSerialText(string text)
    {
        _serialRxBuffer.Append(text);
        while (true)
        {
            var current = _serialRxBuffer.ToString();
            var newlineIndex = current.IndexOf('\n');
            if (newlineIndex < 0)
            {
                break;
            }

            var line = current.Substring(0, newlineIndex).TrimEnd('\r');
            _serialRxBuffer.Remove(0, newlineIndex + 1);
            if (line.Length > 0)
            {
                AppendLog($"RX {line}");
                HandleIncomingLine(line);
            }
        }
    }

    private void HandleIncomingLine(string line)
    {
        try
        {
            using var doc = JsonDocument.Parse(line);
            var root = doc.RootElement;
            if (!root.TryGetProperty("event", out var eventProp))
            {
                return;
            }

            var eventName = eventProp.GetString();
            if (string.Equals(eventName, "calib_result", StringComparison.OrdinalIgnoreCase))
            {
                if (root.TryGetProperty("pulses_per_mm", out var ppmProp) && ppmProp.TryGetDouble(out var ppm))
                {
                    _suspendDirtyTracking = true;
                    _pulsesPerMm.Value = ClampDecimal((decimal)ppm, _pulsesPerMm.Minimum, _pulsesPerMm.Maximum);
                    _suspendDirtyTracking = false;
                    _hasUnsentChanges = false;
                    _hasUnsavedChanges = true;
                    UpdateProgramStateUi();
                    AppendLog($"Calibration complete. Updated pulses/mm to {ppm:0.####}");
                }
                return;
            }

            if (string.Equals(eventName, "ack", StringComparison.OrdinalIgnoreCase)
                && root.TryGetProperty("cmd", out var cmdProp)
                && string.Equals(cmdProp.GetString(), "calib_arm", StringComparison.OrdinalIgnoreCase))
            {
                AppendLog("Calibration armed. Measurement starts on photocell falling edge and ends on rising edge.");
            }
        }
        catch
        {
            // Not JSON or partial JSON line; already logged as RX text.
        }
    }

    private void SendSetConfig()
    {
        var payload = new
        {
            cmd = "set_config",
            pulses_per_mm = (double)_pulsesPerMm.Value,
            max_ms_per_mm = (int)_maxMsPerMm.Value,
            photocell_offset_mm = (double)_photocellOffset.Value,
            debounce_ms = (int)_debounceMs.Value
        };

        SendJson(payload);
    }

    private void SendPatterns()
    {
        NormalizeAllPatterns();
        RefreshPatternViolationHighlights();

        if (!ValidateAllPatterns(out var validationError))
        {
            MessageBox.Show(validationError, "Pattern Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog($"VALIDATION ERROR {validationError}");
            return;
        }

        SendSetPattern(1, _gun1Lines);
        SendSetPattern(2, _gun2Lines);

        if (_serial?.IsOpen == true)
        {
            _hasUnsentChanges = false;
            UpdateProgramStateUi();
        }
    }

    private void SendCalibrationArm()
    {
        var paperLength = (double)_calibPaperLength.Value;
        if (paperLength <= 0.001)
        {
            MessageBox.Show("Calibration paper length must be greater than 0.", "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SendJson(new { cmd = "calib_arm", paper_length_mm = paperLength });
        AppendLog("Calibration request sent. Move sheet so photocell goes LOW then HIGH over the measured length.");
    }

    private void SendSetPattern(int gun, IEnumerable<PatternLine> lines)
    {
        var patternLines = lines.Select(l => new { start = l.StartMm, end = l.EndMm }).ToArray();
        var payload = new
        {
            cmd = "set_pattern",
            gun,
            lines = patternLines
        };
        SendJson(payload);
    }

    private void SendJson<T>(T payload)
    {
        var json = JsonSerializer.Serialize(payload);

        if (_serial?.IsOpen != true)
        {
            AppendLog($"TX (not sent, disconnected) {json}");
            return;
        }

        try
        {
            _serial.WriteLine(json);
            AppendLog($"TX {json}");
        }
        catch (Exception ex)
        {
            AppendLog($"TX ERROR {ex.Message}");
        }
    }

    private void SaveProgramToFile()
    {
        NormalizeAllPatterns();
        RefreshPatternViolationHighlights();

        if (!ValidateAllPatterns(out var validationError))
        {
            MessageBox.Show(validationError, "Pattern Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog($"VALIDATION ERROR {validationError}");
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "Glue Program (*.json)|*.json|All files (*.*)|*.*",
            FileName = "glue_program.json"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        var program = new GlueProgramFile
        {
            Config = GetConfigFromUi(),
            Gun1Lines = _gun1Lines.Select(l => l.Clone()).ToList(),
            Gun2Lines = _gun2Lines.Select(l => l.Clone()).ToList()
        };

        var options = new JsonSerializerOptions { WriteIndented = true };
        File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(program, options));
        AppendLog($"Saved program to {dialog.FileName}");
        _hasUnsavedChanges = false;
        UpdateProgramStateUi();
    }

    private void LoadProgramFromFile()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "Glue Program (*.json)|*.json|All files (*.*)|*.*"
        };

        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(dialog.FileName);
            var program = JsonSerializer.Deserialize<GlueProgramFile>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            if (program is null)
            {
                throw new InvalidDataException("File does not contain a valid program.");
            }

            ApplyProgramToUi(program);
            if (!ValidateAllPatterns(out var validationError))
            {
                MessageBox.Show($"Program loaded with validation warnings:{Environment.NewLine}{validationError}", "Pattern Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                AppendLog($"VALIDATION WARNING {validationError}");
            }
            _hasUnsavedChanges = false;
            _hasUnsentChanges = true;
            UpdateProgramStateUi();
            AppendLog($"Loaded program from {dialog.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load program: {ex.Message}", "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private DeviceConfig GetConfigFromUi()
    {
        return new DeviceConfig
        {
            PulsesPerMm = (double)_pulsesPerMm.Value,
            MaxMsPerMm = (int)_maxMsPerMm.Value,
            PhotocellOffsetMm = (double)_photocellOffset.Value,
            DebounceMs = (int)_debounceMs.Value
        };
    }

    private void ApplyProgramToUi(GlueProgramFile program)
    {
        var previousSuspend = _suspendDirtyTracking;
        _suspendDirtyTracking = true;
        if (program.Config is not null)
        {
            _pulsesPerMm.Value = ClampDecimal((decimal)program.Config.PulsesPerMm, _pulsesPerMm.Minimum, _pulsesPerMm.Maximum);
            _maxMsPerMm.Value = ClampDecimal(program.Config.MaxMsPerMm, _maxMsPerMm.Minimum, _maxMsPerMm.Maximum);
            _photocellOffset.Value = ClampDecimal((decimal)program.Config.PhotocellOffsetMm, _photocellOffset.Minimum, _photocellOffset.Maximum);
            _debounceMs.Value = ClampDecimal(program.Config.DebounceMs, _debounceMs.Minimum, _debounceMs.Maximum);
        }

        ReplaceLines(_gun1Lines, program.Gun1Lines);
        ReplaceLines(_gun2Lines, program.Gun2Lines);

        NormalizeAllPatterns();

        RefreshPreviews();
        _suspendDirtyTracking = previousSuspend;
    }

    private static decimal ClampDecimal(decimal value, decimal min, decimal max)
    {
        if (value < min) return min;
        if (value > max) return max;
        return value;
    }

    private static void ReplaceLines(BindingList<PatternLine> target, IList<PatternLine>? source)
    {
        target.Clear();
        if (source is null)
        {
            return;
        }

        foreach (var line in source)
        {
            target.Add(line.Clone());
        }
    }

    private static void RemoveSelectedLine(DataGridView grid, BindingList<PatternLine> lines)
    {
        if (grid.SelectedRows.Count == 0)
        {
            return;
        }

        var index = grid.SelectedRows[0].Index;
        if (index >= 0 && index < lines.Count)
        {
            lines.RemoveAt(index);
        }
    }

    private void RefreshPreviews()
    {
        _gun1Preview.SetLines(_gun1Lines);
        _gun2Preview.SetLines(_gun2Lines);
        RefreshPatternViolationHighlights();
    }

    private void TryAddLine(BindingList<PatternLine> lines, string gunName)
    {
        if (lines.Count >= MaxLinesPerGun)
        {
            var msg = $"{gunName} supports up to {MaxLinesPerGun} lines.";
            MessageBox.Show(msg, "Pattern Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog($"VALIDATION ERROR {msg}");
            return;
        }

        NormalizePattern(lines);

        var lastEnd = 0.0;
        for (var i = 0; i < lines.Count; i++)
        {
            lastEnd = Math.Max(lastEnd, Math.Max(lines[i].StartMm, lines[i].EndMm));
        }

        var start = Math.Min(MaxPatternMm, lastEnd + 1.0);
        if (start >= MaxPatternMm)
        {
            var msg = $"{gunName} cannot add another line because the pattern already reached {MaxPatternMm:0.##} mm.";
            MessageBox.Show(msg, "Pattern Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            AppendLog($"VALIDATION ERROR {msg}");
            return;
        }

        var end = Math.Min(MaxPatternMm, start + 10.0);
        lines.Add(new PatternLine { StartMm = start, EndMm = end });
    }

    private void NormalizeAllPatterns()
    {
        NormalizePattern(_gun1Lines);
        NormalizePattern(_gun2Lines);
        RefreshPreviews();
    }

    private void RefreshPatternViolationHighlights()
    {
        ApplyPatternViolationHighlight(_gun1Grid, _gun1Lines);
        ApplyPatternViolationHighlight(_gun2Grid, _gun2Lines);
    }

    private static void ApplyPatternViolationHighlight(DataGridView grid, IList<PatternLine> lines)
    {
        var invalidRows = ComputeInvalidRows(lines);
        for (var row = 0; row < grid.Rows.Count; row++)
        {
            var rowStyle = grid.Rows[row].DefaultCellStyle;
            var isInvalid = row < invalidRows.Length && invalidRows[row];
            rowStyle.BackColor = isInvalid ? Color.MistyRose : Color.White;
            rowStyle.ForeColor = Color.Black;
        }
    }

    private static bool[] ComputeInvalidRows(IList<PatternLine> lines)
    {
        var invalid = new bool[lines.Count];
        var ranges = new List<(double Lo, double Hi, int Row)>();

        for (var i = 0; i < lines.Count; i++)
        {
            var a = lines[i].StartMm;
            var b = lines[i].EndMm;
            if (double.IsNaN(a) || double.IsInfinity(a) || double.IsNaN(b) || double.IsInfinity(b))
            {
                invalid[i] = true;
                continue;
            }

            var lo = Math.Min(a, b);
            var hi = Math.Max(a, b);
            if (lo < MinPatternMm || hi > MaxPatternMm)
            {
                invalid[i] = true;
                continue;
            }

            ranges.Add((lo, hi, i));
        }

        ranges.Sort((x, y) => x.Lo.CompareTo(y.Lo));
        for (var i = 1; i < ranges.Count; i++)
        {
            if (ranges[i].Lo < ranges[i - 1].Hi)
            {
                invalid[ranges[i].Row] = true;
                invalid[ranges[i - 1].Row] = true;
            }
        }

        return invalid;
    }

    private void MarkConfigChanged()
    {
        if (_suspendDirtyTracking)
        {
            return;
        }

        _hasUnsavedChanges = true;
        _hasUnsentChanges = true;
        UpdateProgramStateUi();
    }

    private void MarkPatternChanged()
    {
        if (_suspendDirtyTracking)
        {
            return;
        }

        _hasUnsavedChanges = true;
        _hasUnsentChanges = true;
        UpdateProgramStateUi();
    }

    private void UpdateProgramStateUi()
    {
        if (_hasUnsavedChanges || _hasUnsentChanges)
        {
            var parts = new List<string>();
            if (_hasUnsentChanges) parts.Add("not sent");
            if (_hasUnsavedChanges) parts.Add("not saved");
            _programStateLabel.Text = "Modified (" + string.Join(", ", parts) + ")";
            _programStateLabel.ForeColor = Color.DarkOrange;
            return;
        }

        _programStateLabel.Text = "Synced";
        _programStateLabel.ForeColor = Color.ForestGreen;
    }

    private GlueProgramFile BuildProgramFromUi()
    {
        return new GlueProgramFile
        {
            Config = GetConfigFromUi(),
            Gun1Lines = _gun1Lines.Select(l => l.Clone()).ToList(),
            Gun2Lines = _gun2Lines.Select(l => l.Clone()).ToList()
        };
    }

    private void SaveLastAppState()
    {
        try
        {
            var state = new AppStateFile
            {
                Program = BuildProgramFromUi(),
                LastPort = _portCombo.SelectedItem?.ToString(),
                CalibrationPaperLengthMm = (double)_calibPaperLength.Value,
                HasUnsentChanges = _hasUnsentChanges,
                HasUnsavedChanges = _hasUnsavedChanges
            };

            Directory.CreateDirectory(AppStateDirectory);
            var options = new JsonSerializerOptions { WriteIndented = true };
            File.WriteAllText(AppStateFilePath, JsonSerializer.Serialize(state, options));
        }
        catch (Exception ex)
        {
            AppendLog($"State save warning: {ex.Message}");
        }
    }

    private void LoadLastAppState()
    {
        if (!File.Exists(AppStateFilePath))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(AppStateFilePath);
            var state = JsonSerializer.Deserialize<AppStateFile>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (state is null)
            {
                return;
            }

            _suspendDirtyTracking = true;

            if (state.Program is not null)
            {
                ApplyProgramToUi(state.Program);
            }

            _calibPaperLength.Value = ClampDecimal((decimal)state.CalibrationPaperLengthMm, _calibPaperLength.Minimum, _calibPaperLength.Maximum);

            if (!string.IsNullOrWhiteSpace(state.LastPort) && _portCombo.Items.Contains(state.LastPort))
            {
                _portCombo.SelectedItem = state.LastPort;
            }

            _hasUnsentChanges = state.HasUnsentChanges;
            _hasUnsavedChanges = state.HasUnsavedChanges;
            _suspendDirtyTracking = false;
            UpdateProgramStateUi();
            RefreshPreviews();
            AppendLog("Restored previous session state.");
        }
        catch (Exception ex)
        {
            _suspendDirtyTracking = false;
            AppendLog($"State load warning: {ex.Message}");
        }
    }

    private static void NormalizePattern(BindingList<PatternLine> lines)
    {
        var normalized = lines
            .Select(l => new PatternLine
            {
                StartMm = Math.Min(l.StartMm, l.EndMm),
                EndMm = Math.Max(l.StartMm, l.EndMm)
            })
            .OrderBy(l => l.StartMm)
            .ThenBy(l => l.EndMm)
            .ToList();

        lines.Clear();
        foreach (var line in normalized)
        {
            lines.Add(line);
        }
    }

    private bool ValidateAllPatterns(out string error)
    {
        if (!ValidatePattern("Gun 1", _gun1Lines, out error))
        {
            return false;
        }

        if (!ValidatePattern("Gun 2", _gun2Lines, out error))
        {
            return false;
        }

        error = string.Empty;
        return true;
    }

    private static bool ValidatePattern(string gunName, IList<PatternLine> lines, out string error)
    {
        if (lines.Count > MaxLinesPerGun)
        {
            error = $"{gunName}: line count ({lines.Count}) exceeds limit of {MaxLinesPerGun}.";
            return false;
        }

        var ranges = new List<(double Lo, double Hi)>();

        for (var i = 0; i < lines.Count; i++)
        {
            var line = lines[i];
            var a = line.StartMm;
            var b = line.EndMm;

            if (double.IsNaN(a) || double.IsInfinity(a) || double.IsNaN(b) || double.IsInfinity(b))
            {
                error = $"{gunName}: line {i + 1} contains an invalid numeric value.";
                return false;
            }

            var lo = Math.Min(a, b);
            var hi = Math.Max(a, b);

            if (lo < MinPatternMm || hi > MaxPatternMm)
            {
                error = $"{gunName}: line {i + 1} must stay within {MinPatternMm:0.##}..{MaxPatternMm:0.##} mm.";
                return false;
            }

            ranges.Add((lo, hi));
        }

        ranges.Sort((x, y) => x.Lo.CompareTo(y.Lo));
        for (var i = 1; i < ranges.Count; i++)
        {
            if (ranges[i].Lo < ranges[i - 1].Hi)
            {
                error = $"{gunName}: lines overlap. Adjust intervals so they do not overlap.";
                return false;
            }
        }

        error = string.Empty;
        return true;
    }

    private void AppendLog(string text)
    {
        _logBox.AppendText($"{DateTime.Now:HH:mm:ss} | {text}{Environment.NewLine}");
    }
}

public sealed class PatternPreviewPanel : Panel
{
    private readonly List<PatternLine> _lines = new();

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Color PatternColor { get; set; } = Color.DarkOrange;

    public void SetLines(IEnumerable<PatternLine> lines)
    {
        _lines.Clear();
        _lines.AddRange(lines.Select(l => l.Clone()));
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);

        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        var rect = ClientRectangle;
        if (rect.Width < 40 || rect.Height < 30)
        {
            return;
        }

        using var bg = new SolidBrush(Color.FromArgb(245, 245, 245));
        e.Graphics.FillRectangle(bg, rect);

        var pad = 10;
        var chart = new Rectangle(rect.X + pad, rect.Y + pad, rect.Width - (pad * 2), rect.Height - (pad * 2));

        using var borderPen = new Pen(Color.DarkGray, 1);
        e.Graphics.DrawRectangle(borderPen, chart);

        var max = 100.0;
        foreach (var line in _lines)
        {
            max = Math.Max(max, Math.Max(line.StartMm, line.EndMm));
        }
        if (max < 1.0)
        {
            max = 1.0;
        }

        using var fill = new SolidBrush(Color.FromArgb(180, PatternColor));
        foreach (var line in _lines)
        {
            var lo = Math.Min(line.StartMm, line.EndMm);
            var hi = Math.Max(line.StartMm, line.EndMm);

            var x1 = chart.Left + (float)(lo / max) * chart.Width;
            var x2 = chart.Left + (float)(hi / max) * chart.Width;
            var w = Math.Max(2, x2 - x1);

            var glueRect = new RectangleF(x1, chart.Top + 6, w, chart.Height - 12);
            e.Graphics.FillRectangle(fill, glueRect);
        }

        using var font = new Font("Segoe UI", 8F);
        using var textBrush = new SolidBrush(Color.Black);
        e.Graphics.DrawString($"0 mm", font, textBrush, chart.Left, chart.Bottom - 14);
        e.Graphics.DrawString($"{max:0.#} mm", font, textBrush, chart.Right - 60, chart.Bottom - 14);
    }
}

public sealed class PatternLine
{
    public double StartMm { get; set; }
    public double EndMm { get; set; }

    public PatternLine Clone() => new() { StartMm = StartMm, EndMm = EndMm };
}

public sealed class DeviceConfig
{
    public double PulsesPerMm { get; set; } = 10.0;
    public int MaxMsPerMm { get; set; } = 100;
    public double PhotocellOffsetMm { get; set; }
    public int DebounceMs { get; set; } = 20;
}

public sealed class GlueProgramFile
{
    public DeviceConfig? Config { get; set; } = new();
    public List<PatternLine> Gun1Lines { get; set; } = new();
    public List<PatternLine> Gun2Lines { get; set; } = new();
}

public sealed class AppStateFile
{
    public GlueProgramFile? Program { get; set; } = new();
    public string? LastPort { get; set; }
    public double CalibrationPaperLengthMm { get; set; } = 297.0;
    public bool HasUnsentChanges { get; set; }
    public bool HasUnsavedChanges { get; set; }
}
