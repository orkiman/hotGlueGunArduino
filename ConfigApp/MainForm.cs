using System.ComponentModel;
using System.Globalization;
using System.IO.Ports;
using System.Text.Json;
using System.Text;

namespace GlueConfigApp;

public sealed class MainForm : Form
{
    private const int MaxLinesPerGun = 32;
    private const double MinPatternMm = 0.0;
    private const double MaxPatternMm = 100000.0;

    // ── Connection bar ──
    private readonly ComboBox _portCombo = new() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Button _refreshPortsButton = new() { Text = "Refresh", Width = 70 };
    private readonly Button _connectButton = new() { Text = "Connect", Width = 90 };
    private readonly Label _statusLabel = new() { Text = "Disconnected", AutoSize = true, ForeColor = Color.Gray };
    private readonly Label _activeIndicator = new()
    {
        Text = "  INACTIVE  ",
        AutoSize = true,
        Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        ForeColor = Color.White,
        BackColor = Color.FromArgb(200, 60, 60),
        Padding = new Padding(6, 3, 6, 3),
        Margin = new Padding(6, 2, 0, 0)
    };

    // ── Program selector ──
    private readonly ComboBox _programCombo = new() { Width = 200, DropDownStyle = ComboBoxStyle.DropDown };
    private readonly Button _deleteProgramButton = new() { Text = "Delete", Width = 55 };

    // ── Config controls ──
    private readonly NumericUpDown _pulsesPerMm = new() { DecimalPlaces = 4, Minimum = 0.0001M, Maximum = 100000M, Value = 10M, Increment = 0.1M, Width = 110 };
    private readonly NumericUpDown _maxMsPerMm = new() { DecimalPlaces = 0, Minimum = 1, Maximum = 60000, Value = 100, Width = 110 };
    private readonly NumericUpDown _photocellOffset = new() { DecimalPlaces = 2, Minimum = -5000, Maximum = 5000, Value = 0, Increment = 1, Width = 110 };
    private readonly NumericUpDown _debounceMs = new() { DecimalPlaces = 0, Minimum = 0, Maximum = 1000, Value = 20, Width = 110 };

    // ── Calibration ──
    private readonly NumericUpDown _calibPaperLength = new() { DecimalPlaces = 2, Minimum = 1, Maximum = 10000, Value = 297, Increment = 1, Width = 110 };
    private readonly Button _calibArmButton = new() { Text = "Calibrate Enc" };

    // ── Machine control ──
    private readonly Button _activateButton = new() { Text = "Activate", BackColor = Color.FromArgb(220, 255, 220) };
    private readonly Button _deactivateButton = new() { Text = "Deactivate", BackColor = Color.FromArgb(255, 220, 220) };
    private readonly Button _testOpenGun1Button = new() { Text = "Test 1" };
    private readonly Button _testCloseGun1Button = new() { Text = "Close 1" };
    private readonly Button _testOpenGun2Button = new() { Text = "Test 2" };
    private readonly Button _testCloseGun2Button = new() { Text = "Close 2" };
    private readonly Button _testOpenBothButton = new() { Text = "Test Both" };
    private readonly Button _testCloseBothButton = new() { Text = "Close Both" };
    private readonly Button _swTriggerButton = new() { Text = "SW Trigger", BackColor = Color.FromArgb(220, 230, 255) };

    // ── Pattern grids ──
    private readonly DataGridView _gun1Grid = new();
    private readonly DataGridView _gun2Grid = new();
    private readonly Button _gun1AddLineButton = new() { Text = "+", Width = 36 };
    private readonly Button _gun1RemoveLineButton = new() { Text = "−", Width = 36 };
    private readonly Button _gun2AddLineButton = new() { Text = "+", Width = 36 };
    private readonly Button _gun2RemoveLineButton = new() { Text = "−", Width = 36 };
    private readonly PatternPreviewPanel _gun1Preview = new() { PatternColor = Color.OrangeRed };
    private readonly PatternPreviewPanel _gun2Preview = new() { PatternColor = Color.SteelBlue };
    private readonly BindingList<PatternLine> _gun1Lines = new();
    private readonly BindingList<PatternLine> _gun2Lines = new();

    // ── Log ──
    private readonly TextBox _logBox = new()
    {
        Multiline = true,
        ScrollBars = ScrollBars.Vertical,
        ReadOnly = true,
        Dock = DockStyle.Fill,
        Font = new Font("Consolas", 8.5F),
        BackColor = Color.FromArgb(30, 30, 30),
        ForeColor = Color.FromArgb(200, 220, 200)
    };

    // ── State ──
    private SerialPort? _serial;
    private readonly StringBuilder _serialRxBuffer = new();
    private readonly System.Windows.Forms.Timer _serialPollTimer = new() { Interval = 100 };
    private readonly System.Windows.Forms.Timer _autoSendTimer = new() { Interval = 600 };
    private bool _hasUnsavedChanges;
    private bool _hasUnsentChanges;
    private bool _suspendDirtyTracking;

    private static readonly string AppStateDirectory = Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory,
        "Config");
    private static readonly string ProgramsDirectory = Path.Combine(AppStateDirectory, "Programs");
    private static readonly string AppStateFilePath = Path.Combine(AppStateDirectory, "last_state.json");

    public MainForm()
    {
        Text = "Hot Glue Gun Controller";
        Width = 1280;
        Height = 800;
        WindowState = FormWindowState.Maximized;
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9F);
        BackColor = Color.FromArgb(245, 246, 250);

        Directory.CreateDirectory(ProgramsDirectory);

        BuildLayout();
        WireEvents();
        ConfigurePatternGrid(_gun1Grid, _gun1Lines);
        ConfigurePatternGrid(_gun2Grid, _gun2Lines);

        _gun1Lines.ListChanged += (_, _) => { RefreshPreviews(); MarkPatternChanged(); };
        _gun2Lines.ListChanged += (_, _) => { RefreshPreviews(); MarkPatternChanged(); };

        _pulsesPerMm.ValueChanged += (_, _) => MarkConfigChanged();
        _maxMsPerMm.ValueChanged += (_, _) => MarkConfigChanged();
        _photocellOffset.ValueChanged += (_, _) => MarkConfigChanged();
        _debounceMs.ValueChanged += (_, _) => MarkConfigChanged();
        _calibPaperLength.ValueChanged += (_, _) => { RefreshPreviews(); MarkConfigChanged(); };

        RefreshSerialPorts();
        RefreshProgramList();
        RefreshPreviews();
        LoadLastAppState();
        UpdateProgramStateUi();

        Shown += (_, _) => TryAutoConnectOnStartup();
    }

    // ═══════════════════════════════════════════════════════════════
    //  LAYOUT
    // ═══════════════════════════════════════════════════════════════

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(6)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // toolbar
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // config bar
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // main area
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 160));  // log

        root.Controls.Add(BuildToolbar(), 0, 0);
        root.Controls.Add(BuildConfigBar(), 0, 1);
        root.Controls.Add(BuildMainArea(), 0, 2);
        root.Controls.Add(BuildLogPanel(), 0, 3);

        Controls.Add(root);
    }

    private Control BuildToolbar()
    {
        var bar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = false,
            Padding = new Padding(0, 2, 0, 4)
        };

        bar.Controls.Add(MakeLabel("Port:"));
        bar.Controls.Add(_portCombo);
        bar.Controls.Add(_refreshPortsButton);
        bar.Controls.Add(_connectButton);
        bar.Controls.Add(MakeSeparator());
        bar.Controls.Add(MakeLabel("Program:"));
        bar.Controls.Add(_programCombo);
        bar.Controls.Add(_deleteProgramButton);
        bar.Controls.Add(MakeSeparator());
        bar.Controls.Add(_statusLabel);
        bar.Controls.Add(MakeSeparator());
        bar.Controls.Add(_activeIndicator);
        return bar;
    }

    private Control BuildConfigBar()
    {
        var bar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true,
            Padding = new Padding(0, 0, 0, 2),
            BackColor = Color.FromArgb(235, 237, 242)
        };

        bar.Controls.Add(MakeLabel("Pulses/mm:"));
        bar.Controls.Add(_pulsesPerMm);
        bar.Controls.Add(MakeLabel("Max ms/mm:"));
        bar.Controls.Add(_maxMsPerMm);
        bar.Controls.Add(MakeLabel("Photocell offset (mm):"));
        bar.Controls.Add(_photocellOffset);
        bar.Controls.Add(MakeLabel("Debounce (ms):"));
        bar.Controls.Add(_debounceMs);
        bar.Controls.Add(MakeSeparator());
        bar.Controls.Add(MakeLabel("Calib length (mm):"));
        bar.Controls.Add(_calibPaperLength);
        bar.Controls.Add(_calibArmButton);

        return bar;
    }

    private Control BuildMainArea()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

        layout.Controls.Add(BuildGunPanel("Gun 1", _gun1Grid, _gun1Preview, _gun1AddLineButton, _gun1RemoveLineButton), 0, 0);
        layout.Controls.Add(BuildGunPanel("Gun 2", _gun2Grid, _gun2Preview, _gun2AddLineButton, _gun2RemoveLineButton), 0, 1);

        var controlPanel = BuildControlPanel();
        layout.Controls.Add(controlPanel, 1, 0);
        layout.SetRowSpan(controlPanel, 2);

        return layout;
    }

    private Control BuildGunPanel(string title, DataGridView grid, PatternPreviewPanel preview, Button addBtn, Button removeBtn)
    {
        var group = new GroupBox { Text = title, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(60, 60, 60) };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.FixedSingle;

        preview.Dock = DockStyle.Fill;
        preview.MinimumSize = new Size(100, 60);

        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        btnPanel.Controls.Add(addBtn);
        btnPanel.Controls.Add(removeBtn);

        layout.Controls.Add(grid, 0, 0);
        layout.Controls.Add(preview, 1, 0);
        layout.Controls.Add(btnPanel, 0, 1);
        layout.SetColumnSpan(btnPanel, 2);

        group.Controls.Add(layout);
        return group;
    }

    private Control BuildControlPanel()
    {
        var group = new GroupBox { Text = "Machine Control", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(60, 60, 60) };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 8,
            Padding = new Padding(6)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        for (int i = 0; i < 8; i++)
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _activateButton.Dock = DockStyle.Fill;
        _deactivateButton.Dock = DockStyle.Fill;
        _testOpenGun1Button.Dock = DockStyle.Fill;
        _testCloseGun1Button.Dock = DockStyle.Fill;
        _testOpenGun2Button.Dock = DockStyle.Fill;
        _testCloseGun2Button.Dock = DockStyle.Fill;
        _testOpenBothButton.Dock = DockStyle.Fill;
        _testCloseBothButton.Dock = DockStyle.Fill;

        layout.Controls.Add(_activateButton, 0, 0);
        layout.Controls.Add(_deactivateButton, 1, 0);
        layout.Controls.Add(new Label { Text = " ", AutoSize = true }, 0, 1);
        layout.Controls.Add(_testOpenGun1Button, 0, 2);
        layout.Controls.Add(_testCloseGun1Button, 1, 2);
        layout.Controls.Add(_testOpenGun2Button, 0, 3);
        layout.Controls.Add(_testCloseGun2Button, 1, 3);
        layout.Controls.Add(_testOpenBothButton, 0, 4);
        layout.Controls.Add(_testCloseBothButton, 1, 4);
        layout.Controls.Add(new Label { Text = " ", AutoSize = true }, 0, 5);
        _swTriggerButton.Dock = DockStyle.Fill;
        layout.Controls.Add(_swTriggerButton, 0, 6);
        layout.SetColumnSpan(_swTriggerButton, 2);

        group.Controls.Add(layout);
        return group;
    }

    private Control BuildLogPanel()
    {
        var group = new GroupBox { Text = "Serial Log", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(60, 60, 60) };
        group.Controls.Add(_logBox);
        return group;
    }

    private static Label MakeLabel(string text) =>
        new() { Text = text, AutoSize = true, Margin = new Padding(6, 7, 2, 0) };

    private static Label MakeSeparator() =>
        new() { Text = "|", AutoSize = true, ForeColor = Color.LightGray, Margin = new Padding(8, 7, 8, 0) };

    // ═══════════════════════════════════════════════════════════════
    //  GRID SETUP
    // ═══════════════════════════════════════════════════════════════

    private static void ConfigurePatternGrid(DataGridView grid, BindingList<PatternLine> source)
    {
        grid.AutoGenerateColumns = false;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.MultiSelect = false;
        grid.RowHeadersVisible = false;
        grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(230, 232, 238);
        grid.EnableHeadersVisualStyles = false;
        grid.DefaultCellStyle.Font = new Font("Segoe UI", 9F);

        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "Start (mm)",
            DataPropertyName = nameof(PatternLine.StartMm),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });
        grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            HeaderText = "End (mm)",
            DataPropertyName = nameof(PatternLine.EndMm),
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        });
        grid.DataSource = source;
    }

    // ═══════════════════════════════════════════════════════════════
    //  EVENTS
    // ═══════════════════════════════════════════════════════════════

    private void WireEvents()
    {
        _refreshPortsButton.Click += (_, _) => RefreshSerialPorts();
        _connectButton.Click += (_, _) => ToggleConnection();

        _activateButton.Click += (_, _) => { SendJson(new { cmd = "set_active", active = true }); SetActiveIndicator(true); };
        _deactivateButton.Click += (_, _) => { SendJson(new { cmd = "set_active", active = false }); SetActiveIndicator(false); };

        _testOpenGun1Button.Click += (_, _) => SendJson(new { cmd = "test_open", gun = 1, timeout_ms = 30000 });
        _testCloseGun1Button.Click += (_, _) => SendJson(new { cmd = "test_close", gun = 1 });
        _testOpenGun2Button.Click += (_, _) => SendJson(new { cmd = "test_open", gun = 2, timeout_ms = 30000 });
        _testCloseGun2Button.Click += (_, _) => SendJson(new { cmd = "test_close", gun = 2 });
        _testOpenBothButton.Click += (_, _) => SendJson(new { cmd = "test_open", gun = "both", timeout_ms = 30000 });
        _testCloseBothButton.Click += (_, _) => SendJson(new { cmd = "test_close", gun = "both" });
        _calibArmButton.Click += (_, _) => SendCalibrationArm();

        _deleteProgramButton.Click += (_, _) => DeleteProgram();
        _swTriggerButton.Click += (_, _) => SendJson(new { cmd = "sw_trigger" });
        _programCombo.SelectedIndexChanged += (_, _) => LoadSelectedProgram();

        _gun1AddLineButton.Click += (_, _) => TryAddLine(_gun1Lines, "Gun 1");
        _gun2AddLineButton.Click += (_, _) => TryAddLine(_gun2Lines, "Gun 2");
        _gun1RemoveLineButton.Click += (_, _) => RemoveSelectedLine(_gun1Grid, _gun1Lines);
        _gun2RemoveLineButton.Click += (_, _) => RemoveSelectedLine(_gun2Grid, _gun2Lines);

        _gun1Grid.CellEndEdit += (_, _) => { SortPattern(_gun1Lines); RefreshPreviews(); MarkPatternChanged(); };
        _gun2Grid.CellEndEdit += (_, _) => { SortPattern(_gun2Lines); RefreshPreviews(); MarkPatternChanged(); };
        _gun1Grid.CellValueChanged += (_, _) => { RefreshPreviews(); MarkPatternChanged(); };
        _gun2Grid.CellValueChanged += (_, _) => { RefreshPreviews(); MarkPatternChanged(); };
        _gun1Grid.DataError += (_, e) => e.ThrowException = false;
        _gun2Grid.DataError += (_, e) => e.ThrowException = false;

        _serialPollTimer.Tick += (_, _) => PollSerial();
        _autoSendTimer.Tick += (_, _) => AutoSendIfValid();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        base.OnFormClosing(e);
        SaveLastAppState();
        DisconnectSerial();
    }

    // ═══════════════════════════════════════════════════════════════
    //  SERIAL
    // ═══════════════════════════════════════════════════════════════

    private void PollSerial()
    {
        try
        {
            if (_serial?.IsOpen != true) return;
            var text = _serial.ReadExisting();
            if (!string.IsNullOrEmpty(text)) AppendSerialText(text);
        }
        catch { }
    }

    private void RefreshSerialPorts()
    {
        var current = _portCombo.SelectedItem?.ToString();
        var ports = SerialPort.GetPortNames().OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToArray();
        _portCombo.Items.Clear();
        _portCombo.Items.AddRange(ports);
        if (!string.IsNullOrWhiteSpace(current) && ports.Contains(current, StringComparer.OrdinalIgnoreCase))
            _portCombo.SelectedItem = current;
        else if (_portCombo.Items.Count > 0)
            _portCombo.SelectedIndex = 0;
    }

    private void ToggleConnection()
    {
        if (_serial?.IsOpen == true) { DisconnectSerial(); return; }
        if (_portCombo.SelectedItem is not string port || string.IsNullOrWhiteSpace(port))
        {
            MessageBox.Show("Select a COM port.", "Missing Port", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        ConnectSerial(port, showErrors: true);
    }

    private bool ConnectSerial(string port, bool showErrors)
    {
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
            _statusLabel.ForeColor = Color.Green;
            AppendLog($"Connected to {port}");
            return true;
        }
        catch (Exception ex)
        {
            if (showErrors)
                MessageBox.Show($"Could not open serial port: {ex.Message}", "Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                AppendLog($"Auto-connect failed on {port}: {ex.Message}");
            DisconnectSerial();
            return false;
        }
    }

    private void TryAutoConnectOnStartup()
    {
        if (_serial?.IsOpen == true) return;
        RefreshSerialPorts();
        if (_portCombo.SelectedItem is not string port || string.IsNullOrWhiteSpace(port))
        {
            if (_portCombo.Items.Count == 0) { AppendLog("Auto-connect skipped: no COM ports."); return; }
            _portCombo.SelectedIndex = 0;
            port = _portCombo.SelectedItem?.ToString() ?? string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(port)) _ = ConnectSerial(port, showErrors: false);
    }

    private void DisconnectSerial()
    {
        if (_serial is null) { _connectButton.Text = "Connect"; _statusLabel.Text = "Disconnected"; _statusLabel.ForeColor = Color.Gray; return; }
        try
        {
            _serialPollTimer.Stop();
            _serial.DataReceived -= SerialDataReceived;
            if (_serial.IsOpen) _serial.Close();
            _serial.Dispose();
        }
        catch { }
        finally
        {
            _serial = null;
            _connectButton.Text = "Connect";
            _statusLabel.Text = "Disconnected";
            _statusLabel.ForeColor = Color.Gray;
            AppendLog("Disconnected");
        }
    }

    private void SerialDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        try
        {
            var text = _serial?.ReadExisting();
            if (!string.IsNullOrEmpty(text)) BeginInvoke(() => AppendSerialText(text));
        }
        catch { }
    }

    private void AppendSerialText(string text)
    {
        _serialRxBuffer.Append(text);
        while (true)
        {
            var current = _serialRxBuffer.ToString();
            var idx = current.IndexOf('\n');
            if (idx < 0) break;
            var line = current.Substring(0, idx).TrimEnd('\r');
            _serialRxBuffer.Remove(0, idx + 1);
            if (line.Length > 0) { AppendLog($"RX {line}"); HandleIncomingLine(line); }
        }
    }

    private void HandleIncomingLine(string line)
    {
        foreach (var frag in ExtractJsonObjects(line))
        {
            try
            {
                using var doc = JsonDocument.Parse(frag);
                var root = doc.RootElement;
                if (!root.TryGetProperty("event", out var eventProp)) continue;
                var ev = eventProp.GetString();

                if (string.Equals(ev, "calib_result", StringComparison.OrdinalIgnoreCase))
                {
                    if (TryGetJsonDouble(root, "pulses_per_mm", out var ppm))
                    {
                        _suspendDirtyTracking = true;
                        _pulsesPerMm.Value = ClampDecimal((decimal)ppm, _pulsesPerMm.Minimum, _pulsesPerMm.Maximum);
                        _suspendDirtyTracking = false;
                        _hasUnsentChanges = false;
                        _hasUnsavedChanges = true;
                        UpdateProgramStateUi();
                        AppendLog($"Calibration complete. pulses/mm = {ppm:0.####}");
                    }
                    continue;
                }

                if (string.Equals(ev, "ack", StringComparison.OrdinalIgnoreCase)
                    && root.TryGetProperty("cmd", out var cmdProp))
                {
                    var cmdStr = cmdProp.GetString();
                    if (string.Equals(cmdStr, "calib_arm", StringComparison.OrdinalIgnoreCase))
                        AppendLog("Calibration armed. Starts on photocell falling edge, ends on rising edge.");
                    if (string.Equals(cmdStr, "set_active", StringComparison.OrdinalIgnoreCase))
                    {
                        // Arduino confirmed; we already set indicator optimistically on click
                    }
                }
            }
            catch { }
        }
    }

    private static IEnumerable<string> ExtractJsonObjects(string text)
    {
        int i = 0;
        while (i < text.Length)
        {
            int start = text.IndexOf('{', i);
            if (start < 0) yield break;
            int depth = 0;
            for (int j = start; j < text.Length; j++)
            {
                if (text[j] == '{') depth++;
                else if (text[j] == '}') depth--;
                if (depth == 0) { yield return text.Substring(start, j - start + 1); i = j + 1; goto next; }
            }
            yield break;
            next:;
        }
    }

    private static bool TryGetJsonDouble(JsonElement root, string prop, out double value)
    {
        value = 0;
        if (!root.TryGetProperty(prop, out var p)) return false;
        if (p.ValueKind == JsonValueKind.Number) return p.TryGetDouble(out value);
        if (p.ValueKind == JsonValueKind.String)
            return double.TryParse(p.GetString(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value);
        return false;
    }

    // ═══════════════════════════════════════════════════════════════
    //  SEND / AUTO-SEND
    // ═══════════════════════════════════════════════════════════════

    private void AutoSendIfValid()
    {
        _autoSendTimer.Stop();

        // Auto-save to current program name
        AutoSaveProgram();

        if (!_hasUnsentChanges || _serial?.IsOpen != true) return;
        if (!ValidateAllPatterns(out _)) return;

        SendSetConfig();
        SendSetPattern(1, _gun1Lines);
        SendSetPattern(2, _gun2Lines);
        _hasUnsentChanges = false;
        UpdateProgramStateUi();
        AppendLog("Auto-sent config + patterns.");
    }

    private void AutoSaveProgram()
    {
        var name = _programCombo.Text.Trim();
        if (string.IsNullOrWhiteSpace(name)) return;
        if (!_hasUnsavedChanges) return;
        if (!ValidateAllPatterns(out _)) return;

        try
        {
            var path = Path.Combine(ProgramsDirectory, name + ".json");
            var program = BuildProgramFromUi();
            File.WriteAllText(path, JsonSerializer.Serialize(program, new JsonSerializerOptions { WriteIndented = true }));
            _hasUnsavedChanges = false;
            UpdateProgramStateUi();
            RefreshProgramList();
            _programCombo.Text = name;
        }
        catch (Exception ex) { AppendLog($"Auto-save warning: {ex.Message}"); }
    }

    private void SendSetConfig()
    {
        SendJson(new
        {
            cmd = "set_config",
            pulses_per_mm = (double)_pulsesPerMm.Value,
            max_ms_per_mm = (int)_maxMsPerMm.Value,
            photocell_offset_mm = (double)_photocellOffset.Value,
            debounce_ms = (int)_debounceMs.Value
        });
    }

    private void SendCalibrationArm()
    {
        var len = (double)_calibPaperLength.Value;
        if (len <= 0.001) { MessageBox.Show("Paper length must be > 0.", "Calibration", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        SendJson(new { cmd = "calib_arm", paper_length_mm = len });
        AppendLog("Calibration request sent. Move sheet so photocell goes LOW then HIGH.");
    }

    private void SendSetPattern(int gun, IEnumerable<PatternLine> lines)
    {
        SendJson(new { cmd = "set_pattern", gun, lines = lines.Select(l => new { start = l.StartMm, end = l.EndMm }).ToArray() });
    }

    private void SendJson<T>(T payload)
    {
        var json = JsonSerializer.Serialize(payload);
        if (_serial?.IsOpen != true) { AppendLog($"TX (offline) {json}"); return; }
        try { _serial.WriteLine(json); AppendLog($"TX {json}"); }
        catch (Exception ex) { AppendLog($"TX ERROR {ex.Message}"); }
    }

    // ═══════════════════════════════════════════════════════════════
    //  PROGRAMS (save / load / delete via dropdown)
    // ═══════════════════════════════════════════════════════════════

    private void RefreshProgramList()
    {
        _suspendDirtyTracking = true;
        var current = _programCombo.Text;
        _programCombo.Items.Clear();
        if (Directory.Exists(ProgramsDirectory))
        {
            foreach (var f in Directory.GetFiles(ProgramsDirectory, "*.json").OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
                _programCombo.Items.Add(Path.GetFileNameWithoutExtension(f));
        }
        _programCombo.Text = current;
        _suspendDirtyTracking = false;
    }

    private void DeleteProgram()
    {
        var name = _programCombo.Text.Trim();
        if (string.IsNullOrWhiteSpace(name)) return;
        var path = Path.Combine(ProgramsDirectory, name + ".json");
        if (!File.Exists(path)) { MessageBox.Show("Program not found.", "Delete", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        if (MessageBox.Show($"Delete program \"{name}\"?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        File.Delete(path);
        RefreshProgramList();
        _programCombo.Text = "";
        AppendLog($"Deleted program \"{name}\".");
    }

    private void LoadSelectedProgram()
    {
        if (_suspendDirtyTracking) return;
        var name = _programCombo.SelectedItem?.ToString();
        if (string.IsNullOrWhiteSpace(name)) return;
        var path = Path.Combine(ProgramsDirectory, name + ".json");
        if (!File.Exists(path)) return;

        try
        {
            var json = File.ReadAllText(path);
            var program = JsonSerializer.Deserialize<GlueProgramFile>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (program is null) return;
            ApplyProgramToUi(program);
            _hasUnsavedChanges = false;
            _hasUnsentChanges = true;
            UpdateProgramStateUi();
            ScheduleAutoSend();
            AppendLog($"Loaded program \"{name}\".");
        }
        catch (Exception ex) { AppendLog($"Load error: {ex.Message}"); }
    }

    // ═══════════════════════════════════════════════════════════════
    //  PATTERN EDITING
    // ═══════════════════════════════════════════════════════════════

    private void TryAddLine(BindingList<PatternLine> lines, string gunName)
    {
        if (lines.Count >= MaxLinesPerGun)
        {
            MessageBox.Show($"{gunName}: max {MaxLinesPerGun} lines.", "Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var lastEnd = 0.0;
        foreach (var l in lines) lastEnd = Math.Max(lastEnd, Math.Max(l.StartMm, l.EndMm));
        var start = Math.Min(MaxPatternMm, lastEnd + 1.0);
        if (start >= MaxPatternMm) { MessageBox.Show($"{gunName}: pattern reached max.", "Limit", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        lines.Add(new PatternLine { StartMm = start, EndMm = Math.Min(MaxPatternMm, start + 10.0) });
    }

    private static void RemoveSelectedLine(DataGridView grid, BindingList<PatternLine> lines)
    {
        if (grid.SelectedRows.Count == 0) return;
        var idx = grid.SelectedRows[0].Index;
        if (idx >= 0 && idx < lines.Count) lines.RemoveAt(idx);
    }

    private static void SortPattern(BindingList<PatternLine> lines)
    {
        var sorted = lines.OrderBy(l => l.StartMm).ThenBy(l => l.EndMm).ToList();
        bool changed = false;
        for (int i = 0; i < sorted.Count; i++)
        {
            if (lines[i].StartMm != sorted[i].StartMm || lines[i].EndMm != sorted[i].EndMm) { changed = true; break; }
        }
        if (!changed) return;
        lines.RaiseListChangedEvents = false;
        lines.Clear();
        foreach (var l in sorted) lines.Add(l);
        lines.RaiseListChangedEvents = true;
        lines.ResetBindings();
    }

    // ═══════════════════════════════════════════════════════════════
    //  VALIDATION & HIGHLIGHTS
    // ═══════════════════════════════════════════════════════════════

    private void RefreshPreviews()
    {
        var paperLen = (double)_calibPaperLength.Value;
        _gun1Preview.SetLines(_gun1Lines, paperLen);
        _gun2Preview.SetLines(_gun2Lines, paperLen);
        RefreshPatternViolationHighlights();
    }

    private void RefreshPatternViolationHighlights()
    {
        ApplyHighlight(_gun1Grid, _gun1Lines);
        ApplyHighlight(_gun2Grid, _gun2Lines);
    }

    private static void ApplyHighlight(DataGridView grid, IList<PatternLine> lines)
    {
        var inv = ComputeInvalidRows(lines);
        for (var r = 0; r < grid.Rows.Count; r++)
        {
            var bad = r < inv.Length && inv[r];
            grid.Rows[r].DefaultCellStyle.BackColor = bad ? Color.MistyRose : Color.White;
            grid.Rows[r].DefaultCellStyle.ForeColor = bad ? Color.DarkRed : Color.Black;
        }
    }

    private static bool[] ComputeInvalidRows(IList<PatternLine> lines)
    {
        var inv = new bool[lines.Count];
        var ranges = new List<(double Lo, double Hi, int Row)>();

        for (var i = 0; i < lines.Count; i++)
        {
            var a = lines[i].StartMm;
            var b = lines[i].EndMm;
            if (double.IsNaN(a) || double.IsInfinity(a) || double.IsNaN(b) || double.IsInfinity(b)) { inv[i] = true; continue; }
            if (a > b) { inv[i] = true; continue; }
            if (a < MinPatternMm || b > MaxPatternMm) { inv[i] = true; continue; }
            ranges.Add((a, b, i));
        }

        ranges.Sort((x, y) => x.Lo.CompareTo(y.Lo));
        for (var i = 1; i < ranges.Count; i++)
        {
            if (ranges[i].Lo < ranges[i - 1].Hi) { inv[ranges[i].Row] = true; inv[ranges[i - 1].Row] = true; }
        }
        return inv;
    }

    private bool ValidateAllPatterns(out string error)
    {
        if (!ValidatePattern("Gun 1", _gun1Lines, out error)) return false;
        if (!ValidatePattern("Gun 2", _gun2Lines, out error)) return false;
        error = string.Empty;
        return true;
    }

    private static bool ValidatePattern(string gunName, IList<PatternLine> lines, out string error)
    {
        if (lines.Count > MaxLinesPerGun) { error = $"{gunName}: exceeds {MaxLinesPerGun} lines."; return false; }
        var ranges = new List<(double Lo, double Hi)>();
        for (var i = 0; i < lines.Count; i++)
        {
            var a = lines[i].StartMm; var b = lines[i].EndMm;
            if (double.IsNaN(a) || double.IsInfinity(a) || double.IsNaN(b) || double.IsInfinity(b))
            { error = $"{gunName}: line {i + 1} has invalid value."; return false; }
            if (a > b) { error = $"{gunName}: line {i + 1} start > end."; return false; }
            if (a < MinPatternMm || b > MaxPatternMm) { error = $"{gunName}: line {i + 1} out of range."; return false; }
            ranges.Add((a, b));
        }
        ranges.Sort((x, y) => x.Lo.CompareTo(y.Lo));
        for (var i = 1; i < ranges.Count; i++)
        {
            if (ranges[i].Lo < ranges[i - 1].Hi) { error = $"{gunName}: lines overlap."; return false; }
        }
        error = string.Empty;
        return true;
    }

    // ═══════════════════════════════════════════════════════════════
    //  DIRTY TRACKING & AUTO-SEND SCHEDULING
    // ═══════════════════════════════════════════════════════════════

    private void MarkConfigChanged()
    {
        if (_suspendDirtyTracking) return;
        _hasUnsavedChanges = true;
        _hasUnsentChanges = true;
        UpdateProgramStateUi();
        ScheduleAutoSend();
    }

    private void MarkPatternChanged()
    {
        if (_suspendDirtyTracking) return;
        _hasUnsavedChanges = true;
        _hasUnsentChanges = true;
        UpdateProgramStateUi();
        ScheduleAutoSend();
    }

    private void ScheduleAutoSend()
    {
        _autoSendTimer.Stop();
        _autoSendTimer.Start();
    }

    private void UpdateProgramStateUi()
    {
    }

    private void SetActiveIndicator(bool active)
    {
        if (active)
        {
            _activeIndicator.Text = "  ACTIVE  ";
            _activeIndicator.BackColor = Color.FromArgb(50, 160, 50);
            _activeIndicator.ForeColor = Color.White;
        }
        else
        {
            _activeIndicator.Text = "  INACTIVE  ";
            _activeIndicator.BackColor = Color.FromArgb(200, 60, 60);
            _activeIndicator.ForeColor = Color.White;
        }
    }

    // ═══════════════════════════════════════════════════════════════
    //  PROGRAM STATE PERSISTENCE
    // ═══════════════════════════════════════════════════════════════

    private GlueProgramFile BuildProgramFromUi() => new()
    {
        Config = new DeviceConfig
        {
            PulsesPerMm = (double)_pulsesPerMm.Value,
            MaxMsPerMm = (int)_maxMsPerMm.Value,
            PhotocellOffsetMm = (double)_photocellOffset.Value,
            DebounceMs = (int)_debounceMs.Value
        },
        Gun1Lines = _gun1Lines.Select(l => l.Clone()).ToList(),
        Gun2Lines = _gun2Lines.Select(l => l.Clone()).ToList()
    };

    private void ApplyProgramToUi(GlueProgramFile program)
    {
        var prev = _suspendDirtyTracking;
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
        RefreshPreviews();
        _suspendDirtyTracking = prev;
    }

    private static decimal ClampDecimal(decimal v, decimal min, decimal max) => v < min ? min : v > max ? max : v;

    private static void ReplaceLines(BindingList<PatternLine> target, IList<PatternLine>? source)
    {
        target.Clear();
        if (source is null) return;
        foreach (var l in source) target.Add(l.Clone());
    }

    private void SaveLastAppState()
    {
        try
        {
            var state = new AppStateFile
            {
                Program = BuildProgramFromUi(),
                LastPort = _portCombo.SelectedItem?.ToString(),
                LastProgramName = _programCombo.Text,
                CalibrationPaperLengthMm = (double)_calibPaperLength.Value,
                HasUnsentChanges = _hasUnsentChanges,
                HasUnsavedChanges = _hasUnsavedChanges
            };
            Directory.CreateDirectory(AppStateDirectory);
            File.WriteAllText(AppStateFilePath, JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex) { AppendLog($"State save warning: {ex.Message}"); }
    }

    private void LoadLastAppState()
    {
        if (!File.Exists(AppStateFilePath)) return;
        try
        {
            var json = File.ReadAllText(AppStateFilePath);
            var state = JsonSerializer.Deserialize<AppStateFile>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (state is null) return;

            _suspendDirtyTracking = true;
            if (state.Program is not null) ApplyProgramToUi(state.Program);
            _calibPaperLength.Value = ClampDecimal((decimal)state.CalibrationPaperLengthMm, _calibPaperLength.Minimum, _calibPaperLength.Maximum);
            if (!string.IsNullOrWhiteSpace(state.LastPort) && _portCombo.Items.Contains(state.LastPort))
                _portCombo.SelectedItem = state.LastPort;
            if (!string.IsNullOrWhiteSpace(state.LastProgramName))
                _programCombo.Text = state.LastProgramName;
            _hasUnsentChanges = state.HasUnsentChanges;
            _hasUnsavedChanges = state.HasUnsavedChanges;
            _suspendDirtyTracking = false;
            UpdateProgramStateUi();
            RefreshPreviews();
            AppendLog("Restored previous session.");
        }
        catch (Exception ex) { _suspendDirtyTracking = false; AppendLog($"State load warning: {ex.Message}"); }
    }

    private void AppendLog(string text)
    {
        _logBox.AppendText($"{DateTime.Now:HH:mm:ss} | {text}{Environment.NewLine}");
    }
}

// ═══════════════════════════════════════════════════════════════
//  PATTERN PREVIEW (double-buffered)
// ═══════════════════════════════════════════════════════════════

public sealed class PatternPreviewPanel : Panel
{
    private readonly List<PatternLine> _lines = new();

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public Color PatternColor { get; set; } = Color.DarkOrange;

    public PatternPreviewPanel()
    {
        DoubleBuffered = true;
        ResizeRedraw = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }

    [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public double PaperLengthMm { get; set; } = 297.0;

    public void SetLines(IEnumerable<PatternLine> lines, double paperLengthMm = 297.0)
    {
        _lines.Clear();
        _lines.AddRange(lines.Select(l => l.Clone()));
        PaperLengthMm = paperLengthMm;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        var rect = ClientRectangle;
        if (rect.Width < 40 || rect.Height < 20) return;

        using var bg = new SolidBrush(Color.FromArgb(250, 251, 253));
        g.FillRectangle(bg, rect);

        var pad = 8;
        var chart = new Rectangle(rect.X + pad, rect.Y + pad, rect.Width - pad * 2, rect.Height - pad * 2 - 16);
        if (chart.Height < 6) return;

        using var axisPen = new Pen(Color.FromArgb(180, 180, 180), 1);
        g.DrawLine(axisPen, chart.Left, chart.Bottom, chart.Right, chart.Bottom);

        var paperLen = Math.Max(1.0, PaperLengthMm);
        var patternMax = 0.0;
        foreach (var l in _lines) patternMax = Math.Max(patternMax, Math.Max(l.StartMm, l.EndMm));
        var max = Math.Max(paperLen, patternMax * 1.05);
        if (max < 1.0) max = 1.0;

        // Draw end-of-paper line
        var paperX = chart.Left + (float)(paperLen / max) * chart.Width;
        using var paperPen = new Pen(Color.FromArgb(200, 220, 50, 50), 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
        g.DrawLine(paperPen, paperX, chart.Top, paperX, chart.Bottom);

        // Shade area beyond paper
        if (paperX < chart.Right - 1)
        {
            using var beyondBrush = new SolidBrush(Color.FromArgb(30, 255, 0, 0));
            g.FillRectangle(beyondBrush, paperX, chart.Top, chart.Right - paperX, chart.Height);
        }

        using var fill = new SolidBrush(Color.FromArgb(180, PatternColor));
        using var fillBeyond = new SolidBrush(Color.FromArgb(180, Color.IndianRed));
        using var outline = new Pen(Color.FromArgb(220, PatternColor), 1);
        using var outlineBeyond = new Pen(Color.FromArgb(220, Color.IndianRed), 1);
        using var labelFont = new Font("Segoe UI", 7F);
        using var labelBrush = new SolidBrush(Color.FromArgb(100, 100, 100));

        foreach (var line in _lines)
        {
            var lo = Math.Min(line.StartMm, line.EndMm);
            var hi = Math.Max(line.StartMm, line.EndMm);
            var x1 = chart.Left + (float)(lo / max) * chart.Width;
            var x2 = chart.Left + (float)(hi / max) * chart.Width;
            var w = Math.Max(3, x2 - x1);
            var beyond = hi > paperLen;
            var r = new RectangleF(x1, chart.Top + 2, w, chart.Height - 2);
            g.FillRectangle(beyond ? fillBeyond : fill, r);
            g.DrawRectangle(beyond ? outlineBeyond : outline, r.X, r.Y, r.Width, r.Height);

            if (w > 28)
            {
                var txt = $"{lo:0.#}-{hi:0.#}";
                var sz = g.MeasureString(txt, labelFont);
                if (sz.Width < w - 2)
                    g.DrawString(txt, labelFont, labelBrush, r.X + (w - sz.Width) / 2, r.Y + (r.Height - sz.Height) / 2);
            }
        }

        using var axisFont = new Font("Segoe UI", 7.5F);
        using var axisBrush = new SolidBrush(Color.FromArgb(80, 80, 80));
        g.DrawString("0", axisFont, axisBrush, chart.Left, chart.Bottom + 1);

        // Paper length label at the paper line
        var paperStr = $"{paperLen:0.#} mm";
        var paperSz = g.MeasureString(paperStr, axisFont);
        using var paperLabelBrush = new SolidBrush(Color.FromArgb(200, 60, 60));
        g.DrawString(paperStr, axisFont, paperLabelBrush, paperX - paperSz.Width / 2, chart.Bottom + 1);

        if (patternMax > paperLen)
        {
            var maxStr = $"{patternMax:0.#} mm";
            var maxSz = g.MeasureString(maxStr, axisFont);
            g.DrawString(maxStr, axisFont, axisBrush, chart.Right - maxSz.Width, chart.Bottom + 1);
        }
    }
}

// ═══════════════════════════════════════════════════════════════
//  DATA MODELS
// ═══════════════════════════════════════════════════════════════

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
    public string? LastProgramName { get; set; }
    public double CalibrationPaperLengthMm { get; set; } = 297.0;
    public bool HasUnsentChanges { get; set; }
    public bool HasUnsavedChanges { get; set; }
}
