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
    private readonly Label _statusLabel = new()
    {
        Text = "  Disconnected  ",
        AutoSize = true,
        Font = new Font("Segoe UI", 10F, FontStyle.Bold),
        ForeColor = Color.White,
        BackColor = Color.FromArgb(140, 140, 140),
        Padding = new Padding(6, 3, 6, 3),
        Margin = new Padding(6, 2, 0, 0)
    };
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

    // ── Tooltips ──
    private readonly ToolTip _toolTip = new() { AutoPopDelay = 10000, InitialDelay = 400, ReshowDelay = 200 };

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
    private readonly CombinedPreviewPanel _combinedPreview = new();
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
    private readonly Button _logToggleButton = new()
    {
        Text = "▶ Serial Log",
        Dock = DockStyle.Top,
        Height = 24,
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.FromArgb(220, 222, 228),
        TextAlign = ContentAlignment.MiddleLeft,
        Padding = new Padding(4, 0, 0, 0)
    };
    private Panel _logPanel = null!;
    private bool _logExpanded = false;
    private TableLayoutPanel _rootLayout = null!;

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
        SetupTooltips();

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

    private void SetupTooltips()
    {
        // Config parameters
        _toolTip.SetToolTip(_pulsesPerMm, "Encoder pulses per millimeter of paper travel.\nSet by calibration or manually.");
        _toolTip.SetToolTip(_maxMsPerMm, "Maximum time (ms) allowed per mm of movement.\nIf paper moves slower than this, glue guns are blocked to prevent dripping.");
        _toolTip.SetToolTip(_photocellOffset, "Distance (mm) between the photocell sensor and the glue nozzles.\nPositive = nozzles are downstream of the sensor.");
        _toolTip.SetToolTip(_debounceMs, "Input debounce time (ms) for the photocell sensor.\nFilters out electrical noise on the signal.");
        _toolTip.SetToolTip(_calibPaperLength, "Known paper length (mm) used during encoder calibration.\nAlso sets the paper boundary shown in the pattern preview.");
        _toolTip.SetToolTip(_calibArmButton, "Arm encoder calibration.\nPass a sheet of known length past the photocell to measure pulses/mm.");

        // Machine control
        _toolTip.SetToolTip(_activateButton, "Activate the system — enables photocell detection and glue pattern firing.");
        _toolTip.SetToolTip(_deactivateButton, "Deactivate the system — stops all guns and ignores triggers.");
        _toolTip.SetToolTip(_testOpenGun1Button, "Manually open Gun 1 for testing (30s timeout).");
        _toolTip.SetToolTip(_testCloseGun1Button, "Manually close Gun 1.");
        _toolTip.SetToolTip(_testOpenGun2Button, "Manually open Gun 2 for testing (30s timeout).");
        _toolTip.SetToolTip(_testCloseGun2Button, "Manually close Gun 2.");
        _toolTip.SetToolTip(_testOpenBothButton, "Manually open both guns for testing (30s timeout).");
        _toolTip.SetToolTip(_testCloseBothButton, "Manually close both guns.");
        _toolTip.SetToolTip(_swTriggerButton, "Software trigger — simulates a photocell event to start a glue cycle.\nSystem must be active and encoder must be moving.");

        // Connection
        _toolTip.SetToolTip(_portCombo, "Select the COM port connected to the Arduino controller.");
        _toolTip.SetToolTip(_refreshPortsButton, "Refresh the list of available COM ports.");
        _toolTip.SetToolTip(_connectButton, "Connect or disconnect from the selected COM port.");
        _toolTip.SetToolTip(_programCombo, "Select or type a program name.\nPrograms store pattern and config settings.");
        _toolTip.SetToolTip(_deleteProgramButton, "Delete the currently selected program file.");

        // Pattern grids
        _toolTip.SetToolTip(_gun1Grid, "Glue pattern for Gun 1.\nEach row defines a start and end position (mm) where glue is applied.");
        _toolTip.SetToolTip(_gun2Grid, "Glue pattern for Gun 2.\nEach row defines a start and end position (mm) where glue is applied.");
    }

    // ═══════════════════════════════════════════════════════════════
    //  LAYOUT
    // ═══════════════════════════════════════════════════════════════

    private void BuildLayout()
    {
        _rootLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(6)
        };
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // toolbar
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // config bar
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));   // main area
        _rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // log (collapsed by default)

        _rootLayout.Controls.Add(BuildToolbar(), 0, 0);
        _rootLayout.Controls.Add(BuildConfigBar(), 0, 1);
        _rootLayout.Controls.Add(BuildMainArea(), 0, 2);
        _rootLayout.Controls.Add(BuildLogPanel(), 0, 3);

        Controls.Add(_rootLayout);
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

        bar.Controls.Add(MakeLabel("Max ms/mm:"));
        bar.Controls.Add(_maxMsPerMm);
        bar.Controls.Add(MakeLabel("Photocell offset (mm):"));
        bar.Controls.Add(_photocellOffset);
        bar.Controls.Add(MakeLabel("input Debounce (ms):"));
        bar.Controls.Add(_debounceMs);
        bar.Controls.Add(MakeSeparator());
        bar.Controls.Add(MakeLabel("paper length (mm):"));
        bar.Controls.Add(_calibPaperLength);
        bar.Controls.Add(_calibArmButton);
        bar.Controls.Add(MakeLabel("Pulses/mm:"));
        bar.Controls.Add(_pulsesPerMm);
        

        return bar;
    }

    private Control BuildMainArea()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 2
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));     // gun 1
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));     // gun 2
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));   // machine control
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));          // grids + control
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110));         // preview

        layout.Controls.Add(BuildGunGridPanel("Gun 1", _gun1Grid, _gun1AddLineButton, _gun1RemoveLineButton), 0, 0);
        layout.Controls.Add(BuildGunGridPanel("Gun 2", _gun2Grid, _gun2AddLineButton, _gun2RemoveLineButton), 1, 0);
        layout.Controls.Add(BuildControlPanel(), 2, 0);

        var previewGroup = new GroupBox { Text = "Pattern Preview", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(60, 60, 60) };
        _combinedPreview.Dock = DockStyle.Fill;
        _combinedPreview.MinimumSize = new Size(100, 40);
        previewGroup.Controls.Add(_combinedPreview);
        layout.Controls.Add(previewGroup, 0, 1);
        layout.SetColumnSpan(previewGroup, 3);

        return layout;
    }

    private static Control BuildGunGridPanel(string title, DataGridView grid, Button addBtn, Button removeBtn)
    {
        var group = new GroupBox { Text = title, Dock = DockStyle.Fill, ForeColor = Color.FromArgb(60, 60, 60) };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        grid.Dock = DockStyle.Fill;
        grid.BackgroundColor = Color.White;
        grid.BorderStyle = BorderStyle.FixedSingle;

        var btnPanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };
        btnPanel.Controls.Add(addBtn);
        btnPanel.Controls.Add(removeBtn);

        layout.Controls.Add(grid, 0, 0);
        layout.Controls.Add(btnPanel, 0, 1);

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
        _logPanel = new Panel { Dock = DockStyle.Fill, Height = 28 };
        _logBox.Visible = false;
        _logToggleButton.Click += (_, _) => ToggleLog();
        _logPanel.Controls.Add(_logBox);
        _logPanel.Controls.Add(_logToggleButton);
        return _logPanel;
    }

    private void ToggleLog()
    {
        _logExpanded = !_logExpanded;
        if (_logExpanded)
        {
            _logToggleButton.Text = "▼ Serial Log";
            _logBox.Visible = true;
            _rootLayout.RowStyles[3] = new RowStyle(SizeType.Absolute, 180);
        }
        else
        {
            _logToggleButton.Text = "▶ Serial Log";
            _logBox.Visible = false;
            _rootLayout.RowStyles[3] = new RowStyle(SizeType.AutoSize);
        }
        _rootLayout.PerformLayout();
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
            _statusLabel.Text = $"  Connected ({port})  ";
            _statusLabel.ForeColor = Color.White;
            _statusLabel.BackColor = Color.FromArgb(50, 160, 50);
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
        if (_serial is null) { _connectButton.Text = "Connect"; _statusLabel.Text = "  Disconnected  "; _statusLabel.ForeColor = Color.White; _statusLabel.BackColor = Color.FromArgb(140, 140, 140); return; }
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
            _statusLabel.Text = "  Disconnected  ";
            _statusLabel.ForeColor = Color.White;
            _statusLabel.BackColor = Color.FromArgb(140, 140, 140);
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
        _combinedPreview.SetData(_gun1Lines, _gun2Lines, paperLen);
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

public sealed class CombinedPreviewPanel : Panel
{
    private readonly List<PatternLine> _gun1Lines = new();
    private readonly List<PatternLine> _gun2Lines = new();
    private double _paperLengthMm = 297.0;

    private static readonly Color Gun1Color = Color.OrangeRed;
    private static readonly Color Gun2Color = Color.SteelBlue;
    private const int LineThickness = 16;
    private const int Gun1Y = 0;   // row index
    private const int Gun2Y = 1;

    public CombinedPreviewPanel()
    {
        DoubleBuffered = true;
        ResizeRedraw = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
    }

    public void SetData(IEnumerable<PatternLine> gun1, IEnumerable<PatternLine> gun2, double paperLengthMm)
    {
        _gun1Lines.Clear();
        _gun1Lines.AddRange(gun1.Select(l => l.Clone()));
        _gun2Lines.Clear();
        _gun2Lines.AddRange(gun2.Select(l => l.Clone()));
        _paperLengthMm = paperLengthMm;
        Invalidate();
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        var rect = ClientRectangle;
        if (rect.Width < 40 || rect.Height < 30) return;

        using var bg = new SolidBrush(Color.FromArgb(250, 251, 253));
        g.FillRectangle(bg, rect);

        var pad = 10;
        var topPad = 22;
        var chart = new Rectangle(rect.X + pad, rect.Y + topPad, rect.Width - pad * 2, rect.Height - topPad - pad - 18);
        if (chart.Height < 20) return;

        // Compute scale
        var paperLen = Math.Max(1.0, _paperLengthMm);
        var patternMax = 0.0;
        foreach (var l in _gun1Lines) patternMax = Math.Max(patternMax, Math.Max(l.StartMm, l.EndMm));
        foreach (var l in _gun2Lines) patternMax = Math.Max(patternMax, Math.Max(l.StartMm, l.EndMm));
        var max = Math.Max(paperLen, patternMax * 1.05);
        if (max < 1.0) max = 1.0;

        // Axis
        using var axisPen = new Pen(Color.FromArgb(180, 180, 180), 1);
        g.DrawLine(axisPen, chart.Left, chart.Bottom, chart.Right, chart.Bottom);

        // End-of-paper dashed line
        var paperX = chart.Left + (float)(paperLen / max) * chart.Width;
        using var paperPen = new Pen(Color.FromArgb(200, 220, 50, 50), 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
        g.DrawLine(paperPen, paperX, chart.Top - 4, paperX, chart.Bottom);

        // Shade beyond paper
        if (paperX < chart.Right - 1)
        {
            using var beyondBrush = new SolidBrush(Color.FromArgb(25, 255, 0, 0));
            g.FillRectangle(beyondBrush, paperX, chart.Top, chart.Right - paperX, chart.Height);
        }

        // Gun rows: split chart height into two lanes
        var laneH = chart.Height / 2;
        var gun1CenterY = chart.Top + laneH / 2;
        var gun2CenterY = chart.Top + laneH + laneH / 2;

        // Lane divider
        using var lanePen = new Pen(Color.FromArgb(60, 180, 180, 180), 1) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dot };
        g.DrawLine(lanePen, chart.Left, chart.Top + laneH, chart.Right, chart.Top + laneH);

        // Draw gun labels
        using var labelFont = new Font("Segoe UI", 7.5F, FontStyle.Bold);
        using var gun1LabelBrush = new SolidBrush(Color.FromArgb(180, Gun1Color));
        using var gun2LabelBrush = new SolidBrush(Color.FromArgb(180, Gun2Color));
        g.DrawString("Gun 1", labelFont, gun1LabelBrush, chart.Left, chart.Top - 16);
        g.DrawString("Gun 2", labelFont, gun2LabelBrush, chart.Left, chart.Top + laneH - 1);

        // Draw pattern lines
        DrawGunLines(g, _gun1Lines, gun1CenterY, max, paperLen, chart, Gun1Color, labelsBelow: false);
        DrawGunLines(g, _gun2Lines, gun2CenterY, max, paperLen, chart, Gun2Color, labelsBelow: true);

        // Axis labels
        using var axisFont = new Font("Segoe UI", 7.5F);
        using var axisBrush = new SolidBrush(Color.FromArgb(80, 80, 80));
        g.DrawString("0", axisFont, axisBrush, chart.Left, chart.Bottom + 2);

        using var paperLabelBrush = new SolidBrush(Color.FromArgb(200, 60, 60));
        var paperStr = $"{paperLen:0.#} mm";
        var paperSz = g.MeasureString(paperStr, axisFont);
        g.DrawString(paperStr, axisFont, paperLabelBrush, paperX - paperSz.Width / 2, chart.Bottom + 2);

        if (patternMax > paperLen)
        {
            var maxStr = $"{patternMax:0.#} mm";
            var maxSz = g.MeasureString(maxStr, axisFont);
            g.DrawString(maxStr, axisFont, axisBrush, chart.Right - maxSz.Width, chart.Bottom + 2);
        }
    }

    private static void DrawGunLines(Graphics g, List<PatternLine> lines, int centerY, double max, double paperLen, Rectangle chart, Color color, bool labelsBelow)
    {
        using var pen = new Pen(color, LineThickness);
        using var penBeyond = new Pen(Color.IndianRed, LineThickness);
        using var labelFont = new Font("Segoe UI", 6.5F);
        using var labelBrush = new SolidBrush(Color.FromArgb(140, 60, 60, 60));

        foreach (var line in lines)
        {
            var lo = Math.Min(line.StartMm, line.EndMm);
            var hi = Math.Max(line.StartMm, line.EndMm);
            var x1 = chart.Left + (float)(lo / max) * chart.Width;
            var x2 = chart.Left + (float)(hi / max) * chart.Width;
            if (x2 - x1 < 2) x2 = x1 + 2;
            var beyond = hi > paperLen;
            g.DrawLine(beyond ? penBeyond : pen, x1, centerY, x2, centerY);

            var w = x2 - x1;
            if (w > 30)
            {
                var txt = $"{lo:0.#}-{hi:0.#}";
                var sz = g.MeasureString(txt, labelFont);
                if (sz.Width < w)
                {
                    var labelY = labelsBelow
                        ? centerY + LineThickness / 2 + 2
                        : centerY - LineThickness / 2 - sz.Height - 1;
                    g.DrawString(txt, labelFont, labelBrush, x1 + (w - sz.Width) / 2, labelY);
                }
            }
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
