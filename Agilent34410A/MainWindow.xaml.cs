using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Media;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Threading;
using System.Windows.Threading;
using Microsoft.Win32;
using NationalInstruments.Visa;

namespace Agilent34410A;

// Data model for readings
public class Reading
{
    public int Number { get; set; }
    public string Timestamp { get; set; } = "";
    public double Value { get; set; }
    public string ValueStr { get; set; } = "";
    public string Unit { get; set; } = "";
    public string Function { get; set; } = "";
}

// Measurement function definition
public class MeasFunc
{
    public string Name { get; set; } = "";
    public string MeasCmd { get; set; } = "";
    public string ConfCmd { get; set; } = "";
    public string SensePrefix { get; set; } = "";
    public string Unit { get; set; } = "";
    public string[] Ranges { get; set; } = [];
    public bool HasNplc { get; set; }
    public bool HasImpedance { get; set; }
    public bool HasNull { get; set; }
    public bool HasBandwidth { get; set; }
    public bool HasAperture { get; set; }
}

public partial class MainWindow : Window
{
    private MessageBasedSession? _session;
    private bool _connected;
    private bool _streaming;
    private bool _recording;
    private CancellationTokenSource? _streamCts;
    private StreamWriter? _csvWriter;
    private string _currentFunc = "DC Voltage";
    private int _readingCount;
    private double _statSum;
    private double? _statMin, _statMax;

    // Chart data
    private readonly List<(DateTime time, double value)> _chartData = new();
    private const int MaxChartPoints = 200;

    // Rate tracking
    private DateTime _rateWindowStart = DateTime.Now;
    private int _rateCount;
    private double _currentRate;

    // Alarm state
    private bool _alarmEnabled;
    private double? _alarmHi, _alarmLo;
    private bool _alarmTriggered;

    // Scope digitize data
    private bool _scopeRunning;
    private CancellationTokenSource? _scopeCts;

    // Relative mode
    private bool _relEnabled;
    private double _relReference;

    private readonly ObservableCollection<Reading> _readings = new();
    private readonly Dictionary<string, Button> _funcButtons = new();

    private static readonly Dictionary<string, MeasFunc> Functions = new()
    {
        ["DC Voltage"] = new MeasFunc {
            Name = "DC Voltage", MeasCmd = "MEAS:VOLT:DC?", ConfCmd = "CONF:VOLT:DC",
            SensePrefix = "VOLT:DC", Unit = "V DC",
            Ranges = ["AUTO", "0.1", "1", "10", "100", "1000"],
            HasNplc = true, HasImpedance = true, HasNull = true },

        ["AC Voltage"] = new MeasFunc {
            Name = "AC Voltage", MeasCmd = "MEAS:VOLT:AC?", ConfCmd = "CONF:VOLT:AC",
            SensePrefix = "VOLT:AC", Unit = "V AC",
            Ranges = ["AUTO", "0.1", "1", "10", "100", "750"],
            HasBandwidth = true, HasNull = true },

        ["DC Current"] = new MeasFunc {
            Name = "DC Current", MeasCmd = "MEAS:CURR:DC?", ConfCmd = "CONF:CURR:DC",
            SensePrefix = "CURR:DC", Unit = "A DC",
            Ranges = ["AUTO", "0.0001", "0.001", "0.01", "0.1", "1", "3"],
            HasNplc = true, HasNull = true },

        ["AC Current"] = new MeasFunc {
            Name = "AC Current", MeasCmd = "MEAS:CURR:AC?", ConfCmd = "CONF:CURR:AC",
            SensePrefix = "CURR:AC", Unit = "A AC",
            Ranges = ["AUTO", "0.0001", "0.001", "0.01", "0.1", "1", "3"],
            HasBandwidth = true, HasNull = true },

        ["2W Resistance"] = new MeasFunc {
            Name = "2W Resistance", MeasCmd = "MEAS:RES?", ConfCmd = "CONF:RES",
            SensePrefix = "RES", Unit = "\u03a9",
            Ranges = ["AUTO", "100", "1e3", "10e3", "100e3", "1e6", "10e6", "100e6", "1e9"],
            HasNplc = true, HasNull = true },

        ["4W Resistance"] = new MeasFunc {
            Name = "4W Resistance", MeasCmd = "MEAS:FRES?", ConfCmd = "CONF:FRES",
            SensePrefix = "FRES", Unit = "\u03a9 4W",
            Ranges = ["AUTO", "100", "1e3", "10e3", "100e3", "1e6", "10e6", "100e6", "1e9"],
            HasNplc = true, HasNull = true },

        ["Frequency"] = new MeasFunc {
            Name = "Frequency", MeasCmd = "MEAS:FREQ?", ConfCmd = "CONF:FREQ",
            SensePrefix = "FREQ", Unit = "Hz",
            Ranges = ["AUTO", "0.1", "1", "10", "100", "750"],
            HasNull = true, HasAperture = true },

        ["Period"] = new MeasFunc {
            Name = "Period", MeasCmd = "MEAS:PER?", ConfCmd = "CONF:PER",
            SensePrefix = "PER", Unit = "s",
            Ranges = ["AUTO", "0.1", "1", "10", "100", "750"],
            HasNull = true, HasAperture = true },

        ["Continuity"] = new MeasFunc {
            Name = "Continuity", MeasCmd = "MEAS:CONT?", ConfCmd = "CONF:CONT",
            SensePrefix = "CONT", Unit = "\u03a9", Ranges = [] },

        ["Diode"] = new MeasFunc {
            Name = "Diode", MeasCmd = "MEAS:DIOD?", ConfCmd = "CONF:DIOD",
            SensePrefix = "DIOD", Unit = "V", Ranges = [] },

        ["Temperature"] = new MeasFunc {
            Name = "Temperature", MeasCmd = "MEAS:TEMP?", ConfCmd = "CONF:TEMP",
            SensePrefix = "TEMP", Unit = "\u00b0C", Ranges = [],
            HasNplc = true, HasNull = true },

        ["Capacitance"] = new MeasFunc {
            Name = "Capacitance", MeasCmd = "MEAS:CAP?", ConfCmd = "CONF:CAP",
            SensePrefix = "CAP", Unit = "F",
            Ranges = ["AUTO", "1e-9", "10e-9", "100e-9", "1e-6", "10e-6", "100e-6"],
            HasNull = true },
    };

    public MainWindow()
    {
        InitializeComponent();
        dataGrid.ItemsSource = _readings;
        InitControls();
        PopulateResources();

        // Auto-connect to GPIB after load
        Loaded += (_, _) => Dispatcher.BeginInvoke(() =>
        {
            var resource = cmbResource.SelectedItem as string;
            if (!string.IsNullOrEmpty(resource) && resource.Contains("GPIB"))
                Connect(resource);
        }, DispatcherPriority.Background);
    }

    private void InitControls()
    {
        // NPLC values
        cmbNplc.ItemsSource = new[] { "0.006", "0.02", "0.06", "0.2", "1", "2", "10", "100" };
        cmbNplc.SelectedItem = "10";

        // Trigger sources
        cmbTrigSource.ItemsSource = new[] { "IMM", "BUS", "EXT", "INT" };
        cmbTrigSource.SelectedIndex = 0;

        // Math functions
        cmbMath.ItemsSource = new[] { "OFF", "NULL", "DB", "DBM", "AVER", "LIM" };
        cmbMath.SelectedIndex = 0;

        // Temperature config
        cmbProbe.ItemsSource = new[] { "RTD", "FRTD", "THER" };
        cmbProbe.SelectedIndex = 0;
        cmbProbeType.ItemsSource = new[] { "PT100", "2252", "5000", "10000" };
        cmbProbeType.SelectedIndex = 0;
        cmbTempUnit.ItemsSource = new[] { "C", "F", "K" };
        cmbTempUnit.SelectedIndex = 0;

        // Map function buttons
        _funcButtons["DC Voltage"] = btnDCV;
        _funcButtons["AC Voltage"] = btnACV;
        _funcButtons["DC Current"] = btnDCI;
        _funcButtons["AC Current"] = btnACI;
        _funcButtons["2W Resistance"] = btnR2W;
        _funcButtons["4W Resistance"] = btnR4W;
        _funcButtons["Frequency"] = btnFreq;
        _funcButtons["Period"] = btnPer;
        _funcButtons["Continuity"] = btnCont;
        _funcButtons["Diode"] = btnDiode;
        _funcButtons["Temperature"] = btnTemp;
        _funcButtons["Capacitance"] = btnCap;

        // Initial range
        UpdateRangeCombo();
        UpdateConfigVisibility();
    }

    // ─── VISA Communication ─────────────────────────────────

    private bool _scopeLogging; // only log during scope debug

    private void Log(string msg)
    {
        try
        {
            var ts = DateTime.Now.ToString("HH:mm:ss.fff");
            var line = $"[{ts}] {msg}";
            // Write to file for easy reading
            try { File.AppendAllText(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "scope_debug.log"), line + "\n"); } catch { }
            Dispatcher.InvokeAsync(() =>
            {
                lstLog.Items.Add(line);
                if (lstLog.Items.Count > 500)
                    lstLog.Items.RemoveAt(0);
                lstLog.ScrollIntoView(lstLog.Items[lstLog.Items.Count - 1]);
            });
        }
        catch { }
    }

    private void BtnClearLog_Click(object sender, RoutedEventArgs e)
    {
        lstLog.Items.Clear();
    }

    private void BtnToggleLog_Click(object sender, RoutedEventArgs e)
    {
        logPanel.Visibility = logPanel.Visibility == Visibility.Visible
            ? Visibility.Collapsed : Visibility.Visible;
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    private void BtnScopeScreenshot_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            SetProcessDPIAware();
            var dlg = new SaveFileDialog
            {
                Filter = "PNG Image|*.png",
                FileName = $"scope_{DateTime.Now:yyyyMMdd_HHmmss}.png"
            };
            if (dlg.ShowDialog() == true)
            {
                // Render the scope area to bitmap
                var dpi = VisualTreeHelper.GetDpi(this);
                // Navigate up: chartCanvas -> Grid -> Grid (scope) -> Border
                var scopeGrid = VisualTreeHelper.GetParent(chartCanvas);
                var outerGrid = VisualTreeHelper.GetParent(scopeGrid);
                var scopeBorder = (FrameworkElement)VisualTreeHelper.GetParent(outerGrid);
                var renderW = scopeBorder.ActualWidth * dpi.DpiScaleX;
                var renderH = scopeBorder.ActualHeight * dpi.DpiScaleY;
                var rtb = new System.Windows.Media.Imaging.RenderTargetBitmap(
                    (int)renderW, (int)renderH, dpi.PixelsPerInchX, dpi.PixelsPerInchY,
                    PixelFormats.Pbgra32);
                rtb.Render(scopeBorder);
                var encoder = new System.Windows.Media.Imaging.PngBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(rtb));
                using var stream = File.Create(dlg.FileName);
                encoder.Save(stream);
                Log($"Screenshot saved: {dlg.FileName}");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Screenshot error: {ex.Message}");
        }
    }

    private void Write(string cmd)
    {
        if (_scopeLogging) Log($">> {cmd}");
        _session?.RawIO.Write(cmd + "\n");
    }

    private string Query(string cmd)
    {
        if (_scopeLogging) Log($">> {cmd}");
        _session?.RawIO.Write(cmd + "\n");
        var result = _session?.RawIO.ReadString().Trim() ?? "";
        if (_scopeLogging) Log($"<< {result}");
        return result;
    }

    private double QueryFloat(string cmd)
    {
        var result = Query(cmd);
        return double.Parse(result, NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    // ─── Connection ─────────────────────────────────────────

    private void PopulateResources()
    {
        try
        {
            using var rm = new ResourceManager();
            var resources = rm.Find("?*::INSTR");
            cmbResource.ItemsSource = resources;

            foreach (var r in resources)
            {
                if (r.Contains("GPIB"))
                {
                    cmbResource.SelectedItem = r;
                    break;
                }
            }
            if (cmbResource.SelectedItem == null && resources.Any())
                cmbResource.SelectedItem = resources.First();
        }
        catch (Exception ex)
        {
            txtStatus.Text = $"VISA Error: {ex.Message}";
        }
    }

    private void Connect(string resource)
    {
        try
        {
            using var rm = new ResourceManager();
            _session = (MessageBasedSession)rm.Open(resource);
            _session.TimeoutMilliseconds = 15000;
            _session.TerminationCharacterEnabled = true;

            var idn = Query("*IDN?");
            _connected = true;
            txtStatus.Text = "Connected";
            txtStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x4c, 0xaf, 0x50));
            txtIdn.Text = idn;
            btnConnect.Content = "Disconnect";
            btnConnect.Style = (Style)FindResource("BtnStop");

            try { txtTerminals.Text = $"[{Query("ROUT:TERM?")}]"; } catch { }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Connection failed:\n{ex.Message}", "Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void Disconnect()
    {
        StopStream();
        StopRecording();
        _session?.Dispose();
        _session = null;
        _connected = false;
        txtStatus.Text = "Disconnected";
        txtStatus.Foreground = new SolidColorBrush(Color.FromRgb(0xef, 0x53, 0x50));
        txtIdn.Text = "";
        txtTerminals.Text = "";
        btnConnect.Content = "Connect";
        btnConnect.Style = (Style)FindResource("BtnAccent");
        txtValue.Text = "---";
        txtUnit.Text = "";
    }

    // ─── Measurement ────────────────────────────────────────

    private (double value, string unit, string timestamp) TakeReading()
    {
        var func = Functions[_currentFunc];
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var value = QueryFloat(func.MeasCmd);
        return (value, func.Unit, timestamp);
    }

    private string FormatValue(double value)
    {
        var abs = Math.Abs(value);
        if (abs >= 9.9e37) return "OVL";
        if (abs == 0) return "0.000000";
        if (abs >= 1e6) return $"{value / 1e6:F6} M";
        if (abs >= 1e3) return $"{value / 1e3:F6} k";
        if (abs >= 1) return $"{value:F6}";
        if (abs >= 1e-3) return $"{value * 1e3:F6} m";
        if (abs >= 1e-6) return $"{value * 1e6:F6} \u00b5";
        if (abs >= 1e-9) return $"{value * 1e9:F6} n";
        return $"{value:+0.000000E+00}";
    }

    private string FormatForLog(double value)
    {
        var abs = Math.Abs(value);
        if (abs >= 9.9e37) return "OVL";
        if (abs == 0) return "0.00000000";
        if (abs >= 1) return $"{value:F8}";
        if (abs >= 1e-3) return $"{value:F11}";
        return $"{value:+0.00000000E+00}";
    }

    private void AddReading(double value, string unit, string timestamp)
    {
        _readingCount++;
        var reading = new Reading
        {
            Number = _readingCount,
            Timestamp = timestamp,
            Value = value,
            ValueStr = FormatForLog(value),
            Unit = unit,
            Function = _currentFunc
        };

        _readings.Add(reading);

        // Auto-scroll
        if (_readings.Count > 0)
            dataGrid.ScrollIntoView(_readings[^1]);

        // Update display
        var displayValue = value;
        if (_relEnabled)
        {
            var delta = value - _relReference;
            displayValue = delta;
            txtDelta.Text = $"\u0394 {FormatValue(delta)}";
            txtRelDelta.Text = $"Ref: {FormatForLog(_relReference)}  \u0394: {FormatForLog(delta)}";
        }
        else
        {
            txtDelta.Text = "";
        }

        txtValue.Text = FormatValue(displayValue);
        txtUnit.Text = unit;

        // Update stats
        _statSum += value;
        _statMin = _statMin.HasValue ? Math.Min(_statMin.Value, value) : value;
        _statMax = _statMax.HasValue ? Math.Max(_statMax.Value, value) : value;

        txtStatCount.Text = $"Count: {_readingCount}";
        txtStatMin.Text = $"Min: {FormatForLog(_statMin ?? 0)}";
        txtStatMax.Text = $"Max: {FormatForLog(_statMax ?? 0)}";
        txtStatAvg.Text = $"Avg: {FormatForLog(_statSum / _readingCount)}";
        txtStatPtp.Text = $"P-P: {FormatForLog((_statMax ?? 0) - (_statMin ?? 0))}";

        // Update rate
        UpdateRate();

        // Update chart
        UpdateChart(value);

        // Check alarms
        CheckAlarm(value);

        // Continuity beep
        if (_currentFunc == "Continuity" && value < 50 && Math.Abs(value) < 9.9e37)
        {
            try { SystemSounds.Beep.Play(); } catch { }
        }

        // Write CSV
        if (_recording && _csvWriter != null)
        {
            _csvWriter.WriteLine($"{_readingCount},\"{timestamp}\",{FormatForLog(value)},\"{unit}\",\"{_currentFunc}\"");
            _csvWriter.Flush();
        }
    }

    // ─── Rate Monitor ────────────────────────────────────────

    private void UpdateRate()
    {
        _rateCount++;
        var elapsed = (DateTime.Now - _rateWindowStart).TotalSeconds;
        if (elapsed >= 1.0)
        {
            _currentRate = _rateCount / elapsed;
            _rateCount = 0;
            _rateWindowStart = DateTime.Now;
            txtRateDisplay.Text = $"[{_currentRate:F1} rdg/s]";
        }
    }

    // ─── Oscilloscope Display ────────────────────────────────

    private void UpdateChart(double value)
    {
        // Skip overload values
        if (Math.Abs(value) >= 9.9e37) return;

        _chartData.Add((DateTime.Now, value));
        if (_chartData.Count > MaxChartPoints)
            _chartData.RemoveAt(0);

        DrawScope();
    }

    private string FormatAxisValue(double value)
    {
        var abs = Math.Abs(value);
        if (abs == 0) return "0";
        if (abs >= 1e6) return $"{value / 1e6:G4}M";
        if (abs >= 1e3) return $"{value / 1e3:G4}k";
        if (abs >= 1) return $"{value:G5}";
        if (abs >= 1e-3) return $"{value * 1e3:G4}m";
        if (abs >= 1e-6) return $"{value * 1e6:G4}\u00b5";
        if (abs >= 1e-9) return $"{value * 1e9:G4}n";
        return $"{value:G3}";
    }

    private string FormatTimeSpan(double seconds)
    {
        if (seconds >= 60) return $"{seconds / 60:F1}m";
        if (seconds >= 1) return $"{seconds:F1}s";
        return $"{seconds * 1000:F0}ms";
    }

    private void DrawScope()
    {
        chartCanvas.Children.Clear();
        yAxisCanvas.Children.Clear();
        xAxisCanvas.Children.Clear();

        if (_chartData.Count < 2) return;

        var w = chartCanvas.ActualWidth;
        var h = chartCanvas.ActualHeight;
        if (w < 20 || h < 20) return;

        // Calculate amplitude range
        var values = _chartData.Select(d => d.value).ToArray();
        var minVal = values.Min();
        var maxVal = values.Max();
        var ampRange = maxVal - minVal;

        // Auto-scale amplitude with nice divisions
        if (ampRange < 1e-12) ampRange = Math.Max(Math.Abs(maxVal) * 0.1, 1e-6);
        var divY = CalculateNiceDivision(ampRange, 8);
        var yBottom = Math.Floor(minVal / divY) * divY;
        var yTop = Math.Ceiling(maxVal / divY) * divY;
        if (yTop <= yBottom) yTop = yBottom + divY;
        var fullYRange = yTop - yBottom;

        // Calculate time range
        var tStart = _chartData[0].time;
        var tEnd = _chartData[^1].time;
        var totalSeconds = (tEnd - tStart).TotalSeconds;
        if (totalSeconds < 0.001) totalSeconds = 1;
        var divX = CalculateNiceDivision(totalSeconds, 10);

        int numGridX = 10;
        int numGridY = (int)Math.Round(fullYRange / divY);
        if (numGridY < 2) numGridY = 2;
        if (numGridY > 12) numGridY = 12;

        // Draw graticule (scope grid)
        var gridBrush = new SolidColorBrush(Color.FromRgb(0x0a, 0x30, 0x20));
        var centerBrush = new SolidColorBrush(Color.FromRgb(0x10, 0x40, 0x28));
        var dotBrush = new SolidColorBrush(Color.FromRgb(0x14, 0x50, 0x30));

        // Horizontal grid lines + Y-axis labels
        for (int i = 0; i <= numGridY; i++)
        {
            var y = h - (h * i / (double)numGridY);
            var yVal = yBottom + (fullYRange * i / numGridY);

            // Grid line
            var isCenterish = Math.Abs(i - numGridY / 2.0) < 0.6;
            chartCanvas.Children.Add(new Line
            {
                X1 = 0, Y1 = y, X2 = w, Y2 = y,
                Stroke = isCenterish ? centerBrush : gridBrush,
                StrokeThickness = isCenterish ? 1 : 0.5
            });

            // Tick dots along the line
            for (int d = 1; d < numGridX; d++)
            {
                var dotX = w * d / numGridX;
                var tick = new Ellipse { Width = 2, Height = 2, Fill = dotBrush };
                Canvas.SetLeft(tick, dotX - 1);
                Canvas.SetTop(tick, y - 1);
                chartCanvas.Children.Add(tick);
            }

            // Y-axis label
            var label = new System.Windows.Controls.TextBlock
            {
                Text = FormatAxisValue(yVal),
                Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xe6, 0x76)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 9,
                TextAlignment = TextAlignment.Right,
                Width = 54
            };
            Canvas.SetRight(label, 2);
            Canvas.SetTop(label, y - 7);
            yAxisCanvas.Children.Add(label);
        }

        // Vertical grid lines
        for (int i = 0; i <= numGridX; i++)
        {
            var x = w * i / numGridX;
            var isCenterish = Math.Abs(i - numGridX / 2.0) < 0.6;
            chartCanvas.Children.Add(new Line
            {
                X1 = x, Y1 = 0, X2 = x, Y2 = h,
                Stroke = isCenterish ? centerBrush : gridBrush,
                StrokeThickness = isCenterish ? 1 : 0.5
            });

            // Tick dots along the line
            for (int d = 1; d < numGridY; d++)
            {
                var dotY = h * d / numGridY;
                var tick = new Ellipse { Width = 2, Height = 2, Fill = dotBrush };
                Canvas.SetLeft(tick, x - 1);
                Canvas.SetTop(tick, dotY - 1);
                chartCanvas.Children.Add(tick);
            }
        }

        // X-axis time labels
        var xAxisW = xAxisCanvas.ActualWidth > 0 ? xAxisCanvas.ActualWidth : w;
        for (int i = 0; i <= numGridX; i += 2)
        {
            var tVal = totalSeconds * i / numGridX;
            var label = new System.Windows.Controls.TextBlock
            {
                Text = FormatTimeSpan(tVal),
                Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0xa5, 0xf5)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 9,
                TextAlignment = TextAlignment.Center
            };
            var xPos = xAxisW * i / numGridX;
            Canvas.SetLeft(label, xPos - 18);
            Canvas.SetTop(label, 1);
            xAxisCanvas.Children.Add(label);
        }

        // Draw alarm limit lines
        if (_alarmEnabled)
        {
            if (_alarmHi.HasValue)
            {
                var yHi = h - ((_alarmHi.Value - yBottom) / fullYRange * h);
                if (yHi >= 0 && yHi <= h)
                    chartCanvas.Children.Add(new Line
                    {
                        X1 = 0, Y1 = yHi, X2 = w, Y2 = yHi,
                        Stroke = new SolidColorBrush(Color.FromRgb(0xef, 0x53, 0x50)),
                        StrokeThickness = 1,
                        StrokeDashArray = new DoubleCollection([4.0, 4.0])
                    });
            }
            if (_alarmLo.HasValue)
            {
                var yLo = h - ((_alarmLo.Value - yBottom) / fullYRange * h);
                if (yLo >= 0 && yLo <= h)
                    chartCanvas.Children.Add(new Line
                    {
                        X1 = 0, Y1 = yLo, X2 = w, Y2 = yLo,
                        Stroke = new SolidColorBrush(Color.FromRgb(0x42, 0xa5, 0xf5)),
                        StrokeThickness = 1,
                        StrokeDashArray = new DoubleCollection([4.0, 4.0])
                    });
            }
        }

        // Draw reference line
        if (_relEnabled)
        {
            var yRef = h - ((_relReference - yBottom) / fullYRange * h);
            if (yRef >= 0 && yRef <= h)
                chartCanvas.Children.Add(new Line
                {
                    X1 = 0, Y1 = yRef, X2 = w, Y2 = yRef,
                    Stroke = new SolidColorBrush(Color.FromRgb(0xff, 0xab, 0x40)),
                    StrokeThickness = 1,
                    StrokeDashArray = new DoubleCollection([6.0, 3.0])
                });
        }

        // Draw waveform trace (phosphor green)
        var polyline = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0xff, 0x88)),
            StrokeThickness = 1.8,
            StrokeLineJoin = PenLineJoin.Round
        };

        for (int i = 0; i < _chartData.Count; i++)
        {
            var tSec = (_chartData[i].time - tStart).TotalSeconds;
            var x = (tSec / totalSeconds) * w;
            var y = h - ((_chartData[i].value - yBottom) / fullYRange * h);
            polyline.Points.Add(new Point(x, Math.Max(0, Math.Min(h, y))));
        }
        chartCanvas.Children.Add(polyline);

        // Draw glow effect (dimmer, wider trace underneath)
        var glow = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromArgb(0x30, 0x00, 0xff, 0x88)),
            StrokeThickness = 5,
            StrokeLineJoin = PenLineJoin.Round,
            Points = polyline.Points
        };
        chartCanvas.Children.Insert(chartCanvas.Children.Count - 1, glow);

        // Latest value dot
        if (_chartData.Count > 0)
        {
            var lastT = (_chartData[^1].time - tStart).TotalSeconds;
            var lastX = (lastT / totalSeconds) * w;
            var lastY = h - ((_chartData[^1].value - yBottom) / fullYRange * h);
            lastY = Math.Max(0, Math.Min(h, lastY));

            var dot = new Ellipse
            {
                Width = 8, Height = 8,
                Fill = new SolidColorBrush(Color.FromRgb(0x00, 0xff, 0x88)),
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = Color.FromRgb(0x00, 0xff, 0x88),
                    BlurRadius = 8, ShadowDepth = 0, Opacity = 0.6
                }
            };
            Canvas.SetLeft(dot, lastX - 4);
            Canvas.SetTop(dot, lastY - 4);
            chartCanvas.Children.Add(dot);
        }

        // Update scope header info
        var unit = Functions[_currentFunc].Unit;
        txtScopeFunc.Text = _currentFunc;
        txtScopeAmpl.Text = $"Y: {FormatAxisValue(divY)}{unit}/div";
        txtScopeTime.Text = $"X: {FormatTimeSpan(divX)}/div";
        txtScopePoints.Text = $"[{_chartData.Count} pts]";
    }

    private static double CalculateNiceDivision(double range, int targetDivisions)
    {
        var rawDiv = range / targetDivisions;
        var magnitude = Math.Pow(10, Math.Floor(Math.Log10(rawDiv)));
        var normalized = rawDiv / magnitude;

        double niceDiv;
        if (normalized <= 1) niceDiv = 1;
        else if (normalized <= 2) niceDiv = 2;
        else if (normalized <= 5) niceDiv = 5;
        else niceDiv = 10;

        return niceDiv * magnitude;
    }

    // ─── Alarm / Limits ──────────────────────────────────────

    private void CheckAlarm(double value)
    {
        if (!_alarmEnabled) return;
        if (Math.Abs(value) >= 9.9e37) return;

        bool hiTrip = _alarmHi.HasValue && value > _alarmHi.Value;
        bool loTrip = _alarmLo.HasValue && value < _alarmLo.Value;

        if (hiTrip || loTrip)
        {
            _alarmTriggered = true;
            var alarmType = hiTrip ? "HIGH" : "LOW";
            txtAlarmStatus.Text = $"ALARM: {alarmType} ({FormatForLog(value)})";
            txtAlarmStatus.Foreground = new SolidColorBrush(Color.FromRgb(0xef, 0x53, 0x50));
            txtAlarmIndicator.Text = $"ALARM {alarmType}";
            alarmIndicator.Background = new SolidColorBrush(Color.FromRgb(0xef, 0x53, 0x50));

            if (chkAlarmBeep.IsChecked == true)
            {
                try { SystemSounds.Exclamation.Play(); } catch { }
            }
        }
        else
        {
            if (_alarmTriggered)
            {
                txtAlarmStatus.Text = "OK - Within limits";
                txtAlarmStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x66, 0xbb, 0x6a));
                txtAlarmIndicator.Text = "";
                alarmIndicator.Background = Brushes.Transparent;
            }
            _alarmTriggered = false;
        }
    }

    private void ChkAlarmEnable_Click(object sender, RoutedEventArgs e)
    {
        _alarmEnabled = chkAlarmEnable.IsChecked == true;
        if (_alarmEnabled)
        {
            ParseAlarmLimits();
            txtAlarmStatus.Text = "Armed";
            txtAlarmStatus.Foreground = new SolidColorBrush(Color.FromRgb(0xff, 0xab, 0x40));
        }
        else
        {
            txtAlarmStatus.Text = "";
            txtAlarmIndicator.Text = "";
            alarmIndicator.Background = Brushes.Transparent;
            _alarmTriggered = false;
        }
    }

    private void ParseAlarmLimits()
    {
        _alarmHi = double.TryParse(txtAlarmHi.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var hi) ? hi : null;
        _alarmLo = double.TryParse(txtAlarmLo.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var lo) ? lo : null;
    }

    // ─── Relative Mode ───────────────────────────────────────

    private void ChkRelEnable_Click(object sender, RoutedEventArgs e)
    {
        _relEnabled = chkRelEnable.IsChecked == true;
        if (_relEnabled)
        {
            if (double.TryParse(txtRelRef.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var r))
                _relReference = r;
            txtRelLabel.Text = "[REL]";
        }
        else
        {
            txtRelLabel.Text = "";
            txtDelta.Text = "";
            txtRelDelta.Text = "";
        }
    }

    private void BtnRelCapture_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) { ShowNotConnected(); return; }
        Task.Run(() =>
        {
            try
            {
                var (value, _, _) = TakeReading();
                Dispatcher.Invoke(() =>
                {
                    _relReference = value;
                    txtRelRef.Text = FormatForLog(value);
                    if (chkRelEnable.IsChecked == true)
                        _relEnabled = true;
                    chkRelEnable.IsChecked = true;
                    txtRelLabel.Text = "[REL]";
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => ShowError("Capture reference", ex));
            }
        });
    }

    // ─── Streaming ──────────────────────────────────────────

    private async void StartStream()
    {
        if (!_connected) { ShowNotConnected(); return; }

        _streaming = true;
        _streamCts = new CancellationTokenSource();
        btnStream.Content = "\u25a0 Stop";
        btnStream.Style = (Style)FindResource("BtnStop");
        btnSingle.IsEnabled = false;

        // Reset rate counter
        _rateWindowStart = DateTime.Now;
        _rateCount = 0;

        // Re-parse alarm limits in case they changed
        if (_alarmEnabled) ParseAlarmLimits();

        var token = _streamCts.Token;

        await Task.Run(async () =>
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
                    var (value, unit, timestamp) = TakeReading();
                    Dispatcher.Invoke(() => AddReading(value, unit, timestamp));
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"Stream error: {ex.Message}", "Error");
                        StopStream();
                    });
                    break;
                }

                var interval = Dispatcher.Invoke(() =>
                {
                    if (double.TryParse(txtInterval.Text, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var v))
                        return v;
                    return 0.5;
                });

                try { await Task.Delay(TimeSpan.FromSeconds(interval), token); }
                catch (TaskCanceledException) { break; }
            }
        });
    }

    private void StopStream()
    {
        _streaming = false;
        _streamCts?.Cancel();
        _streamCts?.Dispose();
        _streamCts = null;
        btnStream.Content = "\u25b6 Stream";
        btnStream.Style = (Style)FindResource("BtnAccent");
        btnSingle.IsEnabled = true;
        txtRateDisplay.Text = "";
    }

    // ─── Recording ──────────────────────────────────────────

    private void StartRecording()
    {
        var dlg = new SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"34410A_{_currentFunc.Replace(" ", "_")}_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            DefaultExt = ".csv"
        };

        if (dlg.ShowDialog() != true) return;

        _csvWriter = new StreamWriter(dlg.FileName, false, Encoding.UTF8);
        _csvWriter.WriteLine("#,Timestamp,Value,Unit,Function");
        _recording = true;
        btnRecord.Content = "\u25a0 Stop Rec";
        btnRecord.Style = (Style)FindResource("BtnStop");
        txtRecording.Text = $"REC: {System.IO.Path.GetFileName(dlg.FileName)}";
    }

    private void StopRecording()
    {
        _recording = false;
        _csvWriter?.Close();
        _csvWriter?.Dispose();
        _csvWriter = null;
        btnRecord.Content = "\u25cf Record CSV";
        btnRecord.Style = (Style)FindResource("BtnBase");
        txtRecording.Text = "";
    }

    // ─── Function Selection ─────────────────────────────────

    private void SelectFunction(string funcName)
    {
        _currentFunc = funcName;
        var func = Functions[funcName];

        // Update button styles
        foreach (var (name, btn) in _funcButtons)
            btn.Style = (Style)FindResource(name == funcName ? "BtnFuncActive" : "BtnFunc");

        txtFunction.Text = funcName;
        txtUnit.Text = func.Unit;
        UpdateRangeCombo();
        UpdateConfigVisibility();

        // Configure instrument
        if (_connected)
        {
            try
            {
                if (funcName == "Temperature")
                    Write($"CONF:TEMP {cmbProbe.SelectedItem},{cmbProbeType.SelectedItem}");
                else if (funcName is "Continuity" or "Diode")
                    Write(func.ConfCmd);
                else
                    Write($"{func.ConfCmd} AUTO");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Configure failed: {ex.Message}", "Error");
            }
        }
    }

    private void UpdateRangeCombo()
    {
        var func = Functions[_currentFunc];
        cmbRange.ItemsSource = func.Ranges;
        if (func.Ranges.Length > 0)
            cmbRange.SelectedIndex = 0;
    }

    private void UpdateConfigVisibility()
    {
        var func = Functions[_currentFunc];

        grpImpedance.Visibility = func.HasImpedance ? Visibility.Visible : Visibility.Collapsed;
        grpNull.Visibility = func.HasNull ? Visibility.Visible : Visibility.Collapsed;
        grpTemp.Visibility = _currentFunc == "Temperature" ? Visibility.Visible : Visibility.Collapsed;

        if (func.HasBandwidth)
        {
            grpBandwidth.Visibility = Visibility.Visible;
            grpBandwidth.Header = "AC Bandwidth";
            txtBwLabel.Text = "BW (Hz):";
            cmbBandwidth.ItemsSource = new[] { "3", "20", "200" };
            cmbBandwidth.SelectedItem = "20";
        }
        else if (func.HasAperture)
        {
            grpBandwidth.Visibility = Visibility.Visible;
            grpBandwidth.Header = "Gate Time / Aperture";
            txtBwLabel.Text = "Aperture (s):";
            cmbBandwidth.ItemsSource = new[] { "0.01", "0.1", "1" };
            cmbBandwidth.SelectedItem = "0.1";
        }
        else
        {
            grpBandwidth.Visibility = Visibility.Collapsed;
        }
    }

    // ─── Event Handlers ─────────────────────────────────────

    private void BtnRefresh_Click(object sender, RoutedEventArgs e) => PopulateResources();

    private void BtnConnect_Click(object sender, RoutedEventArgs e)
    {
        if (_connected)
            Disconnect();
        else
        {
            var resource = cmbResource.SelectedItem as string;
            if (string.IsNullOrEmpty(resource))
            { MessageBox.Show("Select a VISA resource.", "No Resource"); return; }
            Connect(resource);
        }
    }

    private void BtnFunc_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button btn && btn.Tag is string funcName)
            SelectFunction(funcName);
    }

    private void BtnAutoDetect_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) { ShowNotConnected(); return; }

        btnAutoDetect.IsEnabled = false;
        btnAutoDetect.Content = "Scanning...";
        txtValue.Text = "DETECTING...";

        Task.Run(() =>
        {
            string? detectedFunc = null;
            double detectedValue = 0;
            string detectedUnit = "";
            string detectedTimestamp = "";

            try
            {
                // Use CONF + fast NPLC + READ? for quick probing
                // Save/restore timeout
                var origTimeout = _session!.TimeoutMilliseconds;
                _session.TimeoutMilliseconds = 5000;

                // Probe DC Voltage first
                Dispatcher.Invoke(() => txtFunction.Text = "Probing DC Voltage...");
                double dcv = 0;
                bool hasDcv = false;
                try
                {
                    Write("CONF:VOLT:DC AUTO");
                    Write("VOLT:DC:NPLC 0.2");
                    var dcvResult = QueryFloat("READ?");
                    if (Math.Abs(dcvResult) < 9.8e37 && Math.Abs(dcvResult) > 0.0001)
                    {
                        dcv = dcvResult;
                        hasDcv = true;
                    }
                }
                catch { }

                // If DCV detected something, also check ACV to distinguish AC from DC
                if (hasDcv)
                {
                    Dispatcher.Invoke(() => txtFunction.Text = "Probing AC Voltage...");
                    try
                    {
                        Write("CONF:VOLT:AC AUTO");
                        Write("VOLT:AC:BAND 200");
                        _session.TimeoutMilliseconds = 8000; // AC needs more settling
                        var acvResult = QueryFloat("READ?");
                        if (Math.Abs(acvResult) < 9.8e37 && acvResult > 0.1)
                        {
                            // Significant AC component - it's an AC signal
                            detectedFunc = "AC Voltage";
                            detectedValue = acvResult;
                            detectedUnit = Functions["AC Voltage"].Unit;
                        }
                        else
                        {
                            // No significant AC - it's DC
                            detectedFunc = "DC Voltage";
                            detectedValue = dcv;
                            detectedUnit = Functions["DC Voltage"].Unit;
                        }
                    }
                    catch
                    {
                        // AC probe failed, use DC result
                        detectedFunc = "DC Voltage";
                        detectedValue = dcv;
                        detectedUnit = Functions["DC Voltage"].Unit;
                    }
                }

                _session.TimeoutMilliseconds = 5000;

                // If no voltage, try other functions
                if (detectedFunc == null)
                {
                    var otherProbes = new (string funcName, string confCmd, string nplcCmd, double overload, Func<double, bool> isValid)[]
                    {
                        ("2W Resistance", "CONF:RES AUTO", "RES:NPLC 0.2", 9.8e37, v => Math.Abs(v) < 9.8e37 && Math.Abs(v) > 0),
                        ("DC Current",    "CONF:CURR:DC AUTO", "CURR:DC:NPLC 0.2", 9.8e37, v => Math.Abs(v) < 9.8e37 && Math.Abs(v) > 0.000001),
                        ("Frequency",     "CONF:FREQ", "", 0, v => v > 0.5),
                        ("Capacitance",   "CONF:CAP AUTO", "", 9.8e37, v => Math.Abs(v) < 9.8e37 && Math.Abs(v) > 0),
                        ("Continuity",    "CONF:CONT", "", 9.8e37, v => Math.Abs(v) < 1000),
                        ("Diode",         "CONF:DIOD", "", 9.8e37, v => v > 0.1 && v < 5.0),
                    };

                    foreach (var (funcName, confCmd, nplcCmd, overload, isValid) in otherProbes)
                    {
                        try
                        {
                            Dispatcher.Invoke(() => txtFunction.Text = $"Probing {funcName}...");
                            Write(confCmd);
                            if (!string.IsNullOrEmpty(nplcCmd))
                                Write(nplcCmd);
                            var result = QueryFloat("READ?");

                            if (isValid(result))
                            {
                                detectedFunc = funcName;
                                detectedValue = result;
                                detectedUnit = Functions[funcName].Unit;
                                break;
                            }
                        }
                        catch { }
                    }
                }

                detectedTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                _session.TimeoutMilliseconds = origTimeout;
            }
            catch { }

            Dispatcher.Invoke(() =>
            {
                if (detectedFunc != null)
                {
                    SelectFunction(detectedFunc);
                    AddReading(detectedValue, detectedUnit, detectedTimestamp);

                    // Auto-start streaming after successful detection
                    if (!_streaming)
                        StartStream();
                }
                else
                {
                    SelectFunction("DC Voltage");
                    txtValue.Text = "No Signal";
                }

                btnAutoDetect.IsEnabled = true;
                btnAutoDetect.Content = "\U0001f50d AUTO";
            });
        });
    }

    private void BtnSingle_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) { ShowNotConnected(); return; }
        Task.Run(() =>
        {
            try
            {
                var (value, unit, timestamp) = TakeReading();
                Dispatcher.Invoke(() => AddReading(value, unit, timestamp));
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => MessageBox.Show($"Read error: {ex.Message}", "Error"));
            }
        });
    }

    private void BtnDigitize_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) { ShowNotConnected(); return; }

        if (_scopeRunning)
        {
            _scopeCts?.Cancel();
            return;
        }

        if (_streaming) StopStream();

        _scopeRunning = true;
        _scopeCts = new CancellationTokenSource();
        btnDigitize.Content = "\u25a0 STOP";
        btnDigitize.Background = new SolidColorBrush(Color.FromRgb(0xe5, 0x39, 0x35));
        btnSingle.IsEnabled = false;
        btnStream.IsEnabled = false;

        var token = _scopeCts.Token;

        Task.Run(async () =>
        {
            var origTimeout = _session!.TimeoutMilliseconds;
            try
            {
                var cycles = Dispatcher.Invoke(() =>
                {
                    if (int.TryParse(txtScopeCycles.Text, out var c) && c > 0 && c <= 50)
                        return c;
                    return 3;
                });

                _session.TimeoutMilliseconds = 30000;

                // Enable logging for scope debug
                _scopeLogging = true;
                Log("=== SCOPE START ===");
                Log($"Requested cycles: {cycles}");

                // Configure digitize mode - maximize sample rate
                ConfigureDigitize(cycles);

                // Sample rate is fixed by timer setting

                // Known sample rate from timer setting
                double sampleRate = 1.0 / 0.0002; // 5000 S/s
                // Acquisition time = samples * interval
                int acqTimeMs = (int)(_scopeSampleCount * 0.0002 * 1000) + 200;
                Log($"Acquisition time: {acqTimeMs}ms for {_scopeSampleCount} samples at {sampleRate} S/s");

                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        // Capture: INIT, wait for ALL samples, then read all data
                        Write("ABORT");
                        Write("INITIATE");
                        Thread.Sleep(acqTimeMs);

                        // Read ALL samples - the GPIB buffer may chunk the response
                        // Use FETCH? and read in a loop until we have all data
                        _session!.RawIO.Write("FETCH?\n");
                        var sb = new StringBuilder();
                        int totalRead = 0;
                        while (totalRead < _scopeSampleCount)
                        {
                            var chunk = _session.RawIO.ReadString();
                            sb.Append(chunk);
                            // Count commas to know how many readings we got
                            totalRead = sb.ToString().Split(',').Length;
                            if (chunk.Contains("\n") || chunk.EndsWith("\n")) break;
                        }
                        var rawData = sb.ToString().Trim();
                        Log($">> FETCH? (read {rawData.Split(',').Length} values)");

                        var samples = ParseSamples(rawData);
                        if (samples.Length < 10) { Log($"Too few samples: {samples.Length}"); continue; }

                        Log($"Captured {samples.Length} samples (expected {_scopeSampleCount}) at {sampleRate:F0} S/s");
                        Log($"Min={samples.Min():F2} Max={samples.Max():F2} Range={samples.Max()-samples.Min():F2}");

                        var freq = DetectFrequency(samples, sampleRate);
                        Log($"Detected freq: {freq:F1} Hz");
                        var rms = Math.Sqrt(samples.Select(s => s * s).Average());
                        var pk = samples.Max() - samples.Min();

                        // Trim to whole cycles, then apply time base zoom
                        double[] displaySamples = samples;
                        if (freq > 10 && freq < 1000)
                        {
                            int samplesPerCycle = (int)(sampleRate / freq);
                            int wholeCycles = samples.Length / samplesPerCycle;
                            if (wholeCycles >= 1)
                            {
                                int trimCount = wholeCycles * samplesPerCycle;
                                displaySamples = samples.Take(trimCount).ToArray();
                            }
                        }

                        Dispatcher.Invoke(() =>
                        {
                            // Time base slider: 100 = show all, lower = zoom in
                            var zoom = sldTimeBase.Value / 100.0;
                            if (zoom < 0.05) zoom = 0.05;
                            int showCount = (int)(displaySamples.Length * zoom);
                            if (showCount < 20) showCount = 20;
                            if (showCount > displaySamples.Length) showCount = displaySamples.Length;

                            var zoomedSamples = displaySamples.Take(showCount).ToArray();
                            double displayTimeMs = showCount / sampleRate * 1000;

                            // Update time base label
                            if (zoom >= 0.99)
                                txtTimeBaseVal.Text = "Auto";
                            else
                                txtTimeBaseVal.Text = $"{displayTimeMs / 10:F1}ms";

                            txtValue.Text = FormatValue(rms);
                            txtUnit.Text = "V RMS";
                            txtFunction.Text = $"Scope - {freq:F1} Hz  ({sampleRate:F0} S/s)";
                            txtRateDisplay.Text = $"[Pk-Pk: {FormatValue(pk)}]";
                            DrawScopeWaveform(zoomedSamples, sampleRate, freq);
                        });
                    }
                    catch (Exception ex)
                    {
                        Log($"SCOPE ERROR: {ex.Message}");
                        Dispatcher.Invoke(() =>
                        {
                            txtValue.Text = "SCOPE ERR";
                            txtScopeFunc.Text = ex.Message;
                        });
                        break;
                    }

                }
            }
            finally
            {
                Log("=== SCOPE STOP ===");
                _scopeLogging = false;
                // Restore instrument to normal mode
                try
                {
                    Write("ABORT");
                    _session!.TimeoutMilliseconds = origTimeout;
                    Write("SAMPLE:COUNT 1");
                    Write("SAMPLE:SOURCE IMMEDIATE");
                    Write("SYSTEM:BEEPER:STATE ON");    // Restore beeper
                    Write("DISPLAY:ENABLE ON");          // Restore display
                }
                catch { }

                Dispatcher.Invoke(() =>
                {
                    _scopeRunning = false;
                    btnDigitize.Content = "\u2588 SCOPE";
                    btnDigitize.Background = new SolidColorBrush(Color.FromRgb(0x00, 0xc8, 0x53));
                    btnSingle.IsEnabled = true;
                    btnStream.IsEnabled = true;
                });
            }
        });
    }

    private int _scopeSampleCount;

    private void ConfigureDigitize(int cycles)
    {
        Write("*RST");
        Thread.Sleep(500);                         // Wait for reset to complete
        Write("*CLS");

        // Silence the meter
        Write("DISP OFF");
        Write("SYST:BEEP:STAT OFF");

        // Configure DC voltage - this resets measurement subsystem
        Write("CONF:VOLT:DC 1000");                // Fixed 1000V range

        // Speed settings - MUST come AFTER CONF
        Write("VOLT:DC:NPLC 0.006");               // Fastest NPLC
        Write("VOLT:DC:ZERO:AUTO OFF");             // No auto-zero

        // Verify NPLC was set
        var nplc = Query("VOLT:DC:NPLC?");
        System.Diagnostics.Debug.WriteLine($"NPLC set to: {nplc}");

        // Trigger - immediate, no delay
        Write("TRIG:SOUR IMM");
        Write("TRIG:DEL:AUTO OFF");
        Write("TRIG:DEL 0");
        Write("TRIG:COUN 1");

        // Timer-based sampling - buffers ALL samples internally before read
        // Timer MUST be > aperture time. NPLC 0.006 at 50Hz = 120us aperture
        // Use 200us (0.0002s) = 5000 readings/sec - safely above aperture
        Write("SAMP:SOUR TIM");
        Write("SAMP:TIM 0.0002");

        // At 5000 S/s, 50Hz = 100 samples/cycle
        _scopeSampleCount = 100 * cycles;
        if (_scopeSampleCount > 50000) _scopeSampleCount = 50000;
        if (_scopeSampleCount < 200) _scopeSampleCount = 200;
        Write($"SAMP:COUN {_scopeSampleCount}");

        // Verify settings
        var nplcVal = Query("VOLT:DC:NPLC?");
        var sampSrc = Query("SAMP:SOUR?");
        var sampTim = Query("SAMP:TIM?");
        var sampCnt = Query("SAMP:COUN?");
        var err = Query("SYST:ERR?");
        Log($"Verified: NPLC={nplcVal}, SAMP:SRC={sampSrc}, SAMP:TIM={sampTim}, SAMP:COUN={sampCnt}, ERR={err}");
    }

    private double[] ParseSamples(string rawData)
    {
        var parts = rawData.Split(',');
        var samples = new double[parts.Length];
        for (int i = 0; i < parts.Length; i++)
        {
            if (double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var v))
                samples[i] = v;
        }
        return samples;
    }

    private double DetectFrequency(double[] samples, double sampleRate)
    {
        if (samples.Length < 10) return 0;

        // Remove DC offset
        var mean = samples.Average();
        var zeroCrossings = 0;
        for (int i = 1; i < samples.Length; i++)
        {
            if ((samples[i - 1] - mean) * (samples[i] - mean) < 0)
                zeroCrossings++;
        }

        // Each full cycle has 2 zero crossings
        var duration = samples.Length / sampleRate;
        var freq = zeroCrossings / (2.0 * duration);
        return freq;
    }

    private void DrawScopeWaveform(double[] samples, double sampleRate, double frequency)
    {
        chartCanvas.Children.Clear();
        yAxisCanvas.Children.Clear();
        xAxisCanvas.Children.Clear();

        if (samples.Length < 2) return;

        var w = chartCanvas.ActualWidth;
        var h = chartCanvas.ActualHeight;
        if (w < 20 || h < 20) return;

        // Amplitude range
        var minVal = samples.Min();
        var maxVal = samples.Max();
        var ampRange = maxVal - minVal;
        if (ampRange < 0.001) ampRange = 1;

        // Nice Y-axis divisions with generous padding (25%)
        var divY = CalculateNiceDivision(ampRange, 6);
        var center = (minVal + maxVal) / 2.0;
        // Always add at least 1 extra division above and below the signal
        var halfDivs = (int)Math.Ceiling(((maxVal - center) / divY)) + 1;
        var yTop = center + halfDivs * divY;
        var yBottom = center - halfDivs * divY;
        if (yTop <= yBottom) yTop = yBottom + divY;
        var fullYRange = yTop - yBottom;

        // Time range
        var totalTime = samples.Length / sampleRate;
        var divX = CalculateNiceDivision(totalTime, 10);
        int numGridX = 10;
        int numGridY = (int)Math.Round(fullYRange / divY);
        if (numGridY < 2) numGridY = 4;
        if (numGridY > 12) numGridY = 8;

        // Scope graticule colors
        var gridBrush = new SolidColorBrush(Color.FromRgb(0x0c, 0x2a, 0x1e));
        var centerBrush = new SolidColorBrush(Color.FromRgb(0x14, 0x4a, 0x30));
        var dotBrush = new SolidColorBrush(Color.FromRgb(0x18, 0x55, 0x38));
        var borderBrush = new SolidColorBrush(Color.FromRgb(0x0a, 0x3a, 0x28));

        // Outer border ticks on all edges
        for (int i = 0; i <= numGridX; i++)
        {
            var x = w * i / numGridX;
            // Top tick
            chartCanvas.Children.Add(new Line { X1 = x, Y1 = 0, X2 = x, Y2 = 4, Stroke = borderBrush, StrokeThickness = 1 });
            // Bottom tick
            chartCanvas.Children.Add(new Line { X1 = x, Y1 = h - 4, X2 = x, Y2 = h, Stroke = borderBrush, StrokeThickness = 1 });
        }
        for (int i = 0; i <= numGridY; i++)
        {
            var y = h * i / (double)numGridY;
            // Left tick
            chartCanvas.Children.Add(new Line { X1 = 0, Y1 = y, X2 = 4, Y2 = y, Stroke = borderBrush, StrokeThickness = 1 });
            // Right tick
            chartCanvas.Children.Add(new Line { X1 = w - 4, Y1 = y, X2 = w, Y2 = y, Stroke = borderBrush, StrokeThickness = 1 });
        }

        // Horizontal grid lines + Y-axis labels
        for (int i = 0; i <= numGridY; i++)
        {
            var y = h - (h * i / (double)numGridY);
            var yVal = yBottom + (fullYRange * i / numGridY);

            var isCenter = Math.Abs(i - numGridY / 2.0) < 0.6;
            chartCanvas.Children.Add(new Line
            {
                X1 = 0, Y1 = y, X2 = w, Y2 = y,
                Stroke = isCenter ? centerBrush : gridBrush,
                StrokeThickness = isCenter ? 1.2 : 0.4,
                StrokeDashArray = isCenter ? null : new DoubleCollection([2.0, 4.0])
            });

            // Graticule dots at intersections
            for (int d = 1; d < numGridX; d++)
            {
                var dotX = w * d / numGridX;
                var tick = new Ellipse { Width = 2, Height = 2, Fill = dotBrush };
                Canvas.SetLeft(tick, dotX - 1);
                Canvas.SetTop(tick, y - 1);
                chartCanvas.Children.Add(tick);
            }

            // Y-axis label
            var label = new System.Windows.Controls.TextBlock
            {
                Text = FormatAxisValue(yVal),
                Foreground = new SolidColorBrush(Color.FromRgb(0x00, 0xe6, 0x76)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 10,
                TextAlignment = TextAlignment.Right,
                Width = 54
            };
            Canvas.SetRight(label, 2);
            Canvas.SetTop(label, y - 8);
            yAxisCanvas.Children.Add(label);
        }

        // Vertical grid lines
        for (int i = 0; i <= numGridX; i++)
        {
            var x = w * i / numGridX;
            var isCenter = Math.Abs(i - numGridX / 2.0) < 0.6;
            chartCanvas.Children.Add(new Line
            {
                X1 = x, Y1 = 0, X2 = x, Y2 = h,
                Stroke = isCenter ? centerBrush : gridBrush,
                StrokeThickness = isCenter ? 1.2 : 0.4,
                StrokeDashArray = isCenter ? null : new DoubleCollection([2.0, 4.0])
            });

            for (int d = 1; d < numGridY; d++)
            {
                var dotY = h * d / numGridY;
                var tick = new Ellipse { Width = 2, Height = 2, Fill = dotBrush };
                Canvas.SetLeft(tick, x - 1);
                Canvas.SetTop(tick, dotY - 1);
                chartCanvas.Children.Add(tick);
            }
        }

        // X-axis time labels
        var xAxisW = xAxisCanvas.ActualWidth > 0 ? xAxisCanvas.ActualWidth : w;
        for (int i = 0; i <= numGridX; i += 2)
        {
            var tVal = totalTime * i / numGridX;
            var label = new System.Windows.Controls.TextBlock
            {
                Text = $"{tVal * 1000:F1}ms",
                Foreground = new SolidColorBrush(Color.FromRgb(0x42, 0xa5, 0xf5)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 10,
                TextAlignment = TextAlignment.Center
            };
            Canvas.SetLeft(label, xAxisW * i / numGridX - 20);
            Canvas.SetTop(label, 1);
            xAxisCanvas.Children.Add(label);
        }

        // Zero line (0V reference) - brighter dashed line
        var zeroY = h - ((0 - yBottom) / fullYRange * h);
        if (zeroY > 5 && zeroY < h - 5)
        {
            chartCanvas.Children.Add(new Line
            {
                X1 = 0, Y1 = zeroY, X2 = w, Y2 = zeroY,
                Stroke = new SolidColorBrush(Color.FromRgb(0x30, 0x70, 0x55)),
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection([6.0, 3.0])
            });
            // "0V" label on zero line
            var zLabel = new System.Windows.Controls.TextBlock
            {
                Text = "0V",
                Foreground = new SolidColorBrush(Color.FromArgb(0x80, 0x00, 0xe6, 0x76)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 8
            };
            Canvas.SetLeft(zLabel, 4);
            Canvas.SetTop(zLabel, zeroY - 12);
            chartCanvas.Children.Add(zLabel);
        }

        // Waveform rendering - triple-layer phosphor glow effect
        var outerGlow = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromArgb(0x10, 0x00, 0xff, 0x88)),
            StrokeThickness = 10,
            StrokeLineJoin = PenLineJoin.Round
        };
        var midGlow = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromArgb(0x28, 0x00, 0xff, 0x88)),
            StrokeThickness = 5,
            StrokeLineJoin = PenLineJoin.Round
        };
        var innerGlow = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromArgb(0x50, 0x40, 0xff, 0xa0)),
            StrokeThickness = 3,
            StrokeLineJoin = PenLineJoin.Round
        };
        var trace = new Polyline
        {
            Stroke = new SolidColorBrush(Color.FromRgb(0x00, 0xff, 0x88)),
            StrokeThickness = 2,
            StrokeLineJoin = PenLineJoin.Round
        };

        // Build waveform points - downsample if too many
        var step = Math.Max(1, samples.Length / (int)w);
        for (int i = 0; i < samples.Length; i += step)
        {
            var x = (i / (double)(samples.Length - 1)) * w;
            var y = h - ((samples[i] - yBottom) / fullYRange * h);
            y = Math.Max(1, Math.Min(h - 1, y));
            var pt = new Point(x, y);
            outerGlow.Points.Add(pt);
            midGlow.Points.Add(pt);
            innerGlow.Points.Add(pt);
            trace.Points.Add(pt);
        }

        chartCanvas.Children.Add(outerGlow);
        chartCanvas.Children.Add(midGlow);
        chartCanvas.Children.Add(innerGlow);
        chartCanvas.Children.Add(trace);

        // Trigger level line
        var trigPct = sldTrigLevel.Value / 100.0;
        var trigVal = yBottom + trigPct * fullYRange;
        var trigY = h - ((trigVal - yBottom) / fullYRange * h);
        if (trigY > 0 && trigY < h)
        {
            chartCanvas.Children.Add(new Line
            {
                X1 = 0, Y1 = trigY, X2 = w, Y2 = trigY,
                Stroke = new SolidColorBrush(Color.FromRgb(0xff, 0x60, 0x20)),
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection([4.0, 3.0])
            });
            // Trigger arrow on left edge
            var trigLabel = new System.Windows.Controls.TextBlock
            {
                Text = $"T {FormatAxisValue(trigVal)}V",
                Foreground = new SolidColorBrush(Color.FromRgb(0xff, 0x60, 0x20)),
                FontFamily = new FontFamily("Consolas"),
                FontSize = 8
            };
            Canvas.SetLeft(trigLabel, 4);
            Canvas.SetTop(trigLabel, trigY + 2);
            chartCanvas.Children.Add(trigLabel);
        }

        // Cursors
        if (chkCursors.IsChecked == true)
        {
            var c1Pct = sldCursor1.Value / 100.0;
            var c2Pct = sldCursor2.Value / 100.0;
            var c1X = c1Pct * w;
            var c2X = c2Pct * w;

            // Cursor 1 - yellow
            chartCanvas.Children.Add(new Line
            {
                X1 = c1X, Y1 = 0, X2 = c1X, Y2 = h,
                Stroke = new SolidColorBrush(Color.FromRgb(0xff, 0xd5, 0x40)),
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection([3.0, 2.0])
            });
            // Cursor 2 - cyan
            chartCanvas.Children.Add(new Line
            {
                X1 = c2X, Y1 = 0, X2 = c2X, Y2 = h,
                Stroke = new SolidColorBrush(Color.FromRgb(0x40, 0xe0, 0xff)),
                StrokeThickness = 1,
                StrokeDashArray = new DoubleCollection([3.0, 2.0])
            });

            // Calculate cursor measurements
            int c1Idx = (int)(c1Pct * (samples.Length - 1));
            int c2Idx = (int)(c2Pct * (samples.Length - 1));
            c1Idx = Math.Clamp(c1Idx, 0, samples.Length - 1);
            c2Idx = Math.Clamp(c2Idx, 0, samples.Length - 1);
            double c1V = samples[c1Idx];
            double c2V = samples[c2Idx];
            double dt = Math.Abs(c2Idx - c1Idx) / sampleRate * 1000; // ms
            double dv = c2V - c1V;

            txtCursorInfo.Text = $"\u0394t={dt:F1}ms \u0394V={dv:F1}V";
        }
        else
        {
            txtCursorInfo.Text = "";
        }

        // Scope header info
        txtScopeFunc.Text = $"{frequency:F1} Hz";
        txtScopeAmpl.Text = $"{FormatAxisValue(divY)}V/div";
        txtScopeTime.Text = $"{totalTime * 1000 / numGridX:F1}ms/div";
        txtScopePoints.Text = $"{samples.Length} pts | {sampleRate:F0} S/s";
    }

    private void BtnStream_Click(object sender, RoutedEventArgs e)
    {
        if (_streaming) StopStream(); else StartStream();
    }

    private void BtnRecord_Click(object sender, RoutedEventArgs e)
    {
        if (_recording) StopRecording(); else StartRecording();
    }

    private void BtnClear_Click(object sender, RoutedEventArgs e)
    {
        _readings.Clear();
        _readingCount = 0;
        _statSum = 0;
        _statMin = null;
        _statMax = null;
        _chartData.Clear();
        chartCanvas.Children.Clear();
        txtStatCount.Text = "Count: 0";
        txtStatMin.Text = "Min: ---";
        txtStatMax.Text = "Max: ---";
        txtStatAvg.Text = "Avg: ---";
        txtStatPtp.Text = "P-P: ---";
        txtScopeFunc.Text = "";
        txtScopeAmpl.Text = "";
        txtScopeTime.Text = "";
        txtScopePoints.Text = "";
        txtRateDisplay.Text = "";
        _rateCount = 0;
        _rateWindowStart = DateTime.Now;
    }

    private void BtnExport_Click(object sender, RoutedEventArgs e)
    {
        if (_readings.Count == 0) { MessageBox.Show("No data to export."); return; }

        var dlg = new SaveFileDialog
        {
            Filter = "CSV Files|*.csv|All Files|*.*",
            FileName = $"34410A_export_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
            DefaultExt = ".csv"
        };

        if (dlg.ShowDialog() != true) return;

        using var writer = new StreamWriter(dlg.FileName, false, Encoding.UTF8);
        writer.WriteLine("#,Timestamp,Value,Unit,Function");
        foreach (var r in _readings)
            writer.WriteLine($"{r.Number},\"{r.Timestamp}\",{r.ValueStr},\"{r.Unit}\",\"{r.Function}\"");

        MessageBox.Show($"Exported {_readings.Count} readings to:\n{dlg.FileName}", "Export Complete");
    }

    // Config handlers
    private void BtnSetRange_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        var range = cmbRange.SelectedItem as string;
        if (string.IsNullOrEmpty(range)) return;
        try
        {
            if (range == "AUTO")
            {
                Write($"{func.SensePrefix}:RANG:AUTO ON");
                chkAutoRange.IsChecked = true;
            }
            else
            {
                Write($"{func.SensePrefix}:RANG:AUTO OFF");
                Write($"{func.SensePrefix}:RANG {range}");
                chkAutoRange.IsChecked = false;
            }
        }
        catch (Exception ex) { ShowError("Set range", ex); }
    }

    private void ChkAutoRange_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        var state = chkAutoRange.IsChecked == true ? "ON" : "OFF";
        try { Write($"{func.SensePrefix}:RANG:AUTO {state}"); }
        catch (Exception ex) { ShowError("Auto range", ex); }
    }

    private void BtnSetNplc_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        if (!func.HasNplc) return;
        try { Write($"{func.SensePrefix}:NPLC {cmbNplc.SelectedItem}"); }
        catch (Exception ex) { ShowError("Set NPLC", ex); }
    }

    private void ChkAutoZero_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        var state = chkAutoZero.IsChecked == true ? "ON" : "OFF";
        try { Write($"{func.SensePrefix}:ZERO:AUTO {state}"); }
        catch (Exception ex) { ShowError("Auto zero", ex); }
    }

    private void ChkHighZ_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var state = chkHighZ.IsChecked == true ? "ON" : "OFF";
        try { Write($"VOLT:DC:IMP:AUTO {state}"); }
        catch (Exception ex) { ShowError("Impedance", ex); }
    }

    private void BtnSetBandwidth_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        var val = cmbBandwidth.SelectedItem as string;
        if (string.IsNullOrEmpty(val)) return;
        try
        {
            if (func.HasBandwidth) Write($"{func.SensePrefix}:BAND {val}");
            else if (func.HasAperture) Write($"{func.SensePrefix}:APER {val}");
        }
        catch (Exception ex) { ShowError("Bandwidth/Aperture", ex); }
    }

    private void ChkNull_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        var state = chkNull.IsChecked == true ? "ON" : "OFF";
        try { Write($"{func.SensePrefix}:NULL {state}"); }
        catch (Exception ex) { ShowError("Null", ex); }
    }

    private void BtnSetNull_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var func = Functions[_currentFunc];
        try { Write($"{func.SensePrefix}:NULL:VAL {txtNullValue.Text}"); }
        catch (Exception ex) { ShowError("Set null value", ex); }
    }

    private void BtnApplyTemp_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        try
        {
            Write($"CONF:TEMP {cmbProbe.SelectedItem},{cmbProbeType.SelectedItem}");
            Write($"UNIT:TEMP {cmbTempUnit.SelectedItem}");
            var unit = cmbTempUnit.SelectedItem as string;
            Functions["Temperature"].Unit = unit switch
            {
                "F" => "\u00b0F",
                "K" => "K",
                _ => "\u00b0C"
            };
            txtUnit.Text = Functions["Temperature"].Unit;
        }
        catch (Exception ex) { ShowError("Temperature config", ex); }
    }

    private void BtnApplyTrigger_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        try
        {
            Write($"TRIG:SOUR {cmbTrigSource.SelectedItem}");
            var delay = txtTrigDelay.Text.Trim();
            if (delay.Equals("AUTO", StringComparison.OrdinalIgnoreCase))
                Write("TRIG:DEL:AUTO ON");
            else
            {
                Write("TRIG:DEL:AUTO OFF");
                Write($"TRIG:DEL {delay}");
            }
            Write($"TRIG:COUN {txtTrigCount.Text}");
            Write($"SAMP:COUN {txtSampCount.Text}");
        }
        catch (Exception ex) { ShowError("Trigger config", ex); }
    }

    private void BtnApplyMath_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        try
        {
            var func = cmbMath.SelectedItem as string;
            if (func == "OFF")
                Write("CALC:STAT OFF");
            else
            {
                Write($"CALC:FUNC {func}");
                Write("CALC:STAT ON");
            }
        }
        catch (Exception ex) { ShowError("Math config", ex); }
    }

    // System handlers
    private void BtnReset_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        StopStream();
        Write("*RST");
        Write("*CLS");
        txtValue.Text = "RESET";
        var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.5) };
        timer.Tick += (_, _) => { txtValue.Text = "---"; timer.Stop(); };
        timer.Start();
    }

    private void BtnSelfTest_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        txtValue.Text = "TESTING...";
        Task.Run(() =>
        {
            try
            {
                var result = Query("*TST?");
                var pass = result.Trim() == "0";
                Dispatcher.Invoke(() =>
                {
                    var status = pass ? "PASS" : $"FAIL ({result})";
                    txtValue.Text = status;
                    MessageBox.Show($"Self-test: {status}", "Self Test");
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => ShowError("Self test", ex));
            }
        });
    }

    private void BtnInstInfo_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) { ShowNotConnected(); return; }

        Task.Run(() =>
        {
            try
            {
                var idn = Query("*IDN?");
                var parts = idn.Split(',');
                var manufacturer = parts.Length > 0 ? parts[0].Trim() : "Unknown";
                var model = parts.Length > 1 ? parts[1].Trim() : "Unknown";
                var serial = parts.Length > 2 ? parts[2].Trim() : "Unknown";
                var firmware = parts.Length > 3 ? parts[3].Trim() : "Unknown";

                string calStr;
                try { calStr = Query("CAL:STR?"); }
                catch { calStr = "N/A"; }

                string calCount;
                try { calCount = Query("CAL:COUN?"); }
                catch { calCount = "N/A"; }

                string scpiVer;
                try { scpiVer = Query("SYST:VERS?"); }
                catch { scpiVer = "N/A"; }

                string terminals;
                try { terminals = Query("ROUT:TERM?"); }
                catch { terminals = "N/A"; }

                Dispatcher.Invoke(() =>
                {
                    var info = $"Manufacturer: {manufacturer}\n" +
                               $"Model: {model}\n" +
                               $"Serial: {serial}\n" +
                               $"Firmware: {firmware}\n" +
                               $"SCPI Version: {scpiVer}\n" +
                               $"Terminal Block: {terminals}\n" +
                               $"Cal String: {calStr}\n" +
                               $"Cal Count: {calCount}";
                    MessageBox.Show(info, "Instrument Information", MessageBoxButton.OK, MessageBoxImage.Information);
                });
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => ShowError("Instrument info", ex));
            }
        });
    }

    private void BtnErrors_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        var errors = new List<string>();
        try
        {
            while (true)
            {
                var err = Query("SYST:ERR?");
                if (err.StartsWith("+0") || err.StartsWith("0")) break;
                errors.Add(err);
                if (errors.Count > 20) break;
            }
        }
        catch (Exception ex) { ShowError("Check errors", ex); return; }

        MessageBox.Show(errors.Count > 0 ? string.Join("\n", errors) : "No errors.",
            "Error Queue", MessageBoxButton.OK,
            errors.Count > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
    }

    private void BtnBeep_Click(object sender, RoutedEventArgs e)
    {
        if (_connected) try { Write("SYST:BEEP"); } catch { }
    }

    private void BtnSetDisplay_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        try { Write($"DISP:WIND:TEXT \"{txtDisplayText.Text}\""); }
        catch (Exception ex) { ShowError("Display text", ex); }
    }

    private void BtnClearDisplay_Click(object sender, RoutedEventArgs e)
    {
        if (!_connected) return;
        try { Write("DISP:WIND:TEXT:CLE"); }
        catch (Exception ex) { ShowError("Clear display", ex); }
    }

    // Helpers
    private void ShowNotConnected() =>
        MessageBox.Show("Connect to the instrument first.", "Not Connected",
            MessageBoxButton.OK, MessageBoxImage.Warning);

    private void ShowError(string action, Exception ex) =>
        MessageBox.Show($"{action} failed:\n{ex.Message}", "Error",
            MessageBoxButton.OK, MessageBoxImage.Error);

    private void Window_Closing(object? sender, CancelEventArgs e)
    {
        _scopeCts?.Cancel();
        StopStream();
        StopRecording();
        _session?.Dispose();
    }
}
