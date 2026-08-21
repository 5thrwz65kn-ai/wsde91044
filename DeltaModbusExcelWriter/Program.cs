using ClosedXML.Excel;
using System.Buffers.Binary;
using System.Globalization;
using System.Net.Sockets;

namespace DeltaModbusExcelWriter;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }
}

internal sealed class WriteItem
{
    public int ExcelRow { get; init; }
    public int SourceAddress { get; init; }
    public ushort RegisterAddress { get; init; }
    public int Value { get; init; }
    public string Status { get; set; } = "等待";
}

internal sealed class MainForm : Form
{
    private readonly TextBox ipBox = new() { Text = "192.168.1.10", Width = 135 };
    private readonly NumericUpDown portBox = new() { Minimum = 1, Maximum = 65535, Value = 502, Width = 75 };
    private readonly NumericUpDown unitBox = new() { Minimum = 0, Maximum = 247, Value = 1, Width = 65 };
    private readonly NumericUpDown timeoutBox = new() { Minimum = 1, Maximum = 60, Value = 3, Width = 60 };
    private readonly NumericUpDown delayBox = new() { Minimum = 0, Maximum = 10000, Value = 50, Width = 75 };
    private readonly TextBox fileBox = new() { ReadOnly = true, Dock = DockStyle.Fill };
    private readonly ComboBox addressMode = new() { DropDownStyle = ComboBoxStyle.DropDownList, Width = 260 };
    private readonly DataGridView grid = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        AllowUserToAddRows = false,
        AllowUserToDeleteRows = false,
        RowHeadersVisible = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect,
        AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
    };
    private readonly Button importButton = new() { Text = "選擇 Excel", AutoSize = true };
    private readonly Button startButton = new() { Text = "開始逐列寫入", AutoSize = true };
    private readonly Button stopButton = new() { Text = "停止", AutoSize = true, Enabled = false };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Fill };
    private readonly Label statusLabel = new() { Text = "尚未匯入 Excel", AutoSize = true, TextAlign = ContentAlignment.MiddleRight };
    private readonly List<WriteItem> items = [];
    private CancellationTokenSource? cancellation;

    private static readonly string[] AddressModes =
    [
        "自動（40001→0，其餘不偏移）",
        "0-based（0 = 第一個暫存器）",
        "1-based（1 = 第一個暫存器）",
        "4xxxx（40001 = 第一個暫存器）"
    ];

    public MainForm()
    {
        Text = "Excel → Modbus TCP 寫入工具";
        Width = 980;
        Height = 680;
        MinimumSize = new Size(850, 560);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Microsoft JhengHei UI", 9F);
        BuildUi();
        addressMode.Items.AddRange(AddressModes);
        addressMode.SelectedIndex = 0;
        importButton.Click += (_, _) => ChooseExcel();
        startButton.Click += async (_, _) => await StartWritingAsync();
        stopButton.Click += (_, _) => cancellation?.Cancel();
        addressMode.SelectedIndexChanged += (_, _) =>
        {
            if (File.Exists(fileBox.Text)) LoadExcel(fileBox.Text);
        };
    }

    private void BuildUi()
    {
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), RowCount = 5, ColumnCount = 1 };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var connection = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), WrapContents = true };
        AddPair(connection, "目標 IP", ipBox);
        AddPair(connection, "Port", portBox);
        AddPair(connection, "站號", unitBox);
        AddPair(connection, "逾時(秒)", timeoutBox);
        AddPair(connection, "逐筆間隔(ms)", delayBox);
        root.Controls.Add(WrapGroup("連線設定", connection), 0, 0);

        var importLayout = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), ColumnCount = 4 };
        importLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        importLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        importLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        importLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        importLayout.Controls.Add(fileBox, 0, 0);
        importLayout.Controls.Add(importButton, 1, 0);
        importLayout.Controls.Add(new Label { Text = "地址格式", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 4, 0) }, 2, 0);
        importLayout.Controls.Add(addressMode, 3, 0);
        root.Controls.Add(WrapGroup("Excel 匯入", importLayout), 0, 1);

        root.Controls.Add(new Label
        {
            Text = "Excel：A欄＝通訊地址，B欄＝寫入數值；第一列可放標題。固定使用 Holding Register / FC06。",
            Dock = DockStyle.Fill,
            Padding = new Padding(4, 7, 4, 7)
        }, 0, 2);

        grid.Columns.Add("ExcelRow", "Excel列");
        grid.Columns.Add("SourceAddress", "原始地址");
        grid.Columns.Add("RegisterAddress", "實際位址 (0-based)");
        grid.Columns.Add("Value", "寫入數值");
        grid.Columns.Add("Status", "狀態");
        grid.Columns[0].FillWeight = 60;
        grid.Columns[1].FillWeight = 90;
        grid.Columns[2].FillWeight = 120;
        grid.Columns[3].FillWeight = 90;
        grid.Columns[4].FillWeight = 180;
        root.Controls.Add(grid, 0, 3);

        var bottom = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 5, Padding = new Padding(0, 8, 0, 0) };
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottom.Controls.Add(startButton, 0, 0);
        bottom.Controls.Add(stopButton, 1, 0);
        bottom.Controls.Add(new Panel(), 2, 0);
        bottom.Controls.Add(progress, 3, 0);
        bottom.Controls.Add(statusLabel, 4, 0);
        root.Controls.Add(bottom, 0, 4);

        Controls.Add(root);
    }

    private static GroupBox WrapGroup(string title, Control content)
    {
        var box = new GroupBox { Text = title, Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6) };
        box.Controls.Add(content);
        return box;
    }

    private static void AddPair(FlowLayoutPanel panel, string text, Control control)
    {
        panel.Controls.Add(new Label { Text = text, AutoSize = true, Margin = new Padding(8, 7, 4, 0) });
        panel.Controls.Add(control);
    }

    private void ChooseExcel()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "選擇 Excel",
            Filter = "Excel 活頁簿 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
            CheckFileExists = true
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            fileBox.Text = dialog.FileName;
            LoadExcel(dialog.FileName);
        }
    }

    private void LoadExcel(string path)
    {
        try
        {
            var loaded = ParseExcel(path, addressMode.SelectedItem?.ToString() ?? AddressModes[0]);
            items.Clear();
            items.AddRange(loaded);
            RefreshGrid();
            statusLabel.Text = $"已載入 {items.Count} 筆";
        }
        catch (Exception ex)
        {
            items.Clear();
            RefreshGrid();
            MessageBox.Show(this, ex.Message, "Excel 格式錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static List<WriteItem> ParseExcel(string path, string mode)
    {
        using var workbook = new XLWorkbook(path);
        var sheet = workbook.Worksheets.First();
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var result = new List<WriteItem>();
        var errors = new List<string>();

        for (var row = 1; row <= lastRow; row++)
        {
            var addressText = sheet.Cell(row, 1).GetFormattedString().Trim();
            var valueText = sheet.Cell(row, 2).GetFormattedString().Trim();
            if (addressText.Length == 0 && valueText.Length == 0) continue;

            try
            {
                var sourceAddress = ParseInteger(addressText, "地址");
                var value = ParseInteger(valueText, "數值");
                var register = ConvertAddress(sourceAddress, mode);
                ValidateValue(value);
                result.Add(new WriteItem
                {
                    ExcelRow = row,
                    SourceAddress = sourceAddress,
                    RegisterAddress = checked((ushort)register),
                    Value = value
                });
            }
            catch (Exception ex)
            {
                if (row == 1 && result.Count == 0) continue;
                errors.Add($"第 {row} 列：{ex.Message}");
            }
        }

        if (errors.Count > 0)
        {
            var text = string.Join(Environment.NewLine, errors.Take(10));
            if (errors.Count > 10) text += Environment.NewLine + "……";
            throw new InvalidDataException(text);
        }
        if (result.Count == 0) throw new InvalidDataException("Excel 中沒有可寫入的資料");
        return result;
    }

    private static int ParseInteger(string text, string name)
    {
        if (string.IsNullOrWhiteSpace(text)) throw new FormatException($"{name}不可空白");
        text = text.Trim();
        if (text.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            return Convert.ToInt32(text[2..], 16);
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out var current)) return current;
        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var invariant)) return invariant;
        if (double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out var number) &&
            number == Math.Truncate(number) && number is >= int.MinValue and <= int.MaxValue)
            return (int)number;
        throw new FormatException($"{name}必須是整數");
    }

    private static int ConvertAddress(int address, string mode)
    {
        var result = mode switch
        {
            "1-based（1 = 第一個暫存器）" => address - 1,
            "4xxxx（40001 = 第一個暫存器）" => address - 40001,
            "自動（40001→0，其餘不偏移）" when address >= 40001 => address - 40001,
            _ => address
        };
        if (result is < 0 or > 65535) throw new ArgumentOutOfRangeException(nameof(address), $"換算後位址 {result} 超出 0～65535");
        return result;
    }

    private static ushort ValidateValue(int value)
    {
        if (value is < -32768 or > 65535)
            throw new ArgumentOutOfRangeException(nameof(value), $"數值 {value} 超出單一 16-bit 暫存器範圍");
        return unchecked((ushort)value);
    }

    private void RefreshGrid()
    {
        grid.Rows.Clear();
        foreach (var item in items)
            grid.Rows.Add(item.ExcelRow, item.SourceAddress, item.RegisterAddress, item.Value, item.Status);
        progress.Minimum = 0;
        progress.Maximum = Math.Max(1, items.Count);
        progress.Value = 0;
    }

    private async Task StartWritingAsync()
    {
        if (items.Count == 0)
        {
            MessageBox.Show(this, "請先匯入 Excel。", "尚無資料", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var host = ipBox.Text.Trim();
        if (host.Length == 0)
        {
            MessageBox.Show(this, "請輸入目標 IP。", "設定錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (MessageBox.Show(this,
            $"即將向 {host}:{portBox.Value} 寫入 {items.Count} 筆 Holding Register。\n\n請先確認地址不會觸發危險動作。是否繼續？",
            "確認寫入", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

        foreach (var item in items) item.Status = "等待";
        RefreshGrid();
        SetRunning(true);
        cancellation = new CancellationTokenSource();
        var completed = 0;

        try
        {
            statusLabel.Text = $"正在連線 {host}:{portBox.Value}…";
            using var client = new ModbusTcpClient(host, (int)portBox.Value, (byte)unitBox.Value, TimeSpan.FromSeconds((double)timeoutBox.Value));
            await client.ConnectAsync(cancellation.Token);
            statusLabel.Text = "連線成功，開始寫入";

            for (var i = 0; i < items.Count; i++)
            {
                cancellation.Token.ThrowIfCancellationRequested();
                SetRowStatus(i, "寫入中");
                await client.WriteSingleRegisterAsync(items[i].RegisterAddress, ValidateValue(items[i].Value), cancellation.Token);
                completed++;
                SetRowStatus(i, "成功");
                progress.Value = completed;
                if (delayBox.Value > 0)
                    await Task.Delay((int)delayBox.Value, cancellation.Token);
            }

            statusLabel.Text = $"完成：成功寫入 {completed} 筆";
            MessageBox.Show(this, statusLabel.Text, "寫入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (OperationCanceledException)
        {
            statusLabel.Text = $"已停止；完成 {completed}/{items.Count} 筆";
            MessageBox.Show(this, statusLabel.Text, "已停止", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            var index = Math.Min(completed, items.Count - 1);
            SetRowStatus(index, $"失敗：{ex.Message}");
            statusLabel.Text = $"第 {items[index].ExcelRow} 列失敗";
            MessageBox.Show(this, $"Excel 第 {items[index].ExcelRow} 列寫入失敗：\n{ex.Message}", "Modbus 寫入錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            cancellation.Dispose();
            cancellation = null;
            SetRunning(false);
        }
    }

    private void SetRowStatus(int index, string status)
    {
        items[index].Status = status;
        grid.Rows[index].Cells[4].Value = status;
        grid.FirstDisplayedScrollingRowIndex = Math.Max(0, index);
    }

    private void SetRunning(bool running)
    {
        startButton.Enabled = !running;
        importButton.Enabled = !running;
        addressMode.Enabled = !running;
        ipBox.Enabled = !running;
        portBox.Enabled = !running;
        unitBox.Enabled = !running;
        timeoutBox.Enabled = !running;
        delayBox.Enabled = !running;
        stopButton.Enabled = running;
    }
}

internal sealed class ModbusTcpClient : IDisposable
{
    private readonly string host;
    private readonly int port;
    private readonly byte unitId;
    private readonly TimeSpan timeout;
    private readonly TcpClient client = new();
    private ushort transactionId;

    public ModbusTcpClient(string host, int port, byte unitId, TimeSpan timeout)
    {
        this.host = host;
        this.port = port;
        this.unitId = unitId;
        this.timeout = timeout;
    }

    public async Task ConnectAsync(CancellationToken cancellationToken)
    {
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(timeout);
        await client.ConnectAsync(host, port, timeoutCts.Token);
        client.NoDelay = true;
    }

    public async Task WriteSingleRegisterAsync(ushort address, ushort value, CancellationToken cancellationToken)
    {
        if (!client.Connected) throw new IOException("尚未連線");
        transactionId++;
        Span<byte> request = stackalloc byte[12];
        BinaryPrimitives.WriteUInt16BigEndian(request[0..2], transactionId);
        BinaryPrimitives.WriteUInt16BigEndian(request[2..4], 0);
        BinaryPrimitives.WriteUInt16BigEndian(request[4..6], 6);
        request[6] = unitId;
        request[7] = 0x06;
        BinaryPrimitives.WriteUInt16BigEndian(request[8..10], address);
        BinaryPrimitives.WriteUInt16BigEndian(request[10..12], value);

        var stream = client.GetStream();
        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(timeout);
        await stream.WriteAsync(request.ToArray(), timeoutCts.Token);

        var response = new byte[12];
        await ReadExactlyAsync(stream, response.AsMemory(0, 7), timeoutCts.Token);
        var length = BinaryPrimitives.ReadUInt16BigEndian(response.AsSpan(4, 2));
        if (length < 2 || length > 254) throw new IOException("Modbus 回覆長度錯誤");
        var remaining = length - 1;
        var body = remaining <= 5 ? response.AsMemory(7, remaining) : new byte[remaining].AsMemory();
        await ReadExactlyAsync(stream, body, timeoutCts.Token);

        var responseTransaction = BinaryPrimitives.ReadUInt16BigEndian(response.AsSpan(0, 2));
        if (responseTransaction != transactionId) throw new IOException("收到不相符的 Transaction ID");
        if (BinaryPrimitives.ReadUInt16BigEndian(response.AsSpan(2, 2)) != 0) throw new IOException("Protocol ID 不正確");
        if (response[6] != unitId) throw new IOException("站號不相符");

        var data = body.Span;
        if ((data[0] & 0x80) != 0)
        {
            var code = data.Length > 1 ? data[1] : (byte)0;
            var message = code switch
            {
                1 => "不支援的功能碼",
                2 => "不合法的資料地址",
                3 => "不合法的資料值",
                4 => "從站設備故障",
                _ => "未知錯誤"
            };
            throw new IOException($"Modbus Exception {code}：{message}");
        }
        if (data.Length != 5 || data[0] != 0x06) throw new IOException("Modbus 回覆格式錯誤");
        if (BinaryPrimitives.ReadUInt16BigEndian(data[1..3]) != address ||
            BinaryPrimitives.ReadUInt16BigEndian(data[3..5]) != value)
            throw new IOException("設備回覆的地址或數值不一致");
    }

    private static async Task ReadExactlyAsync(NetworkStream stream, Memory<byte> buffer, CancellationToken cancellationToken)
    {
        var offset = 0;
        while (offset < buffer.Length)
        {
            var read = await stream.ReadAsync(buffer[offset..], cancellationToken);
            if (read == 0) throw new EndOfStreamException("目標設備已關閉連線");
            offset += read;
        }
    }

    public void Dispose() => client.Dispose();
}
