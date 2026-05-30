using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace HealthcareSanatoriumInterface
{
    internal static class Program
    {
        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();

        [STAThread]
        private static void Main(string[] args)
        {
            try { SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string dbPath = Path.Combine(baseDir, "HealthcareSanatoriumSystem.accdb");
            if (args.Length > 0 && args[0] == "--self-test")
            {
                try
                {
                    DbContext testDb = new DbContext(dbPath);
                    if (!testDb.Exists()) Environment.Exit(2);
                    testDb.ScalarInt("SELECT Count(*) FROM tblVoucherIssues");
                    testDb.ScalarInt("SELECT Count(*) FROM tblPensioners");
                    Environment.Exit(0);
                }
                catch (Exception ex)
                {
                    File.WriteAllText(Path.Combine(baseDir, "self-test-error.txt"), ex.ToString());
                    Environment.Exit(1);
                }
            }
            if (args.Length > 0 && args[0] == "--pdf-self-test")
            {
                try
                {
                    DbContext testDb = new DbContext(dbPath);
                    if (!testDb.Exists()) Environment.Exit(2);
                    DataTable table = testDb.Query("SELECT TOP 5 VoucherNo AS [Путевка], IssueDate AS [Регистрация], PensionerFullName AS [Пенсионер], SanatoriumName AS [Санаторий], TotalCost AS [Стоимость] FROM qryReport_VoucherIssuesDetailed ORDER BY IssueDate, VoucherNo");
                    string testPdf = Path.Combine(baseDir, "pdf-self-test.pdf");
                    ProfessionalPdfExporter.ExportDataTableToFile(testPdf, "Тестовый PDF-отчет", "Система здравоохранения", table);
                    FileInfo info = new FileInfo(testPdf);
                    if (!info.Exists || info.Length < 1000) Environment.Exit(3);
                    try { info.Delete(); } catch { }
                    Environment.Exit(0);
                }
                catch (Exception ex)
                {
                    File.WriteAllText(Path.Combine(baseDir, "pdf-self-test-error.txt"), ex.ToString());
                    Environment.Exit(1);
                }
            }
            SplashForm splash = new SplashForm();
            splash.Show();
            Application.DoEvents();
            splash.SetStatus("Загрузка интерфейса...");
            MainForm mainForm = new MainForm(dbPath);
            mainForm.Load += delegate { splash.Close(); };
            Application.Run(mainForm);
        }
    }

    internal static class Theme
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

        public static bool Dark = false;
        public static Color Ink { get { return Dark ? Color.FromArgb(238, 244, 255) : Color.FromArgb(15, 23, 42); } }
        public static Color Muted { get { return Dark ? Color.FromArgb(160, 174, 199) : Color.FromArgb(100, 116, 139); } }
        public static Color Line { get { return Dark ? Color.FromArgb(56, 71, 102) : Color.FromArgb(218, 226, 238); } }
        public static Color Surface { get { return Dark ? Color.FromArgb(19, 27, 44) : Color.FromArgb(255, 255, 255); } }
        public static Color SurfaceAlt { get { return Dark ? Color.FromArgb(25, 35, 56) : Color.FromArgb(248, 251, 255); } }
        public static Color Field { get { return Dark ? Color.FromArgb(29, 40, 62) : Color.FromArgb(255, 255, 255); } }
        public static Color Blue { get { return Dark ? Color.FromArgb(96, 165, 250) : Color.FromArgb(37, 99, 235); } }
        public static Color Blue2 { get { return Dark ? Color.FromArgb(34, 211, 238) : Color.FromArgb(14, 165, 233); } }
        public static Color Green { get { return Dark ? Color.FromArgb(52, 211, 153) : Color.FromArgb(16, 185, 129); } }
        public static Color AccentSoft { get { return Dark ? Color.FromArgb(35, 55, 87) : Color.FromArgb(232, 242, 255); } }
        public static readonly Font BaseFont = new Font("Segoe UI", 10f, FontStyle.Regular);
        public static readonly Font TitleFont = new Font("Segoe UI Semibold", 22f, FontStyle.Bold);
        public static readonly Font H2Font = new Font("Segoe UI Semibold", 15f, FontStyle.Bold);
        public static readonly Font ButtonFont = new Font("Segoe UI Semibold", 10f, FontStyle.Bold);
        public static readonly Font ButtonCaptionFont = new Font("Segoe UI Semibold", 9.6f, FontStyle.Bold);
        public static readonly Font ButtonTitleFont = new Font("Segoe UI Semibold", 10.2f, FontStyle.Bold);
        public static readonly Font ButtonSubFont = new Font("Segoe UI", 8.5f, FontStyle.Regular);
        public static readonly Font GridHeaderFont = new Font("Segoe UI Semibold", 9f, FontStyle.Bold);
        public static readonly Font GridCellFont = new Font("Segoe UI", 9f, FontStyle.Regular);
        public static readonly Font StatsFont = new Font("Segoe UI Semibold", 11f, FontStyle.Bold);

        public static void StyleButton(Button button, bool primary)
        {
            GlassButton glass = button as GlassButton;
            if (glass != null)
            {
                glass.Primary = primary;
                glass.Large = button.Height >= 58 || button.Text.IndexOf('\n') >= 0;
            }
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = primary ? 0 : 1;
            button.FlatAppearance.BorderColor = Line;
            button.FlatAppearance.MouseOverBackColor = Dark ? Color.FromArgb(40, 56, 86) : Color.FromArgb(239, 246, 255);
            button.FlatAppearance.MouseDownBackColor = Dark ? Color.FromArgb(27, 38, 61) : Color.FromArgb(219, 234, 254);
            button.UseVisualStyleBackColor = false;
            button.BackColor = glass == null ? (primary ? Blue : Surface) : Color.Transparent;
            button.ForeColor = primary ? Color.White : Ink;
            button.Font = ButtonFont;
            button.Margin = Padding.Empty;
            button.AutoEllipsis = true;
            if (button.Height < 36) button.Height = 36;
            button.Cursor = Cursors.Hand;
            button.TabStop = true;
            button.Invalidate();
        }

        public static Label Label(string text, Font font, Color color)
        {
            Label label = new Label();
            label.Text = text;
            label.Font = font;
            label.ForeColor = color;
            label.BackColor = Color.Transparent;
            label.AutoSize = true;
            return label;
        }

        public static TextBox TextBox(string placeholder)
        {
            TextBox box = new TextBox();
            box.Font = BaseFont;
            box.BorderStyle = BorderStyle.FixedSingle;
            box.BackColor = Field;
            box.ForeColor = Ink;
            box.Tag = placeholder;
            return box;
        }

        public static void StyleGrid(DataGridView grid)
        {
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.ReadOnly = true;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.MultiSelect = false;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            grid.RowTemplate.Height = 34;
            grid.ColumnHeadersHeight = 38;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ShowCellToolTips = false;
            grid.BackgroundColor = Surface;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Dark ? Color.FromArgb(27, 38, 61) : Color.FromArgb(241, 246, 253);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Ink;
            grid.ColumnHeadersDefaultCellStyle.Font = GridHeaderFont;
            grid.ColumnHeadersDefaultCellStyle.Padding = new Padding(8, 0, 8, 0);
            grid.DefaultCellStyle.Font = GridCellFont;
            grid.DefaultCellStyle.BackColor = Surface;
            grid.DefaultCellStyle.ForeColor = Ink;
            grid.DefaultCellStyle.Padding = new Padding(8, 0, 8, 0);
            grid.DefaultCellStyle.SelectionBackColor = Dark ? Color.FromArgb(37, 63, 105) : Color.FromArgb(219, 234, 254);
            grid.DefaultCellStyle.SelectionForeColor = Ink;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Dark ? Color.FromArgb(22, 31, 50) : Color.FromArgb(248, 251, 255);
            grid.RowHeadersVisible = false;
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.GridColor = Dark ? Color.FromArgb(45, 58, 82) : Color.FromArgb(232, 238, 248);
            grid.Dock = DockStyle.Fill;
            try
            {
                typeof(DataGridView).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, null, grid, new object[] { true });
            }
            catch { }
        }

        // Shared form background — must match GlassForm.FormBgColor()
        public static Color FormBackground { get { return Dark ? Color.FromArgb(10, 16, 28) : Color.FromArgb(229, 238, 250); } }

        public static void ApplyTo(Control root)
        {
            if (root == null) return;
            root.Font = BaseFont;

            if (root is GlassPanel)
            {
                ((GlassPanel)root).InvalidatePanelCache();
            }
            else if (root is Form)
            {
                root.BackColor = FormBackground;
            }

            foreach (Control control in root.Controls)
            {
                if (control is GlassButton)
                {
                    control.Invalidate(); // GlassButton reads Theme.Dark in OnPaint
                }
                else if (control is Button)
                {
                    Button button = (Button)control;
                    StyleButton(button, button.FlatAppearance.BorderSize == 0);
                }
                else if (control is DataGridView)
                {
                    StyleGrid((DataGridView)control);
                }
                else if (control is TextBox)
                {
                    control.BackColor = Field;
                    control.ForeColor = Ink;
                    string ph = control.Tag as string;
                    if (!string.IsNullOrEmpty(ph))
                    {
                        try { SendMessage(control.Handle, 0x1501u, IntPtr.Zero, ph); } catch { }
                    }
                }
                else if (control is ComboBox)
                {
                    control.BackColor = Field;
                    control.ForeColor = Ink;
                    control.Font = BaseFont;
                }
                else if (control is DateTimePicker || control is NumericUpDown)
                {
                    control.BackColor = Field;
                    control.ForeColor = Ink;
                    control.Font = BaseFont;
                }
                else if (control is CheckBox)
                {
                    control.ForeColor = Ink;
                    control.BackColor = Color.Transparent;
                    control.Font = BaseFont;
                }
                else if (control is Label)
                {
                    control.BackColor = Color.Transparent;
                    control.ForeColor = string.Equals(control.Tag as string, "accent", StringComparison.OrdinalIgnoreCase) ? Blue : (control.Font.Bold ? Ink : Muted);
                }
                ApplyTo(control);
            }
        }
    }


    internal static class UiShapes
    {
        public static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int diameter = Math.Max(2, radius * 2);
            Rectangle arc = new Rectangle(bounds.Location, new Size(diameter, diameter));
            GraphicsPath path = new GraphicsPath();

            path.AddArc(arc, 180, 90);
            arc.X = bounds.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = bounds.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = bounds.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }
    }



internal static class ErrorLogger
{
    private static readonly object LogLock = new object();
    private static string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error_log.txt");

    public static void Log(string message)
    {
        try
        {
            lock (LogLock)
            {
                string logEntry = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " | " + message;
                File.AppendAllText(LogPath, logEntry + Environment.NewLine);
            }
        }
        catch { }
    }

    public static void LogException(Exception ex)
    {
        string message = "EXCEPTION: " + ex.GetType().Name + " - " + ex.Message + "\n" + ex.StackTrace;
        if (ex.InnerException != null)
        {
            message += "\nInner: " + ex.InnerException.Message;
        }
        Log(message);
    }
}

internal static class Debouncer
{
    private static readonly Dictionary<string, System.Windows.Forms.Timer> Timers = new Dictionary<string, System.Windows.Forms.Timer>();

    public static void Debounce(string key, System.Action action, int delayMs = 300)
    {
        System.Windows.Forms.Timer existing;
        if (Timers.TryGetValue(key, out existing))
        {
            existing.Stop();
            existing.Dispose();
            Timers.Remove(key);
        }
        System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
        timer.Interval = delayMs;
        timer.Tick += delegate
        {
            timer.Stop();
            timer.Dispose();
            Timers.Remove(key);
            try { action(); }
            catch (Exception ex) { ErrorLogger.LogException(ex); }
        };
        Timers[key] = timer;
        timer.Start();
    }
}

public static class UiPerformance
{
    public static void Optimize(Control root)
    {
        if (root == null) return;

        root.SuspendLayout();
        try
        {
            EnableDoubleBuffer(root);
            foreach (Control control in root.Controls)
            {
                Optimize(control);
            }
        }
        finally
        {
            root.ResumeLayout(false);
        }
    }

    public static void BindGrid(DataGridView grid, DataTable table)
    {
        if (grid == null) return;

        CurrencyManager manager = null;
        try
        {
            manager = grid.BindingContext == null || grid.DataSource == null ? null : (CurrencyManager)grid.BindingContext[grid.DataSource];
        }
        catch
        {
            manager = null;
        }

        grid.SuspendLayout();
        try
        {
            if (manager != null) manager.SuspendBinding();
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            grid.DataSource = table;
            if (grid.Columns.Count > 0) grid.Columns[0].Visible = false;
            // Enable sorting by column header click
            foreach (DataGridViewColumn col in grid.Columns)
                if (col.Visible) col.SortMode = DataGridViewColumnSortMode.Automatic;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            grid.ClearSelection();
        }
        finally
        {
            if (manager != null) manager.ResumeBinding();
            grid.ResumeLayout(false);
        }
    }

    public static void EnableDoubleBuffer(Control control)
    {
        if (control == null) return;
        try
        {
            System.Reflection.PropertyInfo doubleBuffered = control.GetType().GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            if (doubleBuffered != null)
            {
                doubleBuffered.SetValue(control, true, null);
            }
        }
        catch
        {
        }
    }
}
public sealed class DbContext
{
    private static readonly string[] AccdbProviders = new[] { "Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0" };
    private static readonly string[] MdbProviders = new[] { "Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.16.0", "Microsoft.ACE.OLEDB.12.0" };
    private string resolvedProvider;
    private string lastProviderError;
    private readonly object syncRoot = new object();
    private readonly Dictionary<string, DataTable> lookupCache = new Dictionary<string, DataTable>();
    private readonly object cacheLock = new object();

    public DbContext(string databasePath)
    {
        DatabasePath = databasePath;
    }

    public string DatabasePath { get; private set; }

    public bool Exists()
    {
        return !string.IsNullOrWhiteSpace(DatabasePath) && File.Exists(DatabasePath);
    }

    public void SetDatabasePath(string path)
    {
        lock (syncRoot)
        {
            DatabasePath = path;
            resolvedProvider = null;
            lastProviderError = null;
            ClearLookupCache();
        }
    }

    public DataTable Query(string sql, params OleDbParameter[] parameters)
    {
        using (OleDbConnection connection = CreateOpenConnection())
        using (OleDbCommand command = new OleDbCommand(sql, connection))
        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
        {
            command.CommandTimeout = 30;
            if (parameters != null && parameters.Length > 0)
            {
                command.Parameters.AddRange(parameters);
            }

            DataTable table = new DataTable();
            adapter.Fill(table);
            return table;
        }
    }

    public int Execute(string sql, params OleDbParameter[] parameters)
    {
        using (OleDbConnection connection = CreateOpenConnection())
        using (OleDbCommand command = new OleDbCommand(sql, connection))
        {
            command.CommandTimeout = 30;
            if (parameters != null && parameters.Length > 0)
            {
                command.Parameters.AddRange(parameters);
            }

            return command.ExecuteNonQuery();
        }
    }

    public int ScalarInt(string sql)
    {
        using (OleDbConnection connection = CreateOpenConnection())
        using (OleDbCommand command = new OleDbCommand(sql, connection))
        {
            command.CommandTimeout = 30;
            object value = command.ExecuteScalar();
            if (value == null || value == DBNull.Value) return 0;
            return Convert.ToInt32(value);
        }
    }

    public DataTable Lookup(string table, string idField, string nameField)
    {
        return Query("SELECT " + idField + ", " + nameField + " FROM " + table + " ORDER BY " + nameField);
    }

    public DataTable LookupCached(string table, string idField, string nameField)
    {
        string cacheKey = table + "|" + idField + "|" + nameField;
        lock (cacheLock)
        {
            DataTable cached;
            if (lookupCache.TryGetValue(cacheKey, out cached))
            {
                return cached;
            }
        }

        DataTable result = Lookup(table, idField, nameField);
        lock (cacheLock)
        {
            lookupCache[cacheKey] = result;
        }
        return result;
    }

    public void ClearLookupCache()
    {
        lock (cacheLock)
        {
            lookupCache.Clear();
        }
    }

    private OleDbConnection CreateOpenConnection()
    {
        if (!Exists())
        {
            throw new FileNotFoundException("Файл базы данных не найден", DatabasePath);
        }

        string provider = ResolveProvider();
        if (string.IsNullOrWhiteSpace(provider))
        {
            throw new InvalidOperationException(BuildProviderError());
        }

        OleDbConnection connection = new OleDbConnection(BuildConnectionString(provider));
        try
        {
            connection.Open();
            return connection;
        }
        catch (Exception ex)
        {
            connection.Dispose();
            throw new InvalidOperationException(BuildConnectionError(provider, ex), ex);
        }
    }

    private string BuildConnectionString(string provider)
    {
        OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
        builder["Provider"] = provider;
        builder["Data Source"] = DatabasePath;
        builder["Persist Security Info"] = false;
        builder["Mode"] = "Share Deny None";
        builder["OLE DB Services"] = -1;
        builder["Jet OLEDB:Database Locking Mode"] = 1;
        builder["Jet OLEDB:Engine Type"] = 5;
        builder["Jet OLEDB:Global Partial Bulk Ops"] = 2;
        builder["Jet OLEDB:Registry Path"] = "";
        builder["Jet OLEDB:System database"] = "";
        return builder.ConnectionString;
    }

    private string ResolveProvider()
    {
        lock (syncRoot)
        {
            if (!string.IsNullOrWhiteSpace(resolvedProvider))
            {
                return resolvedProvider;
            }

        if (!Exists())
        {
            return null;
        }

        List<string> failures = new List<string>();
        foreach (string provider in GetProviderCandidates())
        {
            string failure;
            if (CanOpenProvider(provider, out failure))
            {
                resolvedProvider = provider;
                lastProviderError = null;
                return resolvedProvider;
            }

            if (!string.IsNullOrWhiteSpace(failure))
            {
                failures.Add(failure);
            }
        }

        resolvedProvider = null;
        lastProviderError = failures.Count > 0 ? string.Join(" | ", failures.ToArray()) : "Ни один провайдер OLE DB не смог открыть базу.";
        return null;
        }
    }

    private bool CanOpenProvider(string provider, out string failure)
    {
        failure = null;
        try
        {
            using (OleDbConnection connection = new OleDbConnection(BuildConnectionString(provider)))
            {
                connection.Open();
                connection.Close();
                return true;
            }
        }
        catch (Exception ex)
        {
            failure = ex.Message;
            return false;
        }
    }

    private string BuildProviderError()
    {
        return "Ошибка подключения к базе данных. Проверьте:\n" +
               "1. Файл БД не поврежден\n" +
               "2. БД не открыта в другом приложении\n" +
               "3. Права доступа к папке";
    }

    private string BuildConnectionError(string provider, Exception ex)
    {
        return "Не удалось открыть БД: " + ex.Message;
    }

    private string[] GetProviderCandidates()
    {
        string extension = string.IsNullOrWhiteSpace(DatabasePath) ? string.Empty : Path.GetExtension(DatabasePath).ToLowerInvariant();
        if (extension == ".mdb")
        {
            return MdbProviders;
        }

        return AccdbProviders;
    }
}


    internal sealed class LookupItem
    {
        public LookupItem(int id, string text)
        {
            Id = id;
            Text = text;
        }

        public int Id { get; private set; }
        public string Text { get; private set; }

        public override string ToString()
        {
            return Text;
        }
    }

    
public class GlassForm : Form
{
    protected readonly DbContext Db;

    protected GlassForm(DbContext db)
    {
        Db = db;
        Font = Theme.BaseFont;
        Icon = TryLoadAppIcon();
        BackColor = FormBgColor();
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        DoubleBuffered = true;
        AutoScroll = true;
        Padding = new Padding(16);
        MinimumSize = new Size(1024, 680);
        MaximizeBox = true;
        Shown += delegate { ApplyTheme(); };
    }

    private static Color FormBgColor()
    {
        return Theme.FormBackground;
    }

    // No custom OnPaintBackground — WinForms fills with BackColor instantly

    protected static Icon TryLoadAppIcon()
    {
        string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app_icon.ico");
        if (!File.Exists(iconPath))
        {
            return null;
        }

        try
        {
            return new Icon(iconPath);
        }
        catch
        {
            return null;
        }
    }

    protected void Guard(Action action)
    {
        try
        {
            if (!Db.Exists())
            {
                ErrorLogger.Log("Database not found: " + Db.DatabasePath);
                ToastNotifier.Show(this, "Файл базы данных не найден:\n" + Db.DatabasePath, false);
                return;
            }

            action();
        }
        catch (Exception ex)
        {
            ErrorLogger.LogException(ex);
            ToastNotifier.Show(this, "Операция не выполнена: " + ex.Message, false);
        }
    }

    protected static TableLayoutPanel MakeRootLayout()
    {
        TableLayoutPanel t = new TableLayoutPanel();
        t.Dock = DockStyle.Fill;
        t.BackColor = Color.Transparent;
        t.ColumnCount = 1;
        t.RowCount = 4;
        t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        t.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // 0: header
        t.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // 1: filters
        t.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // 2: grid
        t.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // 3: actions
        return t;
    }

    protected static TableLayoutPanel MakeHeaderRow()
    {
        TableLayoutPanel h = new TableLayoutPanel();
        h.Dock = DockStyle.Fill;
        h.BackColor = Color.Transparent;
        h.ColumnCount = 2;
        h.RowCount = 1;
        h.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        h.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        h.Margin = new Padding(0, 2, 0, 12);
        h.Padding = new Padding(0, 0, 0, 10);
        return h;
    }

    protected static FlowLayoutPanel MakeHeaderButtons()
    {
        FlowLayoutPanel f = new FlowLayoutPanel();
        f.AutoSize = true;
        f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        f.FlowDirection = FlowDirection.LeftToRight;
        f.WrapContents = false;
        f.BackColor = Color.Transparent;
        f.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        return f;
    }

    protected static FlowLayoutPanel MakeFilterRow()
    {
        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill;
        f.AutoSize = true;
        f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        f.FlowDirection = FlowDirection.LeftToRight;
        f.WrapContents = true;
        f.BackColor = Color.Transparent;
        f.Margin = new Padding(0, 0, 0, 8);
        return f;
    }

    protected static TableLayoutPanel MakeActionBar()
    {
        TableLayoutPanel t = new TableLayoutPanel();
        t.Dock = DockStyle.Fill;
        t.BackColor = Color.Transparent;
        t.ColumnCount = 2;
        t.RowCount = 1;
        t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        t.Margin = new Padding(0, 4, 0, 0);
        return t;
    }

    protected static FlowLayoutPanel MakeLeftActions()
    {
        FlowLayoutPanel f = new FlowLayoutPanel();
        f.AutoSize = true;
        f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        f.FlowDirection = FlowDirection.LeftToRight;
        f.WrapContents = false;
        f.BackColor = Color.Transparent;
        return f;
    }

    protected static FlowLayoutPanel MakeRightActions()
    {
        FlowLayoutPanel f = new FlowLayoutPanel();
        f.AutoSize = true;
        f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        f.FlowDirection = FlowDirection.RightToLeft;
        f.WrapContents = false;
        f.BackColor = Color.Transparent;
        f.Anchor = AnchorStyles.Right | AnchorStyles.Top;
        return f;
    }

    protected static Button MakeButton(string text, bool primary)
    {
        GlassButton b = new GlassButton();
        b.Text = text;
        b.Size = new Size(116, 38);
        Theme.StyleButton(b, primary);
        return b;
    }

    protected void FillLookup(ComboBox combo, DataTable table, string idField, string nameField, bool addAll, string allLabel = "— Все —")
    {
        combo.BeginUpdate();
        try
        {
            combo.Items.Clear();
            if (addAll)
            {
                combo.Items.Add(new LookupItem(0, allLabel));
            }

            foreach (DataRow row in table.Rows)
            {
                combo.Items.Add(new LookupItem(Convert.ToInt32(row[idField]), Convert.ToString(row[nameField])));
            }

            if (combo.Items.Count > 0)
            {
                combo.SelectedIndex = 0;
            }

            combo.DropDownStyle = ComboBoxStyle.DropDownList;
        }
        finally
        {
            combo.EndUpdate();
        }
    }

    protected void ApplyTheme()
    {
        try
        {
            SuspendLayout();
            BackColor = FormBgColor();
            RefreshGlassPanels(this);
            Theme.ApplyTo(this);
            UiPerformance.Optimize(this);
        }
        finally
        {
            ResumeLayout(true);
            Refresh();
        }
    }

    private static void RefreshGlassPanels(Control root)
    {
        foreach (Control c in root.Controls)
        {
            GlassPanel gp = c as GlassPanel;
            if (gp != null)
                gp.InvalidatePanelCache();

            // Update embedded (non-toplevel) form background
            Form f = c as Form;
            if (f != null && !f.TopLevel)
                f.BackColor = Theme.FormBackground;

            RefreshGlassPanels(c); // always recurse — catches nested panels + embedded forms
        }
    }
}


    internal sealed class GlassPanel : Panel
    {
        private const int Radius = 18;

        public GlassPanel()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.UserPaint, true);
            Padding = new Padding(20);
            UpdatePanelColor();
        }

        public void InvalidatePanelCache()
        {
            UpdatePanelColor();
            ApplyRoundedRegion();
            Invalidate();
        }

        private void UpdatePanelColor()
        {
            BackColor = Theme.Surface;
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            ApplyRoundedRegion();
        }

        private void ApplyRoundedRegion()
        {
            if (Width <= 1 || Height <= 1) return;
            Region oldRegion = Region;
            using (GraphicsPath path = UiShapes.RoundedRect(new Rectangle(0, 0, Width, Height), Radius))
            {
                Region = new Region(path);
            }
            if (oldRegion != null)
            {
                oldRegion.Dispose();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            if (r.Width <= 0 || r.Height <= 0) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color top = Theme.Dark ? Color.FromArgb(27, 38, 61) : Color.FromArgb(255, 255, 255);
            Color bottom = Theme.Dark ? Color.FromArgb(18, 27, 45) : Color.FromArgb(247, 251, 255);
            Color borderColor = Theme.Dark ? Color.FromArgb(58, 74, 108) : Color.FromArgb(210, 222, 240);

            using (GraphicsPath path = UiShapes.RoundedRect(r, Radius))
            using (LinearGradientBrush brush = new LinearGradientBrush(r, top, bottom, LinearGradientMode.Vertical))
            using (Pen border = new Pen(borderColor, 1f))
            {
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(border, path);
            }

            using (Pen glow = new Pen(Theme.Dark ? Color.FromArgb(36, Theme.Blue) : Color.FromArgb(38, Theme.Blue), 2f))
                e.Graphics.DrawLine(glow, Radius, 1, Width - Radius, 1);

            using (Pen shadow = new Pen(Theme.Dark ? Color.FromArgb(38, 0, 0, 0) : Color.FromArgb(26, 99, 125, 170), 2f))
                e.Graphics.DrawLine(shadow, Radius, Height - 2, Width - Radius, Height - 2);
        }
    }

    internal sealed class GlassButton : Button
    {
        private const int Radius = 14;
        private bool _hot;
        private bool _pressed;

        public GlassButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.Selectable | ControlStyles.UserPaint, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
        }

        public bool Primary { get; set; }
        public bool Large { get; set; }

        protected override void OnMouseEnter(EventArgs e) { _hot = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hot = false; _pressed = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { _pressed = true; Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { _pressed = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle r = new Rectangle(0, 0, Width - 1, Height - 1);
            if (r.Width <= 0 || r.Height <= 0) return;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color top, bottom, fg, borderColor;

            if (!Enabled)
            {
                top = bottom = Theme.Dark ? Color.FromArgb(26, 34, 50) : Color.FromArgb(241, 245, 249);
                fg = Theme.Muted;
                borderColor = Theme.Line;
            }
            else if (Primary)
            {
                top = _pressed ? Color.FromArgb(30, 64, 175) : (_hot ? Color.FromArgb(59, 130, 246) : Color.FromArgb(37, 99, 235));
                bottom = _pressed ? Color.FromArgb(29, 78, 216) : (_hot ? Color.FromArgb(14, 165, 233) : Color.FromArgb(29, 78, 216));
                fg = Color.White;
                borderColor = Color.Transparent;
            }
            else if (Theme.Dark)
            {
                top = _pressed ? Color.FromArgb(24, 35, 56) : (_hot ? Color.FromArgb(40, 56, 86) : Color.FromArgb(29, 40, 62));
                bottom = _pressed ? Color.FromArgb(20, 29, 48) : (_hot ? Color.FromArgb(32, 46, 72) : Color.FromArgb(24, 34, 54));
                fg = Theme.Ink;
                borderColor = Color.FromArgb(60, 78, 114);
            }
            else
            {
                top = _pressed ? Color.FromArgb(226, 234, 246) : (_hot ? Color.FromArgb(245, 249, 255) : Color.FromArgb(255, 255, 255));
                bottom = _pressed ? Color.FromArgb(214, 226, 244) : (_hot ? Color.FromArgb(235, 244, 255) : Color.FromArgb(248, 251, 255));
                fg = Theme.Ink;
                borderColor = Color.FromArgb(205, 217, 235);
            }

            using (GraphicsPath path = UiShapes.RoundedRect(r, Radius))
            using (LinearGradientBrush brush = new LinearGradientBrush(r, top, bottom, LinearGradientMode.Vertical))
            {
                e.Graphics.FillPath(brush, path);
                if (borderColor != Color.Transparent)
                {
                    using (Pen border = new Pen(borderColor, 1f))
                        e.Graphics.DrawPath(border, path);
                }
            }

            if (Large && !Primary && Enabled)
            {
                Rectangle accent = new Rectangle(1, 12, 4, Math.Max(12, Height - 24));
                using (GraphicsPath accentPath = UiShapes.RoundedRect(accent, 2))
                using (LinearGradientBrush accentBrush = new LinearGradientBrush(accent, Theme.Blue, Theme.Blue2, LinearGradientMode.Vertical))
                    e.Graphics.FillPath(accentBrush, accentPath);
            }

            if (Focused && ShowFocusCues)
            {
                Rectangle focus = new Rectangle(3, 3, Width - 7, Height - 7);
                using (GraphicsPath focusPath = UiShapes.RoundedRect(focus, Radius - 3))
                using (Pen focusPen = new Pen(Color.FromArgb(120, Theme.Blue), 1f))
                    e.Graphics.DrawPath(focusPen, focusPath);
            }

            DrawCaption(e.Graphics, r, fg);
        }

        private void DrawCaption(Graphics g, Rectangle r, Color fg)
        {
            string[] lines = Text.Replace("\r", "").Split('\n');
            using (StringFormat sf = new StringFormat())
            {
                sf.Trimming = StringTrimming.EllipsisCharacter;
                sf.Alignment = TextAlign == ContentAlignment.MiddleLeft ? StringAlignment.Near : StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                int pad = Large ? 20 : 12;
                RectangleF tr = new RectangleF(r.X + pad, r.Y + 2, r.Width - pad * 2, r.Height - 4);

                if (lines.Length > 1)
                {
                    Color subColor = Primary ? Color.FromArgb(218, 232, 255) : Theme.Muted;
                    float y = r.Y + (r.Height - 42) / 2f;
                    using (SolidBrush tb = new SolidBrush(fg))
                    using (SolidBrush sb = new SolidBrush(subColor))
                    {
                        g.DrawString(lines[0], Theme.ButtonTitleFont, tb, new RectangleF(tr.X, y, tr.Width, 22), sf);
                        g.DrawString(lines[1], Theme.ButtonSubFont, sb, new RectangleF(tr.X, y + 24, tr.Width, 18), sf);
                    }
                }
                else
                {
                    using (SolidBrush tb = new SolidBrush(fg))
                        g.DrawString(Text, Theme.ButtonCaptionFont, tb, tr, sf);
                }
            }
        }
    }

    
internal sealed class MainForm : GlassForm
{
    private readonly Label stats;
    private readonly Label sectionTitle;
    private readonly Label sectionHint;
    private readonly Dictionary<string, GlassButton> navButtons = new Dictionary<string, GlassButton>();
    private readonly TableLayoutPanel navLayout;
    private readonly FlowLayoutPanel manualBar;
    private readonly GlassPanel contentPanel;
    private readonly TableLayoutPanel homeLayout;
    private Form activeEmbeddedForm;

    public MainForm(string dbPath)
        : base(new DbContext(dbPath))
    {
        Text = "Система здравоохранения";
        Size = new Size(1200, 760);
        MinimumSize = new Size(1080, 700);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 2;
        root.RowCount = 1;
        root.BackColor = Color.Transparent;
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 344f));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        Controls.Add(root);

        GlassPanel sidebar = new GlassPanel();
        sidebar.Dock = DockStyle.Fill;
        sidebar.Padding = new Padding(20);
        sidebar.Margin = new Padding(0, 0, 14, 0);
        root.Controls.Add(sidebar, 0, 0);

        contentPanel = new GlassPanel();
        contentPanel.Dock = DockStyle.Fill;
        contentPanel.Padding = new Padding(28);
        contentPanel.Margin = Padding.Empty;
        root.Controls.Add(contentPanel, 1, 0);

        TableLayoutPanel sidebarLayout = new TableLayoutPanel();
        sidebarLayout.Dock = DockStyle.Fill;
        sidebarLayout.BackColor = Color.Transparent;
        sidebarLayout.ColumnCount = 1;
        sidebarLayout.RowCount = 5;
        sidebarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebar.Controls.Add(sidebarLayout);

        Label eyebrow = Theme.Label("MEDCARE · SANATORIUM", new Font("Segoe UI Semibold", 8.7f, FontStyle.Bold), Theme.Blue);
        eyebrow.Margin = new Padding(4, 2, 4, 8);
        sidebarLayout.Controls.Add(eyebrow, 0, 0);

        Label title = Theme.Label("Система здравоохранения", Theme.TitleFont, Theme.Ink);
        title.Margin = new Padding(4, 0, 4, 8);
        title.MaximumSize = new Size(292, 0);
        sidebarLayout.Controls.Add(title, 0, 1);

        Label subtitle = Theme.Label("Санаторные путевки, пенсионеры, отчеты и сервис данных", Theme.BaseFont, Theme.Muted);
        subtitle.Margin = new Padding(4, 0, 4, 20);
        subtitle.MaximumSize = new Size(292, 0);
        sidebarLayout.Controls.Add(subtitle, 0, 2);

        Panel navHost = new Panel();
        navHost.Dock = DockStyle.Fill;
        navHost.AutoScroll = true;
        navHost.BackColor = Color.Transparent;
        sidebarLayout.Controls.Add(navHost, 0, 3);

        navLayout = new TableLayoutPanel();
        navLayout.Dock = DockStyle.Top;
        navLayout.AutoSize = true;
        navLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        navLayout.BackColor = Color.Transparent;
        navLayout.ColumnCount = 1;
        navLayout.RowCount = 0;
        navLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        navHost.Controls.Add(navLayout);

        TableLayoutPanel sidebarActions = new TableLayoutPanel();
        sidebarActions.Dock = DockStyle.Fill;
        sidebarActions.AutoSize = true;
        sidebarActions.ColumnCount = 1;
        sidebarActions.RowCount = 2;
        sidebarActions.BackColor = Color.Transparent;
        sidebarActions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        sidebarActions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarActions.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.Controls.Add(sidebarActions, 0, 4);

        Button theme = new GlassButton();
        theme.Text = Theme.Dark ? "Светлая тема" : "Темная тема";
        theme.Height = 42;
        Theme.StyleButton(theme, false);
        theme.Click += delegate
        {
            Theme.Dark = !Theme.Dark;
            theme.Text = Theme.Dark ? "Светлая тема" : "Темная тема";
            ApplyTheme();
        };
        sidebarActions.Controls.Add(theme, 0, 0);

        Button close = new GlassButton();
        close.Text = "Закрыть";
        close.Height = 42;
        Theme.StyleButton(close, false);
        close.Click += delegate { Close(); };
        sidebarActions.Controls.Add(close, 0, 1);

        homeLayout = new TableLayoutPanel();
        homeLayout.Dock = DockStyle.Fill;
        homeLayout.BackColor = Color.Transparent;
        homeLayout.ColumnCount = 1;
        homeLayout.RowCount = 5;
        homeLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        homeLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        contentPanel.Controls.Add(homeLayout);

        sectionTitle = Theme.Label("Панель управления", Theme.H2Font, Theme.Ink);
        sectionTitle.Margin = new Padding(0, 0, 0, 6);
        homeLayout.Controls.Add(sectionTitle, 0, 0);

        sectionHint = Theme.Label("Современная рабочая область: быстрый доступ к пенсионерам, путевкам, санаториям и отчетности без лишних действий.", Theme.BaseFont, Theme.Muted);
        sectionHint.Margin = new Padding(0, 0, 0, 10);
        sectionHint.MaximumSize = new Size(760, 0);
        homeLayout.Controls.Add(sectionHint, 0, 1);

        stats = Theme.Label("", Theme.StatsFont, Theme.Blue);
        stats.Tag = "accent";
        stats.Margin = new Padding(0, 0, 0, 16);
        homeLayout.Controls.Add(stats, 0, 2);

        Panel spacer = new Panel();
        spacer.Dock = DockStyle.Fill;
        spacer.BackColor = Color.Transparent;
        homeLayout.Controls.Add(spacer, 0, 3);

        manualBar = new FlowLayoutPanel();
        manualBar.Dock = DockStyle.Fill;
        manualBar.AutoSize = true;
        manualBar.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        manualBar.FlowDirection = FlowDirection.LeftToRight;
        manualBar.WrapContents = false;
        manualBar.Padding = new Padding(0, 12, 0, 0);
        homeLayout.Controls.Add(manualBar, 0, 4);

        Button manualPdf = new GlassButton();
        manualPdf.Text = "Открыть PDF";
        manualPdf.Size = new Size(140, 40);
        manualPdf.Margin = new Padding(0, 0, 8, 0);
        Theme.StyleButton(manualPdf, false);
        manualPdf.Click += delegate { OpenManual("pdf"); };
        manualBar.Controls.Add(manualPdf);

        Button manualRtf = new GlassButton();
        manualRtf.Text = "Открыть RTF";
        manualRtf.Size = new Size(140, 40);
        manualRtf.Margin = new Padding(0, 0, 8, 0);
        Theme.StyleButton(manualRtf, false);
        manualRtf.Click += delegate { OpenManual("rtf"); };
        manualBar.Controls.Add(manualRtf);

        Button manualTxt = new GlassButton();
        manualTxt.Text = "Открыть TXT";
        manualTxt.Size = new Size(140, 40);
        manualTxt.Margin = new Padding(0, 0, 16, 0);
        Theme.StyleButton(manualTxt, false);
        manualTxt.Click += delegate { OpenManual("txt"); };
        manualBar.Controls.Add(manualTxt);

        Button allPdf = new GlassButton();
        allPdf.Text = "PDF: все таблицы";
        allPdf.Size = new Size(162, 40);
        Theme.StyleButton(allPdf, true);
        allPdf.Click += delegate { ExportAllTablesPdf(); };
        manualBar.Controls.Add(allPdf);

        AddNavButton("Пенсионеры", "Карточки, фильтры, поиск", delegate
        {
            SetActiveSection("Пенсионеры", "Добавляйте, ищите и удаляйте записи пенсионеров в одном месте.");
            ShowEmbeddedForm(new PensionersForm(Db));
        });

        AddNavButton("Журнал путевок", "Выдача, статусы, печать", delegate
        {
            SetActiveSection("Журнал путевок", "Работайте с путевками, просматривайте историю и формируйте отчеты.");
            ShowEmbeddedForm(new VouchersForm(Db));
        });

        AddNavButton("Санатории", "Профили, регионы, коечный фонд", delegate
        {
            SetActiveSection("Санатории", "Содержите справочник санаториев и их параметры.");
            ShowEmbeddedForm(new SanatoriumsForm(Db));
        });

        AddNavButton("Сервис данных", "Очистка и контроль записей", delegate
        {
            SetActiveSection("Сервис данных", "Контролируйте данные и одновременно следите за состоянием базы.");
            ShowEmbeddedForm(new DataToolsForm(Db));
        });

        AddNavButton("Отчет: журнал", "Предпросмотр отчета Access", delegate
        {
            SetActiveSection("Отчет: журнал", "Откройте стандартный Access-отчет напрямую из приложения.");
            ReportRunner.OpenReport(Db.DatabasePath, "rptVoucherIssues", false);
        });

        AddNavButton("Выбрать БД", "Подключить другой .accdb", delegate
        {
            SetActiveSection("Выбор базы данных", "Подключите другую базу и сразу обновите статистику.");
            ChooseDatabase(this, EventArgs.Empty);
        });

        SelectNavButton("Пенсионеры");
        Load += delegate { RefreshStats(); };
    }

    private void AddNavButton(string title, string subtitle, EventHandler click)
    {
        GlassButton button = new GlassButton();
        button.Text = title + "\n" + subtitle;
        button.TextAlign = ContentAlignment.MiddleLeft;
        button.Dock = DockStyle.Fill;
        button.Height = 88;
        button.Margin = new Padding(0, 0, 0, 12);
        button.Padding = new Padding(12, 8, 12, 8);
        Theme.StyleButton(button, false);
        button.Click += delegate
        {
            SelectNavButton(title);
            click(button, EventArgs.Empty);
        };

        navLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        navLayout.RowCount = navLayout.RowCount + 1;
        navLayout.Controls.Add(button, 0, navLayout.RowCount - 1);
        navButtons[title] = button;
    }

    private void SelectNavButton(string title)
    {
        GlassButton selected = null;
        foreach (KeyValuePair<string, GlassButton> pair in navButtons)
        {
            pair.Value.Primary = string.Equals(pair.Key, title, StringComparison.OrdinalIgnoreCase);
            pair.Value.Invalidate();
            if (string.Equals(pair.Key, title, StringComparison.OrdinalIgnoreCase))
            {
                selected = pair.Value;
            }
        }

        if (selected != null)
        {
            selected.Focus();
        }
    }

    private void RefreshStats()
    {
        if (!Db.Exists()) return;
        stats.Text = "Загрузка...";
        System.ComponentModel.BackgroundWorker worker = new System.ComponentModel.BackgroundWorker();
        worker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs ev)
        {
            int[] counts = new int[3];
            counts[0] = Db.ScalarInt("SELECT Count(*) FROM tblPensioners");
            counts[1] = Db.ScalarInt("SELECT Count(*) FROM tblVoucherIssues");
            counts[2] = Db.ScalarInt("SELECT Count(*) FROM tblSanatoriums");
            ev.Result = counts;
        };
        worker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs ev)
        {
            if (ev.Error != null) { ErrorLogger.LogException(ev.Error); stats.Text = ""; return; }
            int[] counts = ev.Result as int[];
            if (counts != null && !IsDisposed)
                stats.Text = counts[0] + " пенсионеров  ·  " + counts[1] + " путевок  ·  " + counts[2] + " санаториев";
        };
        worker.RunWorkerAsync();
    }

    private void ShowEmbeddedForm(Form form)
    {
        if (activeEmbeddedForm != null)
        {
            Form oldForm = activeEmbeddedForm;
            activeEmbeddedForm = null;
            oldForm.Close();
            oldForm.Dispose();
        }

        activeEmbeddedForm = form;
        form.TopLevel = false;
        form.FormBorderStyle = FormBorderStyle.None;
        form.Dock = DockStyle.Fill;
        form.StartPosition = FormStartPosition.Manual;
        form.FormClosed += delegate
        {
            if (ReferenceEquals(activeEmbeddedForm, form))
            {
                activeEmbeddedForm = null;
                contentPanel.Controls.Remove(form);
                homeLayout.Visible = true;
                RefreshStats();
            }
        };

        homeLayout.Visible = false;
        contentPanel.Controls.Add(form);
        form.BringToFront();
        form.Show();
    }

    private void SetActiveSection(string title, string hint)
    {
        sectionTitle.Text = title;
        sectionHint.Text = hint;
    }

    private void OpenManual(string extension)
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string path = Path.Combine(baseDir, "UserGuide." + extension);
        if (!File.Exists(path))
        {
            ToastNotifier.Show(this, "Файл руководства не найден:\n" + path, false);
            return;
        }

        if (extension.Equals("pdf", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                ShowEmbeddedForm(new DocumentViewerForm(path, "Руководство пользователя — PDF"));
            }
            catch (Exception ex)
            {
                ToastNotifier.Show(this, "Не удалось открыть PDF: " + ex.Message, false);
            }
        }
        else if (extension.Equals("rtf", StringComparison.OrdinalIgnoreCase) || extension.Equals("txt", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                ShowEmbeddedForm(new DocumentViewerForm(path, "Руководство пользователя"));
            }
            catch (Exception ex)
            {
                ToastNotifier.Show(this, "Не удалось открыть документ: " + ex.Message, false);
            }
        }
    }

    private void ExportAllTablesPdf()
    {
        Guard(delegate
        {
            if (ReportTables.ExportAllTablesPdf(this, Db))
            {
                ToastNotifier.Show(this, "PDF со всеми таблицами сохранен", true);
            }
        });
    }

    private void ChooseDatabase(object sender, EventArgs e)
    {
        using (OpenFileDialog dialog = new OpenFileDialog())
        {
            dialog.Filter = "Access database (*.accdb;*.mdb)|*.accdb;*.mdb";
            dialog.Title = "Выберите файл базы данных";
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                Db.SetDatabasePath(dialog.FileName);
                RefreshStats();
                ToastNotifier.Show(this, "База подключена и данные обновлены", true);
            }
        }
    }
}


    internal sealed class PensionersForm : GlassForm
    {
        private readonly TextBox search;
        private readonly ComboBox region;
        private readonly CheckBox onlyWithVouchers;
        private readonly DataGridView grid;
        private readonly Label statusLabel;
        private bool _loading;

        public PensionersForm(DbContext db) : base(db)
        {
            Text = "Пенсионеры";
            Size = new Size(1100, 700);
            MinimumSize = new Size(820, 520);
            KeyPreview = true;

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            Controls.Add(panel);

            TableLayoutPanel root = MakeRootLayout();
            panel.Controls.Add(root);

            TableLayoutPanel header = MakeHeaderRow();
            root.Controls.Add(header, 0, 0);
            header.Controls.Add(Theme.Label("Пенсионеры", Theme.TitleFont, Theme.Ink), 0, 0);
            FlowLayoutPanel hBtns = MakeHeaderButtons();
            header.Controls.Add(hBtns, 1, 0);
            Button exportPdf = MakeButton("Экспорт PDF", false); exportPdf.Click += delegate { ExportPdf(); };
            Button load = MakeButton("Обновить", true); load.Click += delegate { LoadGrid(); };
            hBtns.Controls.Add(exportPdf); hBtns.Controls.Add(load);

            FlowLayoutPanel filters = MakeFilterRow();
            root.Controls.Add(filters, 0, 1);
            search = Theme.TextBox("Поиск по ФИО и удостоверению");
            search.Width = 256; search.Height = 32; search.Margin = new Padding(0, 0, 10, 4);
            region = new ComboBox(); region.Width = 210; region.Margin = new Padding(0, 0, 10, 4);
            onlyWithVouchers = new CheckBox();
            onlyWithVouchers.Text = "Только с путевками";
            onlyWithVouchers.AutoSize = true;
            onlyWithVouchers.BackColor = Color.Transparent;
            onlyWithVouchers.Margin = new Padding(4, 6, 0, 4);
            filters.Controls.Add(search); filters.Controls.Add(region); filters.Controls.Add(onlyWithVouchers);

            grid = new DataGridView();
            grid.Margin = new Padding(0, 0, 0, 10);
            Theme.StyleGrid(grid);
            grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs ev)
            {
                if (ev.RowIndex >= 0) EditPensioner(s, EventArgs.Empty);
            };
            new ToolTip().SetToolTip(grid, "Двойной клик — редактировать  |  F5 — обновить  |  Ins — добавить  |  Del — удалить");
            root.Controls.Add(grid, 0, 2);

            TableLayoutPanel actions = MakeActionBar();
            root.Controls.Add(actions, 0, 3);
            FlowLayoutPanel left = MakeLeftActions();
            actions.Controls.Add(left, 0, 0);
            Button add = MakeButton("Добавить", true); add.Margin = new Padding(0, 0, 8, 0); add.Click += AddPensioner;
            Button editBtn = MakeButton("Изменить", false); editBtn.Margin = new Padding(0, 0, 8, 0); editBtn.Click += EditPensioner;
            Button del = MakeButton("Удалить", false); del.Click += DeletePensioner;
            left.Controls.Add(add); left.Controls.Add(editBtn); left.Controls.Add(del);
            statusLabel = Theme.Label("", Theme.BaseFont, Theme.Muted);
            statusLabel.Margin = new Padding(12, 8, 0, 0);
            left.Controls.Add(statusLabel);
            Button close = MakeButton("Назад", false); close.Click += delegate { Close(); };
            FlowLayoutPanel right = MakeRightActions();
            right.Controls.Add(close);
            actions.Controls.Add(right, 1, 0);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(region, Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName", true, "Регион: все");
                    LoadGrid();
                });
            };
            search.TextChanged += delegate { Debouncer.Debounce("pensioners_search", LoadGrid, 350); };
            region.SelectedIndexChanged += delegate { Debouncer.Debounce("pensioners_region", LoadGrid, 200); };
            onlyWithVouchers.CheckedChanged += delegate { Debouncer.Debounce("pensioners_vouchers", LoadGrid, 150); };
            search.KeyDown += delegate(object s, KeyEventArgs ke) {
                if (ke.KeyCode == Keys.Enter) { ke.SuppressKeyPress = true; LoadGrid(); }
            };
            KeyDown += delegate(object s, KeyEventArgs ke) {
                if (ke.KeyCode == Keys.F5) { LoadGrid(); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Delete && grid.Focused && grid.SelectedRows.Count > 0) { DeletePensioner(s, EventArgs.Empty); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Insert) { AddPensioner(s, EventArgs.Empty); ke.Handled = true; }
            };
        }

        private void LoadGrid()
        {
            if (_loading || !Db.Exists()) return;
            _loading = true;
            statusLabel.Text = "Загрузка...";
            grid.Enabled = false;

            List<string> where = new List<string>();
            List<OleDbParameter> parameters = new List<OleDbParameter>();
            string text = search.Text.Trim();
            if (text.Length > 0)
            {
                where.Add("(p.LastName LIKE ? OR p.FirstName LIKE ? OR p.MiddleName LIKE ? OR p.PensionCertificateNo LIKE ?)");
                string value = "%" + text + "%";
                parameters.Add(new OleDbParameter("p1", value));
                parameters.Add(new OleDbParameter("p2", value));
                parameters.Add(new OleDbParameter("p3", value));
                parameters.Add(new OleDbParameter("p4", value));
            }
            LookupItem selectedRegion = region.SelectedItem as LookupItem;
            if (selectedRegion != null && selectedRegion.Id > 0)
            {
                where.Add("p.RegionID = ?");
                parameters.Add(new OleDbParameter("region", selectedRegion.Id));
            }
            if (onlyWithVouchers.Checked)
                where.Add("EXISTS (SELECT 1 FROM tblVoucherIssues AS vi WHERE vi.PensionerID = p.PensionerID)");

            string sql =
                "SELECT p.PensionerID AS [Код], p.LastName AS [Фамилия], p.FirstName AS [Имя], p.MiddleName AS [Отчество], " +
                "r.RegionName AS [Регион], c.CategoryName AS [Категория], p.BirthDate AS [Дата рождения], " +
                "p.PensionCertificateNo AS [Удостоверение], p.Phone AS [Телефон] " +
                "FROM (tblPensioners AS p INNER JOIN tblRegions AS r ON p.RegionID = r.RegionID) " +
                "INNER JOIN tblPensionerCategories AS c ON p.CategoryID = c.CategoryID ";
            if (where.Count > 0) sql += "WHERE " + string.Join(" AND ", where.ToArray()) + " ";
            sql += "ORDER BY p.LastName, p.FirstName";
            OleDbParameter[] pArr = parameters.ToArray();
            string sqlFinal = sql;

            System.ComponentModel.BackgroundWorker worker = new System.ComponentModel.BackgroundWorker();
            worker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs ev) { ev.Result = Db.Query(sqlFinal, pArr); };
            worker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs ev)
            {
                _loading = false;
                grid.Enabled = true;
                if (ev.Error != null) { ErrorLogger.LogException(ev.Error); statusLabel.Text = "Ошибка"; return; }
                DataTable table = (DataTable)ev.Result;
                UiPerformance.BindGrid(grid, table);
                statusLabel.Text = "Записей: " + table.Rows.Count;
            };
            worker.RunWorkerAsync();
        }

        private void ExportPdf()
        {
            Guard(delegate
            {
                if (grid.Rows.Count == 0) LoadGrid();
                bool saved = ProfessionalPdfExporter.ExportGridWithDialog(this, "Пенсионеры - текущая таблица", grid, false, "pensioners.pdf");
                if (saved) ToastNotifier.Show(this, "PDF пенсионеров сохранен", true);
            });
        }

        private void AddPensioner(object sender, EventArgs e)
        {
            Guard(delegate
            {
                using (PensionerEditForm form = new PensionerEditForm(Db))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Пенсионер добавлен. База обновлена", true);
                    }
                }
            });
        }

        private void EditPensioner(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите пенсионера для редактирования.", "Редактирование", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                using (PensionerEditForm form = new PensionerEditForm(Db, id))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Данные пенсионера обновлены", true);
                    }
                }
            });
        }

        private void DeletePensioner(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите пенсионера для удаления.", "Удаление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (MessageBox.Show("Удалить выбранного пенсионера и все его путевки?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                Db.Execute("DELETE FROM tblVoucherIssues WHERE PensionerID = ?", new OleDbParameter("pensioner", id));
                Db.Execute("DELETE FROM tblPensioners WHERE PensionerID = ?", new OleDbParameter("pensioner", id));
                LoadGrid();
                ToastNotifier.Show(this, "Пенсионер удален. База обновлена", true);
            });
        }
    }

    internal sealed class VouchersForm : GlassForm
    {
        private readonly ComboBox region;
        private readonly ComboBox status;
        private readonly CheckBox hideCanceled;
        private readonly DataGridView grid;
        private readonly Label statusLabel;
        private bool _loading;

        public VouchersForm(DbContext db) : base(db)
        {
            Text = "Журнал путевок";
            Size = new Size(1140, 700);
            MinimumSize = new Size(860, 520);
            KeyPreview = true;

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            Controls.Add(panel);

            TableLayoutPanel root = MakeRootLayout();
            panel.Controls.Add(root);

            TableLayoutPanel header = MakeHeaderRow();
            root.Controls.Add(header, 0, 0);
            header.Controls.Add(Theme.Label("Журнал путевок", Theme.TitleFont, Theme.Ink), 0, 0);
            FlowLayoutPanel hBtns = MakeHeaderButtons();
            header.Controls.Add(hBtns, 1, 0);
            Button preview = MakeButton("Предпросмотр", false); preview.Width = 140; preview.Click += delegate { ShowPrintOptions(false); };
            Button print = MakeButton("Печать", false); print.Width = 90; print.Margin = new Padding(8, 0, 8, 0); print.Click += delegate { ShowPrintOptions(true); };
            Button load = MakeButton("Обновить", true); load.Click += delegate { LoadGrid(); };
            hBtns.Controls.Add(preview); hBtns.Controls.Add(print); hBtns.Controls.Add(load);

            FlowLayoutPanel filters = MakeFilterRow();
            root.Controls.Add(filters, 0, 1);
            region = new ComboBox(); region.Width = 210; region.Margin = new Padding(0, 0, 10, 4);
            status = new ComboBox(); status.Width = 190; status.Margin = new Padding(0, 0, 10, 4);
            hideCanceled = new CheckBox();
            hideCanceled.Text = "Скрыть отмененные";
            hideCanceled.Checked = true;
            hideCanceled.AutoSize = true;
            hideCanceled.BackColor = Color.Transparent;
            hideCanceled.Margin = new Padding(4, 6, 0, 4);
            filters.Controls.Add(region); filters.Controls.Add(status); filters.Controls.Add(hideCanceled);

            grid = new DataGridView();
            grid.Margin = new Padding(0, 0, 0, 10);
            Theme.StyleGrid(grid);
            grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs ev)
            {
                if (ev.RowIndex >= 0) EditVoucher(s, EventArgs.Empty);
            };
            new ToolTip().SetToolTip(grid, "Двойной клик — редактировать  |  F5 — обновить  |  Ins — добавить  |  Del — удалить");
            root.Controls.Add(grid, 0, 2);

            TableLayoutPanel actions = MakeActionBar();
            root.Controls.Add(actions, 0, 3);
            FlowLayoutPanel left = MakeLeftActions();
            actions.Controls.Add(left, 0, 0);
            Button add = MakeButton("Добавить", true); add.Margin = new Padding(0, 0, 8, 0); add.Click += AddVoucher;
            Button editBtn = MakeButton("Изменить", false); editBtn.Margin = new Padding(0, 0, 8, 0); editBtn.Click += EditVoucher;
            Button del = MakeButton("Удалить", false); del.Click += DeleteVoucher;
            left.Controls.Add(add); left.Controls.Add(editBtn); left.Controls.Add(del);
            statusLabel = Theme.Label("", Theme.BaseFont, Theme.Muted);
            statusLabel.Margin = new Padding(12, 8, 0, 0);
            left.Controls.Add(statusLabel);
            Button close = MakeButton("Назад", false); close.Click += delegate { Close(); };
            FlowLayoutPanel right = MakeRightActions();
            right.Controls.Add(close);
            actions.Controls.Add(right, 1, 0);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(region, Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName", true, "Регион: все");
                    FillLookup(status, Db.LookupCached("tblVoucherStatuses", "StatusID", "StatusName"), "StatusID", "StatusName", true, "Статус: все");
                    LoadGrid();
                });
            };
            region.SelectedIndexChanged += delegate { Debouncer.Debounce("vouchers_region", LoadGrid, 200); };
            status.SelectedIndexChanged += delegate { Debouncer.Debounce("vouchers_status", LoadGrid, 200); };
            hideCanceled.CheckedChanged += delegate { Debouncer.Debounce("vouchers_canceled", LoadGrid, 150); };
            KeyDown += delegate(object s, KeyEventArgs ke) {
                if (ke.KeyCode == Keys.F5) { LoadGrid(); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Delete && grid.Focused && grid.SelectedRows.Count > 0) { DeleteVoucher(s, EventArgs.Empty); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Insert) { AddVoucher(s, EventArgs.Empty); ke.Handled = true; }
            };
        }

        private void LoadGrid()
        {
            if (_loading || !Db.Exists()) return;
            _loading = true;
            statusLabel.Text = "Загрузка...";
            grid.Enabled = false;

            List<string> where = new List<string>();
            List<OleDbParameter> parameters = new List<OleDbParameter>();
            LookupItem selectedRegion = region.SelectedItem as LookupItem;
            if (selectedRegion != null && selectedRegion.Id > 0)
            {
                where.Add("p.RegionID = ?");
                parameters.Add(new OleDbParameter("region", selectedRegion.Id));
            }
            LookupItem selectedStatus = status.SelectedItem as LookupItem;
            if (selectedStatus != null && selectedStatus.Id > 0)
            {
                where.Add("vi.StatusID = ?");
                parameters.Add(new OleDbParameter("status", selectedStatus.Id));
            }
            if (hideCanceled.Checked)
                where.Add("vs.StatusName <> 'Отменена'");

            string sql =
                "SELECT vi.IssueID, vi.VoucherNo AS [Путевка], vi.IssueDate AS [Регистрация], " +
                "p.LastName & ' ' & p.FirstName AS [Пенсионер], r.RegionName AS [Регион], " +
                "s.SanatoriumName AS [Санаторий], vi.StartDate AS [Заезд], vi.EndDate AS [Выезд], " +
                "DateDiff('d',[vi].[StartDate],[vi].[EndDate]) + 1 AS [Дней], " +
                "(DateDiff('d',[vi].[StartDate],[vi].[EndDate]) + 1) * [s].[PricePerDay] AS [Стоимость], " +
                "vs.StatusName AS [Статус] " +
                "FROM (((tblVoucherIssues AS vi INNER JOIN tblPensioners AS p ON vi.PensionerID = p.PensionerID) " +
                "INNER JOIN tblRegions AS r ON p.RegionID = r.RegionID) " +
                "INNER JOIN tblSanatoriums AS s ON vi.SanatoriumID = s.SanatoriumID) " +
                "INNER JOIN tblVoucherStatuses AS vs ON vi.StatusID = vs.StatusID ";
            if (where.Count > 0) sql += "WHERE " + string.Join(" AND ", where.ToArray()) + " ";
            sql += "ORDER BY vi.IssueDate DESC, vi.VoucherNo";
            OleDbParameter[] pArr = parameters.ToArray();
            string sqlFinal = sql;

            System.ComponentModel.BackgroundWorker worker = new System.ComponentModel.BackgroundWorker();
            worker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs ev) { ev.Result = Db.Query(sqlFinal, pArr); };
            worker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs ev)
            {
                _loading = false;
                grid.Enabled = true;
                if (ev.Error != null) { ErrorLogger.LogException(ev.Error); statusLabel.Text = "Ошибка"; return; }
                DataTable table = (DataTable)ev.Result;
                UiPerformance.BindGrid(grid, table);
                statusLabel.Text = "Записей: " + table.Rows.Count;
            };
            worker.RunWorkerAsync();
        }

        private void AddVoucher(object sender, EventArgs e)
        {
            Guard(delegate
            {
                using (VoucherEditForm form = new VoucherEditForm(Db))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Путевка добавлена. База обновлена", true);
                    }
                }
            });
        }

        private void EditVoucher(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите путевку для редактирования.", "Редактирование", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                using (VoucherEditForm form = new VoucherEditForm(Db, id))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Данные путевки обновлены", true);
                    }
                }
            });
        }

        private void DeleteVoucher(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите путевку для удаления.", "Удаление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (MessageBox.Show("Удалить выбранную запись журнала?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                Db.Execute("DELETE FROM tblVoucherIssues WHERE IssueID = ?", new OleDbParameter("issue", id));
                LoadGrid();
                ToastNotifier.Show(this, "Запись журнала удалена. База обновлена", true);
            });
        }

        private void ShowPrintOptions(bool preferPrint)
        {
            Guard(delegate
            {
                using (PrintOptionsForm options = new PrintOptionsForm(preferPrint))
                {
                    if (options.ShowDialog(this) != DialogResult.OK) return;
                    if (options.Target == PrintTarget.AllTables)
                    {
                        if (ReportTables.ExportAllTablesPdf(this, Db))
                        {
                            ToastNotifier.Show(this, "PDF со всеми таблицами сохранен", true);
                        }
                        return;
                    }

                    if (options.ExportPdf)
                    {
                        bool saved = false;
                        if (options.Target == PrintTarget.AccessJournal)
                        {
                            DataTable table = Db.Query("SELECT VoucherNo AS [Путевка], IssueDate AS [Регистрация], PensionerFullName AS [Пенсионер], PensionerRegion AS [Регион], SanatoriumName AS [Санаторий], StartDate AS [Заезд], EndDate AS [Выезд], DaysCount AS [Дней], TotalCost AS [Стоимость], StatusName AS [Статус] FROM qryReport_VoucherIssuesDetailed ORDER BY IssueDate, VoucherNo");
                            saved = ProfessionalPdfExporter.ExportDataTableWithDialog(this, "Журнал регистрации выдачи путевок", "Система здравоохранения", table, "journal_vouchers.pdf");
                        }
                        else if (options.Target == PrintTarget.AccessRegionTotals)
                        {
                            DataTable table = Db.Query("SELECT SanatoriumRegion AS [Регион санатория], SanatoriumName AS [Санаторий], VoucherCount AS [Путевок], TotalDays AS [Дней], TotalPlannedCost AS [Плановая стоимость] FROM qry04_Totals_BySanatorium ORDER BY SanatoriumRegion, SanatoriumName");
                            saved = ProfessionalPdfExporter.ExportDataTableWithDialog(this, "Итоги по санаториям и регионам", "Система здравоохранения", table, "region_totals.pdf");
                        }
                        else if (options.Target == PrintTarget.VisibleGrid)
                        {
                            saved = ProfessionalPdfExporter.ExportGridWithDialog(this, "Журнал путевок - видимые строки", grid, false, "visible_vouchers.pdf");
                        }
                        else if (options.Target == PrintTarget.SelectedGridRow)
                        {
                            if (grid.SelectedRows.Count == 0)
                            {
                                MessageBox.Show("Выберите строку журнала для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                return;
                            }
                            saved = ProfessionalPdfExporter.ExportGridWithDialog(this, "Журнал путевок - выбранная строка", grid, true, "selected_voucher.pdf");
                        }
                        if (saved) ToastNotifier.Show(this, "PDF-файл красиво собран и сохранен", true);
                        return;
                    }

                    if (options.Target == PrintTarget.AccessJournal)
                    {
                        if (ReportRunner.OpenReport(Db.DatabasePath, "rptVoucherIssues", options.DirectPrint))
                            ToastNotifier.Show(this, options.DirectPrint ? "Отчет журнала отправлен на печать" : "Открыт предпросмотр журнала", true);
                    }
                    else if (options.Target == PrintTarget.AccessRegionTotals)
                    {
                        if (ReportRunner.OpenReport(Db.DatabasePath, "rptRegionGroupedTotals", options.DirectPrint))
                            ToastNotifier.Show(this, options.DirectPrint ? "Итоговый отчет отправлен на печать" : "Открыт предпросмотр итогов", true);
                    }
                    else if (options.Target == PrintTarget.VisibleGrid)
                    {
                        GridPrinter.PrintGrid("Журнал путевок - видимые строки", grid, false, options.DirectPrint);
                        ToastNotifier.Show(this, options.DirectPrint ? "Видимые строки отправлены на печать" : "Открыт предпросмотр видимых строк", true);
                    }
                    else if (options.Target == PrintTarget.SelectedGridRow)
                    {
                        if (grid.SelectedRows.Count == 0)
                        {
                            MessageBox.Show("Выберите строку журнала для печати.", "Печать", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }
                        GridPrinter.PrintGrid("Журнал путевок - выбранная строка", grid, true, options.DirectPrint);
                        ToastNotifier.Show(this, options.DirectPrint ? "Выбранная строка отправлена на печать" : "Открыт предпросмотр выбранной строки", true);
                    }
                }
            });
        }
    }

    internal sealed class SanatoriumsForm : GlassForm
    {
        private readonly ComboBox profile;
        private readonly CheckBox onlyActive;
        private readonly DataGridView grid;
        private readonly Label statusLabel;
        private bool _loading;

        public SanatoriumsForm(DbContext db) : base(db)
        {
            Text = "Санатории";
            Size = new Size(1060, 680);
            MinimumSize = new Size(780, 500);
            KeyPreview = true;

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            Controls.Add(panel);

            TableLayoutPanel root = MakeRootLayout();
            panel.Controls.Add(root);

            TableLayoutPanel header = MakeHeaderRow();
            root.Controls.Add(header, 0, 0);
            header.Controls.Add(Theme.Label("Санатории", Theme.TitleFont, Theme.Ink), 0, 0);
            FlowLayoutPanel hBtns = MakeHeaderButtons();
            header.Controls.Add(hBtns, 1, 0);
            Button exportPdf = MakeButton("Экспорт PDF", false); exportPdf.Click += delegate { ExportPdf(); };
            Button load = MakeButton("Обновить", true); load.Margin = new Padding(8, 0, 0, 0); load.Click += delegate { LoadGrid(); };
            hBtns.Controls.Add(exportPdf); hBtns.Controls.Add(load);

            FlowLayoutPanel filters = MakeFilterRow();
            root.Controls.Add(filters, 0, 1);
            profile = new ComboBox(); profile.Width = 240; profile.Margin = new Padding(0, 0, 10, 4);
            onlyActive = new CheckBox();
            onlyActive.Text = "Только работающие";
            onlyActive.Checked = true;
            onlyActive.AutoSize = true;
            onlyActive.BackColor = Color.Transparent;
            onlyActive.Margin = new Padding(4, 6, 0, 4);
            filters.Controls.Add(profile); filters.Controls.Add(onlyActive);

            grid = new DataGridView();
            grid.Margin = new Padding(0, 0, 0, 10);
            Theme.StyleGrid(grid);
            grid.CellDoubleClick += delegate(object s, DataGridViewCellEventArgs ev)
            {
                if (ev.RowIndex >= 0) EditSanatorium(s, EventArgs.Empty);
            };
            new ToolTip().SetToolTip(grid, "Двойной клик — редактировать  |  F5 — обновить  |  Ins — добавить  |  Del — удалить");
            root.Controls.Add(grid, 0, 2);

            TableLayoutPanel actions = MakeActionBar();
            root.Controls.Add(actions, 0, 3);
            FlowLayoutPanel left = MakeLeftActions();
            actions.Controls.Add(left, 0, 0);
            Button add = MakeButton("Добавить", true); add.Margin = new Padding(0, 0, 8, 0); add.Click += AddSanatorium;
            Button editBtn = MakeButton("Изменить", false); editBtn.Margin = new Padding(0, 0, 8, 0); editBtn.Click += EditSanatorium;
            Button del = MakeButton("Удалить", false); del.Click += DeleteSanatorium;
            left.Controls.Add(add); left.Controls.Add(editBtn); left.Controls.Add(del);
            statusLabel = Theme.Label("", Theme.BaseFont, Theme.Muted);
            statusLabel.Margin = new Padding(12, 8, 0, 0);
            left.Controls.Add(statusLabel);
            Button close = MakeButton("Назад", false); close.Click += delegate { Close(); };
            FlowLayoutPanel right = MakeRightActions();
            right.Controls.Add(close);
            actions.Controls.Add(right, 1, 0);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(profile, Db.LookupCached("tblMedicalProfiles", "ProfileID", "ProfileName"), "ProfileID", "ProfileName", true, "Профиль: все");
                    LoadGrid();
                });
            };
            profile.SelectedIndexChanged += delegate { Debouncer.Debounce("sanatoriums_profile", LoadGrid, 200); };
            onlyActive.CheckedChanged += delegate { Debouncer.Debounce("sanatoriums_active", LoadGrid, 150); };
            KeyDown += delegate(object s, KeyEventArgs ke) {
                if (ke.KeyCode == Keys.F5) { LoadGrid(); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Delete && grid.Focused && grid.SelectedRows.Count > 0) { DeleteSanatorium(s, EventArgs.Empty); ke.Handled = true; }
                else if (ke.KeyCode == Keys.Insert) { AddSanatorium(s, EventArgs.Empty); ke.Handled = true; }
            };
        }

        private void LoadGrid()
        {
            if (_loading || !Db.Exists()) return;
            _loading = true;
            statusLabel.Text = "Загрузка...";
            grid.Enabled = false;

            List<string> where = new List<string>();
            List<OleDbParameter> parameters = new List<OleDbParameter>();
            LookupItem selectedProfile = profile.SelectedItem as LookupItem;
            if (selectedProfile != null && selectedProfile.Id > 0)
            {
                where.Add("s.ProfileID = ?");
                parameters.Add(new OleDbParameter("profile", selectedProfile.Id));
            }
            if (onlyActive.Checked)
                where.Add("s.IsActive = True");

            string sql =
                "SELECT s.SanatoriumID, s.SanatoriumName AS [Санаторий], r.RegionName AS [Регион], mp.ProfileName AS [Профиль], " +
                "s.CapacityBeds AS [Коек], s.PricePerDay AS [Цена дня], s.Phone AS [Телефон], s.Address AS [Адрес] " +
                "FROM (tblSanatoriums AS s INNER JOIN tblRegions AS r ON s.RegionID = r.RegionID) " +
                "INNER JOIN tblMedicalProfiles AS mp ON s.ProfileID = mp.ProfileID ";
            if (where.Count > 0) sql += "WHERE " + string.Join(" AND ", where.ToArray()) + " ";
            sql += "ORDER BY s.SanatoriumName";
            OleDbParameter[] pArr = parameters.ToArray();
            string sqlFinal = sql;

            System.ComponentModel.BackgroundWorker worker = new System.ComponentModel.BackgroundWorker();
            worker.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs ev) { ev.Result = Db.Query(sqlFinal, pArr); };
            worker.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs ev)
            {
                _loading = false;
                grid.Enabled = true;
                if (ev.Error != null) { ErrorLogger.LogException(ev.Error); statusLabel.Text = "Ошибка"; return; }
                DataTable table = (DataTable)ev.Result;
                UiPerformance.BindGrid(grid, table);
                statusLabel.Text = "Записей: " + table.Rows.Count;
            };
            worker.RunWorkerAsync();
        }

        private void ExportPdf()
        {
            Guard(delegate
            {
                if (grid.Rows.Count == 0) LoadGrid();
                bool saved = ProfessionalPdfExporter.ExportGridWithDialog(this, "Санатории - текущая таблица", grid, false, "sanatoriums.pdf");
                if (saved) ToastNotifier.Show(this, "PDF санаториев сохранен", true);
            });
        }

        private void AddSanatorium(object sender, EventArgs e)
        {
            Guard(delegate
            {
                using (SanatoriumEditForm form = new SanatoriumEditForm(Db))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Санаторий добавлен. База обновлена", true);
                    }
                }
            });
        }

        private void EditSanatorium(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите санаторий для редактирования.", "Редактирование", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                using (SanatoriumEditForm form = new SanatoriumEditForm(Db, id))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Данные санатория обновлены", true);
                    }
                }
            });
        }

        private void DeleteSanatorium(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    MessageBox.Show("Выберите санаторий для удаления.", "Удаление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (MessageBox.Show("Удалить санаторий и связанные путевки?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
                Db.Execute("DELETE FROM tblVoucherIssues WHERE SanatoriumID = ?", new OleDbParameter("sanatorium", id));
                Db.Execute("DELETE FROM tblSanatoriums WHERE SanatoriumID = ?", new OleDbParameter("sanatorium", id));
                LoadGrid();
                ToastNotifier.Show(this, "Санаторий удален. База обновлена", true);
            });
        }
    }

    
internal abstract class RecordFormBase : GlassForm
{
    protected readonly GlassPanel Panel;
    private readonly TableLayoutPanel fieldLayout;
    private readonly FlowLayoutPanel actionBar;
    private int rowIndex;

    protected RecordFormBase(DbContext db, string title, Size size)
        : base(db)
    {
        Text = title;
        Size = size;
        MinimumSize = new Size(Math.Max(680, size.Width), Math.Max(480, size.Height));
        MaximizeBox = false;

        Panel = new GlassPanel();
        Panel.Dock = DockStyle.Fill;
        Panel.Padding = new Padding(24);
        Panel.AutoScroll = true;
        Controls.Add(Panel);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 3;
        root.BackColor = Color.Transparent;
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Panel.Controls.Add(root);

        Label header = Theme.Label(title, Theme.H2Font, Theme.Ink);
        header.Margin = new Padding(4, 0, 4, 12);
        root.Controls.Add(header, 0, 0);

        fieldLayout = new TableLayoutPanel();
        fieldLayout.Dock = DockStyle.Fill;
        fieldLayout.AutoSize = true;
        fieldLayout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        fieldLayout.BackColor = Color.Transparent;
        fieldLayout.ColumnCount = 2;
        fieldLayout.RowCount = 0;
        fieldLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 200f));
        fieldLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.Controls.Add(fieldLayout, 0, 1);

        actionBar = new FlowLayoutPanel();
        actionBar.Dock = DockStyle.Fill;
        actionBar.AutoSize = true;
        actionBar.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        actionBar.FlowDirection = FlowDirection.RightToLeft;
        actionBar.WrapContents = false;
        actionBar.Padding = new Padding(0, 14, 0, 0);
        root.Controls.Add(actionBar, 0, 2);
    }

    protected TextBox AddText(string label, string value, int width)
    {
        AddFieldLabel(label);
        TextBox box = Theme.TextBox(label);
        box.Text = value;
        box.Dock = DockStyle.Fill;
        box.MinimumSize = new Size(width, 28);
        box.Margin = new Padding(0, 4, 0, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(box, 1, rowIndex);
        rowIndex++;
        return box;
    }

    protected ComboBox AddLookup(string label, DataTable data, string idField, string nameField)
    {
        AddFieldLabel(label);
        ComboBox combo = new ComboBox();
        combo.Dock = DockStyle.Fill;
        combo.MinimumSize = new Size(330, 28);
        combo.Margin = new Padding(0, 4, 0, 4);
        FillLookup(combo, data, idField, nameField, false);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(combo, 1, rowIndex);
        rowIndex++;
        return combo;
    }

    protected ComboBox AddSimpleCombo(string label, string[] values)
    {
        AddFieldLabel(label);
        ComboBox combo = new ComboBox();
        combo.Dock = DockStyle.Left;
        combo.Width = 220;
        combo.DropDownStyle = ComboBoxStyle.DropDownList;
        combo.Items.AddRange(values);
        if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        combo.Margin = new Padding(0, 4, 0, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(combo, 1, rowIndex);
        rowIndex++;
        return combo;
    }

    protected DateTimePicker AddDate(string label, DateTime value)
    {
        AddFieldLabel(label);
        DateTimePicker picker = new DateTimePicker();
        picker.Format = DateTimePickerFormat.Short;
        picker.Value = value;
        picker.Dock = DockStyle.Left;
        picker.Width = 180;
        picker.Margin = new Padding(0, 4, 0, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(picker, 1, rowIndex);
        rowIndex++;
        return picker;
    }

    protected NumericUpDown AddNumber(string label, decimal value, decimal max)
    {
        AddFieldLabel(label);
        NumericUpDown number = new NumericUpDown();
        number.Minimum = 0;
        number.Maximum = max;
        number.DecimalPlaces = max > 1000 ? 2 : 0;
        number.ThousandsSeparator = true;
        number.Value = value;
        number.Dock = DockStyle.Left;
        number.Width = 180;
        number.Margin = new Padding(0, 4, 0, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(number, 1, rowIndex);
        rowIndex++;
        return number;
    }

    protected CheckBox AddCheck(string text, bool value)
    {
        CheckBox check = new CheckBox();
        check.Text = text;
        check.Checked = value;
        check.AutoSize = true;
        check.BackColor = Color.Transparent;
        check.Margin = new Padding(0, 8, 0, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(check, 0, rowIndex);
        fieldLayout.SetColumnSpan(check, 2);
        rowIndex++;
        return check;
    }

    protected void AddActions(EventHandler save)
    {
        Button ok = new GlassButton();
        ok.Text = "Сохранить";
        ok.Size = new Size(130, 38);
        ok.Margin = new Padding(8, 0, 0, 0);
        Theme.StyleButton(ok, true);
        ok.Click += save;

        Button cancel = new GlassButton();
        cancel.Text = "Отмена";
        cancel.Size = new Size(120, 38);
        Theme.StyleButton(cancel, false);
        cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };

        actionBar.Controls.Add(ok);
        actionBar.Controls.Add(cancel);
    }

    protected int SelectedId(ComboBox combo)
    {
        LookupItem item = combo.SelectedItem as LookupItem;
        return item == null ? 0 : item.Id;
    }

    protected object Optional(string value)
    {
        return string.IsNullOrWhiteSpace(value) ? (object)DBNull.Value : value.Trim();
    }

    protected bool Require(params TextBox[] boxes)
    {
        foreach (TextBox box in boxes)
        {
            if (string.IsNullOrWhiteSpace(box.Text))
            {
                MessageBox.Show("Заполните обязательные поля.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                box.Focus();
                return false;
            }
        }

        return true;
    }

    protected void SelectLookup(ComboBox combo, int id)
    {
        foreach (object item in combo.Items)
        {
            LookupItem lookupItem = item as LookupItem;
            if (lookupItem != null && lookupItem.Id == id)
            {
                combo.SelectedItem = item;
                return;
            }
        }
    }

    protected void SelectComboText(ComboBox combo, string value)
    {
        for (int i = 0; i < combo.Items.Count; i++)
        {
            if (string.Equals(Convert.ToString(combo.Items[i]), value, StringComparison.Ordinal))
            {
                combo.SelectedIndex = i;
                return;
            }
        }
    }

    protected string DbString(DataRow row, string column)
    {
        object value = row[column];
        return value == null || value == DBNull.Value ? "" : Convert.ToString(value);
    }

    private void AddFieldLabel(string text)
    {
        Label label = Theme.Label(text, Theme.BaseFont, Theme.Muted);
        label.Anchor = AnchorStyles.Left;
        label.Margin = new Padding(0, 8, 12, 4);
        fieldLayout.RowCount = rowIndex + 1;
        fieldLayout.Controls.Add(label, 0, rowIndex);
    }
}


    internal sealed class PensionerEditForm : RecordFormBase
    {
        private readonly int editId;
        private readonly ComboBox region;
        private readonly ComboBox category;
        private readonly TextBox last;
        private readonly TextBox first;
        private readonly TextBox middle;
        private readonly ComboBox gender;
        private readonly DateTimePicker birth;
        private readonly TextBox certificate;
        private readonly TextBox snils;
        private readonly TextBox phone;
        private readonly TextBox address;

        public PensionerEditForm(DbContext db, int editId = 0)
            : base(db, editId == 0 ? "Добавление пенсионера" : "Редактирование пенсионера", new Size(720, 690))
        {
            this.editId = editId;
            region = AddLookup("Регион", Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName");
            category = AddLookup("Категория", Db.LookupCached("tblPensionerCategories", "CategoryID", "CategoryName"), "CategoryID", "CategoryName");
            last = AddText("Фамилия *", "", 330);
            first = AddText("Имя *", "", 330);
            middle = AddText("Отчество", "", 330);
            gender = AddSimpleCombo("Пол", new string[] { "М", "Ж" });
            birth = AddDate("Дата рождения", DateTime.Today.AddYears(-65));
            certificate = AddText("Удостоверение *", editId == 0 ? "PEN-" + DateTime.Now.ToString("yyyyMMddHHmmss") : "", 330);
            snils = AddText("СНИЛС", "", 250);
            phone = AddText("Телефон", "", 250);
            address = AddText("Адрес", "", 380);
            AddActions(Save);
            if (editId > 0)
            {
                Load += delegate { LoadExisting(); };
            }
        }

        private void LoadExisting()
        {
            Guard(delegate
            {
                DataTable dt = Db.Query("SELECT * FROM tblPensioners WHERE PensionerID = ?", new OleDbParameter("id", editId));
                if (dt.Rows.Count == 0) return;
                DataRow row = dt.Rows[0];
                SelectLookup(region, Convert.ToInt32(row["RegionID"]));
                SelectLookup(category, Convert.ToInt32(row["CategoryID"]));
                last.Text = DbString(row, "LastName");
                first.Text = DbString(row, "FirstName");
                middle.Text = DbString(row, "MiddleName");
                SelectComboText(gender, DbString(row, "Gender"));
                object birthVal = row["BirthDate"];
                if (birthVal != null && birthVal != DBNull.Value)
                    birth.Value = Convert.ToDateTime(birthVal);
                certificate.Text = DbString(row, "PensionCertificateNo");
                snils.Text = DbString(row, "SNILS");
                phone.Text = DbString(row, "Phone");
                address.Text = DbString(row, "Address");
            });
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(last, first, certificate)) return;
                if (editId > 0)
                {
                    Db.Execute(
                        "UPDATE tblPensioners SET RegionID=?, CategoryID=?, LastName=?, FirstName=?, MiddleName=?, Gender=?, BirthDate=?, PensionCertificateNo=?, SNILS=?, Phone=?, Address=? WHERE PensionerID=?",
                        new OleDbParameter("region", SelectedId(region)),
                        new OleDbParameter("category", SelectedId(category)),
                        new OleDbParameter("last", last.Text.Trim()),
                        new OleDbParameter("first", first.Text.Trim()),
                        new OleDbParameter("middle", Optional(middle.Text)),
                        new OleDbParameter("gender", Convert.ToString(gender.SelectedItem)),
                        new OleDbParameter("birth", birth.Value.Date),
                        new OleDbParameter("certificate", certificate.Text.Trim()),
                        new OleDbParameter("snils", Optional(snils.Text)),
                        new OleDbParameter("phone", Optional(phone.Text)),
                        new OleDbParameter("address", Optional(address.Text)),
                        new OleDbParameter("id", editId));
                }
                else
                {
                    Db.Execute(
                        "INSERT INTO tblPensioners (RegionID, CategoryID, LastName, FirstName, MiddleName, Gender, BirthDate, PensionCertificateNo, SNILS, Phone, Address) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        new OleDbParameter("region", SelectedId(region)),
                        new OleDbParameter("category", SelectedId(category)),
                        new OleDbParameter("last", last.Text.Trim()),
                        new OleDbParameter("first", first.Text.Trim()),
                        new OleDbParameter("middle", Optional(middle.Text)),
                        new OleDbParameter("gender", Convert.ToString(gender.SelectedItem)),
                        new OleDbParameter("birth", birth.Value.Date),
                        new OleDbParameter("certificate", certificate.Text.Trim()),
                        new OleDbParameter("snils", Optional(snils.Text)),
                        new OleDbParameter("phone", Optional(phone.Text)),
                        new OleDbParameter("address", Optional(address.Text)));
                }
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    internal sealed class VoucherEditForm : RecordFormBase
    {
        private readonly int editId;
        private readonly TextBox number;
        private readonly ComboBox pensioner;
        private readonly ComboBox sanatorium;
        private readonly ComboBox funding;
        private readonly ComboBox status;
        private readonly DateTimePicker issueDate;
        private readonly DateTimePicker startDate;
        private readonly DateTimePicker endDate;
        private readonly NumericUpDown copay;
        private readonly TextBox notes;

        public VoucherEditForm(DbContext db, int editId = 0)
            : base(db, editId == 0 ? "Добавление путевки" : "Редактирование путевки", new Size(760, 660))
        {
            this.editId = editId;
            number = AddText("Номер путевки *", editId == 0 ? "V-" + DateTime.Now.ToString("yyyyMMddHHmmss") : "", 300);
            pensioner = AddLookup("Пенсионер", Db.Query("SELECT PensionerID, LastName & ' ' & FirstName & ' (' & PensionCertificateNo & ')' AS FullName FROM tblPensioners ORDER BY LastName, FirstName"), "PensionerID", "FullName");
            sanatorium = AddLookup("Санаторий", Db.LookupCached("tblSanatoriums", "SanatoriumID", "SanatoriumName"), "SanatoriumID", "SanatoriumName");
            funding = AddLookup("Источник", Db.LookupCached("tblFundingSources", "FundingSourceID", "FundingSourceName"), "FundingSourceID", "FundingSourceName");
            status = AddLookup("Статус", Db.LookupCached("tblVoucherStatuses", "StatusID", "StatusName"), "StatusID", "StatusName");
            issueDate = AddDate("Дата регистрации", DateTime.Today);
            startDate = AddDate("Дата заезда", DateTime.Today.AddDays(14));
            endDate = AddDate("Дата выезда", DateTime.Today.AddDays(27));
            copay = AddNumber("Доплата, %", 0, 100);
            notes = AddText("Примечание", "", 380);
            AddActions(Save);
            if (editId > 0)
            {
                Load += delegate { LoadExisting(); };
            }
        }

        private void LoadExisting()
        {
            Guard(delegate
            {
                DataTable dt = Db.Query("SELECT * FROM tblVoucherIssues WHERE IssueID = ?", new OleDbParameter("id", editId));
                if (dt.Rows.Count == 0) return;
                DataRow row = dt.Rows[0];
                number.Text = DbString(row, "VoucherNo");
                SelectLookup(pensioner, Convert.ToInt32(row["PensionerID"]));
                SelectLookup(sanatorium, Convert.ToInt32(row["SanatoriumID"]));
                SelectLookup(funding, Convert.ToInt32(row["FundingSourceID"]));
                SelectLookup(status, Convert.ToInt32(row["StatusID"]));
                object issued = row["IssueDate"];
                if (issued != null && issued != DBNull.Value) issueDate.Value = Convert.ToDateTime(issued);
                object started = row["StartDate"];
                if (started != null && started != DBNull.Value) startDate.Value = Convert.ToDateTime(started);
                object ended = row["EndDate"];
                if (ended != null && ended != DBNull.Value) endDate.Value = Convert.ToDateTime(ended);
                object copayVal = row["CoPaymentPercent"];
                if (copayVal != null && copayVal != DBNull.Value)
                    copay.Value = Math.Min(copay.Maximum, Math.Max(copay.Minimum, Convert.ToDecimal(copayVal)));
                notes.Text = DbString(row, "Notes");
            });
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(number)) return;
                if (SelectedId(pensioner) == 0 || SelectedId(sanatorium) == 0)
                {
                    MessageBox.Show("Нужны хотя бы один пенсионер и один санаторий.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (endDate.Value.Date < startDate.Value.Date)
                {
                    MessageBox.Show("Дата выезда не может быть раньше даты заезда.", "Проверка", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (editId > 0)
                {
                    Db.Execute(
                        "UPDATE tblVoucherIssues SET IssueDate=?, VoucherNo=?, PensionerID=?, SanatoriumID=?, FundingSourceID=?, StatusID=?, StartDate=?, EndDate=?, CoPaymentPercent=?, Notes=? WHERE IssueID=?",
                        new OleDbParameter("issueDate", issueDate.Value.Date),
                        new OleDbParameter("voucher", number.Text.Trim()),
                        new OleDbParameter("pensioner", SelectedId(pensioner)),
                        new OleDbParameter("sanatorium", SelectedId(sanatorium)),
                        new OleDbParameter("funding", SelectedId(funding)),
                        new OleDbParameter("status", SelectedId(status)),
                        new OleDbParameter("startDate", startDate.Value.Date),
                        new OleDbParameter("endDate", endDate.Value.Date),
                        new OleDbParameter("copay", Convert.ToInt32(copay.Value)),
                        new OleDbParameter("notes", Optional(notes.Text)),
                        new OleDbParameter("id", editId));
                }
                else
                {
                    Db.Execute(
                        "INSERT INTO tblVoucherIssues (IssueDate, VoucherNo, PensionerID, SanatoriumID, FundingSourceID, StatusID, StartDate, EndDate, CoPaymentPercent, Notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                        new OleDbParameter("issueDate", issueDate.Value.Date),
                        new OleDbParameter("voucher", number.Text.Trim()),
                        new OleDbParameter("pensioner", SelectedId(pensioner)),
                        new OleDbParameter("sanatorium", SelectedId(sanatorium)),
                        new OleDbParameter("funding", SelectedId(funding)),
                        new OleDbParameter("status", SelectedId(status)),
                        new OleDbParameter("startDate", startDate.Value.Date),
                        new OleDbParameter("endDate", endDate.Value.Date),
                        new OleDbParameter("copay", Convert.ToInt32(copay.Value)),
                        new OleDbParameter("notes", Optional(notes.Text)));
                }
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    internal sealed class SanatoriumEditForm : RecordFormBase
    {
        private readonly int editId;
        private readonly ComboBox region;
        private readonly ComboBox profile;
        private readonly TextBox name;
        private readonly TextBox address;
        private readonly TextBox phone;
        private readonly NumericUpDown beds;
        private readonly NumericUpDown price;
        private readonly CheckBox active;

        public SanatoriumEditForm(DbContext db, int editId = 0)
            : base(db, editId == 0 ? "Добавление санатория" : "Редактирование санатория", new Size(720, 600))
        {
            this.editId = editId;
            region = AddLookup("Регион", Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName");
            profile = AddLookup("Профиль", Db.LookupCached("tblMedicalProfiles", "ProfileID", "ProfileName"), "ProfileID", "ProfileName");
            name = AddText("Санаторий *", "", 360);
            address = AddText("Адрес", "", 390);
            phone = AddText("Телефон", "", 240);
            beds = AddNumber("Коек", 100, 5000);
            price = AddNumber("Цена дня", 3500, 100000);
            active = AddCheck("Работает", true);
            AddActions(Save);
            if (editId > 0)
            {
                Load += delegate { LoadExisting(); };
            }
        }

        private void LoadExisting()
        {
            Guard(delegate
            {
                DataTable dt = Db.Query("SELECT * FROM tblSanatoriums WHERE SanatoriumID = ?", new OleDbParameter("id", editId));
                if (dt.Rows.Count == 0) return;
                DataRow row = dt.Rows[0];
                SelectLookup(region, Convert.ToInt32(row["RegionID"]));
                SelectLookup(profile, Convert.ToInt32(row["ProfileID"]));
                name.Text = DbString(row, "SanatoriumName");
                address.Text = DbString(row, "Address");
                phone.Text = DbString(row, "Phone");
                object bedsVal = row["CapacityBeds"];
                if (bedsVal != null && bedsVal != DBNull.Value)
                    beds.Value = Math.Min(beds.Maximum, Math.Max(beds.Minimum, Convert.ToDecimal(bedsVal)));
                object priceVal = row["PricePerDay"];
                if (priceVal != null && priceVal != DBNull.Value)
                    price.Value = Math.Min(price.Maximum, Math.Max(price.Minimum, Convert.ToDecimal(priceVal)));
                object activeVal = row["IsActive"];
                if (activeVal != null && activeVal != DBNull.Value)
                    active.Checked = Convert.ToBoolean(activeVal);
            });
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(name)) return;
                if (editId > 0)
                {
                    Db.Execute(
                        "UPDATE tblSanatoriums SET RegionID=?, ProfileID=?, SanatoriumName=?, Address=?, Phone=?, CapacityBeds=?, PricePerDay=?, IsActive=? WHERE SanatoriumID=?",
                        new OleDbParameter("region", SelectedId(region)),
                        new OleDbParameter("profile", SelectedId(profile)),
                        new OleDbParameter("name", name.Text.Trim()),
                        new OleDbParameter("address", Optional(address.Text)),
                        new OleDbParameter("phone", Optional(phone.Text)),
                        new OleDbParameter("beds", Convert.ToInt32(beds.Value)),
                        new OleDbParameter("price", price.Value),
                        new OleDbParameter("active", active.Checked),
                        new OleDbParameter("id", editId));
                }
                else
                {
                    Db.Execute(
                        "INSERT INTO tblSanatoriums (RegionID, ProfileID, SanatoriumName, Address, Phone, CapacityBeds, PricePerDay, IsActive) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                        new OleDbParameter("region", SelectedId(region)),
                        new OleDbParameter("profile", SelectedId(profile)),
                        new OleDbParameter("name", name.Text.Trim()),
                        new OleDbParameter("address", Optional(address.Text)),
                        new OleDbParameter("phone", Optional(phone.Text)),
                        new OleDbParameter("beds", Convert.ToInt32(beds.Value)),
                        new OleDbParameter("price", price.Value),
                        new OleDbParameter("active", active.Checked));
                }
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    
// ── Chart helpers ────────────────────────────────────────────────────────

    internal struct PieSlice { public string Label; public int Value; }
    internal struct BarEntry { public string Label; public int Value; public int Total; public Color Color; }

    internal sealed class PieChartControl : Panel
    {
        private static readonly Color[] Palette = {
            Color.FromArgb(59, 130, 246), Color.FromArgb(34, 197, 94),
            Color.FromArgb(251, 146, 60), Color.FromArgb(239, 68, 68),
            Color.FromArgb(168, 85, 247),
        };
        private PieSlice[] _slices;
        private string _title = "";

        public PieChartControl()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        public void SetData(string title, PieSlice[] slices) { _title = title; _slices = slices; Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Color bg = Theme.Dark ? Color.FromArgb(24, 34, 54) : Color.White;
            Color border = Theme.Dark ? Color.FromArgb(44, 60, 90) : Color.FromArgb(205, 218, 240);
            using (SolidBrush b = new SolidBrush(bg)) g.FillRectangle(b, ClientRectangle);
            using (Pen p = new Pen(border)) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);

            using (Font tf = new Font("Segoe UI Semibold", 10f, FontStyle.Bold))
            using (SolidBrush b = new SolidBrush(Theme.Ink))
                g.DrawString(_title, tf, b, 14, 12);

            if (_slices == null || _slices.Length == 0)
            {
                using (SolidBrush b = new SolidBrush(Theme.Muted))
                    g.DrawString("Нет данных", Theme.BaseFont, b, 14, 38);
                return;
            }

            int total = 0;
            foreach (PieSlice s in _slices) total += s.Value;
            if (total == 0) return;

            int legendW = 170;
            int chartLeft = 14; int chartTop = 38;
            int chartAreaW = Width - legendW - chartLeft - 10;
            int chartAreaH = Height - chartTop - 10;
            int r = Math.Min(chartAreaW, chartAreaH) / 2 - 6;
            if (r < 12) return;
            int cx = chartLeft + chartAreaW / 2;
            int cy = chartTop + chartAreaH / 2;
            int innerR = r * 56 / 100;

            float angle = -90f;
            for (int i = 0; i < _slices.Length; i++)
            {
                float sweep = 360f * _slices[i].Value / total;
                using (SolidBrush b = new SolidBrush(Palette[i % Palette.Length]))
                    g.FillPie(b, cx - r, cy - r, r * 2, r * 2, angle, sweep);
                angle += sweep;
            }
            using (SolidBrush b = new SolidBrush(bg))
                g.FillEllipse(b, cx - innerR, cy - innerR, innerR * 2, innerR * 2);

            using (Font f = new Font("Segoe UI Semibold", 15f, FontStyle.Bold))
            using (SolidBrush b = new SolidBrush(Theme.Ink))
            {
                string txt = total.ToString();
                SizeF sz = g.MeasureString(txt, f);
                g.DrawString(txt, f, b, cx - sz.Width / 2, cy - sz.Height / 2);
            }

            int lx = Width - legendW + 6;
            int ly = chartTop + 4;
            using (Font f = new Font("Segoe UI", 8.5f))
            {
                for (int i = 0; i < _slices.Length && i < Palette.Length; i++)
                {
                    Color c = Palette[i % Palette.Length];
                    using (SolidBrush b = new SolidBrush(c))
                        g.FillRectangle(b, lx, ly + i * 28 + 6, 11, 11);
                    int pct = _slices[i].Value * 100 / total;
                    using (SolidBrush b = new SolidBrush(Theme.Ink))
                        g.DrawString(_slices[i].Label + ": " + _slices[i].Value + " (" + pct + "%)", f, b, lx + 16, ly + i * 28);
                }
            }
        }
    }

    internal sealed class BarChartControl : Panel
    {
        private BarEntry[] _items;
        private string _title = "";

        public BarChartControl()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);
        }

        public void SetData(string title, BarEntry[] items) { _title = title; _items = items; Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            Color bg = Theme.Dark ? Color.FromArgb(24, 34, 54) : Color.White;
            Color border = Theme.Dark ? Color.FromArgb(44, 60, 90) : Color.FromArgb(205, 218, 240);
            using (SolidBrush b = new SolidBrush(bg)) g.FillRectangle(b, ClientRectangle);
            using (Pen p = new Pen(border)) g.DrawRectangle(p, 0, 0, Width - 1, Height - 1);

            using (Font tf = new Font("Segoe UI Semibold", 10f, FontStyle.Bold))
            using (SolidBrush b = new SolidBrush(Theme.Ink))
                g.DrawString(_title, tf, b, 14, 12);

            if (_items == null || _items.Length == 0)
            {
                using (SolidBrush b = new SolidBrush(Theme.Muted))
                    g.DrawString("Нет данных", Theme.BaseFont, b, 14, 38);
                return;
            }

            int lblW = 155; int valW = 72;
            int barLeft = lblW + 14;
            int barRight = Width - valW - 10;
            int barW = barRight - barLeft;
            if (barW < 10) return;

            int bH = 24; int gap = 16; int top = 40;

            for (int i = 0; i < _items.Length; i++)
            {
                int y = top + i * (bH + gap);
                BarEntry item = _items[i];

                using (Font f = new Font("Segoe UI", 9f))
                using (SolidBrush b = new SolidBrush(Theme.Ink))
                    g.DrawString(item.Label, f, b, 14, y + 3);

                Color track = Theme.Dark ? Color.FromArgb(38, 52, 80) : Color.FromArgb(228, 234, 248);
                using (SolidBrush b = new SolidBrush(track))
                    g.FillRectangle(b, barLeft, y, barW, bH);

                if (item.Total > 0 && item.Value > 0)
                {
                    int fw = Math.Max(4, barW * item.Value / item.Total);
                    using (SolidBrush b = new SolidBrush(item.Color))
                        g.FillRectangle(b, barLeft, y, fw, bH);
                }

                int pct = item.Total > 0 ? item.Value * 100 / item.Total : 0;
                using (Font f = new Font("Segoe UI Semibold", 9f, FontStyle.Bold))
                using (SolidBrush b = new SolidBrush(Theme.Ink))
                    g.DrawString(item.Value + " / " + item.Total, f, b, barRight + 6, y + 1);
                using (Font f = new Font("Segoe UI", 8f))
                using (SolidBrush b = new SolidBrush(Theme.Muted))
                    g.DrawString(pct + "%", f, b, barRight + 6, y + 13);
            }
        }
    }

// ── DataToolsForm ────────────────────────────────────────────────────────

internal sealed class DataToolsForm : GlassForm
{
    private Label _vPensioners, _sPensioners;
    private Label _vVouchers,   _sVouchers;
    private Label _vSanatoriums, _sSanatoriums;
    private Label _vCost,        _sCost;
    private Label _vDays,        _sDays;
    private PieChartControl _pieChart;
    private BarChartControl _barChart;
    private FlowLayoutPanel _cardsFlow;
    private bool _statsLoading;

    public DataToolsForm(DbContext db) : base(db)
    {
        Text = "Сервис данных";
        Size = new Size(1060, 720);
        MinimumSize = new Size(980, 680);

        GlassPanel panel = new GlassPanel();
        panel.Dock = DockStyle.Fill;
        panel.Padding = new Padding(22);
        panel.AutoScroll = true;
        Controls.Add(panel);

        TableLayoutPanel layout = new TableLayoutPanel();
        layout.Dock = DockStyle.Fill;
        layout.BackColor = Color.Transparent;
        layout.ColumnCount = 1;
        layout.RowCount = 6;
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 240f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(layout);

        Label title = Theme.Label("Сервис данных", Theme.TitleFont, Theme.Ink);
        title.Margin = new Padding(0, 0, 0, 6);
        layout.Controls.Add(title, 0, 0);

        Label hint = Theme.Label("Очистка не удаляет файл базы, формы, отчеты и связи. Перед полной очисткой требуется подтверждение.", Theme.BaseFont, Theme.Muted);
        hint.Margin = new Padding(0, 0, 0, 14);
        hint.MaximumSize = new Size(980, 0);
        layout.Controls.Add(hint, 0, 1);

        _cardsFlow = new FlowLayoutPanel();
        _cardsFlow.AutoSize = true;
        _cardsFlow.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        _cardsFlow.FlowDirection = FlowDirection.LeftToRight;
        _cardsFlow.WrapContents = false;
        _cardsFlow.BackColor = Color.Transparent;
        _cardsFlow.Margin = new Padding(0, 0, 0, 14);
        layout.Controls.Add(_cardsFlow, 0, 2);

        _vPensioners  = MkCard("Пенсионеры",    out _sPensioners);
        _vVouchers    = MkCard("Путёвок выдано", out _sVouchers);
        _vSanatoriums = MkCard("Санатории",      out _sSanatoriums);
        _vCost        = MkCard("Стоимость",      out _sCost);
        _vDays        = MkCard("Средний срок",   out _sDays);

        TableLayoutPanel chartsRow = new TableLayoutPanel();
        chartsRow.Dock = DockStyle.Fill;
        chartsRow.ColumnCount = 2;
        chartsRow.RowCount = 1;
        chartsRow.BackColor = Color.Transparent;
        chartsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        chartsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        chartsRow.Margin = new Padding(0, 0, 0, 14);
        layout.Controls.Add(chartsRow, 0, 3);

        _pieChart = new PieChartControl();
        _pieChart.Dock = DockStyle.Fill;
        _pieChart.Margin = new Padding(0, 0, 8, 0);
        chartsRow.Controls.Add(_pieChart, 0, 0);

        _barChart = new BarChartControl();
        _barChart.Dock = DockStyle.Fill;
        _barChart.Margin = new Padding(8, 0, 0, 0);
        chartsRow.Controls.Add(_barChart, 1, 0);

        FlowLayoutPanel actions = new FlowLayoutPanel();
        actions.Dock = DockStyle.Fill;
        actions.AutoSize = true;
        actions.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        actions.FlowDirection = FlowDirection.LeftToRight;
        actions.WrapContents = true;
        actions.Padding = new Padding(0, 8, 0, 8);
        layout.Controls.Add(actions, 0, 4);

        AddAction(actions, "Очистить журнал путевок", ClearJournal);
        AddAction(actions, "Очистить пенсионеров и журнал", ClearPensioners);
        AddAction(actions, "Очистить все демо-данные", ClearAll);
        AddAction(actions, "Обновить статистику", delegate { RefreshStats(); ToastNotifier.Show(this, "Статистика обновлена", true); });

        FlowLayoutPanel footer = new FlowLayoutPanel();
        footer.Dock = DockStyle.Fill;
        footer.AutoSize = true;
        footer.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        footer.FlowDirection = FlowDirection.RightToLeft;
        footer.WrapContents = false;
        footer.Padding = new Padding(0, 12, 0, 0);
        layout.Controls.Add(footer, 0, 5);

        Button close = MakeButton("Назад", false);
        close.Click += delegate { Close(); };
        footer.Controls.Add(close);

        Load += delegate { RefreshStats(); };
    }

    private Label MkCard(string title, out Label sub)
    {
        Panel card = new Panel();
        card.Width = 178; card.Height = 88;
        card.Margin = new Padding(0, 0, 10, 0);
        card.BackColor = Color.Transparent;
        card.Paint += delegate(object s, PaintEventArgs e)
        {
            Color bg = Theme.Dark ? Color.FromArgb(30, 42, 65) : Color.FromArgb(246, 249, 255);
            Color bc = Theme.Dark ? Color.FromArgb(44, 60, 90)  : Color.FromArgb(205, 218, 240);
            using (SolidBrush b = new SolidBrush(bg)) e.Graphics.FillRectangle(b, 0, 0, card.Width, card.Height);
            using (Pen p = new Pen(bc)) e.Graphics.DrawRectangle(p, 0, 0, card.Width - 1, card.Height - 1);
        };
        Label val = new Label();
        val.Text = "..."; val.Font = Theme.H2Font; val.ForeColor = Theme.Blue;
        val.AutoSize = true; val.Location = new Point(12, 8);
        Label lbl = new Label();
        lbl.Text = title; lbl.Font = Theme.BaseFont; lbl.ForeColor = Theme.Muted;
        lbl.AutoSize = false; lbl.Size = new Size(154, 18); lbl.Location = new Point(12, 42);
        Label s2 = new Label();
        s2.Text = ""; s2.Font = new Font("Segoe UI", 8.5f); s2.ForeColor = Theme.Muted;
        s2.AutoSize = false; s2.Size = new Size(154, 28); s2.Location = new Point(12, 60);
        card.Controls.Add(val); card.Controls.Add(lbl); card.Controls.Add(s2);
        _cardsFlow.Controls.Add(card);
        sub = s2; return val;
    }

    private void AddAction(Control parent, string text, EventHandler click)
    {
        Button b = new GlassButton();
        b.Text = text; b.Size = new Size(240, 52);
        b.Margin = new Padding(0, 0, 12, 8);
        Theme.StyleButton(b, text.IndexOf("все", StringComparison.OrdinalIgnoreCase) >= 0);
        b.Click += click;
        parent.Controls.Add(b);
    }

    private void RefreshStats()
    {
        if (_statsLoading || !Db.Exists()) return;
        _statsLoading = true;
        foreach (Label l in new[] { _vPensioners, _vVouchers, _vSanatoriums, _vCost, _vDays })
            if (l != null) { l.Text = "..."; l.ForeColor = Theme.Muted; }
        foreach (Label l in new[] { _sPensioners, _sVouchers, _sSanatoriums, _sCost, _sDays })
            if (l != null) l.Text = "";

        System.ComponentModel.BackgroundWorker w = new System.ComponentModel.BackgroundWorker();
        w.DoWork += delegate(object s, System.ComponentModel.DoWorkEventArgs ev)
        {
            Dictionary<string, object> d = new Dictionary<string, object>();
            d["pen"]    = Db.ScalarInt("SELECT Count(*) FROM tblPensioners");
            d["vch"]    = Db.ScalarInt("SELECT Count(*) FROM tblVoucherIssues");
            d["san"]    = Db.ScalarInt("SELECT Count(*) FROM tblSanatoriums");
            d["sanAct"] = Db.ScalarInt("SELECT Count(*) FROM tblSanatoriums WHERE IsActive = True");
            try { d["penWV"] = Db.ScalarInt("SELECT Count(*) FROM tblPensioners WHERE PensionerID IN (SELECT DISTINCT PensionerID FROM tblVoucherIssues)"); } catch { d["penWV"] = -1; }
            try { d["vchY"]  = Db.ScalarInt("SELECT Count(*) FROM tblVoucherIssues WHERE Year(IssueDate) = Year(Date())"); } catch { d["vchY"] = -1; }

            long cost = -1L;
            try { DataTable ct = Db.Query("SELECT Sum(TotalCost) FROM qryReport_VoucherIssuesDetailed"); cost = (ct.Rows.Count > 0 && ct.Rows[0][0] != DBNull.Value) ? Convert.ToInt64(ct.Rows[0][0]) : 0L; } catch { }
            if (cost < 0) { try { DataTable ct = Db.Query("SELECT Sum((DateDiff('d',vi.StartDate,vi.EndDate)+1)*s.PricePerDay) FROM tblVoucherIssues AS vi INNER JOIN tblSanatoriums AS s ON vi.SanatoriumID=s.SanatoriumID WHERE vi.StartDate IS NOT NULL AND vi.EndDate IS NOT NULL"); cost = (ct.Rows.Count > 0 && ct.Rows[0][0] != DBNull.Value) ? Convert.ToInt64(ct.Rows[0][0]) : 0L; } catch { cost = 0L; } }
            d["cost"] = cost;

            double avgD = -1.0;
            try { DataTable dt = Db.Query("SELECT Avg(CDbl(DaysCount)) FROM qryReport_VoucherIssuesDetailed WHERE DaysCount > 0"); avgD = (dt.Rows.Count > 0 && dt.Rows[0][0] != DBNull.Value) ? Math.Round(Convert.ToDouble(dt.Rows[0][0]), 1) : 0.0; } catch { }
            if (avgD < 0) { try { DataTable dt = Db.Query("SELECT Avg(DateDiff('d',StartDate,EndDate)+1) FROM tblVoucherIssues WHERE StartDate IS NOT NULL AND EndDate IS NOT NULL"); avgD = (dt.Rows.Count > 0 && dt.Rows[0][0] != DBNull.Value) ? Math.Round(Convert.ToDouble(dt.Rows[0][0]), 1) : 0.0; } catch { avgD = 0.0; } }
            d["days"] = avgD;

            try { d["statuses"] = Db.Query("SELECT vs.StatusName, Count(*) AS cnt FROM tblVoucherIssues AS vi INNER JOIN tblVoucherStatuses AS vs ON vi.StatusID=vs.StatusID GROUP BY vs.StatusName ORDER BY Count(*) DESC"); } catch { d["statuses"] = null; }
            ev.Result = d;
        };

        w.RunWorkerCompleted += delegate(object s, System.ComponentModel.RunWorkerCompletedEventArgs ev)
        {
            _statsLoading = false;
            if (IsDisposed) return;
            if (ev.Error != null) { ErrorLogger.LogException(ev.Error); return; }

            Dictionary<string, object> d = (Dictionary<string, object>)ev.Result;
            int pen    = (int)d["pen"];   int vch  = (int)d["vch"];
            int san    = (int)d["san"];   int sanA = (int)d["sanAct"];
            int penWV  = (int)d["penWV"]; int vchY = (int)d["vchY"];
            long cost  = (long)d["cost"]; double days = (double)d["days"];
            DataTable st = d["statuses"] as DataTable;

            _vPensioners.Text = pen.ToString();  _vPensioners.ForeColor = Theme.Blue;
            _sPensioners.Text = penWV >= 0 ? "с путёвками: " + penWV : "";

            _vVouchers.Text   = vch.ToString();  _vVouchers.ForeColor   = Theme.Blue;
            _sVouchers.Text   = vchY >= 0 ? "в этом году: " + vchY : "";

            _vSanatoriums.Text = san.ToString(); _vSanatoriums.ForeColor = Theme.Blue;
            _sSanatoriums.Text = "активных: " + sanA;

            if (cost > 0) { _vCost.Text = FmtRub(cost); _vCost.ForeColor = Theme.Green; _sCost.Text = vch > 0 ? "ср. " + FmtRub(cost / vch) : ""; }
            else          { _vCost.Text = "—";           _vCost.ForeColor = Theme.Muted; _sCost.Text = "нет данных"; }

            if (days > 0) { _vDays.Text = days.ToString("0.#") + " дн."; _vDays.ForeColor = Theme.Blue; _sDays.Text = "средний срок"; }
            else          { _vDays.Text = "—"; _vDays.ForeColor = Theme.Muted; _sDays.Text = ""; }

            if (st != null && st.Rows.Count > 0)
            {
                List<PieSlice> slices = new List<PieSlice>();
                foreach (DataRow row in st.Rows)
                    slices.Add(new PieSlice { Label = Convert.ToString(row[0]), Value = Convert.ToInt32(row[1]) });
                _pieChart.SetData("Статусы путёвок", slices.ToArray());
            }
            else _pieChart.SetData("Статусы путёвок", null);

            List<BarEntry> bars = new List<BarEntry>();
            if (pen > 0) bars.Add(new BarEntry { Label = "С путёвками",    Value = penWV >= 0 ? penWV : 0, Total = pen, Color = Color.FromArgb(59, 130, 246) });
            if (san > 0) bars.Add(new BarEntry { Label = "Активных сан.",  Value = sanA,                   Total = san, Color = Color.FromArgb(34, 197, 94)  });
            if (vch > 0 && vchY >= 0) bars.Add(new BarEntry { Label = "Путёвок (год)", Value = vchY, Total = vch, Color = Color.FromArgb(168, 85, 247) });
            _barChart.SetData("Охват и активность", bars.ToArray());
        };
        w.RunWorkerAsync();
    }

    private static string FmtRub(long v)
    {
        if (v >= 1000000) return (v / 1000000.0).ToString("0.#") + " млн ₽";
        if (v >= 1000)    return (v / 1000.0).ToString("0") + " тыс ₽";
        return v + " ₽";
    }

    private void ClearJournal(object sender, EventArgs e)
    {
        Guard(delegate
        {
            if (MessageBox.Show("Удалить все записи журнала выдачи путевок?", "Очистка", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Db.Execute("DELETE FROM tblVoucherIssues");
            RefreshStats();
            ToastNotifier.Show(this, "Журнал очищен. База обновлена", true);
        });
    }

    private void ClearPensioners(object sender, EventArgs e)
    {
        Guard(delegate
        {
            if (MessageBox.Show("Удалить пенсионеров и все связанные путевки?", "Очистка", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Db.Execute("DELETE FROM tblVoucherIssues");
            Db.Execute("DELETE FROM tblPensioners");
            RefreshStats();
            ToastNotifier.Show(this, "Пенсионеры и журнал очищены. База обновлена", true);
        });
    }

    private void ClearAll(object sender, EventArgs e)
    {
        Guard(delegate
        {
            if (MessageBox.Show("Будут удалены ВСЕ записи кроме справочников. Продолжить?", "Полная очистка", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Db.Execute("DELETE FROM tblVoucherIssues");
            Db.Execute("DELETE FROM tblPensioners");
            Db.Execute("DELETE FROM tblSanatoriums");
            RefreshStats();
            ToastNotifier.Show(this, "Все демонстрационные данные очищены", true);
        });
    }
}


    internal static class ReportTables
    {
        public static bool ExportAllTablesPdf(IWin32Window owner, DbContext db)
        {
            List<ProfessionalPdfExporter.Section> sections = new List<ProfessionalPdfExporter.Section>();
            AddSection(sections, "Пенсионеры", delegate { return Pensioners(db); });
            AddSection(sections, "Журнал путевок", delegate { return Vouchers(db); });
            AddSection(sections, "Санатории", delegate { return Sanatoriums(db); });
            AddSection(sections, "Итоги по санаториям", delegate { return RegionTotals(db); });
            return ProfessionalPdfExporter.ExportSectionsWithDialog(owner, "Сводный отчет по всем таблицам", sections, "all_healthcare_tables.pdf");
        }

        private static void AddSection(List<ProfessionalPdfExporter.Section> sections, string title, Func<DataTable> loader)
        {
            try
            {
                sections.Add(ProfessionalPdfExporter.Section.FromDataTable(title, loader()));
            }
            catch (Exception ex)
            {
                DataTable error = new DataTable();
                error.Columns.Add("Раздел");
                error.Columns.Add("Статус");
                error.Rows.Add(title, "Не удалось загрузить: " + ex.Message);
                sections.Add(ProfessionalPdfExporter.Section.FromDataTable(title, error));
            }
        }

        public static DataTable Pensioners(DbContext db)
        {
            return db.Query("SELECT p.LastName AS [Фамилия], p.FirstName AS [Имя], p.MiddleName AS [Отчество], r.RegionName AS [Регион], c.CategoryName AS [Категория], p.BirthDate AS [Дата рождения], p.PensionCertificateNo AS [Удостоверение], p.Phone AS [Телефон] FROM (tblPensioners AS p INNER JOIN tblRegions AS r ON p.RegionID = r.RegionID) INNER JOIN tblPensionerCategories AS c ON p.CategoryID = c.CategoryID ORDER BY p.LastName, p.FirstName");
        }

        public static DataTable Vouchers(DbContext db)
        {
            return db.Query("SELECT VoucherNo AS [Путевка], IssueDate AS [Регистрация], PensionerFullName AS [Пенсионер], PensionerRegion AS [Регион], SanatoriumName AS [Санаторий], StartDate AS [Заезд], EndDate AS [Выезд], DaysCount AS [Дней], TotalCost AS [Стоимость], StatusName AS [Статус] FROM qryReport_VoucherIssuesDetailed ORDER BY IssueDate, VoucherNo");
        }

        public static DataTable Sanatoriums(DbContext db)
        {
            return db.Query("SELECT s.SanatoriumName AS [Санаторий], r.RegionName AS [Регион], mp.ProfileName AS [Профиль], s.CapacityBeds AS [Коек], s.PricePerDay AS [Цена дня], s.Phone AS [Телефон], s.Address AS [Адрес], IIf(s.IsActive, 'Да', 'Нет') AS [Работает] FROM (tblSanatoriums AS s INNER JOIN tblRegions AS r ON s.RegionID = r.RegionID) INNER JOIN tblMedicalProfiles AS mp ON s.ProfileID = mp.ProfileID ORDER BY s.SanatoriumName");
        }

        public static DataTable RegionTotals(DbContext db)
        {
            return db.Query("SELECT SanatoriumRegion AS [Регион санатория], SanatoriumName AS [Санаторий], VoucherCount AS [Путевок], TotalDays AS [Дней], TotalPlannedCost AS [Плановая стоимость] FROM qry04_Totals_BySanatorium ORDER BY SanatoriumRegion, SanatoriumName");
        }
    }

    internal enum PrintTarget
    {
        AccessJournal,
        AccessRegionTotals,
        VisibleGrid,
        SelectedGridRow,
        AllTables
    }

    
internal sealed class PrintOptionsForm : Form
{
    private readonly ComboBox target;
    private readonly RadioButton preview;
    private readonly RadioButton direct;
    private readonly RadioButton pdf;

    public PrintOptionsForm(bool preferPrint)
    {
        Text = "Параметры печати";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ClientSize = new Size(560, 320);
        MinimumSize = new Size(560, 320);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = Theme.BaseFont;
        BackColor = Theme.Dark ? Color.FromArgb(18, 24, 38) : Color.FromArgb(245, 248, 253);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.BackColor = Color.Transparent;
        root.ColumnCount = 1;
        root.RowCount = 5;
        root.Padding = new Padding(24);
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        Label title = Theme.Label("Что печатать?", Theme.H2Font, Theme.Ink);
        title.Margin = new Padding(0, 0, 0, 10);
        root.Controls.Add(title, 0, 0);

        TableLayoutPanel targetRow = new TableLayoutPanel();
        targetRow.Dock = DockStyle.Fill;
        targetRow.ColumnCount = 2;
        targetRow.RowCount = 1;
        targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));
        targetRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        targetRow.Margin = new Padding(0, 0, 0, 10);
        root.Controls.Add(targetRow, 0, 1);

        Label label = Theme.Label("Раздел", Theme.BaseFont, Theme.Muted);
        label.Anchor = AnchorStyles.Left;
        targetRow.Controls.Add(label, 0, 0);

        target = new ComboBox();
        target.DropDownStyle = ComboBoxStyle.DropDownList;
        target.Dock = DockStyle.Fill;
        target.Items.Add(new PrintChoice(PrintTarget.AccessJournal, "Отчет Access: журнал путевок"));
        target.Items.Add(new PrintChoice(PrintTarget.AccessRegionTotals, "Отчет Access: итоги по регионам"));
        target.Items.Add(new PrintChoice(PrintTarget.VisibleGrid, "Текущая таблица: все видимые строки"));
        target.Items.Add(new PrintChoice(PrintTarget.SelectedGridRow, "Текущая таблица: выбранная строка"));
        target.Items.Add(new PrintChoice(PrintTarget.AllTables, "PDF: все основные таблицы вместе"));
        target.SelectedIndex = 0;
        targetRow.Controls.Add(target, 1, 0);

        FlowLayoutPanel modes = new FlowLayoutPanel();
        modes.Dock = DockStyle.Fill;
        modes.FlowDirection = FlowDirection.LeftToRight;
        modes.WrapContents = true;
        modes.AutoSize = true;
        modes.Padding = new Padding(0, 2, 0, 10);
        root.Controls.Add(modes, 0, 2);

        preview = new RadioButton();
        preview.Text = "Предпросмотр";
        preview.AutoSize = true;
        preview.Checked = !preferPrint;
        preview.BackColor = Color.Transparent;
        preview.ForeColor = Theme.Ink;
        modes.Controls.Add(preview);

        direct = new RadioButton();
        direct.Text = "Печать сразу";
        direct.AutoSize = true;
        direct.Checked = preferPrint;
        direct.BackColor = Color.Transparent;
        direct.ForeColor = Theme.Ink;
        modes.Controls.Add(direct);

        pdf = new RadioButton();
        pdf.Text = "Сохранить красивый PDF";
        pdf.AutoSize = true;
        pdf.BackColor = Color.Transparent;
        pdf.ForeColor = Theme.Ink;
        modes.Controls.Add(pdf);

        Label hint = Theme.Label("PDF собирается с аккуратной шапкой, таблицей, страницами и правильным порядком строк.", Theme.BaseFont, Theme.Muted);
        hint.Margin = new Padding(0, 8, 0, 10);
        hint.MaximumSize = new Size(510, 0);
        root.Controls.Add(hint, 0, 3);

        FlowLayoutPanel buttons = new FlowLayoutPanel();
        buttons.Dock = DockStyle.Fill;
        buttons.FlowDirection = FlowDirection.RightToLeft;
        buttons.WrapContents = false;
        buttons.AutoSize = true;
        buttons.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        buttons.Padding = new Padding(0, 8, 0, 0);
        root.Controls.Add(buttons, 0, 4);

        Button cancel = new GlassButton();
        cancel.Text = "Отмена";
        cancel.SetBounds(0, 0, 90, 38);
        cancel.DialogResult = DialogResult.Cancel;
        Theme.StyleButton(cancel, false);
        buttons.Controls.Add(cancel);

        Button ok = new GlassButton();
        ok.Text = "Продолжить";
        ok.SetBounds(0, 0, 120, 38);
        ok.DialogResult = DialogResult.OK;
        Theme.StyleButton(ok, true);
        buttons.Controls.Add(ok);

        AcceptButton = ok;
        CancelButton = cancel;
    }

    public PrintTarget Target
    {
        get
        {
            PrintChoice choice = target.SelectedItem as PrintChoice;
            return choice == null ? PrintTarget.AccessJournal : choice.Target;
        }
    }

    public bool DirectPrint
    {
        get { return direct.Checked; }
    }

    public bool ExportPdf
    {
        get { return pdf.Checked; }
    }

    private sealed class PrintChoice
    {
        public PrintChoice(PrintTarget target, string text)
        {
            Target = target;
            Text = text;
        }

        public PrintTarget Target { get; private set; }
        public string Text { get; private set; }

        public override string ToString()
        {
            return Text;
        }
    }
}


    internal sealed class GridPrinter
    {
        private readonly string title;
        private readonly List<string> headers;
        private readonly List<string[]> rows;
        private int rowIndex;
        private int pageNumber;

        private GridPrinter(string title, DataGridView grid, bool selectedOnly)
        {
            this.title = title;
            headers = new List<string>();
            rows = new List<string[]>();

            List<DataGridViewColumn> columns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (!column.Visible) continue;
                columns.Add(column);
                headers.Add(column.HeaderText);
            }

            foreach (DataGridViewRow gridRow in grid.Rows)
            {
                if (gridRow.IsNewRow) continue;
                if (selectedOnly && !gridRow.Selected) continue;
                string[] values = new string[columns.Count];
                for (int i = 0; i < columns.Count; i++)
                {
                    object value = gridRow.Cells[columns[i].Index].Value;
                    values[i] = value == null || value == DBNull.Value ? "" : Convert.ToString(value);
                }
                rows.Add(values);
            }
        }

        public static void PrintGrid(string title, DataGridView grid, bool selectedOnly, bool directPrint)
        {
            GridPrinter job = new GridPrinter(title, grid, selectedOnly);
            if (job.rows.Count == 0)
            {
                MessageBox.Show("Нет строк для печати.", "Печать", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (PrintDocument document = new PrintDocument())
            {
                document.DocumentName = title;
                document.DefaultPageSettings.Landscape = true;
                document.PrintPage += job.PrintPage;

                if (directPrint)
                {
                    using (PrintDialog dialog = new PrintDialog())
                    {
                        dialog.Document = document;
                        dialog.UseEXDialog = true;
                        if (dialog.ShowDialog() == DialogResult.OK)
                        {
                            document.Print();
                        }
                    }
                }
                else
                {
                    using (PrintPreviewDialog preview = new PrintPreviewDialog())
                    {
                        preview.Document = document;
                        preview.Width = 1100;
                        preview.Height = 760;
                        preview.ShowDialog();
                    }
                }
            }
        }

        private void PrintPage(object sender, PrintPageEventArgs e)
        {
            pageNumber++;
            Rectangle area = e.MarginBounds;
            using (Font titleFont = new Font("Segoe UI Semibold", 14f, FontStyle.Bold))
            using (Font headerFont = new Font("Segoe UI Semibold", 8f, FontStyle.Bold))
            using (Font rowFont = new Font("Segoe UI", 8f))
            using (Pen line = new Pen(Color.FromArgb(190, 202, 220)))
            using (SolidBrush headerBack = new SolidBrush(Color.FromArgb(235, 242, 252)))
            using (SolidBrush text = new SolidBrush(Color.FromArgb(24, 31, 44)))
            using (StringFormat cellFormat = new StringFormat())
            {
                cellFormat.Trimming = StringTrimming.EllipsisCharacter;
                cellFormat.FormatFlags = StringFormatFlags.NoWrap;

                e.Graphics.DrawString(title, titleFont, text, area.Left, area.Top - 38);
                e.Graphics.DrawString("Страница " + pageNumber, rowFont, text, area.Right - 90, area.Top - 30);

                if (headers.Count == 0)
                {
                    e.HasMorePages = false;
                    return;
                }

                int columnWidth = Math.Max(70, area.Width / headers.Count);
                int headerHeight = 28;
                int rowHeight = 26;
                int y = area.Top;

                for (int i = 0; i < headers.Count; i++)
                {
                    Rectangle cell = new Rectangle(area.Left + i * columnWidth, y, columnWidth, headerHeight);
                    e.Graphics.FillRectangle(headerBack, cell);
                    e.Graphics.DrawRectangle(line, cell);
                    e.Graphics.DrawString(headers[i], headerFont, text, new RectangleF(cell.X + 4, cell.Y + 6, cell.Width - 8, cell.Height - 8), cellFormat);
                }
                y += headerHeight;

                while (rowIndex < rows.Count)
                {
                    if (y + rowHeight > area.Bottom)
                    {
                        e.HasMorePages = true;
                        return;
                    }
                    string[] row = rows[rowIndex];
                    for (int i = 0; i < headers.Count; i++)
                    {
                        Rectangle cell = new Rectangle(area.Left + i * columnWidth, y, columnWidth, rowHeight);
                        e.Graphics.DrawRectangle(line, cell);
                        string value = i < row.Length ? row[i] : "";
                        e.Graphics.DrawString(value, rowFont, text, new RectangleF(cell.X + 4, cell.Y + 6, cell.Width - 8, cell.Height - 8), cellFormat);
                    }
                    y += rowHeight;
                    rowIndex++;
                }
                e.HasMorePages = false;
            }
        }
    }

    public static class ProfessionalPdfExporter
    {
        private const int PagePointWidth = 842;
        private const int PagePointHeight = 595;
        private const int Scale = 3;
        private const int PagePixelWidth = PagePointWidth * Scale;
        private const int PagePixelHeight = PagePointHeight * Scale;

        public sealed class Section
        {
            private Section(string title, List<string> headers, List<string[]> rows)
            {
                Title = title;
                Headers = headers;
                Rows = rows;
            }

            public string Title { get; private set; }
            public List<string> Headers { get; private set; }
            public List<string[]> Rows { get; private set; }

            public static Section FromDataTable(string title, DataTable table)
            {
                List<string> headers = new List<string>();
                List<string[]> rows = new List<string[]>();
                if (table != null)
                {
                    foreach (DataColumn column in table.Columns)
                    {
                        headers.Add(column.ColumnName);
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        string[] values = new string[table.Columns.Count];
                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            values[i] = FormatValue(row[i]);
                        }
                        rows.Add(values);
                    }
                }

                if (headers.Count == 0)
                {
                    headers.Add("Статус");
                }
                if (rows.Count == 0)
                {
                    rows.Add(new string[] { "Нет данных" });
                }
                return new Section(title, headers, rows);
            }
        }

        public static bool ExportSectionsWithDialog(IWin32Window owner, string title, List<Section> sections, string defaultFileName)
        {
            if (sections == null || sections.Count == 0)
            {
                MessageBox.Show("Нет данных для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF файл (*.pdf)|*.pdf";
                dialog.FileName = defaultFileName;
                dialog.Title = "Сохранить PDF";
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;

                List<byte[]> pages = new List<byte[]>();
                foreach (Section section in sections)
                {
                    if (section == null) continue;
                    pages.AddRange(RenderPages(title, section.Title, section.Headers, section.Rows));
                }

                if (pages.Count == 0)
                {
                    MessageBox.Show("Нет страниц для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                WriteImagePdf(dialog.FileName, pages);
                return true;
            }
        }

        public static bool ExportDataTableWithDialog(IWin32Window owner, string title, string subtitle, DataTable table, string defaultFileName)
        {
            if (table == null || table.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF файл (*.pdf)|*.pdf";
                dialog.FileName = defaultFileName;
                dialog.Title = "Сохранить PDF";
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;

                List<string> headers = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    headers.Add(column.ColumnName);
                }

                List<string[]> rows = new List<string[]>();
                foreach (DataRow row in table.Rows)
                {
                    string[] values = new string[table.Columns.Count];
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        values[i] = FormatValue(row[i]);
                    }
                    rows.Add(values);
                }

                Export(dialog.FileName, title, subtitle, headers, rows);
                return true;
            }
        }

        public static void ExportDataTableToFile(string path, string title, string subtitle, DataTable table)
        {
            List<string> headers = new List<string>();
            foreach (DataColumn column in table.Columns)
            {
                headers.Add(column.ColumnName);
            }

            List<string[]> rows = new List<string[]>();
            foreach (DataRow row in table.Rows)
            {
                string[] values = new string[table.Columns.Count];
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    values[i] = FormatValue(row[i]);
                }
                rows.Add(values);
            }

            Export(path, title, subtitle, headers, rows);
        }

        public static bool ExportGridWithDialog(IWin32Window owner, string title, DataGridView grid, bool selectedOnly, string defaultFileName)
        {
            List<DataGridViewColumn> columns = new List<DataGridViewColumn>();
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (!column.Visible) continue;
                columns.Add(column);
            }
            if (columns.Count == 0)
            {
                MessageBox.Show("Нет колонок для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            List<string> headers = new List<string>();
            foreach (DataGridViewColumn column in columns)
            {
                headers.Add(column.HeaderText);
            }

            List<string[]> rows = new List<string[]>();
            foreach (DataGridViewRow gridRow in grid.Rows)
            {
                if (gridRow.IsNewRow) continue;
                if (selectedOnly && !gridRow.Selected) continue;
                string[] values = new string[columns.Count];
                for (int i = 0; i < columns.Count; i++)
                {
                    values[i] = FormatValue(gridRow.Cells[columns[i].Index].Value);
                }
                rows.Add(values);
            }

            if (rows.Count == 0)
            {
                MessageBox.Show("Нет строк для PDF.", "PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "PDF файл (*.pdf)|*.pdf";
                dialog.FileName = defaultFileName;
                dialog.Title = "Сохранить PDF";
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;
                Export(dialog.FileName, title, "Система здравоохранения", headers, rows);
                return true;
            }
        }

        private static string FormatValue(object value)
        {
            if (value == null || value == DBNull.Value) return "";
            if (value is DateTime) return ((DateTime)value).ToString("dd.MM.yyyy");
            if (value is decimal || value is double || value is float)
            {
                decimal number;
                try
                {
                    number = Convert.ToDecimal(value);
                    return number.ToString("#,##0.##");
                }
                catch { }
            }
            return Convert.ToString(value);
        }

        private static void Export(string path, string title, string subtitle, List<string> headers, List<string[]> rows)
        {
            List<byte[]> pages = RenderPages(title, subtitle, headers, rows);
            WriteImagePdf(path, pages);
        }

        private static List<byte[]> RenderPages(string title, string subtitle, List<string> headers, List<string[]> rows)
        {
            List<byte[]> images = new List<byte[]>();
            int margin = 110;
            int yStart = 255;
            int tableHeaderHeight = 58;
            int rowHeight = 48;
            int footerHeight = 82;
            int rowsPerPage = Math.Max(1, (PagePixelHeight - yStart - tableHeaderHeight - footerHeight - margin) / rowHeight);
            int pageCount = Math.Max(1, (int)Math.Ceiling(rows.Count / (double)rowsPerPage));
            int rowIndex = 0;

            for (int page = 1; page <= pageCount; page++)
            {
                using (Bitmap bitmap = new Bitmap(PagePixelWidth, PagePixelHeight))
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                    g.Clear(Color.White);

                    DrawPdfHeader(g, title, subtitle, page, pageCount, rows.Count);

                    int tableWidth = PagePixelWidth - margin * 2;
                    int[] widths = CalculateColumnWidths(headers.Count, tableWidth);
                    int y = yStart;
                    DrawTableHeader(g, headers, widths, margin, y, tableHeaderHeight);
                    y += tableHeaderHeight;

                    int rowsOnPage = 0;
                    while (rowIndex < rows.Count && rowsOnPage < rowsPerPage)
                    {
                        DrawTableRow(g, rows[rowIndex], widths, margin, y, rowHeight, rowIndex % 2 == 0);
                        y += rowHeight;
                        rowIndex++;
                        rowsOnPage++;
                    }

                    DrawPdfFooter(g, page, pageCount);
                    images.Add(EncodeJpeg(bitmap, 92L));
                }
            }

            return images;
        }

        private static void DrawPdfHeader(Graphics g, string title, string subtitle, int page, int pageCount, int rowCount)
        {
            using (LinearGradientBrush band = new LinearGradientBrush(new Rectangle(0, 0, PagePixelWidth, 190), Color.FromArgb(245, 249, 255), Color.White, 90f))
            using (SolidBrush ink = new SolidBrush(Color.FromArgb(20, 28, 42)))
            using (SolidBrush muted = new SolidBrush(Color.FromArgb(92, 106, 128)))
            using (SolidBrush accent = new SolidBrush(Color.FromArgb(0, 122, 255)))
            using (Font titleFont = new Font("Segoe UI Semibold", 36f, FontStyle.Bold))
            using (Font metaFont = new Font("Segoe UI", 20f, FontStyle.Regular))
            {
                g.FillRectangle(band, 0, 0, PagePixelWidth, 190);
                g.FillRectangle(accent, 110, 72, 12, 66);
                g.DrawString(title, titleFont, ink, 145, 58);
                string meta = subtitle + "  •  " + DateTime.Now.ToString("dd.MM.yyyy HH:mm") + "  •  записей: " + rowCount;
                g.DrawString(meta, metaFont, muted, 148, 116);
                g.DrawString("стр. " + page + " / " + pageCount, metaFont, muted, PagePixelWidth - 310, 86);
            }
        }

        private static void DrawTableHeader(Graphics g, List<string> headers, int[] widths, int x, int y, int height)
        {
            using (SolidBrush headerBack = new SolidBrush(Color.FromArgb(28, 43, 68)))
            using (SolidBrush text = new SolidBrush(Color.White))
            using (Pen line = new Pen(Color.FromArgb(210, 222, 238)))
            using (Font font = new Font("Segoe UI Semibold", 17f, FontStyle.Bold))
            using (StringFormat format = CellFormat())
            {
                int currentX = x;
                for (int i = 0; i < headers.Count; i++)
                {
                    Rectangle cell = new Rectangle(currentX, y, widths[i], height);
                    g.FillRectangle(headerBack, cell);
                    g.DrawRectangle(line, cell);
                    g.DrawString(headers[i], font, text, new RectangleF(cell.X + 10, cell.Y + 15, cell.Width - 20, cell.Height - 18), format);
                    currentX += widths[i];
                }
            }
        }

        private static void DrawTableRow(Graphics g, string[] row, int[] widths, int x, int y, int height, bool even)
        {
            using (SolidBrush back = new SolidBrush(even ? Color.FromArgb(250, 252, 255) : Color.White))
            using (SolidBrush text = new SolidBrush(Color.FromArgb(25, 34, 50)))
            using (Pen line = new Pen(Color.FromArgb(224, 231, 242)))
            using (Font font = new Font("Segoe UI", 15f, FontStyle.Regular))
            using (StringFormat format = CellFormat())
            {
                int currentX = x;
                for (int i = 0; i < widths.Length; i++)
                {
                    Rectangle cell = new Rectangle(currentX, y, widths[i], height);
                    g.FillRectangle(back, cell);
                    g.DrawRectangle(line, cell);
                    string value = i < row.Length ? row[i] : "";
                    g.DrawString(value, font, text, new RectangleF(cell.X + 10, cell.Y + 13, cell.Width - 20, cell.Height - 15), format);
                    currentX += widths[i];
                }
            }
        }

        private static void DrawPdfFooter(Graphics g, int page, int pageCount)
        {
            using (Pen line = new Pen(Color.FromArgb(225, 232, 243), 2f))
            using (SolidBrush muted = new SolidBrush(Color.FromArgb(110, 122, 145)))
            using (Font font = new Font("Segoe UI", 16f))
            {
                int y = PagePixelHeight - 92;
                g.DrawLine(line, 110, y, PagePixelWidth - 110, y);
                g.DrawString("Система здравоохранения • санаторные путевки", font, muted, 110, y + 22);
                g.DrawString("Страница " + page + " из " + pageCount, font, muted, PagePixelWidth - 360, y + 22);
            }
        }

        private static int[] CalculateColumnWidths(int count, int tableWidth)
        {
            int[] widths = new int[count];
            if (count == 0) return widths;
            int baseWidth = tableWidth / count;
            for (int i = 0; i < count; i++) widths[i] = baseWidth;
            widths[count - 1] += tableWidth - (baseWidth * count);
            return widths;
        }

        private static StringFormat CellFormat()
        {
            StringFormat format = new StringFormat();
            format.Trimming = StringTrimming.EllipsisCharacter;
            format.FormatFlags = StringFormatFlags.NoWrap;
            format.Alignment = StringAlignment.Near;
            format.LineAlignment = StringAlignment.Near;
            return format;
        }

        private static byte[] EncodeJpeg(Bitmap bitmap, long quality)
        {
            ImageCodecInfo codec = null;
            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
            foreach (ImageCodecInfo info in encoders)
            {
                if (info.FormatID == ImageFormat.Jpeg.Guid)
                {
                    codec = info;
                    break;
                }
            }

            using (MemoryStream stream = new MemoryStream())
            {
                if (codec != null)
                {
                    using (EncoderParameters parameters = new EncoderParameters(1))
                    {
                        parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
                        bitmap.Save(stream, codec, parameters);
                    }
                }
                else
                {
                    bitmap.Save(stream, ImageFormat.Jpeg);
                }
                return stream.ToArray();
            }
        }

        private static void WriteImagePdf(string path, List<byte[]> pageImages)
        {
            using (MemoryStream pdf = new MemoryStream())
            {
                WriteAscii(pdf, "%PDF-1.4\n");
                List<long> offsets = new List<long>();
                int objectCount = 2 + pageImages.Count * 3;

                for (int obj = 1; obj <= objectCount; obj++)
                {
                    offsets.Add(pdf.Position);
                    WriteAscii(pdf, obj + " 0 obj\n");
                    if (obj == 1)
                    {
                        WriteAscii(pdf, "<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
                    }
                    else if (obj == 2)
                    {
                        StringBuilder kids = new StringBuilder();
                        for (int i = 0; i < pageImages.Count; i++)
                        {
                            kids.Append(3 + i * 3).Append(" 0 R ");
                        }
                        WriteAscii(pdf, "<< /Type /Pages /Kids [" + kids + "] /Count " + pageImages.Count + " >>\nendobj\n");
                    }
                    else
                    {
                        int index = (obj - 3) / 3;
                        int kind = (obj - 3) % 3;
                        int pageObj = 3 + index * 3;
                        int contentObj = pageObj + 1;
                        int imageObj = pageObj + 2;
                        if (kind == 0)
                        {
                            WriteAscii(pdf, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + PagePointWidth + " " + PagePointHeight + "] /Resources << /XObject << /Im" + index + " " + imageObj + " 0 R >> >> /Contents " + contentObj + " 0 R >>\nendobj\n");
                        }
                        else if (kind == 1)
                        {
                            string content = "q\n" + PagePointWidth + " 0 0 " + PagePointHeight + " 0 0 cm\n/Im" + index + " Do\nQ\n";
                            byte[] bytes = Encoding.ASCII.GetBytes(content);
                            WriteAscii(pdf, "<< /Length " + bytes.Length + " >>\nstream\n");
                            pdf.Write(bytes, 0, bytes.Length);
                            WriteAscii(pdf, "endstream\nendobj\n");
                        }
                        else
                        {
                            byte[] image = pageImages[index];
                            WriteAscii(pdf, "<< /Type /XObject /Subtype /Image /Width " + PagePixelWidth + " /Height " + PagePixelHeight + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + image.Length + " >>\nstream\n");
                            pdf.Write(image, 0, image.Length);
                            WriteAscii(pdf, "\nendstream\nendobj\n");
                        }
                    }
                }

                long xref = pdf.Position;
                WriteAscii(pdf, "xref\n0 " + (objectCount + 1) + "\n");
                WriteAscii(pdf, "0000000000 65535 f \n");
                foreach (long offset in offsets)
                {
                    WriteAscii(pdf, offset.ToString("0000000000") + " 00000 n \n");
                }
                WriteAscii(pdf, "trailer\n<< /Size " + (objectCount + 1) + " /Root 1 0 R >>\nstartxref\n" + xref + "\n%%EOF");
                File.WriteAllBytes(path, pdf.ToArray());
            }
        }

        private static void WriteAscii(Stream stream, string text)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            stream.Write(bytes, 0, bytes.Length);
        }
    }

    internal sealed class SplashForm : Form
    {
        private readonly System.Windows.Forms.Timer _dotTimer;
        private int _dotCount;
        private string _statusText = "Инициализация...";

        public SplashForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(480, 220);
            BackColor = Color.FromArgb(16, 22, 36);
            DoubleBuffered = true;
            ShowInTaskbar = false;
            TopMost = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint, true);

            _dotTimer = new System.Windows.Forms.Timer();
            _dotTimer.Interval = 420;
            _dotTimer.Tick += delegate { _dotCount = (_dotCount + 1) % 4; Invalidate(); };
        }

        public void SetStatus(string text)
        {
            _statusText = text;
            if (IsHandleCreated) BeginInvoke((Action)(() => Invalidate()));
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _dotTimer.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _dotTimer.Stop();
            _dotTimer.Dispose();
            base.OnFormClosed(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Border
            Rectangle border = new Rectangle(1, 1, Width - 3, Height - 3);
            using (GraphicsPath bp = SplashRound(border, 16))
            using (Pen borderPen = new Pen(Color.FromArgb(55, 85, 130), 1.5f))
                g.DrawPath(borderPen, bp);

            // Title
            using (Font titleFont = new Font("Segoe UI", 19f, FontStyle.Bold, GraphicsUnit.Point))
            using (SolidBrush titleBrush = new SolidBrush(Color.FromArgb(224, 234, 250)))
            {
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                g.DrawString("Система здравоохранения", titleFont, titleBrush, new RectangleF(0, 36, Width, 40), sf);
                sf.Dispose();
            }

            // Subtitle
            using (Font subFont = new Font("Segoe UI", 10f, GraphicsUnit.Point))
            using (SolidBrush subBrush = new SolidBrush(Color.FromArgb(100, 135, 175)))
            {
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                g.DrawString("Санаторные путевки · Пенсионеры · Отчёты", subFont, subBrush, new RectangleF(0, 84, Width, 22), sf);
                sf.Dispose();
            }

            // Progress bar (animated dots)
            int barW = 180;
            int barH = 4;
            int barX = (Width - barW) / 2;
            int barY = 128;
            using (SolidBrush trackBrush = new SolidBrush(Color.FromArgb(35, 52, 78)))
                g.FillRectangle(trackBrush, barX, barY, barW, barH);

            int fillW = (int)(barW * (_dotCount / 3.0));
            if (_dotCount > 0)
            {
                using (LinearGradientBrush fillBrush = new LinearGradientBrush(
                    new Rectangle(barX, barY, Math.Max(1, fillW), barH),
                    Color.FromArgb(74, 144, 255), Color.FromArgb(100, 200, 255), LinearGradientMode.Horizontal))
                    g.FillRectangle(fillBrush, barX, barY, fillW, barH);
            }

            // Status text
            using (Font statusFont = new Font("Segoe UI", 9f, GraphicsUnit.Point))
            using (SolidBrush statusBrush = new SolidBrush(Color.FromArgb(80, 110, 150)))
            {
                StringFormat sf = new StringFormat();
                sf.Alignment = StringAlignment.Center;
                g.DrawString(_statusText, statusFont, statusBrush, new RectangleF(0, 150, Width, 20), sf);
                sf.Dispose();
            }
        }

        private static GraphicsPath SplashRound(Rectangle r, int radius)
        {
            int d = radius * 2;
            GraphicsPath p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
    }

    public static class PluginBootstrap
    {
        public static int RunSelfTest(string dbPath)
        {
            DbContext testDb = new DbContext(dbPath);
            if (!testDb.Exists())
            {
                return 2;
            }

            testDb.ScalarInt("SELECT Count(*) FROM tblVoucherIssues");
            testDb.ScalarInt("SELECT Count(*) FROM tblPensioners");
            return 0;
        }

        public static int RunPdfSelfTest(string dbPath)
        {
            DbContext testDb = new DbContext(dbPath);
            if (!testDb.Exists())
            {
                return 2;
            }

            DataTable table = testDb.Query("SELECT TOP 5 VoucherNo AS [Путевка], IssueDate AS [Регистрация], PensionerFullName AS [Пенсионер], SanatoriumName AS [Санаторий], TotalCost AS [Стоимость] FROM qryReport_VoucherIssuesDetailed ORDER BY IssueDate, VoucherNo");
            string testPdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pdf-self-test.pdf");
            ProfessionalPdfExporter.ExportDataTableToFile(testPdf, "Тестовый PDF-отчет", "Система здравоохранения", table);

            FileInfo info = new FileInfo(testPdf);
            if (!info.Exists || info.Length < 1000)
            {
                return 3;
            }

            try
            {
                info.Delete();
            }
            catch
            {
            }

            return 0;
        }

        public static void GenerateManualPdf(string targetPath)
        {
            if (string.IsNullOrWhiteSpace(targetPath))
            {
                throw new ArgumentException("Путь к PDF руководства не задан.");
            }

            string directory = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            ProfessionalPdfExporter.ExportDataTableToFile(targetPath, "Руководство пользователя", "Система здравоохранения", BuildManualTable());
        }

        public static Form CreateMainForm(string dbPath)
        {
            return new MainForm(dbPath);
        }

        public static Form CreateSplash()
        {
            return new SplashForm();
        }

        private static DataTable BuildManualTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Шаг", typeof(string));
            table.Columns.Add("Действие", typeof(string));

            table.Rows.Add("1", "Запустите Setup.exe и выберите папку установки.");
            table.Rows.Add("2", "Дождитесь завершения установки и проверьте, что файлы UserGuide.txt, UserGuide.rtf и UserGuide.pdf появились в целевой папке.");
            table.Rows.Add("3", "Запустите HealthcareSanatoriumInterface.exe. Откроется главное меню.");
            table.Rows.Add("4", "Нажмите Пенсионеры для работы с записями. Поиск срабатывает автоматически. Для изменения — кнопка Изменить или двойной клик по строке.");
            table.Rows.Add("5", "Нажмите Журнал путевок для работы с путевками. Фильтры применяются мгновенно. Редактирование — двойной клик или кнопка Изменить.");
            table.Rows.Add("6", "Нажмите Санатории для просмотра и редактирования справочника. Двойной клик по строке открывает форму изменения.");
            table.Rows.Add("7", "Нажмите “Сервис данных” для очистки демонстрационных записей и проверки состояния базы.");
            table.Rows.Add("8", "Используйте кнопки “Открыть PDF”, “Открыть RTF” и “Открыть TXT”, чтобы посмотреть руководство пользователя.");
            table.Rows.Add("9", "Для печати отчетов выбирайте режимы PDF или прямую печать в окне журнала путевок.");
            table.Rows.Add("10", "Если база данных находится не в папке установки, выберите её через “Выбрать БД”.");
            table.Rows.Add("11", "В сложных случаях используйте кнопку “Отчет: журнал” для просмотра стандартного Access-отчета.");
            table.Rows.Add("12", "Для обновления статистики на главной форме нажмите “Обновить” в нужном модуле или вернитесь в главное меню.");
            table.Rows.Add("13", "Если появляется предупреждение о базе, проверьте, что файл .accdb не перемещён и доступен для чтения.");
            table.Rows.Add("14", "Скриншоты и визуальные подсказки в этом руководстве рассчитаны на новичков, поэтому старайтесь выполнять шаги по порядку.");

            return table;
        }
    }

    internal static class ToastNotifier
    {
        private static readonly object SyncLock = new object();
        private static readonly List<Form> OpenToasts = new List<Form>();

        public static void Show(Control owner, string message, bool success)
        {
            Form parent = owner == null ? null : owner.FindForm();
            ToastForm toast = new ToastForm(message, success);
            Rectangle bounds = parent == null ? Screen.PrimaryScreen.WorkingArea : parent.RectangleToScreen(parent.ClientRectangle);
            toast.StartPosition = FormStartPosition.Manual;

            lock (SyncLock)
            {
                toast.Location = new Point(bounds.Right - toast.Width - 24, bounds.Bottom - toast.Height - 32 - (OpenToasts.Count * (toast.Height + 12)));
                OpenToasts.Add(toast);
            }

            toast.FormClosed += delegate
            {
                lock (SyncLock)
                {
                    OpenToasts.Remove(toast);
                }
            };
            toast.Show(parent);
        }

        private sealed class ToastForm : Form
        {
            private readonly string message;
            private readonly bool success;
            private readonly Timer timer;

            public ToastForm(string message, bool success)
            {
                this.message = message;
                this.success = success;
                FormBorderStyle = FormBorderStyle.None;
                ShowInTaskbar = false;
                int lineCount = message.Split(new[] { "\n" }, StringSplitOptions.None).Length;
                Width = 420;
                Height = 70 + (lineCount > 1 ? (lineCount - 1) * 18 : 0);
                Height = Math.Min(Height, 150);
                
                BackColor = Color.FromArgb(250, 250, 252);
                DoubleBuffered = true;
                AutoScaleMode = AutoScaleMode.Dpi;

                timer = new Timer();
                timer.Interval = 3500;
                timer.Tick += delegate
                {
                    timer.Stop();
                    Close();
                };
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                timer.Start();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                Rectangle rect = new Rectangle(1, 1, Width - 3, Height - 3);
                Color accent = success ? Theme.Green : Color.FromArgb(255, 149, 0);
                Color bgTop = success ? Color.FromArgb(245, 253, 245) : Color.FromArgb(255, 250, 245);
                Color bgBottom = success ? Color.FromArgb(235, 248, 235) : Color.FromArgb(255, 245, 235);
                
                using (GraphicsPath path = Rounded(rect, 18))
                using (LinearGradientBrush fill = new LinearGradientBrush(rect, Theme.Dark ? Color.FromArgb(235, 28, 36, 54) : bgTop, Theme.Dark ? Color.FromArgb(220, 18, 25, 40) : bgBottom, 90f))
                using (Pen border = new Pen(accent, 1.5f))
                using (SolidBrush accentBrush = new SolidBrush(accent))
                using (SolidBrush textBrush = new SolidBrush(Theme.Ink))
                using (Font titleFont = new Font("Segoe UI Semibold", 10f, FontStyle.Bold))
                using (Font messageFont = new Font("Segoe UI", 9f, FontStyle.Regular))
                {
                    e.Graphics.FillPath(fill, path);
                    e.Graphics.DrawPath(border, path);
                    
                    e.Graphics.FillEllipse(accentBrush, 18, 22, 18, 18);
                    string icon = success ? "✓" : "⚠";
                    using (SolidBrush iconBrush = new SolidBrush(Color.White))
                    using (Font iconFont = new Font("Segoe UI Semibold", 11f, FontStyle.Bold))
                    {
                        e.Graphics.DrawString(icon, iconFont, iconBrush, new RectangleF(20, 23, 14, 16));
                    }
                    
                    string[] lines = message.Split(new[] { "\n" }, StringSplitOptions.None);
                    float textY = 20;
                    
                    if (lines.Length > 0)
                    {
                        e.Graphics.DrawString(lines[0], titleFont, textBrush, new RectangleF(48, textY, Width - 72, 18));
                        textY += 22;
                    }
                    
                    if (lines.Length > 1)
                    {
                        for (int i = 1; i < lines.Length && textY < Height - 10; i++)
                        {
                            e.Graphics.DrawString(lines[i], messageFont, new SolidBrush(Theme.Muted), new RectangleF(48, textY, Width - 72, 18));
                            textY += 18;
                        }
                    }
                }
            }

            private static GraphicsPath Rounded(Rectangle rectangle, int radius)
            {
                int diameter = radius * 2;
                GraphicsPath path = new GraphicsPath();
                path.AddArc(rectangle.X, rectangle.Y, diameter, diameter, 180, 90);
                path.AddArc(rectangle.Right - diameter, rectangle.Y, diameter, diameter, 270, 90);
                path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rectangle.X, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }

    internal static class ReportRunner
    {
        public static bool OpenReport(string databasePath, string reportName, bool print)
        {
            object access = null;
            try
            {
                if (!File.Exists(databasePath))
                {
                    MessageBox.Show("Файл базы данных не найден:\n" + databasePath, "Отчет", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                Type type = Type.GetTypeFromProgID("Access.Application");
                if (type == null)
                {
                    MessageBox.Show("Microsoft Access не найден. Отчет можно открыть вручную из базы данных.", "Отчет", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                access = Activator.CreateInstance(type);
                type.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, access, new object[] { true });
                type.InvokeMember("OpenCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, access, new object[] { databasePath });
                object doCmd = type.InvokeMember("DoCmd", System.Reflection.BindingFlags.GetProperty, null, access, null);
                Type doCmdType = doCmd.GetType();
                int view = print ? 0 : 2;
                doCmdType.InvokeMember("OpenReport", System.Reflection.BindingFlags.InvokeMethod, null, doCmd, new object[] { reportName, view });
                if (print)
                {
                    doCmdType.InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doCmd, new object[] { 3, reportName, 0 });
                    type.InvokeMember("CloseCurrentDatabase", System.Reflection.BindingFlags.InvokeMethod, null, access, null);
                    type.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, access, null);
                    Marshal.ReleaseComObject(access);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Не удалось открыть или напечатать отчет.\n\n" + ex.Message, "Отчет", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                if (access != null)
                {
                    try { Marshal.ReleaseComObject(access); } catch { }
                }
                return false;
            }
        }
    }

    internal sealed class DocumentViewerForm : GlassForm
    {
        public DocumentViewerForm(string filePath, string title)
            : base(null)
        {
            Text = title;
            Width = 900;
            Height = 650;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(400, 300);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.BackColor = Color.Transparent;
            layout.ColumnCount = 1;
            layout.RowCount = 2;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            Controls.Add(layout);

            FlowLayoutPanel header = new FlowLayoutPanel();
            header.Dock = DockStyle.Fill;
            header.AutoSize = true;
            header.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            header.FlowDirection = FlowDirection.LeftToRight;
            header.WrapContents = false;
            header.Padding = new Padding(0, 0, 0, 12);
            layout.Controls.Add(header, 0, 0);

            Button back = new GlassButton();
            back.Text = "Назад";
            back.SetBounds(0, 0, 118, 38);
            Theme.StyleButton(back, false);
            back.Click += delegate { Close(); };
            header.Controls.Add(back);

            Label caption = Theme.Label(Path.GetFileName(filePath), Theme.H2Font, Theme.Ink);
            caption.Margin = new Padding(14, 6, 0, 0);
            header.Controls.Add(caption);

            if (filePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                WebBrowser browser = new WebBrowser();
                browser.Dock = DockStyle.Fill;
                browser.AllowWebBrowserDrop = false;
                browser.IsWebBrowserContextMenuEnabled = true;
                layout.Controls.Add(browser, 0, 1);
                try
                {
                    browser.Navigate(new Uri(filePath));
                }
                catch (Exception ex)
                {
                    browser.DocumentText = "<html><body style='font-family:Segoe UI;padding:24px'>Не удалось открыть PDF внутри приложения: " +
                                           System.Security.SecurityElement.Escape(ex.Message) + "</body></html>";
                }
                return;
            }

            RichTextBox viewer = new RichTextBox();
            viewer.ReadOnly = true;
            viewer.Dock = DockStyle.Fill;
            viewer.BackColor = Color.White;
            viewer.ForeColor = Color.Black;
            viewer.Font = new Font("Segoe UI", 10f);
            viewer.BorderStyle = BorderStyle.None;
            layout.Controls.Add(viewer, 0, 1);

            try
            {
                if (filePath.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase))
                {
                    LoadRtf(viewer, filePath);
                }
                else if (filePath.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    viewer.Text = ReadTextDocument(filePath);
                }
            }
            catch (Exception ex)
            {
                viewer.Text = "Ошибка при загрузке файла: " + ex.Message;
            }
        }

        private static void LoadRtf(RichTextBox viewer, string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    viewer.LoadFile(stream, RichTextBoxStreamType.RichText);
                }
                if (!LooksBroken(viewer.Text)) return;
            }
            catch
            {
            }

            string rtf = ReadTextDocument(filePath);
            if (!rtf.TrimStart().StartsWith("{\\rtf", StringComparison.OrdinalIgnoreCase))
            {
                viewer.Text = rtf;
                return;
            }

            viewer.Rtf = EscapeUnicodeForLegacyRtf(rtf);
        }

        private static string ReadTextDocument(string filePath)
        {
            byte[] bytes = File.ReadAllBytes(filePath);
            Encoding[] encodings = { new UTF8Encoding(false, true), Encoding.GetEncoding(1251), Encoding.GetEncoding(1252), Encoding.UTF8 };
            string best = null;

            foreach (Encoding enc in encodings)
            {
                try
                {
                    string text = enc.GetString(bytes);
                    if (best == null || ScoreText(text) > ScoreText(best)) best = text;
                    if (!LooksBroken(text)) break;
                }
                catch
                {
                }
            }

            if (best == null) best = File.ReadAllText(filePath, Encoding.UTF8);
            if (best.StartsWith("\ufeff")) best = best.Substring(1);
            return best;
        }

        private static bool LooksBroken(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            return text.IndexOf('�') >= 0 ||
                   text.IndexOf("Ð", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("РЎ", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("Рђ", StringComparison.Ordinal) >= 0 ||
                   text.IndexOf("СЃ", StringComparison.Ordinal) >= 0;
        }

        private static int ScoreText(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            int score = 0;
            foreach (char c in text)
            {
                if (c >= 'А' && c <= 'я') score += 2;
                if (c == '�') score -= 8;
            }
            return score;
        }

        private static string EscapeUnicodeForLegacyRtf(string rtf)
        {
            StringBuilder builder = new StringBuilder(rtf.Length + 256);
            foreach (char c in rtf)
            {
                if (c <= 127)
                {
                    builder.Append(c);
                }
                else
                {
                    int value = c;
                    if (value > 32767) value -= 65536;
                    builder.Append("\\u");
                    builder.Append(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    builder.Append('?');
                }
            }
            return builder.ToString();
        }
    }
}
