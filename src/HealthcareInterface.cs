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
            Application.Run(new MainForm(dbPath));
        }
    }

    internal static class Theme
    {
        public static bool Dark = false;
        public static Color Ink { get { return Dark ? Color.FromArgb(232, 238, 248) : Color.FromArgb(24, 31, 44); } }
        public static Color Muted { get { return Dark ? Color.FromArgb(164, 176, 196) : Color.FromArgb(91, 107, 129); } }
        public static Color Line { get { return Dark ? Color.FromArgb(64, 78, 105) : Color.FromArgb(207, 216, 230); } }
        public static Color Panel { get { return Dark ? Color.FromArgb(218, 24, 32, 48) : Color.FromArgb(238, 248, 255, 255); } }
        public static Color Surface { get { return Dark ? Color.FromArgb(20, 27, 42) : Color.White; } }
        public static Color Field { get { return Dark ? Color.FromArgb(28, 37, 56) : Color.White; } }
        public static Color Blue { get { return Dark ? Color.FromArgb(74, 144, 255) : Color.FromArgb(0, 122, 255); } }
        public static Color Green { get { return Dark ? Color.FromArgb(61, 220, 132) : Color.FromArgb(52, 199, 89); } }
        public static readonly Font BaseFont = new Font("Segoe UI", 10f, FontStyle.Regular);
        public static readonly Font TitleFont = new Font("Segoe UI Semibold", 21f, FontStyle.Bold);
        public static readonly Font H2Font = new Font("Segoe UI Semibold", 14f, FontStyle.Bold);
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
            button.FlatAppearance.MouseOverBackColor = Dark ? Color.FromArgb(38, 52, 78) : Color.FromArgb(238, 246, 255);
            button.FlatAppearance.MouseDownBackColor = Dark ? Color.FromArgb(24, 34, 54) : Color.FromArgb(219, 236, 255);
            button.UseVisualStyleBackColor = false;
            button.BackColor = glass == null ? (primary ? Blue : Surface) : Color.Transparent;
            button.ForeColor = primary ? Color.White : Ink;
            button.Font = ButtonFont;
            button.Margin = Padding.Empty;
            button.AutoEllipsis = true;
            button.MinimumSize = new Size(Math.Max(button.MinimumSize.Width, 150), Math.Max(button.MinimumSize.Height, 40));
            if (button.Width < 150) button.Width = 150;
            if (button.Height < 40) button.Height = 40;
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
            grid.RowTemplate.Height = 28;
            grid.ShowCellToolTips = false;
            grid.BackgroundColor = Surface;
            grid.BorderStyle = BorderStyle.None;
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Dark ? Color.FromArgb(35, 46, 68) : Color.FromArgb(246, 248, 252);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Ink;
            grid.ColumnHeadersDefaultCellStyle.Font = GridHeaderFont;
            grid.DefaultCellStyle.Font = GridCellFont;
            grid.DefaultCellStyle.BackColor = Surface;
            grid.DefaultCellStyle.ForeColor = Ink;
            grid.DefaultCellStyle.SelectionBackColor = Dark ? Color.FromArgb(50, 76, 116) : Color.FromArgb(224, 239, 255);
            grid.DefaultCellStyle.SelectionForeColor = Ink;
            grid.AlternatingRowsDefaultCellStyle.BackColor = Dark ? Color.FromArgb(24, 33, 50) : Color.FromArgb(250, 252, 255);
            grid.RowHeadersVisible = false;
            grid.GridColor = Dark ? Color.FromArgb(54, 66, 88) : Color.FromArgb(230, 235, 244);
            grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            try
            {
                typeof(DataGridView).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, null, grid, new object[] { true });
            }
            catch { }
        }

        public static void ApplyTo(Control root)
        {
            if (root == null) return;
            root.Font = BaseFont;
            if (!(root is GlassPanel)) root.BackColor = root is Form ? (Dark ? Color.FromArgb(13, 18, 29) : Color.FromArgb(232, 239, 249)) : root.BackColor;
            foreach (Control control in root.Controls)
            {
                if (control is Button)
                {
                    Button button = (Button)control;
                    GlassButton glass = button as GlassButton;
                    StyleButton(button, glass != null ? glass.Primary : button.FlatAppearance.BorderSize == 0);
                }
                else if (control is DataGridView)
                {
                    StyleGrid((DataGridView)control);
                }
                else if (control is TextBox)
                {
                    control.BackColor = Field;
                    control.ForeColor = Ink;
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
                    control.ForeColor = control.Font.Bold ? Ink : Muted;
                }
                ApplyTo(control);
            }
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
    public static void Debounce(string key, System.Action action, int delayMs = 300)
    {
        var timer = new System.Windows.Forms.Timer();
        timer.Interval = delayMs;
        timer.Tick += delegate
        {
            timer.Stop();
            timer.Dispose();
            try
            {
                action();
            }
            catch (Exception ex)
            {
                ErrorLogger.LogException(ex);
            }
        };
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

    public string GetStatusMessage()
    {
        if (!Exists())
        {
            return "Файл базы данных не найден:\n" + DatabasePath;
        }

        try
        {
            string provider = ResolveProvider();
            if (!string.IsNullOrWhiteSpace(provider))
            {
                return "База данных подключена и готова к работе";
            }
        }
        catch
        {
        }

        return "База данных временно недоступна";
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

    public System.Threading.Tasks.Task<DataTable> QueryAsync(string sql, params OleDbParameter[] parameters)
    {
        return System.Threading.Tasks.Task.Run(() => Query(sql, parameters));
    }

    public System.Threading.Tasks.Task<int> ExecuteAsync(string sql, params OleDbParameter[] parameters)
    {
        return System.Threading.Tasks.Task.Run(() => Execute(sql, parameters));
    }

    public System.Threading.Tasks.Task<int> ScalarIntAsync(string sql)
    {
        return System.Threading.Tasks.Task.Run(() => ScalarInt(sql));
    }

    public DataTable QueryWithRetry(string sql, int retryCount = 3, params OleDbParameter[] parameters)
    {
        for (int attempt = 0; attempt < retryCount; attempt++)
        {
            try
            {
                return Query(sql, parameters);
            }
            catch (Exception ex)
            {
                if (attempt < retryCount - 1 && IsRetryableError(ex))
                {
                    System.Threading.Thread.Sleep((int)Math.Pow(2, attempt) * 100);
                }
                else
                {
                    throw;
                }
            }
        }
        throw new InvalidOperationException("Не удалось выполнить запрос после " + retryCount + " попыток");
    }

    public int ExecuteWithRetry(string sql, int retryCount = 3, params OleDbParameter[] parameters)
    {
        for (int attempt = 0; attempt < retryCount; attempt++)
        {
            try
            {
                return Execute(sql, parameters);
            }
            catch (Exception ex)
            {
                if (attempt < retryCount - 1 && IsRetryableError(ex))
                {
                    System.Threading.Thread.Sleep((int)Math.Pow(2, attempt) * 100);
                }
                else
                {
                    throw;
                }
            }
        }
        throw new InvalidOperationException("Не удалось выполнить операцию после " + retryCount + " попыток");
    }

    private static bool IsRetryableError(Exception ex)
    {
        return ex is OleDbException || ex.InnerException is System.IO.IOException;
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
    private readonly Dictionary<string, DateTime> inlineConfirmations = new Dictionary<string, DateTime>();
    private Bitmap backgroundCache;
    private Size backgroundCacheSize;
    private bool backgroundCacheDark;

    protected GlassForm(DbContext db)
    {
        Db = db;
        Font = Theme.BaseFont;
        Icon = TryLoadAppIcon();
        BackColor = Theme.Dark ? Color.FromArgb(13, 18, 29) : Color.FromArgb(232, 239, 249);
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        DoubleBuffered = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);
        UpdateStyles();
        AutoScroll = true;
        Padding = new Padding(24);
        MinimumSize = new Size(1024, 680);
        DoubleBuffered = true;
        ResizeRedraw = true;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        MaximizeBox = true;
        Shown += delegate { ApplyTheme(); };
    }

    protected override void OnPaintBackground(PaintEventArgs e)
    {
        if (ClientSize.Width <= 0 || ClientSize.Height <= 0) return;

        if (backgroundCache == null || backgroundCacheSize != ClientSize || backgroundCacheDark != Theme.Dark)
        {
            ResetBackgroundCache();
            backgroundCache = new Bitmap(ClientSize.Width, ClientSize.Height);
            backgroundCacheSize = ClientSize;
            backgroundCacheDark = Theme.Dark;
            using (Graphics graphics = Graphics.FromImage(backgroundCache))
            {
                PaintCachedBackground(graphics);
            }
        }

        e.Graphics.DrawImageUnscaled(backgroundCache, Point.Empty);
    }

    protected override void OnClosed(EventArgs e)
    {
        ResetBackgroundCache();
        base.OnClosed(e);
    }

    private void ResetBackgroundCache()
    {
        if (backgroundCache != null)
        {
            backgroundCache.Dispose();
            backgroundCache = null;
        }
        backgroundCacheSize = Size.Empty;
    }

    private void PaintCachedBackground(Graphics graphics)
    {
        Rectangle bounds = new Rectangle(Point.Empty, ClientSize);
        Color from = Theme.Dark ? Color.FromArgb(11, 16, 26) : Color.FromArgb(235, 242, 252);
        Color to = Theme.Dark ? Color.FromArgb(25, 34, 52) : Color.FromArgb(250, 252, 255);
        using (LinearGradientBrush brush = new LinearGradientBrush(bounds, from, to, 45f))
        {
            graphics.FillRectangle(brush, bounds);
        }

        using (SolidBrush blue = new SolidBrush(Theme.Dark ? Color.FromArgb(58, 74, 144, 255) : Color.FromArgb(42, 0, 122, 255)))
        using (SolidBrush green = new SolidBrush(Theme.Dark ? Color.FromArgb(42, 61, 220, 132) : Color.FromArgb(32, 52, 199, 89)))
        {
            graphics.FillEllipse(blue, ClientSize.Width - 320, -90, 420, 260);
            graphics.FillEllipse(green, -120, ClientSize.Height - 220, 340, 260);
        }
    }

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

    protected void ShowChildPage(Form form, Action<bool> closed)
    {
        List<Control> hiddenControls = new List<Control>();
        foreach (Control control in Controls)
        {
            if (control.Visible)
            {
                hiddenControls.Add(control);
                control.Visible = false;
            }
        }

        form.TopLevel = false;
        form.FormBorderStyle = FormBorderStyle.None;
        form.Dock = DockStyle.Fill;
        form.StartPosition = FormStartPosition.Manual;
        form.FormClosed += delegate
        {
            bool accepted = form.DialogResult == DialogResult.OK;
            Controls.Remove(form);
            foreach (Control control in hiddenControls)
            {
                if (!control.IsDisposed) control.Visible = true;
            }
            if (closed != null) closed(accepted);
            form.Dispose();
        };
        Controls.Add(form);
        UiPerformance.Optimize(form);
        form.BringToFront();
        form.Show();
    }

    protected bool ConfirmInline(string key, string message)
    {
        DateTime now = DateTime.Now;
        DateTime previous;
        if (inlineConfirmations.TryGetValue(key, out previous) && (now - previous).TotalSeconds <= 6)
        {
            inlineConfirmations.Remove(key);
            return true;
        }

        inlineConfirmations[key] = now;
        ToastNotifier.Show(this, message + "\nНажмите эту кнопку еще раз в течение 6 секунд для подтверждения.", false);
        return false;
    }

    protected async System.Threading.Tasks.Task GuardAsync(System.Func<System.Threading.Tasks.Task> action)
    {
        try
        {
            if (!Db.Exists())
            {
                ErrorLogger.Log("Database not found (async): " + Db.DatabasePath);
                ToastNotifier.Show(this, "Файл базы данных не найден:\n" + Db.DatabasePath, false);
                return;
            }

            await action();
        }
        catch (Exception ex)
        {
            ErrorLogger.LogException(ex);
            ToastNotifier.Show(this, "Операция не выполнена: " + ex.Message, false);
        }
    }

    protected async System.Threading.Tasks.Task<T> GuardAsync<T>(System.Func<System.Threading.Tasks.Task<T>> action)
    {
        try
        {
            if (!Db.Exists())
            {
                ErrorLogger.Log("Database not found (async generic): " + Db.DatabasePath);
                ToastNotifier.Show(this, "Файл базы данных не найден:\n" + Db.DatabasePath, false);
                return default(T);
            }

            return await action();
        }
        catch (Exception ex)
        {
            ErrorLogger.LogException(ex);
            ToastNotifier.Show(this, "Операция не выполнена: " + ex.Message, false);
            return default(T);
        }
    }

    protected void FillLookup(ComboBox combo, DataTable table, string idField, string nameField, bool addAll)
    {
        combo.BeginUpdate();
        try
        {
            combo.Items.Clear();
            if (addAll)
            {
                combo.Items.Add(new LookupItem(0, "Все"));
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
            BackColor = Theme.Dark ? Color.FromArgb(13, 18, 29) : Color.FromArgb(232, 239, 249);
            ResetBackgroundCache();
            Theme.ApplyTo(this);
            UiPerformance.Optimize(this);
        }
        finally
        {
            ResumeLayout(false);
            Invalidate(false);
        }
    }
}


    internal sealed class GlassPanel : Panel
    {
        public GlassPanel()
        {
            DoubleBuffered = true;
            Padding = new Padding(18);
            BackColor = Color.Transparent;
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle shadowRect = new Rectangle(10, 14, Width - 22, Height - 24);
            Rectangle rect = new Rectangle(6, 4, Width - 13, Height - 12);
            using (GraphicsPath shadow = Rounded(shadowRect, 24))
            using (SolidBrush shadowBrush = new SolidBrush(Theme.Dark ? Color.FromArgb(90, 0, 0, 0) : Color.FromArgb(32, 78, 108, 150)))
            {
                e.Graphics.FillPath(shadowBrush, shadow);
            }
            using (GraphicsPath path = Rounded(rect, 22))
            using (LinearGradientBrush fill = new LinearGradientBrush(rect, Theme.Dark ? Color.FromArgb(215, 30, 40, 62) : Color.FromArgb(230, 255, 255, 255), Theme.Dark ? Color.FromArgb(150, 18, 27, 45) : Color.FromArgb(180, 244, 248, 255), 90f))
            using (Pen border = new Pen(Theme.Dark ? Color.FromArgb(90, 120, 150, 190) : Color.FromArgb(170, 255, 255, 255), 1f))
            {
                e.Graphics.FillPath(fill, path);
                using (LinearGradientBrush sheen = new LinearGradientBrush(new Rectangle(rect.X, rect.Y, rect.Width, Math.Max(24, rect.Height / 3)), Theme.Dark ? Color.FromArgb(42, 255, 255, 255) : Color.FromArgb(150, 255, 255, 255), Color.FromArgb(0, 255, 255, 255), 90f))
                {
                    e.Graphics.FillPath(sheen, path);
                }
                using (Pen inner = new Pen(Theme.Dark ? Color.FromArgb(55, 255, 255, 255) : Color.FromArgb(120, 255, 255, 255), 1f))
                {
                    Rectangle innerRect = new Rectangle(rect.X + 1, rect.Y + 1, rect.Width - 2, rect.Height - 2);
                    using (GraphicsPath innerPath = Rounded(innerRect, 20))
                    {
                        e.Graphics.DrawPath(inner, innerPath);
                    }
                }
                e.Graphics.DrawPath(border, path);
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

    internal sealed class GlassButton : Button
    {
        private bool hot;
        private bool pressed;

        public GlassButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
        }

        public bool Primary { get; set; }
        public bool Large { get; set; }

        protected override void OnMouseEnter(EventArgs e)
        {
            hot = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            hot = false;
            pressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            pressed = true;
            Invalidate();
            base.OnMouseDown(mevent);
        }

        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            pressed = false;
            Invalidate();
            base.OnMouseUp(mevent);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color clearColor = Parent == null || Parent.BackColor.A < 255 ? (Theme.Dark ? Color.FromArgb(13, 18, 29) : Color.FromArgb(232, 239, 249)) : Parent.BackColor;
            e.Graphics.Clear(clearColor);
            Rectangle rect = new Rectangle(1, 1, Width - 3, Height - 3);
            int radius = Math.Max(10, Math.Min(16, Height / 3));
            if (pressed) rect.Offset(0, 1);

            Color top;
            Color bottom;
            Color border;
            Color text;
            Color shineTop;
            if (Primary)
            {
                top = hot ? Color.FromArgb(242, 39, 137, 248) : Color.FromArgb(235, 18, 122, 234);
                bottom = hot ? Color.FromArgb(235, 6, 102, 214) : Color.FromArgb(224, 0, 91, 198);
                border = Color.FromArgb(150, 115, 179, 250);
                text = Color.White;
                shineTop = Color.FromArgb(58, 255, 255, 255);
            }
         else if (Theme.Dark)
            {
                top = hot ? Color.FromArgb(230, 50, 65, 90) : Color.FromArgb(210, 40, 55, 80);
                bottom = hot ? Color.FromArgb(195, 35, 48, 70) : Color.FromArgb(175, 28, 38, 60);
                border = Color.FromArgb(128, 110, 130, 160);
                text = Theme.Ink;
                shineTop = Color.FromArgb(42, 255, 255, 255);
            }
            else
            {
                top = hot ? Color.FromArgb(248, 250, 255, 255) : Color.FromArgb(242, 245, 252, 255);
                bottom = hot ? Color.FromArgb(235, 242, 250, 255) : Color.FromArgb(228, 238, 248, 255);
                border = Color.FromArgb(195, 205, 220, 235);
                text = Theme.Ink;
                shineTop = Color.FromArgb(90, 255, 255, 255);
            }

            Rectangle shadowRect = new Rectangle(rect.X, rect.Y + 2, rect.Width, rect.Height);
            using (GraphicsPath shadowPath = Rounded(shadowRect, radius))
            using (SolidBrush shadow = new SolidBrush(Theme.Dark ? Color.FromArgb(36, 0, 0, 0) : Color.FromArgb(14, 80, 104, 145)))
            {
                e.Graphics.FillPath(shadow, shadowPath);
            }

            using (GraphicsPath path = Rounded(rect, radius))
            using (LinearGradientBrush fill = new LinearGradientBrush(rect, top, bottom, 90f))
            using (Pen pen = new Pen(border, 1f))
            {
                e.Graphics.FillPath(fill, path);
                Rectangle shineRect = new Rectangle(rect.X + 2, rect.Y + 2, rect.Width - 4, Math.Max(8, rect.Height / 3));
                using (GraphicsPath shinePath = Rounded(shineRect, Math.Max(8, radius - 3)))
                using (LinearGradientBrush shine = new LinearGradientBrush(shineRect, shineTop, Color.FromArgb(0, 255, 255, 255), 90f))
                {
                    e.Graphics.FillPath(shine, shinePath);
                }
                using (Pen inner = new Pen(Theme.Dark ? Color.FromArgb(36, 255, 255, 255) : Color.FromArgb(92, 255, 255, 255), 1f))
                {
                    Rectangle innerRect = new Rectangle(rect.X + 1, rect.Y + 1, rect.Width - 2, rect.Height - 2);
                    using (GraphicsPath innerPath = Rounded(innerRect, Math.Max(8, radius - 1)))
                    {
                        e.Graphics.DrawPath(inner, innerPath);
                    }
                }
                e.Graphics.DrawPath(pen, path);
            }

            DrawCaption(e.Graphics, rect, text);
        }

        private void DrawCaption(Graphics g, Rectangle rect, Color textColor)
        {
            string[] lines = Text.Replace("\r", "").Split('\n');
            using (StringFormat format = new StringFormat())
            {
                format.Trimming = StringTrimming.EllipsisCharacter;
                format.Alignment = TextAlign == ContentAlignment.MiddleLeft ? StringAlignment.Near : StringAlignment.Center;
                format.LineAlignment = StringAlignment.Center;
                int leftPadding = Large ? 18 : 9;
                RectangleF textRect = new RectangleF(rect.X + leftPadding, rect.Y + 2, rect.Width - (leftPadding * 2), rect.Height - 4);

                if (lines.Length > 1)
                {
                    using (SolidBrush titleBrush = new SolidBrush(textColor))
                    using (SolidBrush subBrush = new SolidBrush(Primary ? Color.FromArgb(225, 255, 255, 255) : Theme.Muted))
                    {
                        float y = rect.Y + (rect.Height - 42) / 2f;
                        RectangleF titleRect = new RectangleF(textRect.X, y, textRect.Width, 22);
                        RectangleF subRect = new RectangleF(textRect.X, y + 24, textRect.Width, 18);
                        g.DrawString(lines[0], Theme.ButtonTitleFont, titleBrush, titleRect, format);
                        g.DrawString(lines[1], Theme.ButtonSubFont, subBrush, subRect, format);
                    }
                }
                else
                {
                    using (SolidBrush brush = new SolidBrush(textColor))
                    {
                        g.DrawString(Text, Theme.ButtonCaptionFont, brush, textRect, format);
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
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 330f));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        Controls.Add(root);

        GlassPanel sidebar = new GlassPanel();
        sidebar.Dock = DockStyle.Fill;
        sidebar.Padding = new Padding(18);
        sidebar.Margin = Padding.Empty;
        root.Controls.Add(sidebar, 0, 0);

        contentPanel = new GlassPanel();
        contentPanel.Dock = DockStyle.Fill;
        contentPanel.Padding = new Padding(24);
        contentPanel.Margin = Padding.Empty;
        root.Controls.Add(contentPanel, 1, 0);

        TableLayoutPanel sidebarLayout = new TableLayoutPanel();
        sidebarLayout.Dock = DockStyle.Fill;
        sidebarLayout.BackColor = Color.Transparent;
        sidebarLayout.ColumnCount = 1;
        sidebarLayout.RowCount = 4;
        sidebarLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        sidebarLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        sidebar.Controls.Add(sidebarLayout);

        Label title = Theme.Label("Система здравоохранения", Theme.TitleFont, Theme.Ink);
        title.Margin = new Padding(4, 2, 4, 8);
        sidebarLayout.Controls.Add(title, 0, 0);

        Label subtitle = Theme.Label("Санаторные путевки, пенсионеры, отчеты и сервис данных", Theme.BaseFont, Theme.Muted);
        subtitle.Margin = new Padding(4, 0, 4, 18);
        subtitle.MaximumSize = new Size(280, 0);
        sidebarLayout.Controls.Add(subtitle, 0, 1);

        Panel navHost = new Panel();
        navHost.Dock = DockStyle.Fill;
        navHost.AutoScroll = true;
        navHost.BackColor = Color.Transparent;
        sidebarLayout.Controls.Add(navHost, 0, 2);

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
        sidebarLayout.Controls.Add(sidebarActions, 0, 3);

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

        sectionHint = Theme.Label("Современное левое меню, быстрые действия и аккуратные уведомления для всех важных событий.", Theme.BaseFont, Theme.Muted);
        sectionHint.Margin = new Padding(0, 0, 0, 10);
        sectionHint.MaximumSize = new Size(760, 0);
        homeLayout.Controls.Add(sectionHint, 0, 1);

        stats = Theme.Label("", Theme.StatsFont, Theme.Blue);
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
        manualPdf.SetBounds(0, 0, 148, 42);
        Theme.StyleButton(manualPdf, false);
        manualPdf.Click += delegate { OpenManual("pdf"); };
        manualBar.Controls.Add(manualPdf);

        Button manualRtf = new GlassButton();
        manualRtf.Text = "Открыть RTF";
        manualRtf.SetBounds(0, 0, 148, 42);
        Theme.StyleButton(manualRtf, false);
        manualRtf.Click += delegate { OpenManual("rtf"); };
        manualBar.Controls.Add(manualRtf);

        Button manualTxt = new GlassButton();
        manualTxt.Text = "Открыть TXT";
        manualTxt.SetBounds(0, 0, 148, 42);
        Theme.StyleButton(manualTxt, false);
        manualTxt.Click += delegate { OpenManual("txt"); };
        manualBar.Controls.Add(manualTxt);

        Button allPdf = new GlassButton();
        allPdf.Text = "PDF: все таблицы";
        allPdf.SetBounds(0, 0, 170, 42);
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

        AddNavButton("Отчет: журнал", "Сохранить журнал в PDF", delegate
        {
            SetActiveSection("Отчет: журнал", "Сохраните отчет журнала как PDF внутри единого окна приложения.");
            Guard(delegate
            {
                DataTable table = ReportTables.Vouchers(Db);
                bool saved = ProfessionalPdfExporter.ExportDataTableWithDialog(this, "Журнал регистрации выдачи путевок", "Система здравоохранения", table, "journal_vouchers.pdf");
                if (saved) ToastNotifier.Show(this, "Отчет журнала сохранен как PDF", true);
            });
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
        Guard(delegate
        {
            int pensioners = Db.ScalarInt("SELECT Count(*) FROM tblPensioners");
            int vouchers = Db.ScalarInt("SELECT Count(*) FROM tblVoucherIssues");
            int sanatoriums = Db.ScalarInt("SELECT Count(*) FROM tblSanatoriums");
            stats.Text = "Пенсионеры: " + pensioners + "   Путевки: " + vouchers + "   Санатории: " + sanatoriums;
        });
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
        UiPerformance.Optimize(form);
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
        ShowEmbeddedForm(new DatabasePathForm(Db, delegate
        {
            RefreshStats();
            ToastNotifier.Show(this, "База подключена и данные обновлены", true);
        }));
    }
}


    internal sealed class DatabasePathForm : GlassForm
    {
        private readonly TextBox pathBox;
        private readonly Action saved;

        public DatabasePathForm(DbContext db, Action saved) : base(db)
        {
            this.saved = saved;
            Text = "Выбор базы данных";
            Size = new Size(900, 520);

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(24);
            Controls.Add(panel);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.BackColor = Color.Transparent;
            layout.ColumnCount = 1;
            layout.RowCount = 5;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.Controls.Add(layout);

            Label title = Theme.Label("Подключение базы данных", Theme.TitleFont, Theme.Ink);
            title.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(title, 0, 0);

            Label hint = Theme.Label("Введите полный путь к файлу .accdb или .mdb. Все происходит внутри текущего окна без системных диалогов.", Theme.BaseFont, Theme.Muted);
            hint.MaximumSize = new Size(760, 0);
            hint.Margin = new Padding(0, 0, 0, 16);
            layout.Controls.Add(hint, 0, 1);

            pathBox = Theme.TextBox("Путь к базе данных");
            pathBox.Text = db.DatabasePath;
            pathBox.Dock = DockStyle.Top;
            pathBox.Margin = new Padding(0, 0, 0, 12);
            layout.Controls.Add(pathBox, 0, 2);

            Label example = Theme.Label("Например: C:\\HealthcareSanatoriumSystem.accdb", Theme.BaseFont, Theme.Muted);
            example.Margin = new Padding(0, 0, 0, 12);
            layout.Controls.Add(example, 0, 3);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.WrapContents = false;
            buttons.AutoSize = true;
            layout.Controls.Add(buttons, 0, 4);

            Button back = new GlassButton();
            back.Text = "Назад";
            back.SetBounds(0, 0, 120, 40);
            Theme.StyleButton(back, false);
            back.Click += delegate { Close(); };
            buttons.Controls.Add(back);

            Button apply = new GlassButton();
            apply.Text = "Подключить";
            apply.SetBounds(0, 0, 150, 40);
            Theme.StyleButton(apply, true);
            apply.Click += Apply;
            buttons.Controls.Add(apply);
        }

        private void Apply(object sender, EventArgs e)
        {
            string path = pathBox.Text.Trim().Trim('"');
            if (!File.Exists(path))
            {
                ToastNotifier.Show(this, "Файл базы данных не найден:\n" + path, false);
                return;
            }

            string extension = Path.GetExtension(path);
            if (!extension.Equals(".accdb", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".mdb", StringComparison.OrdinalIgnoreCase))
            {
                ToastNotifier.Show(this, "Нужен файл Access .accdb или .mdb", false);
                return;
            }

            Db.SetDatabasePath(path);
            if (saved != null) saved();
            Close();
        }
    }


    internal sealed class PensionersForm : GlassForm
    {
        private readonly TextBox search;
        private readonly ComboBox region;
        private readonly CheckBox onlyWithVouchers;
        private readonly DataGridView grid;

        public PensionersForm(DbContext db) : base(db)
        {
            Text = "Пенсионеры";
            Size = new Size(1120, 720);

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            panel.SetBounds(28, 24, 1048, 610);
            Controls.Add(panel);

            Label title = Theme.Label("Пенсионеры", Theme.TitleFont, Theme.Ink);
            title.Location = new Point(24, 20);
            panel.Controls.Add(title);

            search = Theme.TextBox("Поиск");
            search.SetBounds(28, 88, 240, 30);
            search.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel.Controls.Add(search);

            region = new ComboBox();
            region.SetBounds(286, 88, 250, 30);
            region.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel.Controls.Add(region);

            onlyWithVouchers = new CheckBox();
            onlyWithVouchers.Text = "Только с путевками";
            onlyWithVouchers.SetBounds(554, 88, 190, 30);
            onlyWithVouchers.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            onlyWithVouchers.BackColor = Color.Transparent;
            panel.Controls.Add(onlyWithVouchers);

            Button load = new GlassButton();
            load.Text = "Обновить";
            load.SetBounds(766, 84, 124, 38);
            load.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(load, true);
            load.Click += delegate { LoadGrid(); };
            panel.Controls.Add(load);

            Button exportPdf = new GlassButton();
            exportPdf.Text = "PDF";
            exportPdf.SetBounds(904, 84, 114, 38);
            exportPdf.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(exportPdf, false);
            exportPdf.Click += delegate { ExportPdf(); };
            panel.Controls.Add(exportPdf);

            Button add = new GlassButton();
            add.Text = "Добавить";
            add.SetBounds(28, 552, 132, 40);
            add.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(add, true);
            add.Click += AddPensioner;
            panel.Controls.Add(add);

            Button delete = new GlassButton();
            delete.Text = "Удалить";
            delete.SetBounds(174, 552, 132, 40);
            delete.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(delete, false);
            delete.Click += DeletePensioner;
            panel.Controls.Add(delete);

            grid = new DataGridView();
            grid.SetBounds(28, 142, 990, 398);
            Theme.StyleGrid(grid);
            panel.Controls.Add(grid);

            Button close = new GlassButton();
            close.Text = "Назад";
            close.SetBounds(884, 552, 134, 40);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.StyleButton(close, false);
            close.Click += delegate { Close(); };
            panel.Controls.Add(close);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(region, Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName", true);
                    LoadGrid();
                });
            };
        }

        private void LoadGrid()
        {
            Guard(delegate
            {
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
                {
                    where.Add("EXISTS (SELECT 1 FROM tblVoucherIssues AS vi WHERE vi.PensionerID = p.PensionerID)");
                }
                string sql =
                    "SELECT p.PensionerID AS [Код], p.LastName AS [Фамилия], p.FirstName AS [Имя], p.MiddleName AS [Отчество], " +
                    "r.RegionName AS [Регион], c.CategoryName AS [Категория], p.BirthDate AS [Дата рождения], " +
                    "p.PensionCertificateNo AS [Удостоверение], p.Phone AS [Телефон] " +
                    "FROM (tblPensioners AS p INNER JOIN tblRegions AS r ON p.RegionID = r.RegionID) " +
                    "INNER JOIN tblPensionerCategories AS c ON p.CategoryID = c.CategoryID ";
                if (where.Count > 0) sql += "WHERE " + string.Join(" AND ", where.ToArray()) + " ";
                sql += "ORDER BY p.LastName, p.FirstName";
                UiPerformance.BindGrid(grid, Db.Query(sql, parameters.ToArray()));
            });
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
                ShowChildPage(new PensionerEditForm(Db), delegate(bool accepted)
                {
                    if (accepted)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Пенсионер добавлен. База обновлена", true);
                    }
                });
            });
        }

        private void DeletePensioner(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    ToastNotifier.Show(this, "Выберите пенсионера для удаления.", false);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (!ConfirmInline("delete-pensioner-" + id, "Удалить выбранного пенсионера и все его путевки?")) return;
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

        public VouchersForm(DbContext db) : base(db)
        {
            Text = "Журнал путевок";
            Size = new Size(1160, 740);

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            panel.SetBounds(28, 24, 1088, 632);
            Controls.Add(panel);

            Label title = Theme.Label("Журнал регистрации выдачи путевок", Theme.TitleFont, Theme.Ink);
            title.Location = new Point(24, 20);
            panel.Controls.Add(title);

            region = new ComboBox();
            region.SetBounds(28, 88, 230, 30);
            region.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel.Controls.Add(region);

            status = new ComboBox();
            status.SetBounds(276, 88, 220, 30);
            status.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel.Controls.Add(status);

            hideCanceled = new CheckBox();
            hideCanceled.Text = "Скрыть отмененные";
            hideCanceled.Checked = true;
            hideCanceled.SetBounds(514, 88, 190, 30);
            hideCanceled.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            hideCanceled.BackColor = Color.Transparent;
            panel.Controls.Add(hideCanceled);

            Button load = new GlassButton();
            load.Text = "Обновить";
            load.SetBounds(722, 84, 124, 38);
            load.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(load, true);
            load.Click += delegate { LoadGrid(); };
            panel.Controls.Add(load);

            Button preview = new GlassButton();
            preview.Text = "PDF без открытия окон";
            preview.SetBounds(860, 84, 144, 38);
            preview.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(preview, false);
            preview.Click += delegate { ShowPrintOptions(false); };
            panel.Controls.Add(preview);

            Button print = new GlassButton();
            print.Text = "Печать";
            print.SetBounds(1012, 84, 70, 38);
            print.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(print, false);
            print.Click += delegate { ShowPrintOptions(true); };
            panel.Controls.Add(print);

            Button add = new GlassButton();
            add.Text = "Добавить";
            add.SetBounds(28, 574, 132, 40);
            add.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(add, true);
            add.Click += AddVoucher;
            panel.Controls.Add(add);

            Button delete = new GlassButton();
            delete.Text = "Удалить";
            delete.SetBounds(174, 574, 132, 40);
            delete.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(delete, false);
            delete.Click += DeleteVoucher;
            panel.Controls.Add(delete);

            grid = new DataGridView();
            grid.SetBounds(28, 142, 1030, 420);
            Theme.StyleGrid(grid);
            panel.Controls.Add(grid);

            Button close = new GlassButton();
            close.Text = "Назад";
            close.SetBounds(924, 574, 134, 40);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.StyleButton(close, false);
            close.Click += delegate { Close(); };
            panel.Controls.Add(close);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(region, Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName", true);
                    FillLookup(status, Db.LookupCached("tblVoucherStatuses", "StatusID", "StatusName"), "StatusID", "StatusName", true);
                    LoadGrid();
                });
            };
        }

        private void LoadGrid()
        {
            Guard(delegate
            {
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
                {
                    where.Add("vs.StatusName <> 'Отменена'");
                }
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
                UiPerformance.BindGrid(grid, Db.Query(sql, parameters.ToArray()));
            });
        }

        private void AddVoucher(object sender, EventArgs e)
        {
            Guard(delegate
            {
                ShowChildPage(new VoucherEditForm(Db), delegate(bool accepted)
                {
                    if (accepted)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Путевка добавлена. База обновлена", true);
                    }
                });
            });
        }

        private void DeleteVoucher(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    ToastNotifier.Show(this, "Выберите путевку для удаления.", false);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (!ConfirmInline("delete-voucher-" + id, "Удалить выбранную запись журнала?")) return;
                Db.Execute("DELETE FROM tblVoucherIssues WHERE IssueID = ?", new OleDbParameter("issue", id));
                LoadGrid();
                ToastNotifier.Show(this, "Запись журнала удалена. База обновлена", true);
            });
        }

        private void ShowPrintOptions(bool preferPrint)
        {
            Guard(delegate
            {
                PrintOptionsForm options = new PrintOptionsForm(preferPrint);
                ShowChildPage(options, delegate(bool accepted)
                {
                    if (!accepted) return;
                    ApplyPrintOptions(options);
                });
            });
        }

        private void ApplyPrintOptions(PrintOptionsForm options)
        {
            Guard(delegate
            {
                if (options.Target == PrintTarget.AllTables)
                {
                    if (ReportTables.ExportAllTablesPdf(this, Db))
                    {
                        ToastNotifier.Show(this, "PDF со всеми таблицами сохранен", true);
                    }
                    return;
                }

                bool saved = false;
                if (options.Target == PrintTarget.AccessJournal)
                {
                    DataTable table = ReportTables.Vouchers(Db);
                    saved = ProfessionalPdfExporter.ExportDataTableWithDialog(this, "Журнал регистрации выдачи путевок", "Система здравоохранения", table, "journal_vouchers.pdf");
                }
                else if (options.Target == PrintTarget.AccessRegionTotals)
                {
                    DataTable table = ReportTables.RegionTotals(Db);
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
                        ToastNotifier.Show(this, "Выберите строку журнала для PDF.", false);
                        return;
                    }
                    saved = ProfessionalPdfExporter.ExportGridWithDialog(this, "Журнал путевок - выбранная строка", grid, true, "selected_voucher.pdf");
                }

                if (saved) ToastNotifier.Show(this, "PDF-файл сохранен в папку документов", true);
            });
        }
    }

    internal sealed class SanatoriumsForm : GlassForm
    {
        private readonly ComboBox profile;
        private readonly CheckBox onlyActive;
        private readonly DataGridView grid;

        public SanatoriumsForm(DbContext db) : base(db)
        {
            Text = "Санатории";
            Size = new Size(1080, 700);

            GlassPanel panel = new GlassPanel();
            panel.Dock = DockStyle.Fill;
            panel.SetBounds(28, 24, 1008, 590);
            Controls.Add(panel);

            Label title = Theme.Label("Санатории", Theme.TitleFont, Theme.Ink);
            title.Location = new Point(24, 20);
            panel.Controls.Add(title);

            profile = new ComboBox();
            profile.SetBounds(28, 88, 280, 30);
            profile.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel.Controls.Add(profile);

            onlyActive = new CheckBox();
            onlyActive.Text = "Только работающие";
            onlyActive.Checked = true;
            onlyActive.BackColor = Color.Transparent;
            onlyActive.SetBounds(326, 88, 190, 30);
            onlyActive.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            panel.Controls.Add(onlyActive);

            Button load = new GlassButton();
            load.Text = "Обновить";
            load.SetBounds(536, 84, 124, 38);
            load.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(load, true);
            load.Click += delegate { LoadGrid(); };
            panel.Controls.Add(load);

            Button exportPdf = new GlassButton();
            exportPdf.Text = "PDF";
            exportPdf.SetBounds(674, 84, 114, 38);
            exportPdf.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            Theme.StyleButton(exportPdf, false);
            exportPdf.Click += delegate { ExportPdf(); };
            panel.Controls.Add(exportPdf);

            Button add = new GlassButton();
            add.Text = "Добавить";
            add.SetBounds(28, 520, 132, 40);
            add.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(add, true);
            add.Click += AddSanatorium;
            panel.Controls.Add(add);

            Button delete = new GlassButton();
            delete.Text = "Удалить";
            delete.SetBounds(174, 520, 132, 40);
            delete.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Theme.StyleButton(delete, false);
            delete.Click += DeleteSanatorium;
            panel.Controls.Add(delete);

            grid = new DataGridView();
            grid.SetBounds(28, 142, 952, 365);
            Theme.StyleGrid(grid);
            panel.Controls.Add(grid);

            Button close = new GlassButton();
            close.Text = "Назад";
            close.SetBounds(846, 520, 134, 40);
            close.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            Theme.StyleButton(close, false);
            close.Click += delegate { Close(); };
            panel.Controls.Add(close);

            Load += delegate
            {
                Guard(delegate
                {
                    FillLookup(profile, Db.LookupCached("tblMedicalProfiles", "ProfileID", "ProfileName"), "ProfileID", "ProfileName", true);
                    LoadGrid();
                });
            };
        }

        private void LoadGrid()
        {
            Guard(delegate
            {
                List<string> where = new List<string>();
                List<OleDbParameter> parameters = new List<OleDbParameter>();
                LookupItem selectedProfile = profile.SelectedItem as LookupItem;
                if (selectedProfile != null && selectedProfile.Id > 0)
                {
                    where.Add("s.ProfileID = ?");
                    parameters.Add(new OleDbParameter("profile", selectedProfile.Id));
                }
                if (onlyActive.Checked)
                {
                    where.Add("s.IsActive = True");
                }
                string sql =
                    "SELECT s.SanatoriumID, s.SanatoriumName AS [Санаторий], r.RegionName AS [Регион], mp.ProfileName AS [Профиль], " +
                    "s.CapacityBeds AS [Коек], s.PricePerDay AS [Цена дня], s.Phone AS [Телефон], s.Address AS [Адрес] " +
                    "FROM (tblSanatoriums AS s INNER JOIN tblRegions AS r ON s.RegionID = r.RegionID) " +
                    "INNER JOIN tblMedicalProfiles AS mp ON s.ProfileID = mp.ProfileID ";
                if (where.Count > 0) sql += "WHERE " + string.Join(" AND ", where.ToArray()) + " ";
                sql += "ORDER BY s.SanatoriumName";
                UiPerformance.BindGrid(grid, Db.Query(sql, parameters.ToArray()));
            });
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
                ShowChildPage(new SanatoriumEditForm(Db), delegate(bool accepted)
                {
                    if (accepted)
                    {
                        LoadGrid();
                        ToastNotifier.Show(this, "Санаторий добавлен. База обновлена", true);
                    }
                });
            });
        }

        private void DeleteSanatorium(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (grid.SelectedRows.Count == 0)
                {
                    ToastNotifier.Show(this, "Выберите санаторий для удаления.", false);
                    return;
                }
                int id = Convert.ToInt32(grid.SelectedRows[0].Cells[0].Value);
                if (!ConfirmInline("delete-sanatorium-" + id, "Удалить санаторий и связанные путевки?")) return;
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
        MinimumSize = new Size(Math.Max(920, size.Width), Math.Max(620, size.Height));
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
        ok.SetBounds(0, 0, 130, 38);
        Theme.StyleButton(ok, true);
        ok.Click += save;

        Button cancel = new GlassButton();
        cancel.Text = "Отмена";
        cancel.SetBounds(0, 0, 120, 38);
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
                ToastNotifier.Show(this, "Заполните обязательные поля.", false);
                box.Focus();
                return false;
            }
        }

        return true;
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

        public PensionerEditForm(DbContext db) : base(db, "Добавление пенсионера", new Size(720, 690))
        {
            region = AddLookup("Регион", Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName");
            category = AddLookup("Категория", Db.LookupCached("tblPensionerCategories", "CategoryID", "CategoryName"), "CategoryID", "CategoryName");
            last = AddText("Фамилия *", "", 330);
            first = AddText("Имя *", "", 330);
            middle = AddText("Отчество", "", 330);
            gender = AddSimpleCombo("Пол", new string[] { "М", "Ж" });
            birth = AddDate("Дата рождения", DateTime.Today.AddYears(-65));
            certificate = AddText("Удостоверение *", "PEN-" + DateTime.Now.ToString("yyyyMMddHHmmss"), 330);
            snils = AddText("СНИЛС", "", 250);
            phone = AddText("Телефон", "", 250);
            address = AddText("Адрес", "", 380);
            AddActions(Save);
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(last, first, certificate)) return;
                Db.Execute("INSERT INTO tblPensioners (RegionID, CategoryID, LastName, FirstName, MiddleName, Gender, BirthDate, PensionCertificateNo, SNILS, Phone, Address) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
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
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    internal sealed class VoucherEditForm : RecordFormBase
    {
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

        public VoucherEditForm(DbContext db) : base(db, "Добавление путевки", new Size(760, 660))
        {
            number = AddText("Номер путевки *", "V-" + DateTime.Now.ToString("yyyyMMddHHmmss"), 300);
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
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(number)) return;
                if (SelectedId(pensioner) == 0 || SelectedId(sanatorium) == 0)
                {
                    ToastNotifier.Show(this, "Нужны хотя бы один пенсионер и один санаторий.", false);
                    return;
                }
                if (endDate.Value.Date < startDate.Value.Date)
                {
                    ToastNotifier.Show(this, "Дата выезда не может быть раньше даты заезда.", false);
                    return;
                }
                Db.Execute("INSERT INTO tblVoucherIssues (IssueDate, VoucherNo, PensionerID, SanatoriumID, FundingSourceID, StatusID, StartDate, EndDate, CoPaymentPercent, Notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
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
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    internal sealed class SanatoriumEditForm : RecordFormBase
    {
        private readonly ComboBox region;
        private readonly ComboBox profile;
        private readonly TextBox name;
        private readonly TextBox address;
        private readonly TextBox phone;
        private readonly NumericUpDown beds;
        private readonly NumericUpDown price;
        private readonly CheckBox active;

        public SanatoriumEditForm(DbContext db) : base(db, "Добавление санатория", new Size(720, 600))
        {
            region = AddLookup("Регион", Db.LookupCached("tblRegions", "RegionID", "RegionName"), "RegionID", "RegionName");
            profile = AddLookup("Профиль", Db.LookupCached("tblMedicalProfiles", "ProfileID", "ProfileName"), "ProfileID", "ProfileName");
            name = AddText("Санаторий *", "", 360);
            address = AddText("Адрес", "", 390);
            phone = AddText("Телефон", "", 240);
            beds = AddNumber("Коек", 100, 5000);
            price = AddNumber("Цена дня", 3500, 100000);
            active = AddCheck("Работает", true);
            AddActions(Save);
        }

        private void Save(object sender, EventArgs e)
        {
            Guard(delegate
            {
                if (!Require(name)) return;
                Db.Execute("INSERT INTO tblSanatoriums (RegionID, ProfileID, SanatoriumName, Address, Phone, CapacityBeds, PricePerDay, IsActive) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    new OleDbParameter("region", SelectedId(region)),
                    new OleDbParameter("profile", SelectedId(profile)),
                    new OleDbParameter("name", name.Text.Trim()),
                    new OleDbParameter("address", Optional(address.Text)),
                    new OleDbParameter("phone", Optional(phone.Text)),
                    new OleDbParameter("beds", Convert.ToInt32(beds.Value)),
                    new OleDbParameter("price", price.Value),
                    new OleDbParameter("active", active.Checked));
                DialogResult = DialogResult.OK;
                Close();
            });
        }
    }

    
internal sealed class DataToolsForm : GlassForm
{
    private readonly Label stats;

    public DataToolsForm(DbContext db) : base(db)
    {
        Text = "Сервис данных";
        Size = new Size(980, 640);
        MinimumSize = new Size(920, 620);

        GlassPanel panel = new GlassPanel();
        panel.Dock = DockStyle.Fill;
        panel.Padding = new Padding(24);
        panel.AutoScroll = true;
        Controls.Add(panel);

        TableLayoutPanel layout = new TableLayoutPanel();
        layout.Dock = DockStyle.Fill;
        layout.BackColor = Color.Transparent;
        layout.ColumnCount = 1;
        layout.RowCount = 5;
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(layout);

        Label title = Theme.Label("Сервис данных", Theme.TitleFont, Theme.Ink);
        title.Margin = new Padding(0, 0, 0, 6);
        layout.Controls.Add(title, 0, 0);

        Label hint = Theme.Label("Очистка не удаляет файл базы, формы, отчеты и связи. Перед полной очисткой требуется подтверждение.", Theme.BaseFont, Theme.Muted);
        hint.Margin = new Padding(0, 0, 0, 14);
        hint.MaximumSize = new Size(840, 0);
        layout.Controls.Add(hint, 0, 1);

        stats = Theme.Label("", Theme.H2Font, Theme.Blue);
        stats.Margin = new Padding(0, 0, 0, 10);
        layout.Controls.Add(stats, 0, 2);

        FlowLayoutPanel actions = new FlowLayoutPanel();
        actions.Dock = DockStyle.Fill;
        actions.AutoSize = true;
        actions.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        actions.FlowDirection = FlowDirection.LeftToRight;
        actions.WrapContents = true;
        actions.Padding = new Padding(0, 8, 0, 8);
        layout.Controls.Add(actions, 0, 3);

        AddAction(actions, "Очистить журнал путевок", 0, 0, ClearJournal);
        AddAction(actions, "Очистить пенсионеров и журнал", 0, 0, ClearPensioners);
        AddAction(actions, "Очистить все демо-данные", 0, 0, ClearAll);
        AddAction(actions, "Обновить статистику", 0, 0, delegate { RefreshStats(); ToastNotifier.Show(this, "Статистика обновлена", true); });

        FlowLayoutPanel footer = new FlowLayoutPanel();
        footer.Dock = DockStyle.Fill;
        footer.AutoSize = true;
        footer.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        footer.FlowDirection = FlowDirection.RightToLeft;
        footer.WrapContents = false;
        footer.Padding = new Padding(0, 12, 0, 0);
        layout.Controls.Add(footer, 0, 4);

        Button close = new GlassButton();
        close.Text = "Назад";
        close.SetBounds(0, 0, 134, 40);
        Theme.StyleButton(close, false);
        close.Click += delegate { Close(); };
        footer.Controls.Add(close);

        Load += delegate { RefreshStats(); };
    }

    private void AddAction(Control parent, string text, int x, int y, EventHandler click)
    {
        Button button = new GlassButton();
        button.Text = text;
        button.SetBounds(x, y, 250, 62);
        button.Margin = new Padding(0, 0, 14, 14);
        Theme.StyleButton(button, text.IndexOf("все", StringComparison.OrdinalIgnoreCase) >= 0);
        button.Click += click;
        parent.Controls.Add(button);
    }

    private void RefreshStats()
    {
        Guard(delegate
        {
            stats.Text = "Пенсионеры: " + Db.ScalarInt("SELECT Count(*) FROM tblPensioners") +
                         "   Путевки: " + Db.ScalarInt("SELECT Count(*) FROM tblVoucherIssues") +
                         "   Санатории: " + Db.ScalarInt("SELECT Count(*) FROM tblSanatoriums");
        });
    }

    private void ClearJournal(object sender, EventArgs e)
    {
        Guard(delegate
        {
            if (!ConfirmInline("clear-journal", "Удалить все записи журнала выдачи путевок?")) return;
            Db.Execute("DELETE FROM tblVoucherIssues");
            RefreshStats();
            ToastNotifier.Show(this, "Журнал очищен. База обновлена", true);
        });
    }

    private void ClearPensioners(object sender, EventArgs e)
    {
        Guard(delegate
        {
            if (!ConfirmInline("clear-pensioners", "Удалить пенсионеров и все связанные путевки?")) return;
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
            if (!ConfirmInline("clear-all", "Будут удалены ВСЕ записи кроме справочников. Продолжить?")) return;
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
        preview.Text = "PDF без открытия окон";
        preview.AutoSize = true;
        preview.Checked = !preferPrint;
        preview.BackColor = Color.Transparent;
        preview.ForeColor = Theme.Ink;
        modes.Controls.Add(preview);

        direct = new RadioButton();
        direct.Text = "Быстро сохранить PDF";
        direct.AutoSize = true;
        direct.Checked = preferPrint;
        direct.BackColor = Color.Transparent;
        direct.ForeColor = Theme.Ink;
        modes.Controls.Add(direct);

        pdf = new RadioButton();
        pdf.Text = "Красивый PDF";
        pdf.AutoSize = true;
        pdf.BackColor = Color.Transparent;
        pdf.ForeColor = Theme.Ink;
        modes.Controls.Add(pdf);

        Label hint = Theme.Label("Отчеты сохраняются в PDF автоматически внутри приложения: без системных диалогов, предпросмотров и отдельных окон.", Theme.BaseFont, Theme.Muted);
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
        cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        Theme.StyleButton(cancel, false);
        buttons.Controls.Add(cancel);

        Button ok = new GlassButton();
        ok.Text = "Продолжить";
        ok.SetBounds(0, 0, 120, 38);
        ok.DialogResult = DialogResult.OK;
        ok.Click += delegate { DialogResult = DialogResult.OK; Close(); };
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
                return;
            }

            using (PrintDocument document = new PrintDocument())
            {
                document.DocumentName = title;
                document.DefaultPageSettings.Landscape = true;
                document.PrintPage += job.PrintPage;

                document.PrintController = new StandardPrintController();
                document.Print();
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

        private static string BuildExportPath(string defaultFileName)
        {
            string safeName = string.IsNullOrWhiteSpace(defaultFileName) ? "report.pdf" : defaultFileName;
            if (!safeName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) safeName += ".pdf";
            foreach (char invalid in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(invalid, '_');
            }

            string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "HealthcareExports");
            if (string.IsNullOrWhiteSpace(folder) || folder.StartsWith("HealthcareExports", StringComparison.OrdinalIgnoreCase))
            {
                folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "HealthcareExports");
            }
            Directory.CreateDirectory(folder);

            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = Path.GetFileNameWithoutExtension(safeName) + "_" + stamp + ".pdf";
            return Path.Combine(folder, fileName);
        }

        private static void Notify(IWin32Window owner, string message, bool success)
        {
            ToastNotifier.Show(owner as Control, message, success);
        }

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
                Notify(owner, "Нет данных для PDF.", false);
                return false;
            }

            List<byte[]> pages = new List<byte[]>();
            foreach (Section section in sections)
            {
                if (section == null) continue;
                pages.AddRange(RenderPages(title, section.Title, section.Headers, section.Rows));
            }

            if (pages.Count == 0)
            {
                Notify(owner, "Нет страниц для PDF.", false);
                return false;
            }

            string path = BuildExportPath(defaultFileName);
            WriteImagePdf(path, pages);
            Notify(owner, "PDF сохранен:\n" + path, true);
            return true;
        }

        public static bool ExportDataTableWithDialog(IWin32Window owner, string title, string subtitle, DataTable table, string defaultFileName)
        {
            if (table == null || table.Rows.Count == 0)
            {
                Notify(owner, "Нет данных для PDF.", false);
                return false;
            }

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

            string path = BuildExportPath(defaultFileName);
            Export(path, title, subtitle, headers, rows);
            Notify(owner, "PDF сохранен:\n" + path, true);
            return true;
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
                Notify(owner, "Нет колонок для PDF.", false);
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
                Notify(owner, "Нет строк для PDF.", false);
                return false;
            }

            string path = BuildExportPath(defaultFileName);
            Export(path, title, "Система здравоохранения", headers, rows);
            Notify(owner, "PDF сохранен:\n" + path, true);
            return true;
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

        public static System.Threading.Tasks.Task ExportAsync(string path, string title, string subtitle, List<string> headers, List<string[]> rows)
        {
            return System.Threading.Tasks.Task.Run(() => Export(path, title, subtitle, headers, rows));
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

        private static DataTable BuildManualTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Шаг", typeof(string));
            table.Columns.Add("Действие", typeof(string));

            table.Rows.Add("1", "Запустите Setup.exe и выберите папку установки.");
            table.Rows.Add("2", "Дождитесь завершения установки и проверьте, что файлы UserGuide.txt, UserGuide.rtf и UserGuide.pdf появились в целевой папке.");
            table.Rows.Add("3", "Запустите HealthcareSanatoriumInterface.exe. Откроется главное меню.");
            table.Rows.Add("4", "Нажмите “Пенсионеры”, чтобы добавить или удалить записи пенсионеров. Используйте фильтры и поиск.");
            table.Rows.Add("5", "Нажмите “Журнал путевок” для работы с путевками, фильтрами и печатью отчётов.");
            table.Rows.Add("6", "Нажмите “Санатории”, чтобы просматривать и добавлять новые санатории.");
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
        private static readonly Dictionary<Control, Panel> ActivePanels = new Dictionary<Control, Panel>();

        public static void Show(Control owner, string message, bool success)
        {
            Control host = owner == null ? null : owner.FindForm();
            if (host == null) host = owner;
            if (host == null) return;

            Panel oldPanel;
            if (ActivePanels.TryGetValue(host, out oldPanel) && oldPanel != null && !oldPanel.IsDisposed)
            {
                host.Controls.Remove(oldPanel);
                oldPanel.Dispose();
            }

            Panel panel = new Panel();
            panel.Width = Math.Min(520, Math.Max(320, host.ClientSize.Width - 48));
            panel.Height = message.IndexOf('\n') >= 0 ? 92 : 68;
            panel.Left = Math.Max(12, host.ClientSize.Width - panel.Width - 24);
            panel.Top = Math.Max(12, host.ClientSize.Height - panel.Height - 24);
            panel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
            panel.BackColor = success ? (Theme.Dark ? Color.FromArgb(31, 64, 48) : Color.FromArgb(235, 250, 240)) :
                                      (Theme.Dark ? Color.FromArgb(72, 52, 28) : Color.FromArgb(255, 246, 232));
            panel.Padding = new Padding(14, 10, 14, 10);

            Label label = new Label();
            label.Dock = DockStyle.Fill;
            label.BackColor = Color.Transparent;
            label.ForeColor = Theme.Ink;
            label.Font = Theme.BaseFont;
            label.Text = (success ? "✓ " : "⚠ ") + message;
            label.AutoEllipsis = true;
            panel.Controls.Add(label);

            host.Controls.Add(panel);
            panel.BringToFront();
            ActivePanels[host] = panel;

            Timer timer = new Timer();
            timer.Interval = 3800;
            timer.Tick += delegate
            {
                timer.Stop();
                timer.Dispose();
                if (!panel.IsDisposed)
                {
                    host.Controls.Remove(panel);
                    panel.Dispose();
                }
                if (ActivePanels.ContainsKey(host) && ReferenceEquals(ActivePanels[host], panel))
                {
                    ActivePanels.Remove(host);
                }
            };
            timer.Start();
        }
    }

    internal static class Prompt
    {
        public static string Show(string text, string caption)
        {
            return string.Empty;
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
                    ErrorLogger.Log("Report database not found: " + databasePath);
                    return false;
                }
                Type type = Type.GetTypeFromProgID("Access.Application");
                if (type == null)
                {
                    ErrorLogger.Log("Microsoft Access not found for report: " + reportName);
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
                ErrorLogger.LogException(ex);
                if (access != null)
                {
                    try { Marshal.ReleaseComObject(access); } catch { }
                }
                return false;
            }
        }
    }

    internal sealed class DocumentViewerForm : Form
    {
        public DocumentViewerForm(string filePath, string title)
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
            string rtf = ReadTextDocument(filePath);
            if (!rtf.TrimStart().StartsWith("{\\rtf", StringComparison.OrdinalIgnoreCase))
            {
                viewer.Text = FixMojibake(rtf);
                return;
            }

            string plainText = RtfToPlainText(rtf);
            viewer.Text = FixMojibake(plainText);
        }

        private static string RtfToPlainText(string rtf)
        {
            int contentStart = rtf.IndexOf("\\pard", StringComparison.OrdinalIgnoreCase);
            if (contentStart > 0) rtf = rtf.Substring(contentStart);

            StringBuilder text = new StringBuilder();
            for (int i = 0; i < rtf.Length; i++)
            {
                char c = rtf[i];
                if (c == '{' || c == '}') continue;
                if (c != '\\')
                {
                    text.Append(c);
                    continue;
                }

                if (i + 1 >= rtf.Length) break;
                char next = rtf[++i];
                if (next == '\\' || next == '{' || next == '}')
                {
                    text.Append(next);
                    continue;
                }

                int wordStart = i;
                while (i < rtf.Length && char.IsLetter(rtf[i])) i++;
                string word = rtf.Substring(wordStart, i - wordStart);
                bool negative = i < rtf.Length && rtf[i] == '-';
                if (negative) i++;
                int numberStart = i;
                while (i < rtf.Length && char.IsDigit(rtf[i])) i++;
                string number = rtf.Substring(numberStart, i - numberStart);
                if (i < rtf.Length && rtf[i] == ' ')
                {
                }
                else
                {
                    i--;
                }

                if (word == "par" || word == "line")
                {
                    text.Append(Environment.NewLine);
                }
                else if (word == "tab")
                {
                    text.Append('\t');
                }
                else if (word == "u" && number.Length > 0)
                {
                    int value;
                    if (int.TryParse((negative ? "-" : "") + number, out value))
                    {
                        if (value < 0) value += 65536;
                        text.Append((char)value);
                        if (i + 1 < rtf.Length && rtf[i + 1] == '?') i++;
                    }
                }
            }

            string result = text.ToString();
            while (result.Contains(Environment.NewLine + Environment.NewLine + Environment.NewLine))
            {
                result = result.Replace(Environment.NewLine + Environment.NewLine + Environment.NewLine, Environment.NewLine + Environment.NewLine);
            }
            return result.Trim();
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

        private static string FixMojibake(string text)
        {
            if (!LooksBroken(text)) return text;
            try
            {
                byte[] bytes = Encoding.GetEncoding(1251).GetBytes(text);
                string repaired = Encoding.UTF8.GetString(bytes);
                if (!LooksBroken(repaired)) return repaired;
            }
            catch
            {
            }
            return text;
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
