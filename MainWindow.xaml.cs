using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using ExcelDataReader;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Security.Cryptography;

namespace SuperNapominalka
{
    public class AppConfig
    {
        public string SmtpServer { get; set; } = string.Empty;
        public int SmtpPort { get; set; } = 587;
        public string SenderEmail { get; set; } = string.Empty;
        public string TargetEmail { get; set; } = string.Empty;
        public string MessageTemplate { get; set; } = "Напоминание: через неделю событие у {ФИО} ({Дата})!";
        public string EncryptedPassword { get; set; } = string.Empty;
        public string ExcelFilePath { get; set; } = string.Empty;
        public bool IsAutostartEnabled { get; set; } = false;
        public int ScheduleMode { get; set; } = 0;
        public string ScheduleValue { get; set; } = "12";
        public Dictionary<string, DateTime> SentHistory { get; set; } = new Dictionary<string, DateTime>();
    }

    public partial class MainWindow : System.Windows.Window
    {
        private string EncryptString(string plainText)
        {
            byte[] data = Encoding.UTF8.GetBytes(plainText);
            byte[] encryptedData = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(encryptedData);
        }

        private string DecryptString(string encryptedText)
        {
            try
            {
                byte[] data = Convert.FromBase64String(encryptedText);
                byte[] decryptedData = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decryptedData);
            }
            catch { return ""; }
        }

        private string selectedFilePath = string.Empty;
        private readonly string configFileName = "config.json";
        private readonly string logFileName = "log.txt";
        private System.Timers.Timer? autoCheckTimer;
        private AppConfig currentConfig = new AppConfig();
        private DateTime nextPlannedCheckTime;
        private bool isCheckingNow = false;

        private System.Windows.Forms.NotifyIcon? notifyIcon;

        public MainWindow()
        {
            InitializeComponent();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            LoadConfiguration();
            InitializeAutoCheckTimer();
            InitializeNotifyIcon();
            Log("Приложение успешно запущено.");
        }

        #region Логика Трэя (Сворачивание рядом с часами)

        private void InitializeNotifyIcon()
        {
            notifyIcon = new System.Windows.Forms.NotifyIcon();
            var iconStream = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/ICO.ico")).Stream;
            notifyIcon.Icon = new System.Drawing.Icon(iconStream);
            notifyIcon.Text = "Напоминалка продлений ЭЦП";

            notifyIcon.DoubleClick += (s, args) => RestoreWindow();

            var contextMenu = new System.Windows.Forms.ContextMenuStrip();
            contextMenu.Items.Add("Развернуть", null, (s, args) => RestoreWindow());
            contextMenu.Items.Add("Выход", null, (s, args) => ExitApplication());
            notifyIcon.ContextMenuStrip = contextMenu;

            notifyIcon.Visible = true;
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == WindowState.Minimized)
            {
                this.Visibility = Visibility.Hidden;
                notifyIcon?.ShowBalloonTip(2000, "Напоминалка ЭЦП", "Приложение свернуто в трей и продолжает работу.", System.Windows.Forms.ToolTipIcon.Info);
            }
            base.OnStateChanged(e);
        }

        private void RestoreWindow()
        {
            this.Visibility = Visibility.Visible;
            this.WindowState = WindowState.Normal;
            this.Activate();
        }

        private void ExitApplication()
        {
            if (notifyIcon != null)
            {
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
            }
            System.Windows.Application.Current.Shutdown();
        }

        protected override void OnClosed(EventArgs e)
        {
            if (notifyIcon != null)
            {
                notifyIcon.Visible = false;
                notifyIcon.Dispose();
            }
            base.OnClosed(e);
        }

        #endregion

        private void Log(string message)
        {
            string logLine = $"[{DateTime.Now:dd.MM.yyyy HH:mm:ss}] {message}";

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => Log(message)));
                return;
            }

            LogTextBox.AppendText(logLine + Environment.NewLine);
            LogTextBox.ScrollToEnd();

            try
            {
                File.AppendAllText(logFileName, logLine + Environment.NewLine);
            }
            catch { /* Игнорируем залоченность файла лога */ }
        }

        private void LoadConfiguration()
        {
            try
            {
                if (File.Exists(configFileName))
                {
                    string json = File.ReadAllText(configFileName);
                    var config = JsonConvert.DeserializeObject<AppConfig>(json);

                    if (config != null)
                    {
                        currentConfig = config;
                        SmtpServerTextBox.Text = currentConfig.SmtpServer;
                        SmtpPortTextBox.Text = currentConfig.SmtpPort.ToString();
                        SenderEmailTextBox.Text = currentConfig.SenderEmail;
                        TargetEmailTextBox.Text = currentConfig.TargetEmail;
                        MessageTemplateTextBox.Text = currentConfig.MessageTemplate;
                        AppPasswordBox.Password = DecryptString(currentConfig.EncryptedPassword);
                        AutostartCheckBox.IsChecked = currentConfig.IsAutostartEnabled;

                        ScheduleModeComboBox.SelectedIndex = currentConfig.ScheduleMode;
                        ScheduleValueTextBox.Text = currentConfig.ScheduleValue;

                        if (currentConfig.SentHistory == null)
                            currentConfig.SentHistory = new Dictionary<string, DateTime>();

                        if (!string.IsNullOrEmpty(currentConfig.ExcelFilePath) && File.Exists(currentConfig.ExcelFilePath))
                        {
                            selectedFilePath = currentConfig.ExcelFilePath;
                            FilePathTextBlock.Text = $"Выбран файл: {Path.GetFileName(selectedFilePath)}";
                        }
                    }
                }
                else
                {
                    ScheduleModeComboBox.SelectedIndex = 0;
                    ScheduleValueTextBox.Text = "12";
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка загрузки настроек: {ex.Message}", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void SaveConfigButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int.TryParse(SmtpPortTextBox.Text.Trim(), out int port);
                int mode = ScheduleModeComboBox.SelectedIndex;
                string val = ScheduleValueTextBox.Text.Trim();

                if (mode == 0)
                {
                    if (!int.TryParse(val, out int hours) || hours <= 0)
                    {
                        System.Windows.MessageBox.Show("Для интервала введите целое число часов больше 0!", "Ошибка валидации", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else
                {
                    if (!TimeSpan.TryParse(val, out TimeSpan parsedTime) || parsedTime.Days > 0)
                    {
                        System.Windows.MessageBox.Show("Введите корректное время в формате ЧЧ:ММ!", "Ошибка валидации", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }

                currentConfig.SmtpServer = SmtpServerTextBox.Text.Trim();
                currentConfig.SmtpPort = port == 0 ? 587 : port;
                currentConfig.SenderEmail = SenderEmailTextBox.Text.Trim();
                currentConfig.TargetEmail = TargetEmailTextBox.Text.Trim();
                currentConfig.MessageTemplate = MessageTemplateTextBox.Text;
                currentConfig.EncryptedPassword = EncryptString(AppPasswordBox.Password);
                currentConfig.ExcelFilePath = selectedFilePath;
                currentConfig.IsAutostartEnabled = AutostartCheckBox.IsChecked ?? false;
                currentConfig.ScheduleMode = mode;
                currentConfig.ScheduleValue = val;

                SetWindowsAutostart(currentConfig.IsAutostartEnabled);
                SaveConfigToDisk();

                System.Windows.MessageBox.Show("Настройки успешно сохранены!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                Log("Конфигурация расписания обновлена. Пересчет таймера...");

                InitializeAutoCheckTimer();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка сохранения: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveConfigToDisk()
        {
            try
            {
                string json = JsonConvert.SerializeObject(currentConfig, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(configFileName, json);
            }
            catch (Exception ex)
            {
                Log($"❌ Не удалось записать файл конфигурации: {ex.Message}");
            }
        }

        private void SetWindowsAutostart(bool enable)
        {
            try
            {
                string runKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(runKey, true))
                {
                    if (key != null)
                    {
                        string appName = "SuperNapominalkaApp";
                        if (enable)
                        {
                            string exePath = Environment.ProcessPath ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
                            key.SetValue(appName, $"\"{exePath}\"");
                            Log("Программа добавлена в автозагрузку Windows.");
                        }
                        else
                        {
                            if (key.GetValue(appName) != null)
                            {
                                key.DeleteValue(appName);
                                Log("Программа удалена из автозагрузки Windows.");
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log($"⚠️ Ошибка изменения параметров автозагрузки: {ex.Message}"); }
        }

        private void InitializeAutoCheckTimer()
        {
            if (autoCheckTimer != null)
            {
                autoCheckTimer.Stop();
                autoCheckTimer.Elapsed -= OnAutoCheckTimerElapsed;
                autoCheckTimer.Dispose();
                autoCheckTimer = null;
            }

            double intervalMs = CalculateNextInterval(out DateTime nextTime);
            nextPlannedCheckTime = nextTime;

            autoCheckTimer = new System.Timers.Timer(intervalMs);
            autoCheckTimer.Elapsed += OnAutoCheckTimerElapsed;
            autoCheckTimer.AutoReset = (currentConfig.ScheduleMode == 0);
            autoCheckTimer.Start();

            UpdateTimerStatusText();
        }

        private double CalculateNextInterval(out DateTime nextCheck)
        {
            DateTime now = DateTime.Now;

            if (currentConfig.ScheduleMode == 0)
            {
                int.TryParse(currentConfig.ScheduleValue, out int hours);
                if (hours <= 0) hours = 12;

                nextCheck = now.AddHours(hours);
                return hours * 60 * 60 * 1000;
            }
            else
            {
                TimeSpan.TryParse(currentConfig.ScheduleValue, out TimeSpan targetTime);
                DateTime targetDateTime = DateTime.Today.Add(targetTime);

                if (targetDateTime <= now)
                {
                    targetDateTime = targetDateTime.AddDays(1);
                }

                nextCheck = targetDateTime;
                return (targetDateTime - now).TotalMilliseconds;
            }
        }

        private void OnAutoCheckTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    await System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        Log("Срабатывание таймера по расписанию. Начинаю скан базы...")));

                    await ExecuteCheckAndSendAsync(isManual: false);

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (currentConfig.ScheduleMode == 1)
                        {
                            InitializeAutoCheckTimer();
                        }
                        else
                        {
                            nextPlannedCheckTime = DateTime.Now.AddHours(int.Parse(currentConfig.ScheduleValue));
                            UpdateTimerStatusText();
                        }
                    });
                }
                catch (Exception ex)
                {
                    Log($"❌ Ошибка таймера: {ex.Message}");
                }
            });
        }

        private void UpdateTimerStatusText()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(UpdateTimerStatusText));
                return;
            }
            StatusTextBlock.Text = $"Таймер активен. Следующий автоматический скан запланирован на: {nextPlannedCheckTime:dd.MM.yyyy HH:mm:ss}";
        }

        private void ScheduleModeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ScheduleValueLabel == null || ScheduleValueTextBox == null) return;

            if (ScheduleModeComboBox.SelectedIndex == 0)
            {
                ScheduleValueLabel.Text = "Интервал проверки (в часах):";
                if (string.IsNullOrEmpty(ScheduleValueTextBox.Text) || ScheduleValueTextBox.Text.Contains(":"))
                    ScheduleValueTextBox.Text = "12";
            }
            else
            {
                ScheduleValueLabel.Text = "Время ежедневной проверки (ЧЧ:ММ):";
                if (string.IsNullOrEmpty(ScheduleValueTextBox.Text) || !ScheduleValueTextBox.Text.Contains(":"))
                    ScheduleValueTextBox.Text = "09:00";
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Документы Excel (*.xlsx;*.xls)|*.xlsx;*.xls";
            if (openFileDialog.ShowDialog() == true)
            {
                selectedFilePath = openFileDialog.FileName;
                FilePathTextBlock.Text = $"Выбран файл: {Path.GetFileName(selectedFilePath)}";
                Log($"Выбран новый файл базы данных: {Path.GetFileName(selectedFilePath)}");
            }
        }

        private async void CheckButton_Click(object sender, RoutedEventArgs e)
        {
            Log("Ручной запуск проверки базы данных...");
            CheckButton.IsEnabled = false;

            await ExecuteCheckAndSendAsync(isManual: true);

            CheckButton.IsEnabled = true;
        }

        private async Task ExecuteCheckAndSendAsync(bool isManual)
        {
            if (isCheckingNow) return;

            if (string.IsNullOrEmpty(selectedFilePath) || !File.Exists(selectedFilePath))
            {
                if (isManual) System.Windows.MessageBox.Show("Файл Excel не выбран или отсутствует!", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                Log("Проверка отменена: файл базы данных не задан или отсутствует.");
                return;
            }

            string smtpServer = "";
            string senderEmail = "";
            string appPassword = "";
            string targetEmail = "";
            string template = "";
            string portText = "";

            Dispatcher.Invoke(() =>
            {
                smtpServer = SmtpServerTextBox.Text.Trim();
                senderEmail = SenderEmailTextBox.Text.Trim();
                appPassword = AppPasswordBox.Password;
                targetEmail = TargetEmailTextBox.Text.Trim();
                template = MessageTemplateTextBox.Text;
                portText = SmtpPortTextBox.Text.Trim();
            });

            if (!int.TryParse(portText, out int smtpPort) ||
                string.IsNullOrEmpty(smtpServer) || string.IsNullOrEmpty(senderEmail) ||
                string.IsNullOrEmpty(appPassword) || string.IsNullOrEmpty(template))
            {
                if (isManual) System.Windows.MessageBox.Show("Заполните настройки подключения перед выполнением проверки!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Log("Проверка отменена: не все поля подключения заполнены.");
                return;
            }

            try
            {
                isCheckingNow = true;
                int sentCount = 0;
                int skippedByTimerCount = 0;

                DataTable? dt = await Task.Run<DataTable?>(() =>
                {
                    using (var stream = File.Open(selectedFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });

                            if (result.Tables.Count > 0) return result.Tables[0];
                        }
                    }
                    return null;
                });

                if (dt == null)
                {
                    if (isManual) System.Windows.MessageBox.Show("В выбранном файле Excel отсутствуют таблицы!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    Log("Ошибка: Файл Excel пуст.");
                    return;
                }

                // Изменено: Добавлена проверка на наличие колонки "Email"
                if (!dt.Columns.Contains("ФИО") || !dt.Columns.Contains("Дата") || !dt.Columns.Contains("Email"))
                {
                    if (isManual) System.Windows.MessageBox.Show("Критическая ошибка: Столбцы 'ФИО', 'Дата' или 'Email' не найдены!", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    Log("Критическая ошибка: В таблице Excel не найдены колонки 'ФИО', 'Дата' или 'Email'.");
                    return;
                }

                foreach (DataRow row in dt.Rows)
                {
                    if (row["ФИО"] == DBNull.Value || row["Дата"] == DBNull.Value) continue;

                    string fio = row["ФИО"]?.ToString()?.Trim() ?? string.Empty;
                    string rawDate = row["Дата"]?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrEmpty(fio) || string.IsNullOrEmpty(rawDate)) continue;
                    if (!DateTime.TryParse(rawDate, out DateTime expirationDate)) continue;

                    DateTime today = DateTime.Today;
                    int daysUntilEvent = (expirationDate - today).Days;

                    if (daysUntilEvent == 7)
                    {
                        string historyKey = $"{fio}_{expirationDate:dd.MM.yyyy}";

                        if (currentConfig.SentHistory.ContainsKey(historyKey))
                        {
                            DateTime lastSentTime = currentConfig.SentHistory[historyKey];
                            TimeSpan timePassed = DateTime.Now - lastSentTime;

                            if (timePassed.TotalHours < 48)
                            {
                                skippedByTimerCount++;
                                Log($"Пропуск: Напоминание для {fio} уже отправлялось {lastSentTime:dd.MM HH:mm} (меньше 48 часов назад).");
                                continue;
                            }
                        }

                        // Изменено: Динамическое получение email из Excel
                        string recipientEmail = row["Email"]?.ToString()?.Trim() ?? string.Empty;

                        // Если email в строке пустой, используем адрес из настроек (как резервный)
                        if (string.IsNullOrEmpty(recipientEmail))
                        {
                            recipientEmail = targetEmail;
                        }

                        if (string.IsNullOrEmpty(recipientEmail) || !recipientEmail.Contains("@"))
                        {
                            Log($"⚠️ Пропуск: Для {fio} не найден корректный Email ни в Excel, ни в настройках.");
                            continue;
                        }

                        string mailText = template
                            .Replace("{ФИО}", fio)
                            .Replace("{Дата}", expirationDate.ToString("dd.MM.yyyy"));

                        try
                        {
                            // Изменено: Отправка идет на recipientEmail
                            await SendDynamicEmailAsync(smtpServer, smtpPort, senderEmail, appPassword, recipientEmail, "Напоминание о приближающемся событии", mailText);
                            currentConfig.SentHistory[historyKey] = DateTime.Now;
                            sentCount++;
                            Log($"📨 Письмо успешно отправлено для: {fio} на адрес {recipientEmail}.");
                        }
                        catch (Exception mailEx)
                        {
                            Log($"❌ Не удалось отправить письмо для {fio}: {mailEx.Message}");
                        }
                    }
                }

                if (sentCount > 0) SaveConfigToDisk();

                UpdateTimerStatusText();
                Log($"Сканирование завершено. Отправлено писем: {sentCount}. Пропущено фильтром 48ч: {skippedByTimerCount}.");

                if (isManual)
                {
                    System.Windows.MessageBox.Show($"Проверка успешно завершена.\nОтправлено писем: {sentCount}\nПропущено защитой: {skippedByTimerCount}", "Результат проверки", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                if (isManual) System.Windows.MessageBox.Show($"Ошибка выполнения проверки: \n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                else
                {
                    _ = Dispatcher.BeginInvoke(new Action(() => {
                        StatusTextBlock.Text = $"❌ Ошибка автоматического сканирования.";
                    }));
                }
                Log($"❌ Исключение при выполнении сканирования: {ex.Message}");
            }
            finally
            {
                isCheckingNow = false;
            }
        }

        private async Task SendDynamicEmailAsync(string server, int port, string fromEmail, string password, string toEmail, string subject, string body)
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(fromEmail, "Авто-Напоминалка");
                mail.To.Add(toEmail);
                mail.Subject = subject;
                mail.Body = body;

                using (SmtpClient smtp = new SmtpClient(server, port))
                {
                    smtp.Credentials = new NetworkCredential(fromEmail, password);
                    smtp.EnableSsl = true;
                    await smtp.SendMailAsync(mail);
                }
            }
        }
    }
}