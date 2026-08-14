using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Requests;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using ImageMagick;
using iTextSharp.text.pdf;
using Microsoft.VisualBasic.ApplicationServices;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using Org.BouncyCastle.Asn1.Pkcs;
using PdfiumViewer;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using Tesseract;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Color = System.Drawing.Color;
using Font = System.Drawing.Font;
using Label = System.Windows.Forms.Label;
using LicenseContext = System.ComponentModel.LicenseContext;
using Timer = System.Windows.Forms.Timer;

namespace EmailPDFMatchKeyword
{
    public partial class MainForm : Form
    {
        private GmailService service;
        private DriveService Driveservices;
        private string saveFolder;
        private TextBox txtResults;  // class-level variable
        private System.Threading.Timer pollTimer;
        private SheetsService _sheetsService;
        private static DateTime ProcessingStartDate
        {
            get
            {
                TimeZoneInfo indiaZone;
                try
                {
                    indiaZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                }
                catch
                {
                    indiaZone = TimeZoneInfo.Utc;
                }

                DateTime indiaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, indiaZone);
                return indiaNow.Date;
            }
        }
        private string _spreadsheetId = AppSettingsHelper.Get("GoogleDrive:SpreadsheetId");
        private CancellationTokenSource cancellationTokenSource;
        private GoogleSheetHelper _sheetHelper;
        private bool isPollingInProgress = false;
        private bool stopRequested = false;
        public SheetsService SheetsService => _sheetsService;
        public GmailService Service => service;
        Label lblLoading;
        ProgressBar progressBar;

        // New fields for file logging
        private string _logFilePath;
        private readonly object _logLock = new object();


        //private ExtractMethod _ExtractMethod;
        //public MainForm(ExtractMethod ExtractMethod)
        //{
        //  _ExtractMethod = ExtractMethod;
        //}

        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Text = "MainForm";
        }

        #endregion

        public void InitUI()
        {
            saveFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "InvoiceAttachments");
            Directory.CreateDirectory(saveFolder);

            // Ensure Logs folder exists and create initial log file
            try
            {
                var logsDir = Path.Combine(saveFolder, "Logs");
                Directory.CreateDirectory(logsDir);
                _logFilePath = Path.Combine(logsDir, $"Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.AppendAllText(_logFilePath, $"Log started at {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n");
            }
            catch
            {
                // ignore logging initialization failures
            }

			this.Icon = new Icon("Email_Logo.ico");

			// Start button
			Button btnStart = new Button { Text = "Start", Left = 10, Top = 10 };
			// Stop button
			Button btnStop = new Button { Text = "Stop", Left = 100, Top = 10 };
			// Clear button
			Button btnClear = new Button { Text = "Clear", Left = 200, Top = 10 };
			// Clear button
			Button btnfoldersetting = new Button { Text = "Folder path settings", Left = 300, Top = 10,Size = new Size(150, 23) };

			btnfoldersetting.Click += (s, e) =>
            {
                using (FrmChangePath changePathForm = new FrmChangePath())
                {
					changePathForm.StartPosition = FormStartPosition.CenterParent;
					changePathForm.ShowDialog(this);
                }
            };
			btnStart.Click += (s, e) => {
				StartPolling();
				btnStart.Enabled = false; // Disable Start button
				btnClear.Enabled = false; // Disable Clear button
				btnStop.Enabled = true;   // Enable Stop button
			};

			btnStop.Click += (s, e) =>
			{
				StopPolling();
				btnStop.Enabled = false; // Disable Stop button
				btnStart.Enabled = true; // Enable Start button
				btnClear.Enabled = true; // Enable Clear button
			};

			btnClear.Click += (s, e) => txtResults.Clear();

			// Add buttons to the form
			Controls.Add(btnStart);
			Controls.Add(btnStop);
			Controls.Add(btnClear);
			Controls.Add(btnfoldersetting);

			lblLoading = new Label
            {
                Text = "Processing...",
                Left = 400,
                Top = 14,
                AutoSize = true,
                ForeColor = Color.DarkRed,
                Visible = false
            };
            Controls.Add(lblLoading);

            progressBar = new ProgressBar
            {
                Left = 500,
                Top = 12,
                Width = 200,
                Style = ProgressBarStyle.Marquee,
                Visible = false
            };
            Controls.Add(progressBar);

            // Larger results box
            txtResults = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Left = 10,
                Top = 50,
                Width = 800,   // wider
                Height = 500,  // taller
                Font = new Font("Segoe UI", 13, FontStyle.Regular)
            };
            Controls.Add(txtResults);

            // Redirect Console/Trace/Debug output to the log file and UI
            try
            {
                SetupLoggingRedirects();
            }
            catch
            {
                // ignore redirect failures
            }
        }

        // Setup console/trace redirection to write into the same log file and UI
        private void SetupLoggingRedirects()
        {
            // Ensure _logFilePath exists
            if (string.IsNullOrEmpty(_logFilePath))
            {
                var logsDir = Path.Combine(saveFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "InvoiceAttachments", "Logs");
                Directory.CreateDirectory(logsDir);
                _logFilePath = Path.Combine(logsDir, $"Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                File.AppendAllText(_logFilePath, $"Log started at {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n");
            }

            var writer = new UiAndFileTextWriter(_logFilePath, txtResults, _logLock);

            // Redirect console output
            Console.SetOut(writer);
            Console.SetError(writer);

            // Add trace listeners
            System.Diagnostics.Trace.Listeners.Clear();
            System.Diagnostics.Trace.Listeners.Add(new System.Diagnostics.TextWriterTraceListener(writer));
            System.Diagnostics.Trace.AutoFlush = true;

            // Note: Debug.Listeners is not available in this target framework, Trace covers runtime logging.
        }

        // TextWriter that writes to the UI textbox and to a file (thread-safe)
        private class UiAndFileTextWriter : TextWriter
        {
            private readonly string _filePath;
            private readonly TextBox _ui;
            private readonly object _fileLock;

            public UiAndFileTextWriter(string filePath, TextBox ui, object fileLock)
            {
                _filePath = filePath;
                _ui = ui;
                _fileLock = fileLock ?? new object();
            }

            public override Encoding Encoding => Encoding.UTF8;

            private void WriteInternal(string value)
            {
                if (value == null) return;

                // Ensure every write ends with newline for file clarity
                if (!value.EndsWith(Environment.NewLine))
                    value = value + Environment.NewLine;

                string timestamped = $"{DateTime.Now:dd/MM/yyyy HH:mm:ss} - {value}";

                try
                {
                    lock (_fileLock)
                    {
                        File.AppendAllText(_filePath, timestamped, Encoding.UTF8);
                    }
                }
                catch
                {
                    // ignore file write failures
                }

                try
                {
                    if (_ui != null)
                    {
                        if (_ui.InvokeRequired)
                        {
                            _ui.BeginInvoke(new Action(() => _ui.AppendText(timestamped)));
                        }
                        else
                        {
                            _ui.AppendText(timestamped);
                        }
                    }
                }
                catch
                {
                    // ignore UI failures
                }
            }

            public override void Write(char value)
            {
                WriteInternal(value.ToString());
            }

            public override void Write(string value)
            {
                WriteInternal(value);
            }

            public override void WriteLine(string value)
            {
                WriteInternal(value + Environment.NewLine);
            }

            public override void WriteLine()
            {
                WriteInternal(Environment.NewLine);
            }
        }

        public async Task AuthenticateUserAsync()
        {
            try
            {
                using var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read);
                var secrets = GoogleClientSecrets.FromStream(stream).Secrets;

                var flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                {
                    ClientSecrets = secrets,
                    Scopes = new[]
                    { GmailService.Scope.GmailModify, GmailService.Scope.GmailSend, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets },
                    DataStore = new FileDataStore("token.json", true)
                });

                // ✅ Automatically runs a local web server and handles redirect
                var app = new AuthorizationCodeInstalledApp(flow, new LocalServerCodeReceiver());

                ICredential credential = null;

                try
                {
                    credential = await app.AuthorizeAsync("user", CancellationToken.None);
                }
                catch (TokenResponseException tre) when (tre.Error != null && (tre.Error.Error == "invalid_grant" || (tre.Error.ErrorDescription != null && tre.Error.ErrorDescription.Contains("invalid_grant"))))
                {
                    // Refresh token expired or revoked. Delete local token store and retry once.
                    Log("⚠️ Token refresh failed with invalid_grant. Deleting local token store and retrying authentication...");
                    try
                    {
                        DeleteLocalTokenStore();
                    }
                    catch (Exception ex)
                    {
                        Log($"⚠️ Failed to delete token store: {ex.Message}");
                    }

                    // Recreate the flow & app to ensure clean state
                    flow = new GoogleAuthorizationCodeFlow(new GoogleAuthorizationCodeFlow.Initializer
                    {
                        ClientSecrets = secrets,
                        Scopes = new[]
                        { GmailService.Scope.GmailModify, GmailService.Scope.GmailSend, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets },
                        DataStore = new FileDataStore("token.json", true)
                    });

                    var appRetry = new AuthorizationCodeInstalledApp(flow, new LocalServerCodeReceiver());
                    credential = await appRetry.AuthorizeAsync("user", CancellationToken.None);
                }

                // 6️⃣ Initialize Gmail service
                service = new GmailService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Email Attachment Reader"
                });

                // 7️⃣ Initialize Drive service
                Driveservices = new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "My Gmail + Drive App"
                });

                // 8️⃣ Initialize Sheets service
                _sheetsService = new SheetsService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Peer List Automation"
                });

                _sheetHelper = new GoogleSheetHelper(_sheetsService, _spreadsheetId);

                Log("✅ User authenticated successfully with offline access and consent prompt!");
            }
            catch (Exception ex)
            {
                Log($"❌ Authentication failed: {ex.Message}");
                Console.WriteLine($"❌ Authentication failed: {ex.Message}");
            }
        }

        private void DeleteLocalTokenStore()
        {
            try
            {
                // FileDataStore created with folder name "token.json" in app base directory
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var tokenDir = Path.Combine(baseDir, "token.json");
                if (Directory.Exists(tokenDir))
                {
                    Directory.Delete(tokenDir, true);
                    Log($"Deleted token store at: {tokenDir}");
                }

                // Also try current working directory
                var cwdToken = Path.Combine(Directory.GetCurrentDirectory(), "token.json");
                if (Directory.Exists(cwdToken) && !string.Equals(cwdToken, tokenDir, StringComparison.OrdinalIgnoreCase))
                {
                    Directory.Delete(cwdToken, true);
                    Log($"Deleted token store at: {cwdToken}");
                }

                // Some environments store tokens in user profile - try common fallback
                var userTokenPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), ".credentials", "token.json");
                if (Directory.Exists(userTokenPath))
                {
                    Directory.Delete(userTokenPath, true);
                    Log($"Deleted token store at: {userTokenPath}");
                }
            }
            catch (Exception ex)
            {
                // rethrow to let caller log
                throw new Exception("Failed to delete token store", ex);
            }
        }


		public async void StartPolling()
		{
			try
			{
				_spreadsheetId = AppSettingsHelper.Get("GoogleDrive:SpreadsheetId");
				ShowLoader();

				// Create log file (same as your code – unchanged)
				try
				{
					var logsDir = Path.Combine(saveFolder ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "InvoiceAttachments", "Logs");
					if (!string.IsNullOrEmpty(saveFolder)) logsDir = Path.Combine(saveFolder, "Logs");
					Directory.CreateDirectory(logsDir);
					_logFilePath = Path.Combine(logsDir, $"StartPolling_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
					File.AppendAllText(_logFilePath, $"StartPolling invoked at {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n");
				}
				catch { }

				if (cancellationTokenSource != null)
				{
					Log("Polling is already running.");
					return;
				}

				// New CTS
				cancellationTokenSource = new CancellationTokenSource();
				var token = cancellationTokenSource.Token;

				// Reset flags
				stopRequested = false;
				isPollingInProgress = false;

				// Read interval from config
				int intervalMinutes = 12;
				try
				{
					string intervalStr = AppSettingsHelper.Get("PollingIntervalMinutes");
					if (!string.IsNullOrEmpty(intervalStr) && int.TryParse(intervalStr, out int configured))
						intervalMinutes = configured;
				}
				catch { }

				int intervalMs = Math.Max(1, intervalMinutes) * 60 * 1000;

				// First immediate poll
				Log("Starting initial mailbox check (processing queued messages). This may take some time...");
				isPollingInProgress = true;
				try
				{
					await PollMailboxAsync(token);
					Log("Initial mailbox check completed.");
				}
				catch (Exception ex)
				{
					Log($"Error during initial mailbox check: {ex.Message}");
				}
				finally
				{
					isPollingInProgress = false;
				}

				// ⛔ If user pressed Stop while initial poll was running, do NOT start timer
				if (stopRequested)
				{
					Log("Stop requested during initial poll. Polling will not continue.");
					CleanupPolling();
					return;
				}

				// Background timer for periodic checks
				pollTimer = new System.Threading.Timer(async _ =>
				{
					if (stopRequested || isPollingInProgress)
						return;

					try
					{
						isPollingInProgress = true;
						Log("Scheduled poll: checking mailbox...");
						await PollMailboxAsync(token);
					}
					catch (Exception ex)
					{
						Log($"❌ Error during scheduled polling: {ex.Message}");
					}
					finally
					{
						isPollingInProgress = false;
						if (stopRequested)
						{
							Log("🟥 Stop requested — ending polling after this cycle.");
							CleanupPolling();
						}
					}
				}, null, intervalMs, intervalMs);

				Log($"✅ Polling started. Next checks every {intervalMinutes} minutes.");
			}
			catch (Exception ex)
			{
				Log($"Unexpected error starting polling: {ex.Message}");
			}
			finally
			{
				HideLoader();
			}
		}
		public void CleanupPolling()
        {
            try
            {
                if (pollTimer != null)
                {
                    pollTimer.Dispose();
                    pollTimer = null;
                }

                if (cancellationTokenSource != null)
                {
                    cancellationTokenSource.Dispose();
                    cancellationTokenSource = null;
                }

                isPollingInProgress = false;
                stopRequested = false;

                Log("⛔ Polling fully stopped. No further mailbox checks will run.");
            }
            catch (Exception ex)
            {
                Log($"Cleanup error: {ex.Message}");
            }
        }

		public void StopPolling()
		{
			try
			{
				Log("🛑 Stop requested — waiting for current process to finish.");

				// Tell system: do not start any NEW cycles
				stopRequested = true;

				// Stop timer so no new scheduled PollMailboxAsync runs will start
				if (pollTimer != null)
				{
					pollTimer.Dispose();
					pollTimer = null;
				}

				// ❌ Do NOT cancel token here – we want current PollMailboxAsync to finish gracefully
				// if (cancellationTokenSource != null) { cancellationTokenSource.Cancel(); ... }

				// If nothing is running at this moment, clean up right away
				if (!isPollingInProgress)
				{
					CleanupPolling();
				}

				Log("✅ Stop signal sent. Current email (if any) will finish before exit.");
			}
			catch (Exception ex)
			{
				Log($"Error stopping polling: {ex.Message}");
			}
		}

		private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> apiCall, int maxRetries = 5)
        {
            int retryCount = 0;
            while (true)
            {
                try
                {
                    return await apiCall();
                }
                catch (Google.GoogleApiException ex) when (
                    ex.HttpStatusCode == System.Net.HttpStatusCode.TooManyRequests || (ex.Error?.Errors?.Any(e => e.Reason == "userRateLimitExceeded" || e.Reason == "rateLimitExceeded") ?? false))
                {
                    retryCount++;

                    // Try to read retry-after time if Google included it in error message
                    int delayMs = (int)Math.Pow(2, retryCount) * 1000; // default exponential backoff

                    Log($"⚠️ Gmail rate limit hit ({ex.Error?.Errors?.FirstOrDefault()?.Reason ?? "429"}). " +
                        $"Retry #{retryCount} after {delayMs / 1000.0:F1}s...");

                    await Task.Delay(delayMs);

                    if (retryCount >= maxRetries)
                        throw; // stop retrying if max exceeded
                }
                catch (Exception ex)
                {
                    Log($"❌ Unexpected Gmail API error: {ex.Message}");
                    throw;
                }
            }
        }


		public async Task PollMailboxAsync(CancellationToken cancellationToken)
		{
			try
			{
				// First attempt
				await PollMailboxCoreAsync(cancellationToken);
			}
			catch (TokenResponseException tre) when (tre.Error != null && (tre.Error.Error == "invalid_grant" || (tre.Error.ErrorDescription != null && tre.Error.ErrorDescription.Contains("expired or revoked", StringComparison.OrdinalIgnoreCase))))
			{
				Log("⚠️ Gmail token has been expired or revoked while checking mailbox. Clearing token and re-authenticating...");

				// 1) Clear old token
				DeleteLocalTokenStore();

				// 2) Reset services
				service = null;
				Driveservices = null;
				_sheetsService = null;

				// 3) Re-authenticate
				await AuthenticateUserAsync();

				// 4) Retry ONCE
				try
				{
					Log("🔁 Retrying mailbox check after re-authentication...");
					await PollMailboxCoreAsync(cancellationToken);
				}
				catch (Exception exRetry)
				{
					Log($"❌ Mailbox check failed again after re-authentication: {exRetry.Message}");
				}
			}
			catch (Exception ex)
			{
				Log($"Error checking mailbox: {ex.Message}");
			}
		}



		public async Task PollMailboxCoreAsync(CancellationToken cancellationToken)
        {
            if (service == null || service.HttpClientInitializer is not IConfigurableHttpClientInitializer credential || service.HttpClientInitializer == null)
            {
                await AuthenticateUserAsync();
            }

            Log("Checking mailbox...");

            try
            {
                var labelsResponse = await service.Users.Labels.Get("me", "INBOX").ExecuteAsync(cancellationToken);
                int labelUnread = labelsResponse.MessagesUnread ?? 0;

                Log($"Inbox unread count: {labelUnread}");

                TimeZoneInfo indiaZone;
                try
                {
                    indiaZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                }
                catch (Exception tzEx)
                {
                    Log($"⚠️ Failed to find India Standard Time zone: {tzEx.Message}. Using UTC fallback.");
                    indiaZone = TimeZoneInfo.Utc;
                }
                DateTime indiaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, indiaZone);
                var windowStartLocal = indiaNow.Date.AddDays(-2); 
                var windowEndLocalExclusive = indiaNow.Date;  

                var procStartIst = DateTime.SpecifyKind(windowStartLocal, DateTimeKind.Unspecified);
                var procStartUtc = TimeZoneInfo.ConvertTimeToUtc(procStartIst, indiaZone);
                long epochWindowStart = new DateTimeOffset(procStartUtc).ToUnixTimeSeconds();

                var procEndIst = DateTime.SpecifyKind(windowEndLocalExclusive, DateTimeKind.Unspecified);
                var procEndUtc = TimeZoneInfo.ConvertTimeToUtc(procEndIst, indiaZone);
                long epochWindowEnd = new DateTimeOffset(procEndUtc).ToUnixTimeSeconds();

                long gmailAfter = epochWindowStart > 0 ? epochWindowStart - 1 : 0;
                Log($"Current IST Date: {indiaNow:yyyy-MM-dd}");
               var displayStart = windowStartLocal; 
                var displayEndInclusive = windowEndLocalExclusive.AddTicks(-1);
                Log($"Window Start IST (inclusive): {displayStart:dd/MM/yyyy HH:mm:ss}");
                Log($"Window End IST (exclusive): {windowEndLocalExclusive:dd/MM/yyyy HH:mm:ss}  (includes up to {displayEndInclusive:dd/MM/yyyy HH:mm:ss.fff})");
                Log($"Gmail epoch (after) used: {gmailAfter} (UTC seconds) → converted from {procStartIst:yyyy-MM-dd HH:mm:ss} IST");
                Log($"Gmail epoch (before) used: {epochWindowEnd} (UTC seconds) → converted from {procEndIst:yyyy-MM-dd HH:mm:ss} IST");
                Log($"🔎 Fetching unread threads covering IST window: {windowStartLocal:dd/MM/yyyy} (inclusive) → {windowEndLocalExclusive:dd/MM/yyyy} (exclusive). Gmail after:{gmailAfter} before:{epochWindowEnd}");

                var request = service.Users.Threads.List("me");
                request.Q = $"in:inbox after:{gmailAfter} before:{epochWindowEnd} "  + 
                    "(filename:pdf OR filename:doc OR filename:docx)";

				request.IncludeSpamTrash = false;

				var allThreads = new List<Google.Apis.Gmail.v1.Data.Thread>();
				string pageToken = null;


				do
				{
					request.PageToken = pageToken;

					var response = await ExecuteWithRetryAsync(() =>
						request.ExecuteAsync(cancellationToken)
					);

					if (response?.Threads != null)
						allThreads.AddRange(response.Threads);

					pageToken = response?.NextPageToken;
				}
				while (!string.IsNullOrEmpty(pageToken));

                if (allThreads.Count == 0)
                {
                    Log("No new unread threads found since ProcessingStartDate.");
                    return;
                }

                var fifoMessages = allThreads.AsEnumerable().Reverse().ToList();

                Log($"📨 Loaded {fifoMessages.Count} unread threads for processing since {windowStartLocal:dd/MM/yyyy} IST.");

				int processedEmails = 0;


				foreach (var msgItem in fifoMessages)
                {
					if (stopRequested && processedEmails > 0)
					{
						Log("🟥 Stop requested — finishing after current email. No further emails will be processed in this cycle.");
						break;
					}
					try
                    {
                        ShowLoader();
						// 2) Load full thread: ALL messages in this conversation
						var fullThread = await service.Users.Threads
							.Get("me", msgItem.Id)
							.ExecuteAsync(cancellationToken);

						if (fullThread?.Messages == null || fullThread.Messages.Count == 0)
						{
							Log("======================================================");
							Log($"⚠ Skipping thread {msgItem.Id} because it has no messages.");
							Log("======================================================");
							continue;
						}

						// SCENARIO 1: inspect ALL messages in the thread

                        // Thread must contain at least 2 valid attachments (pdf/doc/docx) in any message of the thread
                        bool threadHasAtLeastTwoUnreadValidAttachments = fullThread.Messages.Any(m =>
                                m.Payload?.Parts != null &&
                                m.Payload.Parts.Count(p =>
                                    !string.IsNullOrEmpty(p.Filename) &&
                                    (
                                        p.Filename.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
                                        p.Filename.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
                                        p.Filename.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                                    )
                                ) >= 2
                            );

						if (!threadHasAtLeastTwoUnreadValidAttachments)
						{
							Log($"⏩ Thread {msgItem.Id} skipped (less than 2 UNREAD PDF/DOC/DOCX attachments).");
							continue;
						}

						var messageInfos = fullThread.Messages
	                        .Select(m =>
	                        {
		                        var (utc, local) = GetMessageReceivedDate(m);
		                        return new
		                        {
			                        Message = m,
			                        Utc = utc,
			                        Local = local
		                        };
	                        })
	                        .ToList();

						var threadLocalDates = messageInfos.Select(i => i.Local).ToList();
						DateTime? threadOldestLocal = threadLocalDates.Count > 0 ? threadLocalDates.Min() : (DateTime?)null;
						DateTime? threadNewestLocal = threadLocalDates.Count > 0 ? threadLocalDates.Max() : (DateTime?)null;

						if (threadOldestLocal.HasValue && threadNewestLocal.HasValue)
						{
							Log($"📌 Thread {msgItem.Id} has {fullThread.Messages.Count} messages. " +
								$"Local date range: {threadOldestLocal:u} → {threadNewestLocal:u}");
						}
						else
						{
							Log($"📌 Thread {msgItem.Id} has {fullThread.Messages.Count} messages (no valid dates).");
						}

                        // SCENARIO 2: check if ANY unread message in thread is within configured business-date window
                        // Consider any message in the thread that falls in the date window (read or unread)
                        bool anyMessageInWindow = messageInfos.Any(i =>
                                i.Local >= windowStartLocal &&
                                i.Local < windowEndLocalExclusive);


						if (!anyMessageInWindow)
						{
							Log("======================================================");
                            Log($"⏩ Skipping thread {msgItem.Id} (no unread messages since configured start date).");
							Log("✔ Moving to next thread...");
							Log("======================================================");

							continue;
						}

						// Now we KNOW this thread has at least one message on those days.

                        // 3) Select unread messages that fall inside configured business-date window
                        var unreadMessagesInWindow = messageInfos
                            .Where(i =>
                                i.Message.LabelIds != null &&
                                i.Message.LabelIds.Contains("UNREAD") &&
                                i.Local >= windowStartLocal &&
                                i.Local < windowEndLocalExclusive)
                            .OrderBy(i => i.Local)
                            .ToList();

                        if (unreadMessagesInWindow.Count == 0)
                        {
                            Log("======================================================");
                            Log($"Thread {msgItem.Id} has no messages in the configured date window.");
                            Log("✔ Moving to next thread...");
                            Log("======================================================");
                            continue;
                        }

                        Log($"✅ Thread {msgItem.Id} has {unreadMessagesInWindow.Count} unread message(s) since configured start date.");


						foreach (var info in unreadMessagesInWindow)
                        {

							var message = info.Message;
                            var msgUtc = info.Utc; 
                            var msgLocal = info.Local; 

                            Log($"📅 Receive Date (UTC): {msgUtc:u}");
                            Log($"📅 Receive Date (IST): {msgLocal:yyyy-MM-dd HH:mm:ss} (IST)");
                            Log($"📌 Reference ProcessingStartDate (IST): {ProcessingStartDate:yyyy-MM-dd}");
                            Log($"📌 Allowed Receive Window (IST): {windowStartLocal:yyyy-MM-dd 00:00:00} inclusive → {windowEndLocalExclusive:yyyy-MM-dd 00:00:00} exclusive");
                            Log($"🔍 Processing message: subject + snippet...");
							Log($"🔍 Processing message: subject + snippet...");

							string subject = message.Payload?.Headers?
								.FirstOrDefault(h => h.Name.Equals("Subject", StringComparison.OrdinalIgnoreCase))
								?.Value ?? "NoSubject";

                            Log($"   Subject: {subject}");
                            Log($"   Snippet: {message.Snippet}");
                         

                            DateTime? subjectDate = null;
                            string subjectDateRaw = "NOT FOUND";

                            string subjectOnly = subject ?? "";

                            var subjectDateMatch = Regex.Match(
                                subjectOnly,
                                @"(?<!\d)\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}(?!\d)"
                            );

                            if (subjectDateMatch.Success)
                            {
                                subjectDateRaw = subjectDateMatch.Value;

                                string[] dateFormats =
                                {
        "M/d/yyyy",
        "MM/d/yyyy",
        "M/dd/yyyy",
        "MM/dd/yyyy",

        "M/d/yy",
        "MM/d/yy",
        "M/dd/yy",
        "MM/dd/yy",

        "M-d-yyyy",
        "MM-d-yyyy",
        "M-dd-yyyy",
        "MM-dd-yyyy",

        "M-d-yy",
        "MM-d-yy",
        "M-dd-yy",
        "MM-dd-yy"
    };

                                if (DateTime.TryParseExact(
                                    subjectDateRaw,
                                    dateFormats,
                                    CultureInfo.InvariantCulture,
                                    DateTimeStyles.None,
                                    out DateTime parsedSubjectDate))
                                {
                                    subjectDate = parsedSubjectDate.Date;

                                    Log($"✅ SUBJECT DATE FOUND: {subjectDate:MM/dd/yyyy}");
                                }
                            }

                            // ============================================================
                            // 2. SNIPPET DUE BACK DATE
                            // ONLY IF SUBJECT DATE WAS NOT FOUND
                            // ============================================================

                            if (!subjectDate.HasValue)
                            {
                                string snippetOnly = message.Snippet ?? "";

                                var dueDateMatch = Regex.Match(
                                    snippetOnly,
                                    @"(?:due\s+(?:back|in\s+hand))\s*:?\s*(?<date>\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})",
                                    RegexOptions.IgnoreCase);

                                if (dueDateMatch.Success)
                                {
                                    string dueDateRaw = dueDateMatch.Groups["date"].Value;

                                    string[] dueDateFormats =
                                    {
            "M/d/yyyy",
            "MM/d/yyyy",
            "M/dd/yyyy",
            "MM/dd/yyyy",

            "M/d/yy",
            "MM/d/yy",
            "M/dd/yy",
            "MM/dd/yy",

            "M-d-yyyy",
            "MM-d-yyyy",
            "M-dd-yyyy",
            "MM-dd-yyyy",

            "M-d-yy",
            "MM-d-yy",
            "M-dd-yy",
            "MM-dd-yy"
        };

                                    if (DateTime.TryParseExact(
                                        dueDateRaw,
                                        dueDateFormats,
                                        CultureInfo.InvariantCulture,
                                        DateTimeStyles.None,
                                        out DateTime parsedDueDate))
                                    {
                                        subjectDate = parsedDueDate.Date;
                                        subjectDateRaw = dueDateRaw;

                                        Log($"✅ DUE BACK DATE FOUND IN SNIPPET: {subjectDate:MM/dd/yyyy}");
                                    }
                                }
                            }

                            // ============================================================
                            // 3. FINAL TARGET SHEET DATE
                            // ============================================================

                            DateTime targetSheetDateForCall;
                            string reasonForSelection;

                            if (subjectDate.HasValue)
                            {
                                // Subject/Snippet date gets priority
                                targetSheetDateForCall = subjectDate.Value.Date;

                                reasonForSelection =
                                    "Subject/Due Back Date found; using that date.";
                            }
                            else
                            {
                                // No date found -> Receive Date
                                targetSheetDateForCall = msgLocal.Date;

                                reasonForSelection =
                                    "Subject/Due Back Date not found; using Receive Date.";
                            }

                            // ============================================================
                            // FINAL LOG
                            // ============================================================

                            Log($"📥 Extracted Date: {subjectDateRaw}");

                            Log(
                                $"📅 Subject/Due Back Date: " +
                                $"{(subjectDate.HasValue ? subjectDate.Value.ToString("MM/dd/yyyy") : "NOT FOUND")}"
                            );

                            Log(
                                $"📄 FINAL TARGET SHEET DATE: " +
                                $"{targetSheetDateForCall:MM/dd/yyyy}"
                            );

                            Log(
                                $"📄 FINAL TARGET SHEET NAME: " +
                                $"{targetSheetDateForCall:MM/dd}"
                            );

                            Log($"ℹ️ Reason: {reasonForSelection}");

                            // Your extracted data placeholders
                            string billCharges = "Not Found", billDate = "Not Found", geicoCharges = "Not Found", geicoDate = "Not Found", caseNumber = "Not Found", CLAIMANTNAME = "Not Found", PROVIDER = "Not Found", INCIDENTDATE = "Not Found", SCRIBETEAM = "Not Found";

							int medsToDocPageCount = 0;
							bool hasBillPdf = false, hasGeicopeerPdf = false;

							// Temporary in-memory storage for attachments
							List<(string FileName, byte[] Data)> attachments = new List<(string FileName, byte[] Data)>();

							var parts = message.Payload.Parts ?? new List<MessagePart>();


							var validAttachmentParts = parts
	                            .Where(p =>
		                            !string.IsNullOrEmpty(p.Filename) &&
		                            p.Body?.AttachmentId != null &&
		                            (
			                            p.Filename.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ||
			                            p.Filename.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) ||
			                            p.Filename.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
		                            )
	                            )
	                            .ToList();

							if (validAttachmentParts.Count < 2)
							{
								Log("⏩ Email skipped (less than 2 PDF/DOC/DOCX attachments).");
								continue; // 🔑 NO Attachments.Get API call
							}

							//foreach (var part in parts)
							//foreach (var part in message.Payload.Parts ?? new List<MessagePart>())
							foreach (var part in validAttachmentParts)
							{
								// Only real attachments: has filename + attachment id
								if (string.IsNullOrEmpty(part.Filename) || part.Body == null || string.IsNullOrEmpty(part.Body.AttachmentId))
								{
									continue;
								}

								string attachId = part.Body.AttachmentId;

								var attach = await service.Users.Messages.Attachments
									.Get("me", message.Id, attachId)
									.ExecuteAsync(cancellationToken);

								if (attach?.Data == null)
								{
									Log($"⚠ Attachment {part.Filename} has no data.");
									continue;
								}

								// Gmail returns Base64 URL-safe encoding
								string base64 = attach.Data.Replace('-', '+').Replace('_', '/');
								// Padding fix (length must be multiple of 4)
								if (base64.Length % 4 != 0)
								{
									base64 = base64.PadRight(base64.Length + (4 - base64.Length % 4), '=');
								}

								byte[] bytes = Convert.FromBase64String(base64);

								// Keep in-memory
								attachments.Add((part.Filename, bytes));

								// Save to temp folder if you need physical files
								string tempFilePath = Path.Combine(Path.GetTempPath(), part.Filename);
								await File.WriteAllBytesAsync(tempFilePath, bytes, cancellationToken);

								Log($"Processed attachment: {part.Filename}");
								Log($"Saved attachment: {tempFilePath}");
								try
								{
                                    string pathToUse = "";
										// fallback to Documents/InvoiceAttachments/Logs
										var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "EmailReadFileLog", "Logs");
										Directory.CreateDirectory(fallbackDir);
										pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
										File.AppendAllText(pathToUse, $"Saved attachment: {tempFilePath}");
								}
								catch
								{
									// ignore logging failures to file to avoid crashing the app
								}

								// BILL PDF
								if (Path.GetExtension(tempFilePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase) &&(Path.GetFileName(tempFilePath).ToLower().Contains("bill") || Path.GetFileName(tempFilePath).ToLower().Contains("bills")))
								{

									ShowLoader();
                                    Log("Bill to Peer PDF detected. Converting to images...");
                                    hasBillPdf = true;

                                    using (var pdfStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    {
                                        Log($"this is the pdf stream from File: {pdfStream}");
                                        var images = await _ExtractMethod.ConvertPdfToImages_2Async(pdfStream);
                                        //var images = _ExtractMethod.ConvertPdfToImages_2(pdfStream);
                                        Log($"Extract Images from PDF to Images: {images.Count}");

                                        int retryCount = 2;   // how many times to retry full scan
                                        int delayMs = 1000;   // wait time between retries (1 seconds)

                                        for (int attempt = 1; attempt <= retryCount; attempt++)
                                        {
                                            Log($"🔄 Attempt {attempt} to extract Bill Charges & Date...");

                                            //foreach (var image in images)
                                            for (int pageIndex = 0; pageIndex < images.Count; pageIndex++) // start from second page
                                            {
                                                if (pageIndex == 1) // PageIndex == 1 is the second page
                                                {
                                                    var image = images[pageIndex];
                                                    //var rows = _ExtractMethod.ExtractTableRowsFromImage(image);
                                                    var rows = await _ExtractMethod.ExtractTableRowsFromImageAsync(image);

                                                    if (billCharges == "Not Found")
                                                        billCharges = _ExtractMethod.ExtractCharges(rows);

                                                    if (billDate == "Not Found")
                                                        billDate = _ExtractMethod.ExtractDateOfService(rows);

                                                    if (billCharges != "Not Found" && billDate != "Not Found")
                                                    {
                                                        Log($"✅ The Bill Charges is : {billCharges}");
                                                        Log($"✅ The Bill Date is : {billDate}");
                                                        break; // stop scanning pages
                                                    }
                                                }
                                            }

                                            if (billCharges != "Not Found" && billDate != "Not Found")
                                            {
                                                break;
                                            }

                                            if (attempt < retryCount)
                                            {
                                                Log($"⚠️ Values not found yet, waiting {delayMs} ms before retry...");
                                                System.Threading.Thread.Sleep(delayMs);
                                            }
                                        }

                                        //if (billCharges == "Not Found" || billDate == "Not Found")
                                        //{
                                        //    Log("Trying to check with OpenAI API Process......");

                                        //    var maxPages = 3;
                                        //    var selectedImages = images.Take(maxPages).ToList();
                                        //    Log($"📄 Selected up to {maxPages} pages for OCR processing.");

                                        //    // Assuming your list to store OCR text
                                        //    var ocrResults = new List<string>();

                                        //    string openAiApiKey = AppSettingsHelper.Get("OpenAIAPIKey");

                                        //    //// Replace with your actual OpenAI API key
                                        //    //string openAiApiKey = "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";

                                        //    // HttpClient for reuse
                                        //    using var httpClient = new HttpClient();
                                        //    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", openAiApiKey);

                                        //    foreach (var APIimage in selectedImages)
                                        //    {
                                        //        Log("🖼️ Converting image to base64 for OpenAI API...");
                                        //        // Convert image to base64
                                        //        using var ms = new MemoryStream();
                                        //        APIimage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        //        string base64Image = Convert.ToBase64String(ms.ToArray());
                                        //        string dataUrl = $"data:image/png;base64,{base64Image}";

                                        //        // Build request payload
                                        //        var payload = new
                                        //        {
                                        //            model = "gpt-4o",
                                        //            messages = new[]
                                        //            {
                                        //                new {
                                        //                    role = "user",
                                        //                    content = new object[]
                                        //                    {
                                        //                        new { type = "text", text = "Extract all text from this image." },
                                        //                        new { type = "image_url", image_url = new { url = dataUrl } }
                                        //                    }
                                        //                    }
                                        //                },
                                        //            max_tokens = 2000
                                        //        };

                                        //        var jsonPayload = JsonSerializer.Serialize(payload);
                                        //        var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                                        //        Log("📤 Sending request to OpenAI API...");

                                        //        // Call OpenAI API
                                        //        var response = await httpClient.PostAsync("https://api.openai.com/v1/chat/completions", content);
                                        //        var responseString = await response.Content.ReadAsStringAsync();

                                        //        if (response.IsSuccessStatusCode)
                                        //        {
                                        //            Log("✅ OpenAI API call succeeded. Extracting response...");
                                        //            // Parse and extract message content
                                        //            using var doc = JsonDocument.Parse(responseString);
                                        //            var extractedText = doc.RootElement
                                        //                .GetProperty("choices")[0]
                                        //                .GetProperty("message")
                                        //                .GetProperty("content")
                                        //                .GetString();

                                        //            ocrResults.Add(extractedText);
                                        //            Log($"📝 Text extracted and added to OCR results.{ocrResults}");
                                        //        }
                                        //        else
                                        //        {
                                        //            Log($"❌ OpenAI API call failed with status: {response.StatusCode}");
                                        //            Console.WriteLine(responseString);
                                        //        }
                                        //    }

                                        //    var wrappedOcrResults = new List<List<string>> { ocrResults };
                                        //    //var rows = await _ExtractMethod.ExtractTableRowsFromImageAsync(image);

                                        //    Log($"This is the result of Text is :{wrappedOcrResults}");

                                        //    if (billCharges == "Not Found")
                                        //        billCharges = _ExtractMethod.ExtractChargesAPI(wrappedOcrResults);

                                        //    if (billDate == "Not Found")
                                        //        billDate = _ExtractMethod.ExtractDateOfServiceAPI(wrappedOcrResults);

                                        //    if (billCharges != "Not Found" && billDate != "Not Found")
                                        //    {
                                        //        Log($"✅ The Bill Charges is : {billCharges}");
                                        //        Log($"✅ The Bill Date is : {billDate}");
                                        //        break; // stop scanning pages
                                        //    }
                                        //    Log("❌ Could not find Bill Charges and/or Bill Date after all retries.");
                                        //}

                                        if (billCharges == "Not Found" || billDate == "Not Found")
                                        {
                                            Log("❌ Could not find Bill Charges and/or Bill Date after all retries.");
                                        }
                                    }
                                    HideLoader();
                                }

                                // Handle GEICOPEER PDF
                                if (Path.GetFileName(tempFilePath).Equals("Geicopeer.pdf", StringComparison.OrdinalIgnoreCase) || Path.GetFileName(tempFilePath).Equals("Allstate Peer.pdf", StringComparison.OrdinalIgnoreCase) || (tempFilePath).Equals("FilmMRRDrCvr.pdf", StringComparison.OrdinalIgnoreCase) || (tempFilePath).Equals("drcover.docx", StringComparison.OrdinalIgnoreCase))
                                {
                                    ShowLoader();
                                    Log("Geicopeer OR Allstate Peer PDF detected. Converting to images...");
                                    hasGeicopeerPdf = true;
                                    //using (var pdfStream = File.OpenRead(tempFilePath))
                                    using (var pdfStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                    {
                                        var images = await _ExtractMethod.ConvertPdfToImagesAsync(pdfStream);
                                        //var images = _ExtractMethod.ConvertPdfToImages(pdfStream);

                                        foreach (var image in images)
                                        {
                                            //var rows = _ExtractMethod.ExtractTableRowsFromImage(image);
                                            var rows = await _ExtractMethod.ExtractTableRowsFromImageAsync(image);

                                            var (_, date, charges) = _ExtractMethod.ExtractFromGeicoPeer(rows);

                                            if (caseNumber == "Not Found")
                                                caseNumber = _ExtractMethod.ExtractCaseNumber(rows);

                                            Log($"The Case Number is : {caseNumber}");

                                            if (CLAIMANTNAME == "Not Found")
                                                CLAIMANTNAME = _ExtractMethod.ExtractClientName(rows);

                                            Log($"The CLAIMANT NAME is : {CLAIMANTNAME}");

                                            if (PROVIDER == "Not Found")
                                                PROVIDER = _ExtractMethod.ExtractProvider(rows);

                                            Log($"The PROVIDER is : {PROVIDER}");

                                            if (INCIDENTDATE == "Not Found")
                                                INCIDENTDATE = _ExtractMethod.ExtractDateOfIncident(rows);

                                            Log($"The INCIDENT DATE is : {INCIDENTDATE}");

                                            if (date != "Not Found") geicoDate = date;

                                            Log($"The GEICO DATE is : {geicoDate}");

                                            if (charges != "Not Found") geicoCharges = charges;

                                            Log($"The GEICO Charges is : {geicoCharges}");


											if (string.IsNullOrEmpty(PROVIDER) || PROVIDER == "Not Found")
											{
												string extractedName = null;

												// 1. Try to extract the single name immediately after "Dr." from subject
												//    This captures only the first token after Dr. (letters, hyphen, apostrophe)
												var drRegex = new Regex(@"Dr\.?\s+([A-Z][a-zA-Z'\-]+)\b", RegexOptions.IgnoreCase);
												var subjectDrMatch = drRegex.Match(subject ?? "");

												if (subjectDrMatch.Success)
												{
													// Group 1 contains the single name immediately after "Dr."
													extractedName = subjectDrMatch.Groups[1].Value.Trim().TrimEnd('.', ',');
													Log($"✅ Found PROVIDER (Dr.) in subject: {extractedName}");
												}
												else
												{
													// 2. If no Dr., try to extract a full name pattern (First Last) from subject
													//    but we'll decide whether to keep first or last name — here we keep the first token or last as needed.
													var nameRegex = new Regex(@"\b([A-Z][a-zA-Z'\-]+)\s+([A-Z][a-zA-Z'\-]+)\b");
													var subjectNameMatch = nameRegex.Match(subject ?? "");

													if (subjectNameMatch.Success)
													{
														// Example: subject "Yen Areina" -> Groups[1]=Yen, Groups[2]=Areina
														// If you specifically want the **last name** continue using Groups[2]
														// If you want **first name**, use Groups[1]. Here I keep Groups[1] by default:
														extractedName = subjectNameMatch.Groups[1].Value.Trim().TrimEnd('.', ',');
														Log($"✅ Found PROVIDER first name in subject: {extractedName} (subject had 2-word match)");
													}
													else
													{
														// 3. Try the same extraction from body if not found in subject
														string emailBody = "";

														if (message.Payload?.Body?.Data != null)
														{
															try
															{
																var decodedData = message.Payload.Body.Data.Replace("-", "+").Replace("_", "/");
																var bodyBytes = Convert.FromBase64String(decodedData);
																emailBody = Encoding.UTF8.GetString(bodyBytes);
															}
															catch (Exception ex)
															{
																Log($"⚠️ Failed to decode body: {ex.Message}");
															}
														}

														if (string.IsNullOrWhiteSpace(emailBody))
															emailBody = message.Snippet ?? "";

														// Try Dr. pattern in body (single name after Dr.)
														var bodyDrMatch = drRegex.Match(emailBody);
														if (bodyDrMatch.Success)
														{
															extractedName = bodyDrMatch.Groups[1].Value.Trim().TrimEnd('.', ',');
															Log($"✅ Found PROVIDER (Dr.) in body: {extractedName}");
														}
														else
														{
															// Try full name pattern in body (two words)
															var bodyNameMatch = nameRegex.Match(emailBody);
															if (bodyNameMatch.Success)
															{
																// choose first name or last name depending on your need:
																extractedName = bodyNameMatch.Groups[1].Value.Trim().TrimEnd('.', ',');
																Log($"✅ Found PROVIDER first name in body: {extractedName} (body had 2-word match)");
															}
															else
															{
																Log("❌ PROVIDER not found in subject or body.");
															}
														}
													}
												}

												if (!string.IsNullOrEmpty(extractedName))
												{
													PROVIDER = extractedName;
												}
											}

											try
											{
                                                SCRIBETEAM = _ExtractMethod.GetFolderPrefixFromDrive(Driveservices, PROVIDER);
                                                Log($"First word from matched folder: {SCRIBETEAM}");
                                            }
                                            catch (Exception ex)
                                            {
                                                Log($"Error finding matching folder: {ex.Message}");
                                            }

                                            if (geicoDate != "Not Found" && geicoCharges != "Not Found" && caseNumber != "Not Found" && CLAIMANTNAME != "Not Found" && PROVIDER != "Not Found" && INCIDENTDATE != "Not Found" && SCRIBETEAM != "Not Found")
                                            {
                                                HideLoader();
                                                Log("✅ Successfully extracted all required data from Geicopeer PDF.");
                                                break; // ✅ This only breaks the *page loop*, not the attachments loop
                                            }
                                        }

									}
									HideLoader();

                                }

                                // Handle MedsToDoc PDFs
                                if (Path.GetFileName(tempFilePath).Replace("_", "").Replace(" ", "").ToLower().Contains("medstodoc") &&
                    Path.GetExtension(tempFilePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                                {
                                    ShowLoader();
                                    Log("MedsToDoc PDF detected. Counting pages...");
                                    try
                                    {
                                        using (var pdfStream = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                        {
                                            ShowLoader();
                                            medsToDocPageCount = _ExtractMethod.GetPdfPageCount_iTextSharp(pdfStream); // use PdfSharp version
                                            Log($"MedsToDoc page count: {medsToDocPageCount}");
                                            HideLoader();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Log($"Failed to count pages for {tempFilePath}: {ex.Message}");
                                        medsToDocPageCount = 0;
                                    }
                                    HideLoader();
                                }
                            }

                            if (!string.IsNullOrEmpty(PROVIDER) || PROVIDER != "Not Found")
                            {
                                ShowLoader();
                                // Final Comparison
                                string cleanBillCharges = NormalizeAmount(billCharges);
                                string cleanGeicoCharges = NormalizeAmount(geicoCharges);

                                // --- Helper Methods ---
                                string NormalizeAmount(string input)
                                {
                                    if (string.IsNullOrWhiteSpace(input))
                                        return "0";

                                    string cleaned = input.Replace("$", "").Replace(",", "").Trim();

                                    if (decimal.TryParse(cleaned, out decimal value))
                                    {
                                        // Round to 1 decimal place to handle .0 vs .00
                                        return value.ToString("0.0");
                                    }

                                    return cleaned;
                                }

                                // Clean and format dates: try to parse and convert to MM/dd/yyyy
                                string cleanBillDate = TryFormatDate(billDate);
                                string cleanGeicoDate = TryFormatDate(geicoDate);

                                // Final Comparison
                                string result =
                                    $"BILL PDF: Charges = {cleanBillCharges}, Date of Service = {cleanBillDate}\r\n" +
                                    $"GEICOPEER PDF: Charges = {cleanGeicoCharges}, Date of Service = {cleanGeicoDate}\r\n";

                                // Check if either charges or date is "Not Found"
                                bool chargesValid = cleanBillCharges != "Not Found" && cleanGeicoCharges != "Not Found";
                                bool dateValid = cleanBillDate != "Not Found" && cleanGeicoDate != "Not Found";

                                // Determine status based on comparison
                                string status = (chargesValid && dateValid && cleanBillCharges == cleanGeicoCharges && cleanBillDate == cleanGeicoDate)
                                    ? "Matched"
                                    : "Not Matched";

                                // Prepare detailed mismatch information
                                string mismatchReason = "";
                                if (status == "Not Matched")
                                {
                                    if (!chargesValid)
                                    {
                                        mismatchReason += "Charges do not match. ";
                                    }
                                    if (!dateValid)
                                    {
                                        mismatchReason += "Dates do not match. ";
                                    }
                                    else if (cleanBillCharges != cleanGeicoCharges)
                                    {
                                        mismatchReason += "Charges do not match. ";
                                    }
                                    else if (cleanBillDate != cleanGeicoDate)
                                    {
                                        mismatchReason += "Dates do not match. ";
                                    }
                                }


                        if (hasBillPdf && hasGeicopeerPdf)
                        {
                            // Try to infer vendor from email subject or attachment filenames.
                            // Default to null if not found so downstream logic behaves the same.
                            string extractedVendor = null;

                            try
                            {
                                if (!string.IsNullOrWhiteSpace(subject) && Regex.IsMatch(subject, "\\bISG\\b", RegexOptions.IgnoreCase))
                                {
                                    extractedVendor = "ISG";
                                }
                                else if (attachments != null && attachments.Any(a => !string.IsNullOrWhiteSpace(a.FileName) && Regex.IsMatch(a.FileName, "ISG", RegexOptions.IgnoreCase)))
                                {
                                    extractedVendor = "ISG";
                                }
                            }
                            catch
                            {
                                // If any unexpected error occurs while detecting vendor, leave as null.
                                extractedVendor = null;
                            }

                            Log($"Extracted VENDOR (pre-upload): {extractedVendor ?? "<null>"}");

                           try
                            {
                                if (string.Equals(extractedVendor?.Trim(), "ISG", StringComparison.OrdinalIgnoreCase))
                                {
                                    DateTime parsedDate;
                                    if (!string.IsNullOrWhiteSpace(billDate) &&
                                        !billDate.Equals("Not Found", StringComparison.OrdinalIgnoreCase) &&
                                        DateTime.TryParse(billDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    {
                                        targetSheetDateForCall = parsedDate.Date;
                                        Log($"ISG vendor: using extracted DATE (billDate) as target sheet date: {targetSheetDateForCall:MM/dd/yyyy}");
                                    }
                                    else if (!string.IsNullOrWhiteSpace(INCIDENTDATE) &&
                                             !INCIDENTDATE.Equals("Not Found", StringComparison.OrdinalIgnoreCase) &&
                                             DateTime.TryParse(INCIDENTDATE, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    {
                                        targetSheetDateForCall = parsedDate.Date;
                                        Log($"ISG vendor: using extracted DATE (INCIDENTDATE) as target sheet date: {targetSheetDateForCall:MM/dd/yyyy}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Log($"Error parsing extracted DATE for ISG vendor: {ex.Message}");
                            }

                            bool vendorPass = string.Equals(extractedVendor?.Trim(), "ISG", StringComparison.OrdinalIgnoreCase);

                            Log($"Vendor Validation: {(vendorPass ? "PASS" : "FAIL")}");

                            if (!vendorPass)
                            {
                                Log("PDF/Drive Processing: SKIPPED. Reason: VENDOR is not ISG or not found.");
                            }
                            else
                            {
                                await _ExtractMethod.ProcessAndUploadFilesAsync(
                                    msgUtc,
                                    caseNumber,
                                    CLAIMANTNAME,
                                    status,
                                    PROVIDER,
                                    attachments,
                                    Driveservices,
                                    targetSheetDateForCall);
                            }
                        }

                                // Compare only if both values are valid
                                if (status == "Matched" && hasBillPdf && hasGeicopeerPdf)
                                {
                                    bool sheetSuccess = await _ExtractMethod.InsertDataIntoSheetORDataBase(PROVIDER, caseNumber, CLAIMANTNAME, msgUtc, INCIDENTDATE, medsToDocPageCount, status, SCRIBETEAM, subject, targetSheetDateForCall);

                                    result += "Values MATCH";

                                    if (sheetSuccess)
                                    {
                                        await _ExtractMethod.MarkMessageAsReadAsync(msgItem.Id);
                                        Log(result);
                                        Log($"Values are Match Successfully & Email subject: {subject} Process Completed.");
                                    }
                                    else
                                    {
                                        Log($"❌ Google Sheets insertion failed for '{subject}'. Keeping email unread for retry.");
                                    }
                                }
                                else if (status == "Not Matched" && hasBillPdf && hasGeicopeerPdf)
                                {
                                    bool sheetSuccess = await _ExtractMethod.InsertDataIntoSheetORDataBase(PROVIDER, caseNumber, CLAIMANTNAME, msgUtc, INCIDENTDATE, medsToDocPageCount, status, SCRIBETEAM, subject, targetSheetDateForCall);
                                    result += "Values DO NOT MATCH. Reason: " + mismatchReason;

                                    // Prepare the email body
                                    //string emailBody = $@"
                                    //    <html>
                                    //    <body style='font-family:Segoe UI, sans-serif; color:#333;'>
                                    //        <p>Hello,</p>
                                    //        <p>
                                    //            This is to inform you that the email bearing subject :
                                    //            <strong>{subject}</strong> 
                                    //            doesn't match the required details. Please check the result printed in the system.
                                    //        </p>
                                    //        <p><strong>Reason:</strong> {mismatchReason}</p>
                                    //        <br/>
                                    //        <p><strong>Comparison Details:</strong></p>
                                    //        <pre>{result}</pre>
                                    //        <br/>
                                    //        <p>Thanks</p>
                                    //    </body>
                                    //    </html>";


                                    ////string ToEmail = AppSettingsHelper.Get("CalculateDataEmail");

                                    //var toList = AppSettingsHelper.Get("EmailTO")
                                    //.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                    //.Select(e => e.Trim());

                                    //var ccList = AppSettingsHelper.Get("EmailCC")
                                    //                ?.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                    //                .Select(e => e.Trim());

                                    //await _ExtractMethod.SendEmailAsync(toList, subject: "Required Details are not matched", emailBody, isHtml: true, ccList);
                                    Log($"Email {subject} Process will completed............");

                                    if (sheetSuccess)
                                    {
                                        await _ExtractMethod.MarkMessageAsReadAsync(msgItem.Id);
                                    }
                                    else
                                    {
                                        Log($"❌ Google Sheets insertion failed for '{subject}'. Keeping email unread for retry.");
                                    }

                                    Log(result);

                                    Log($"Values are Not Match Email subject: {subject} Process Completed.");
                                }

                                Log("======================================================");
                                Log($"Email :-: \"{subject}\" Process will completed............");
                                Log("======================================================");

								await Task.Delay(TimeSpan.FromSeconds(20), cancellationToken);
								Log($"Delay for some time to avoid API Traffic Issue second :- 20.");
							}
                            else
                            {
                                Log("======================================================");
                                Log($"Email :-: \"{subject}\" has not found the Dr.Name \"{PROVIDER}\" . Cannot proceed with this Email.");
                                Log("======================================================");

								await Task.Delay(TimeSpan.FromSeconds(20), cancellationToken);
								Log($"Delay for some time to avoid API Traffic Issue second :- 20.");
							}
							
							processedEmails++;
						}
					}
                    catch (Exception ex)
                    {
                        Log($"Error: {ex.Message}");
                    }
                    // Break out if we need to cancel processing the next message
                    if (cancellationToken.IsCancellationRequested)
                    {
                        Log("Polling canceled. Stopping email processing.");
                        break;
                    }
                    HideLoader();
					Log($"📬 Poll cycle finished. Processed {processedEmails} email(s).");
					 await Task.Delay(TimeSpan.FromSeconds(20), cancellationToken);
					Log($"Delay for some time to avoid API Traffic Issue second :- 20.");

				}
				Log("Mailbox polling completed.");
                HideLoader();
            }
            catch (Exception ex)
            {
                Log($"Error checking mailbox: {ex.Message}");
            }
        }

		private (DateTime Utc, DateTime Local) GetMessageReceivedDate(Google.Apis.Gmail.v1.Data.Message message)
		{
            // 1) Start from InternalDate (milliseconds since epoch)
            long epochMs = message.InternalDate ?? 0;
            var utc = DateTimeOffset.FromUnixTimeMilliseconds(epochMs).UtcDateTime;

            // Always convert to India Standard Time for Local (business) calculations
            TimeZoneInfo indiaZone;
            try
            {
                indiaZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
            }
            catch
            {
                indiaZone = TimeZoneInfo.Utc;
            }

            var localIst = TimeZoneInfo.ConvertTimeFromUtc(utc, indiaZone);

            // 2) Prefer the "Date" header if present and parseable — still interpret/convert to UTC then IST
            string rawDateHeader = message.Payload?.Headers?
                .FirstOrDefault(h => h.Name.Equals("Date", StringComparison.OrdinalIgnoreCase))
                ?.Value;

            if (!string.IsNullOrWhiteSpace(rawDateHeader) && DateTimeOffset.TryParse(rawDateHeader, out var hdr))
            {
                utc = hdr.UtcDateTime;
                localIst = TimeZoneInfo.ConvertTimeFromUtc(utc, indiaZone);
            }

            return (utc, localIst);
		}

		public void CopyTemplateSheet(string filePath, string newSheetName)
        {
            // EPPlus requires a license context
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                // Find the "template" sheet
                var templateSheet = package.Workbook.Worksheets["template"];
                if (templateSheet == null)
                {
                    throw new Exception("Template sheet not found in Excel file.");
                }

                // Check if new sheet already exists
                var existingSheet = package.Workbook.Worksheets[newSheetName];
                if (existingSheet != null)
                {
                    package.Workbook.Worksheets.Delete(existingSheet);
                }

                // Add a copy of the template
                var newSheet = package.Workbook.Worksheets.Copy("template", newSheetName);

                // Save changes back to file
                package.Save();
            }
        }


        public void Log(string message)
        {
            string dateTime = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                                         .Replace("-", "/");
            string logMessage = $"{dateTime} - {message}\r\n";

            // Append to UI
            if (txtResults != null)
            {
                if (txtResults.InvokeRequired)
                {
                    txtResults.Invoke(new Action(() =>
                    {
                        txtResults.AppendText(logMessage);
                        txtResults.SelectionStart = txtResults.Text.Length; // auto scroll
                        txtResults.ScrollToCaret();
                    }));
                }
                else
                {
                    txtResults.AppendText(logMessage);
                    txtResults.SelectionStart = txtResults.Text.Length; // auto scroll
                    txtResults.ScrollToCaret();
                }
            }

            // Append to file (if available)
            try
            {
                string pathToUse = _logFilePath;
                if (string.IsNullOrEmpty(pathToUse))
                {
                    // fallback to Documents/InvoiceAttachments/Logs
                    var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "InvoiceAttachments", "Logs");
                    Directory.CreateDirectory(fallbackDir);
                    pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
                }

                lock (_logLock)
                {
                    File.AppendAllText(pathToUse, logMessage);
                }
            }
            catch
            {
                // ignore logging failures to file to avoid crashing the app
            }
        }


        private string TryFormatDate(string inputDate)
        {
            if (string.IsNullOrWhiteSpace(inputDate))
                return "Not Found";

            // Trim unwanted characters: whitespace, dash, comma, period, etc.
            string cleanInput = inputDate.Trim().Trim('-', '–', '.', ',', ';', ':', ' ');

            // Optional: remove any "Date of Service:" text if it accidentally gets captured
            cleanInput = cleanInput
                .Replace("Date of Service", "", StringComparison.OrdinalIgnoreCase)
                .Replace("Date:", "", StringComparison.OrdinalIgnoreCase)
                .Replace("Service Date", "", StringComparison.OrdinalIgnoreCase)
                .Trim('-', '–', '.', ',', ';', ':', ' ');

            DateTime parsedDate;
            string[] formats = { "MM/dd/yy", "MM/dd/yyyy", "MM-dd-yy", "MM-dd-yyyy" };

            if (DateTime.TryParseExact(cleanInput, formats,
                                       CultureInfo.InvariantCulture,
                                       DateTimeStyles.None, out parsedDate))
            {
                return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

            // Return cleaned input if parsing fails
            return cleanInput;
        }


        public void ShowLoader()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ShowLoader));
                return;
            }
            lblLoading.Visible = true;
            progressBar.Visible = true;
            progressBar.MarqueeAnimationSpeed = 30;
        }

        public void HideLoader()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(HideLoader));
                return;
            }

            lblLoading.Visible = false;
            progressBar.Visible = false;
            progressBar.MarqueeAnimationSpeed = 0;
        }
    }
}
