// MainForm.cs
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MailKit.Security;
using MimeKit;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Path = System.IO.Path;

namespace EmailPDFMatchKeyword
{
  //static class Program
  //{
  //  [STAThread]
  //  static void Main()
  //  {
  //    Application.EnableVisualStyles();
  //    Application.SetCompatibleTextRenderingDefault(false);
  //    Application.Run(new MainForm());
  //  }
  //}

  private System.ComponentModel.IContainer components = null;

  protected override void Dispose(bool disposing)
  {
    if (disposing && (components != null))
    {
      components.Dispose();
    }
    base.Dispose(disposing);
  }

  public partial class MainForm01 : Form
  {
    // UI controls
    private TextBox txtHost, txtPort, txtUser, txtPassword, txtMailbox, txtSaveFolder, txtInterval;
    private Button btnStartStop, btnChooseFolder, btnTest;
    private ListBox lstLog;
    private CheckBox chkSsl, chkSearchPdfText;
    private System.Threading.Timer pollTimer;
    private volatile bool isRunning = false;
    private readonly object logLock = new object();
    private string lastSeenUidFile = "lastSeenUid.txt";

    public MainForm01()
    {
      InitializeComponent();
      Text = "Email Attachment Watcher - Bill to peer";
      Width = 850;
      Height = 600;
      InitUI();
    }
    private void InitializeComponent()
    {
      this.components = new System.ComponentModel.Container();
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(800, 450);
      this.Text = "MainForm";
    }

    private void InitUI()
    {
      var lblHost = new Label() { Left = 10, Top = 14, Text = "IMAP Host:" };
      txtHost = new TextBox() { Left = 110, Top = 10, Width = 180, Text = "imap.gmail.com" };

      var lblPort = new Label() { Left = 300, Top = 14, Text = "Port:" };
      txtPort = new TextBox() { Left = 350, Top = 10, Width = 60, Text = "993" };

      chkSsl = new CheckBox() { Left = 420, Top = 12, Text = "SSL", Checked = true };

      var lblUser = new Label() { Left = 10, Top = 44, Text = "Username:" };
      txtUser = new TextBox() { Left = 110, Top = 40, Width = 300 };

      var lblPass = new Label() { Left = 10, Top = 74, Text = "Password:" };
      txtPassword = new TextBox() { Left = 110, Top = 70, Width = 300, UseSystemPasswordChar = true };

      var lblMailbox = new Label() { Left = 430, Top = 44, Text = "Mailbox:" };
      txtMailbox = new TextBox() { Left = 490, Top = 40, Width = 120, Text = "INBOX" };

      var lblSave = new Label() { Left = 10, Top = 104, Text = "Save Folder:" };
      txtSaveFolder = new TextBox() { Left = 110, Top = 100, Width = 480, Text = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "InvoiceAttachments") };
      btnChooseFolder = new Button() { Left = 600, Top = 98, Width = 80, Text = "Browse" };
      btnChooseFolder.Click += (s, e) =>
      {
        using var fbd = new FolderBrowserDialog();
        if (fbd.ShowDialog() == DialogResult.OK) txtSaveFolder.Text = fbd.SelectedPath;
      };

      var lblInterval = new Label() { Left = 10, Top = 134, Text = "Poll (sec):" };
      txtInterval = new TextBox() { Left = 110, Top = 130, Width = 80, Text = (5 * 60).ToString() }; // default 5 minutes

      chkSearchPdfText = new CheckBox() { Left = 210, Top = 132, Text = "Search inside PDF text (slower)", Checked = true };

      btnStartStop = new Button() { Left = 10, Top = 164, Width = 120, Text = "Start" };
      btnStartStop.Click += BtnStartStop_Click;

      btnTest = new Button() { Left = 140, Top = 164, Width = 120, Text = "Run Now" };
      btnTest.Click += async (s, e) => await PollMailboxAsync();

      lstLog = new ListBox() { Left = 10, Top = 200, Width = 810, Height = 340 };

      Controls.AddRange(new Control[] {
                lblHost, txtHost, lblPort, txtPort, chkSsl,
                lblUser, txtUser, lblPass, txtPassword, lblMailbox, txtMailbox,
                lblSave, txtSaveFolder, btnChooseFolder,
                lblInterval, txtInterval, chkSearchPdfText,
                btnStartStop, btnTest, lstLog
            });
    }

    private void BtnStartStop_Click(object sender, EventArgs e)
    {
      if (!isRunning)
      {
        if (!int.TryParse(txtInterval.Text.Trim(), out int seconds) || seconds <= 0)
        {
          MessageBox.Show("Please enter a valid polling interval (seconds).");
          return;
        }

        // create folder if doesn't exist
        Directory.CreateDirectory(txtSaveFolder.Text);

        pollTimer = new System.Threading.Timer(async _ => await PollMailboxAsync(), null, TimeSpan.Zero, TimeSpan.FromSeconds(seconds));
        isRunning = true;
        btnStartStop.Text = "Stop";
        Log($"Started polling every {seconds} seconds.");
      }
      else
      {
        pollTimer?.Dispose();
        isRunning = false;
        btnStartStop.Text = "Start";
        Log("Stopped polling.");
      }
    }

    private void Log(string message)
    {
      var text = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}";
      lock (logLock)
      {
        if (lstLog.InvokeRequired)
        {
          lstLog.Invoke(new Action(() => { lstLog.Items.Insert(0, text); }));
        }
        else
        {
          lstLog.Items.Insert(0, text);
        }
      }
    }

    public static string ExtractTextFromPdf(string path)
    {
      var sb = new StringBuilder();

      using (PdfReader reader = new PdfReader(path))
      {
        int numberOfPages = reader.NumberOfPages;
        for (int i = 1; i <= numberOfPages; i++)
        {
          // Extract text from each page
          string pageText = PdfTextExtractor.GetTextFromPage(reader, i, new LocationTextExtractionStrategy());
          sb.AppendLine(pageText);
        }
      }

      return sb.ToString();
    }
    private async Task PollMailboxAsync()
    {
      // prevent overlapping runs
      if (Monitor.TryEnter(logLock) == false) return;

      try
      {
        Log("Checking mailbox...");
        var host = txtHost.Text.Trim();
        var port = int.TryParse(txtPort.Text.Trim(), out int p) ? p : 993;
        var useSsl = chkSsl.Checked;
        var user = txtUser.Text.Trim();
        var pass = txtPassword.Text;
        var mailbox = string.IsNullOrEmpty(txtMailbox.Text.Trim()) ? "INBOX" : txtMailbox.Text.Trim();
        var saveFolder = txtSaveFolder.Text.Trim();
        var searchPhrase = "bill to peer"; // phrase to match (lowercase for comparison)
        var searchFilenameKeyword = "bill"; // also allow simpler filename check

        try
        {
          using var client = new ImapClient();

          // Accept all SSL certs (use only for testing; in production validate!)
          client.ServerCertificateValidationCallback = (s, c, h, e) => true;

          var secureSocket = useSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.StartTlsWhenAvailable;
          await client.ConnectAsync(host, port, secureSocket);

          await client.AuthenticateAsync(user, pass);
          Log("Connected and authenticated.");

          var folder = client.GetFolder(mailbox);
          await folder.OpenAsync(FolderAccess.ReadWrite);

          // Search for unseen or recent messages - here we search for NotSeen to avoid re-processing.
          // You might prefer other logic: store highest UID processed and search for UIDs greater than that.
          var query = SearchQuery.NotSeen;
          var uids = await folder.SearchAsync(query);

          if (uids == null || uids.Count == 0)
          {
            Log("No new messages found.");
          }
          else
          {
            Log($"Found {uids.Count} new messages.");
            foreach (var uid in uids)
            {
              try
              {
                var message = await folder.GetMessageAsync(uid);
                Log($"Processing message: {message.Subject} from {message.From}");

                if (message.Attachments != null)
                {
                  foreach (var attachment in message.Attachments)
                  {
                    if (attachment is MimePart part)
                    {
                      var fileName = part.FileName ?? "attachment.pdf";
                      var lowerFileName = fileName.ToLowerInvariant();

                      var extension = Path.GetExtension(fileName).ToLowerInvariant();
                      var savePath = Path.Combine(saveFolder, $"{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("n").Substring(0, 8)}_{fileName}");

                      // If filename suggests it's a bill or is pdf, we will save it, and optionally search inside.
                      bool isPdf = extension == ".pdf" || part.ContentType?.MimeType == "application/pdf";
                      bool filenameMatches = lowerFileName.Contains(searchFilenameKeyword);

                      bool contentMatches = false;

                      // Save to disk first to allow pdf reading
                      using (var stream = File.Create(savePath))
                      {
                        await part.Content.DecodeToAsync(stream);
                      }

                      if (isPdf && chkSearchPdfText.Checked)
                      {
                        try
                        {
                          string allText = ExtractTextFromPdf(savePath).ToLowerInvariant();

                          if (allText.Contains(searchPhrase))
                          {
                            contentMatches = true;
                          }
                        }
                        catch (Exception ex)
                        {
                          Log($"PDF read error for {Path.GetFileName(savePath)}: {ex.Message}");
                        }
                      }

                      // decide if this attachment is a match
                      if (filenameMatches || (isPdf && contentMatches))
                      {
                        // Move or mark matched file to a "matched" folder
                        var matchedFolder = Path.Combine(saveFolder, "Matched");
                        Directory.CreateDirectory(matchedFolder);
                        var finalPath = Path.Combine(matchedFolder, Path.GetFileName(savePath));
                        if (File.Exists(finalPath)) File.Delete(finalPath);
                        File.Move(savePath, finalPath);

                        Log($"MATCHED attachment saved: {finalPath}");
                        // TODO: further processing (OCR, convert to image, extract date, charges) can be added here
                      }
                      else
                      {
                        // Not matched: keep or delete as per preference. We'll keep in Unmatched
                        var unmatchedFolder = Path.Combine(saveFolder, "Unmatched");
                        Directory.CreateDirectory(unmatchedFolder);
                        var finalPath = Path.Combine(unmatchedFolder, Path.GetFileName(savePath));
                        if (File.Exists(finalPath)) File.Delete(finalPath);
                        File.Move(savePath, finalPath);
                        Log($"Saved non-matching attachment to: {finalPath}");
                      }
                    }
                    else if (attachment is MessagePart rfc822)
                    {
                      // Generate a filename for the attached message
                      var fileName = !string.IsNullOrEmpty(rfc822.ContentDisposition?.FileName)
                          ? rfc822.ContentDisposition.FileName
                          : "attached-message.eml";

                      var path = Path.Combine(saveFolder, fileName);

                      using (var stream = File.Create(path))
                      {
                        await rfc822.Message.WriteToAsync(stream);
                      }

                      Log($"Saved attached message: {path}");
                    }
                  }
                }
                else
                {
                  Log("No attachments in message.");
                }

                // Mark message as seen to avoid re-processing
                await folder.AddFlagsAsync(uid, MessageFlags.Seen, true);
              }
              catch (Exception exMessage)
              {
                Log($"Error processing message {uid}: {exMessage.Message}");
              }
            } // end foreach uid
          }

          await client.DisconnectAsync(true);
          Log("Disconnected.");
        }
        catch (AuthenticationException)
        {
          Log("Authentication failed - check username/password or app-specific passwords and IMAP settings.");
        }
        catch (Exception ex)
        {
          Log($"Error connecting or processing mailbox: {ex.Message}");
        }

      }
      finally
      {
        Monitor.Exit(logLock);
      }
    }
  }
}
