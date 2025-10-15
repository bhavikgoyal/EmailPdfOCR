using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
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
    private System.Windows.Forms.Timer pollTimer;
    private SheetsService _sheetsService;
    private string _spreadsheetId = AppSettingsHelper.Get("GoogleDrive:SpreadsheetId");
    private CancellationTokenSource cancellationTokenSource;
    private GoogleSheetHelper _sheetHelper;
    public SheetsService SheetsService => _sheetsService;
    public GmailService Service => service;
    Label lblLoading;
    ProgressBar progressBar;


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

      // Start button
      Button btnStart = new Button { Text = "Start", Left = 10, Top = 10 };
			btnStart.Click += (s, e) => StartPolling();     
      Controls.Add(btnStart);

      // Stop button
      Button btnStop = new Button { Text = "Stop", Left = 100, Top = 10 };
      btnStop.Click += (s, e) => StopPolling();
      Controls.Add(btnStop);

      // Clear button
      Button btnClear = new Button { Text = "Clear", Left = 200, Top = 10 };
      btnClear.Click += (s, e) => txtResults.Clear();
      Controls.Add(btnClear);

      //CheckBox chkSearchPdfText = new CheckBox { Left = 100, Top = 12, Text = "Search inside PDF", Checked = true };
      //chkSearchPdfText.CheckedChanged += (s, e) => searchPdf = chkSearchPdfText.Checked;
      //Controls.Add(chkSearchPdfText);

      // Loader label
      lblLoading = new Label
      {
        Text = "Processing...",
        Left = 320,
        Top = 14,
        AutoSize = true,
        ForeColor = Color.DarkRed,
        Visible = false
      };
      Controls.Add(lblLoading);

      // Optional: Progress bar
      progressBar = new ProgressBar
      {
        Left = 420,
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
        Font = new Font("Segoe UI", 13, FontStyle.Regular), // bigger, cleaner font
        //BackColor = Color.Black,  // optional: nice terminal look
        //ForeColor = Color.Lime,   // optional: makes text pop
      };
      Controls.Add(txtResults);
    }

    public async Task AuthenticateUserAsync()
    {
      using var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read);
      var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
          GoogleClientSecrets.FromStream(stream).Secrets,
          new[] { GmailService.Scope.GmailModify, GmailService.Scope.GmailSend, DriveService.Scope.Drive, SheetsService.Scope.Spreadsheets }, 
          "user",
          CancellationToken.None,
          new FileDataStore("token.json", true)); 

      service = new GmailService(new BaseClientService.Initializer
      {
        HttpClientInitializer = credential,
        ApplicationName = "Email Attachment Reader"
      });

      Driveservices = new DriveService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = "My Gmail + Drive App",
      });

      _sheetsService = new SheetsService(new BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = "Peer List Automation"
      });

      _sheetHelper = new GoogleSheetHelper(_sheetsService, _spreadsheetId);

      Log("User authenticated via Gmail API.");
    }

    public void StartPolling()
    {
      try
      {
        ShowLoader();
        if (cancellationTokenSource != null)
        {
          // If polling is already started, don't start again
          Log("Polling is already running.");
          return;
        }

        // Create a new CancellationTokenSource to manage the cancellation
        cancellationTokenSource = new CancellationTokenSource();
        var token = cancellationTokenSource.Token;

        if (pollTimer == null)
        {
            pollTimer = new Timer();
            // Read interval from appsettings.json
            string intervalStr = AppSettingsHelper.Get("PollingIntervalMinutes");
            if (!int.TryParse(intervalStr, out int intervalMinutes))
            {
                intervalMinutes = 10; // default to 10 if parsing fails
            }
            pollTimer.Interval = intervalMinutes * 60 * 1000; // 5 minutes in milliseconds
            pollTimer.Tick += async (s, e) => await PollMailboxAsync(token);
        }

        // Run once immediately
        _ = PollMailboxAsync(token);

        pollTimer.Start();

        Log("Started polling: first check immediately, then every 5 minutes...");
      }
      catch (Exception ex)
      {
        Log($"Unexpected error: {ex.Message}");
      }
      finally
      {
        HideLoader();
      }
    }
    public void StopPolling()
    {
      if (pollTimer != null && pollTimer.Enabled)
      {
        pollTimer.Stop();
        Log("Polling stopped.");
      }

      // Cancel the polling after the current process completes
      if (cancellationTokenSource != null)
      {
        cancellationTokenSource.Cancel();  // This will stop the email polling after the current email is processed
        Log("Requested to stop polling after current email is processed.");
      }

      // Ensure we reset the token source after canceling it
      cancellationTokenSource = null;
    }

    public async Task PollMailboxAsync(CancellationToken cancellationToken)
		{
			if (service == null)
			{
				await AuthenticateUserAsync();
			}

      Log("Checking mailbox...");

      try
      {
        var labelsResponse = await service.Users.Labels.Get("me", "INBOX").ExecuteAsync(cancellationToken);
        int labelUnread = labelsResponse.MessagesUnread ?? 0;
        Log($"📬 Gmail label unread (INBOX): {labelUnread}");

        if (labelUnread == 0)
        {
          Log("✅ No new messages in INBOX.");
          return;
        }

        // --- Use THREADS list instead of MESSAGES list ---
        var request = service.Users.Threads.List("me");
        request.Q = "in:inbox is:unread";

        var allThreads = new List<Google.Apis.Gmail.v1.Data.Thread>();
        string pageToken = null;

        do
        {
          if (cancellationToken.IsCancellationRequested)
            cancellationToken.ThrowIfCancellationRequested();

          request.PageToken = pageToken;
          var response = await request.ExecuteAsync(cancellationToken);

          if (response?.Threads != null && response.Threads.Count > 0)
            allThreads.AddRange(response.Threads);

          pageToken = response?.NextPageToken;
        }
        while (!string.IsNullOrEmpty(pageToken));

        if (allThreads.Count == 0)
        {
          Log("No new unread threads found.");
          return;
        }

        // Gmail API returns newest first
        var fifoMessages = allThreads.AsEnumerable().Reverse().ToList();

        Log($"📨 Loaded {fifoMessages.Count} unread threads for processing.");

        foreach (var msgItem in fifoMessages)
        {
          try
          {
            ShowLoader();
            var message = await service.Users.Messages.Get("me", msgItem.Id).ExecuteAsync();
            Log($"Processing message: {message.Snippet}");

            string subject = message.Payload.Headers.FirstOrDefault(h => h.Name == "Subject")?.Value ?? "NoSubject";
            
            string billCharges = "Not Found", billDate = "Not Found", geicoCharges = "Not Found", geicoDate = "Not Found", caseNumber = "Not Found", CLAIMANTNAME = "Not Found", PROVIDER = "Not Found" , INCIDENTDATE = "Not Found", SCRIBETEAM = "Not Found";
            int medsToDocPageCount = 0;
            bool hasBillPdf = false , hasGeicopeerPdf = false;

            // --- Temporary storage for attachments ---
            List<(string FileName, byte[] Data)> attachments = new List<(string, byte[])>();

            foreach (var part in message.Payload.Parts ?? new System.Collections.Generic.List<MessagePart>())
            {
              if (!string.IsNullOrEmpty(part.Filename))
              {
                var attachId = part.Body.AttachmentId;
                var attach = await service.Users.Messages.Attachments.Get("me", msgItem.Id, attachId).ExecuteAsync();
                var bytes = Convert.FromBase64String(attach.Data.Replace('-', '+').Replace('_', '/'));
                
                // Save temporarily in memory (we’ll write to disk after folder creation)
                attachments.Add((part.Filename, bytes));
                string tempFilePath = Path.Combine(Path.GetTempPath(), part.Filename);
                File.WriteAllBytes(tempFilePath, bytes);

                Log($"Processed attachment: {part.Filename}");

                Log($"Saved attachment: {tempFilePath}");

                if (Path.GetExtension(tempFilePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase) &&
                    Path.GetFileName(tempFilePath).ToLower().Contains("bill"))
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

                    if (billCharges == "Not Found" || billDate == "Not Found")
                    {
                      Log("❌ Could not find Bill Charges and/or Bill Date after all retries.");
                    }
                  }
                  HideLoader();
                }

                // Handle GEICOPEER PDF
                if (Path.GetFileName(tempFilePath).Equals("Geicopeer.pdf", StringComparison.OrdinalIgnoreCase))
                {
                  ShowLoader();
                  Log("Geicopeer PDF detected. Converting to images...");
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

                      var (_,date, charges) = _ExtractMethod.ExtractFromGeicoPeer(rows);

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

                        // 1. Try to extract "Dr. Name" from subject
                        var drRegex = new Regex(@"Dr\.?\s+([A-Z][a-z]*\.?\s*)+", RegexOptions.IgnoreCase);
                        var subjectDrMatch = drRegex.Match(subject);

                        if (subjectDrMatch.Success)
                        {
                          extractedName = subjectDrMatch.Value.Trim();
                          Log($"✅ Found PROVIDER in subject (Dr.): {extractedName}");
                        }
                        else
                        {
                          // 2. If no Dr., try to extract full name (two words) from subject
                          // Assuming provider names are two words (First Last)
                          var nameRegex = new Regex(@"\b([A-Z][a-z]+)\s([A-Z][a-z]+)\b");
                          var subjectNameMatch = nameRegex.Match(subject);

                          if (subjectNameMatch.Success)
                          {
                            // Extract last name only (second group)
                            extractedName = subjectNameMatch.Groups[2].Value.Trim();
                            Log($"✅ Found PROVIDER last name in subject: {extractedName}");
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
                              emailBody = message.Snippet;

                            // Try Dr. pattern in body
                            var bodyDrMatch = drRegex.Match(emailBody);
                            if (bodyDrMatch.Success)
                            {
                              extractedName = bodyDrMatch.Value.Trim();
                              Log($"✅ Found PROVIDER in body (Dr.): {extractedName}");
                            }
                            else
                            {
                              // Try full name pattern in body
                              var bodyNameMatch = nameRegex.Match(emailBody);
                              if (bodyNameMatch.Success)
                              {
                                extractedName = bodyNameMatch.Groups[2].Value.Trim();
                                Log($"✅ Found PROVIDER last name in body: {extractedName}");
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

              //string cleanBillDate = billDate.Trim();
              //string cleanGeicoDate = geicoDate.Trim();

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
                await _ExtractMethod.ProcessAndUploadFilesAsync(caseNumber, CLAIMANTNAME, status, PROVIDER, attachments, Driveservices);
              }

              // Compare only if both values are valid
              if (status == "Matched" && hasBillPdf && hasGeicopeerPdf)
              {
                _ExtractMethod.InsertDataIntoSheetORDataBase(PROVIDER, caseNumber, CLAIMANTNAME, INCIDENTDATE, medsToDocPageCount, status, SCRIBETEAM);

                result += "Values MATCH";
                
                await _ExtractMethod.MarkMessageAsReadAsync(msgItem.Id);

                Log(result);

                Log($"Providen Values are match Successfully & Email subject: {subject} Process Completed.");
              }
              else if (status == "Not Matched" && hasBillPdf && hasGeicopeerPdf)
              {
                _ExtractMethod.InsertDataIntoSheetORDataBase(PROVIDER, caseNumber, CLAIMANTNAME, INCIDENTDATE, medsToDocPageCount, status, SCRIBETEAM);
                result += "Values DO NOT MATCH. Reason: " + mismatchReason;

                // Prepare the email body
                string emailBody = $@"
                    <html>
                    <body style='font-family:Segoe UI, sans-serif; color:#333;'>
                        <p>Hello,</p>
                        <p>
                            This is to inform you that the email bearing subject :
                            <strong>{subject}</strong> 
                            doesn't match the required details. Please check the result printed in the system.
                        </p>
                        <p><strong>Reason:</strong> {mismatchReason}</p>
                        <br/>
                        <p><strong>Comparison Details:</strong></p>
                        <pre>{result}</pre>
                        <br/>
                        <p>Thanks</p>
                    </body>
                    </html>";


                //string ToEmail = AppSettingsHelper.Get("CalculateDataEmail");

                var toList = AppSettingsHelper.Get("EmailTO")
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

                var ccList = AppSettingsHelper.Get("EmailCC")
                                ?.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                .Select(e => e.Trim());


                await _ExtractMethod.SendEmailAsync( toList, subject: "Required Details are not matched", emailBody, isHtml: true, ccList );
                Log($"Email {subject} Process will completed............");
              }
              Log(result);

              Log("======================================================");
              Log($"Email {subject} Process will completed............");
              Log("======================================================");
            }
            else
            {
              Log("======================================================");
              Log($"Email {subject} has not found the Dr.Name [PROVIDER]. Cannot proceed with this Email.");
              Log("======================================================");
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
        }
        Log("Mailbox polling completed.");
        HideLoader();
      }
			catch (Exception ex)
			{
				Log($"Error checking mailbox: {ex.Message}");
			}
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
      lblLoading.Visible = true;
      progressBar.Visible = true;
      progressBar.MarqueeAnimationSpeed = 30;
    }

    public void HideLoader()
    {
      lblLoading.Visible = false;
      progressBar.Visible = false;
      progressBar.MarqueeAnimationSpeed = 0;
    }



  }
}
