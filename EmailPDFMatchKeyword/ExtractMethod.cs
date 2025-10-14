using DocumentFormat.OpenXml.Spreadsheet;
using Google.Apis.Drive.v3;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ImageMagick;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Configuration;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Tesseract;

namespace EmailPDFMatchKeyword
{
  public class ExtractMethod
  {
    private MainForm _mainForm;
    public ExtractMethod(MainForm mainForm)
    {
      _mainForm = mainForm;
      _spreadsheetId = AppSettingsHelper.Get("GoogleDrive:SpreadsheetId");

      if (string.IsNullOrEmpty(_spreadsheetId))
      {
        throw new Exception("❌ SpreadsheetId is missing from appsettings.json");
      }
    }

    private readonly object logLock = new object();
    private DriveService Driveservices;
    private string _spreadsheetId;  // put your real ID here
    private GoogleSheetHelper _sheetHelper;


    public void InsertDataIntoSheetORDataBase(string provider, string caseNumber, string claimantName, string incidentDate, int pages, string Matchstatus, string SCRIBETEAM)
    {
      try
      {

        // Get current US Eastern Time
        TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
        DateTime usNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZone);
        _mainForm.Log($"⏰ Current US (Eastern) time: {usNow}");

        DateTime targetDate = CalculateTargetSheetDate(usNow);

        _mainForm.Log($"Start inserting Data in Database...");

        SqliteHelper.InsertCopyTemplateSheet(provider, caseNumber, claimantName, incidentDate, pages, Matchstatus, SCRIBETEAM, targetDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));

        _mainForm.Log($"Insert the data into Database succcessfull......");

        string todaySheetName = targetDate.ToString("MM/dd", CultureInfo.InvariantCulture);
        _mainForm.Log($"📄 Target sheet date selected: {todaySheetName}");
        var sheetsService = _mainForm.SheetsService;

        
        var spreadsheet = sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
        var todaySheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == todaySheetName);

        if (todaySheet != null)
          _mainForm.Log($"✅ Using existing sheet: {todaySheetName}");
        else
          _mainForm.Log($"❌ Sheet not found. You may need to copy the template to create {todaySheetName}.");

        if (todaySheet != null)
        {
          _mainForm.Log($"✅ Using existing sheet: {todaySheetName}");
        }
        else
        {
          try
          {
            _mainForm.Log($"❌ No sheet found for {todaySheetName}. Creating new sheet from template...");

            var templateSheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == "TEMPLATE");
            if (templateSheet == null) throw new Exception("❌ Template sheet not found.");

            var copyRequest = new CopySheetToAnotherSpreadsheetRequest
            {
              DestinationSpreadsheetId = _spreadsheetId
            };

            var response = sheetsService.Spreadsheets.Sheets.CopyTo(copyRequest, _spreadsheetId, (int)templateSheet.Properties.SheetId).Execute();

            _mainForm.Log($"Renaming copied sheet to {todaySheetName} and positioning it next to template...");

            // Rename + Move beside template
            var RequestUp = new BatchUpdateSpreadsheetRequest
            {
              Requests = new List<Request>
              {
                new Request
                {
                  UpdateSheetProperties = new UpdateSheetPropertiesRequest
                  {
                    Properties = new Google.Apis.Sheets.v4.Data.SheetProperties
                    {
                        SheetId = response.SheetId,
                        Title = todaySheetName
                    },
                    Fields = "title"
                  }
                },
                new Request
                {
                  UpdateSheetProperties = new UpdateSheetPropertiesRequest
                  {
                    Properties = new Google.Apis.Sheets.v4.Data.SheetProperties
                    {
                        SheetId = response.SheetId,
                        Index = (templateSheet.Properties.Index ?? 0) + 1
                    },
                    Fields = "index"
                  }
                }
              }
            };
            sheetsService.Spreadsheets.BatchUpdate(RequestUp, _spreadsheetId).Execute();

            spreadsheet = sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
            todaySheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.SheetId == response.SheetId);

            // Generate direct sheet link
            string sheetLink = $"https://docs.google.com/spreadsheets/d/{_spreadsheetId}/edit#gid={response.SheetId}";
            _mainForm.Log($"✅ New sheet created: <a href='{sheetLink}' target='_blank'>{todaySheetName}</a>");


            _mainForm.Log("Proceeding to calculate previous sheet data and send email...");
            CalculateAndSendEmail(); // Call the method to calculate and send the email
            _mainForm.Log("Sheet Data Calculated & Email send Successfully");
          }
          catch (Exception ex)
          {
            _mainForm.Log($"❌ Failed to create new sheet: {ex.Message}");
            return;
          }
        }

        _mainForm.Log($"Loading values from {todaySheetName}...");

        try
        {
          // 2. Load all values
          var range = $"{todaySheetName}!A1:Z5000";
          var getRequest = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
          var values = getRequest.Execute().Values ?? new List<IList<object>>();

          // 3. Find provider section
          _mainForm.Log($"Searching for provider section for '{provider}'...");
          int providerSectionRow = -1;
          for (int r = 0; r < values.Count; r++)
          {
            string rowText = string.Join(" ", values[r]).ToUpperInvariant();
            if (rowText.Contains(provider.ToUpperInvariant()))
            {
              providerSectionRow = r;
              break;
            }
          }

          if (providerSectionRow == -1)
            _mainForm.Log($"❌ Provider '{provider}' not found in any section.");


          // 4. Find header row (first row after provider section with "NO.", "DATE", etc.)
          _mainForm.Log("Looking for header row...");
          int headerRow = -1;
          string[] headerKeywords = { "NO", "DATE", "PROVIDER", "CASE", "CLAIMANT", "PAGES", "STATUS" };
          for (int r = providerSectionRow; r < values.Count; r++)
          {
            int matches = headerKeywords.Count(h => values[r].Any(v => v.ToString().ToUpper().Contains(h)));
            if (matches >= 2) { headerRow = r; break; }
          }
          if (headerRow == -1) throw new Exception($"❌ Header row not found for provider {provider}");

          int startDataRow = headerRow + 1;

          // 5. Find first empty row after header
          _mainForm.Log("Finding first empty row after header...");
          int insertRow = values.Count;
          for (int r = startDataRow; r < values.Count; r++)
          {
            bool isEmpty = values[r].All(v => string.IsNullOrWhiteSpace(v?.ToString()));
            if (isEmpty) { insertRow = r; break; }
          }
          if (insertRow == values.Count) insertRow = values.Count + 1;

          // 6. Build new row values (align with columns in screenshot)
          _mainForm.Log("Building new row for insertion...");
          var newRow = new List<object>
          {
              (insertRow - startDataRow + 1).ToString(),           // NO.
              "",                                                 // Initials (leave blank)
              targetDate.ToString("MM/dd/yyyy" , CultureInfo.InvariantCulture),                  // DATE
              provider ?? "",                                     // PROVIDER
              SCRIBETEAM ?? "",                                   // SCRIBE TEAM
              incidentDate ?? "",                                 // DOA
              "ISG",                                              // VENDOR
              caseNumber ?? "",                                   // CASE #
              claimantName ?? "",                                 // CLAIMANT NAME
              pages > 0 ? pages.ToString() : "",                  // PAGES
              "",                                                 // NOTES (blank)
              "",                                     // DATE SUBMITTED
              "",                                                 // TIME SUBMITTED
              "",                                                 // YES/NO
              Matchstatus ?? ""                                   // STATUS
          };

          // 7. Insert row
          _mainForm.Log($"Inserting new row at {todaySheetName}!A{insertRow + 1}...");
          string insertRange = $"{todaySheetName}!A{insertRow + 1}";
          var valueRange = new ValueRange { Values = new List<IList<object>> { newRow } };

          var updateRequest = sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, insertRange);
          updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
          updateRequest.Execute();

          _mainForm.Log($"✅ Row inserted at {todaySheetName}!A{insertRow + 1} for provider {provider}");

        }
        catch (Google.GoogleApiException gEx)
        {
          _mainForm.Log($"❌ Google Sheets API Error while reading sheet '{todaySheetName}': {gEx.Message}");
        }

      }
      catch (Exception ex)
      {
        _mainForm.Log($"EPPlus error: {ex.Message}\r\nCheck if the file is a valid Excel format and not open in another program.");
      }
    }

    public string GetFolderPrefixFromDrive(DriveService driveService, string providerName = null)
    {
      if (driveService == null) throw new ArgumentNullException(nameof(driveService));

      //string parentId = "0AOr8Zxx2A1Y6Uk9PVA"; // "2025 Test Peers" folder ID
      string parentId = AppSettingsHelper.Get("GoogleDrive:ParentFolderId");

      var listRequest = driveService.Files.List();
      listRequest.Q = $"mimeType='application/vnd.google-apps.folder' and trashed=false and '{parentId}' in parents";
      listRequest.Fields = "files(id, name)";
      listRequest.SupportsAllDrives = true;
      listRequest.IncludeItemsFromAllDrives = true;
      var result = listRequest.Execute();


      if (result.Files.Count == 0)
        return null;

      Google.Apis.Drive.v3.Data.File matchedFolder = null;

      if (!string.IsNullOrWhiteSpace(providerName))
      {
        matchedFolder = result.Files
            .FirstOrDefault(f => f.Name.IndexOf(providerName, StringComparison.OrdinalIgnoreCase) >= 0);
      }

      if (matchedFolder == null)
      {
        matchedFolder = result.Files.First(); // fallback: just take the first folder
      }
      var parts = matchedFolder.Name.Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
      return parts.Length > 0 ? parts[0] : matchedFolder.Name;
    }

    public async Task MarkMessageAsReadAsync(string messageId)
    {
      var GServices = _mainForm.Service;

      var message = await GServices.Users.Messages.Get("me", messageId).ExecuteAsync();
      var subjectHeader = message.Payload.Headers
          .FirstOrDefault(header => header.Name == "Subject")?.Value;

      if (subjectHeader != null)
      {
        _mainForm.Log($"Email Subject: {subjectHeader}");
      }
      else
      {
        _mainForm.Log("Subject header not found.");
      }

      var mods = new Google.Apis.Gmail.v1.Data.ModifyMessageRequest
      {
        RemoveLabelIds = new[] { "UNREAD" }
      };

      await GServices.Users.Messages.Modify(mods, "me", messageId).ExecuteAsync();
      _mainForm.Log($"Message {subjectHeader} marked as read.");

    }

    public async Task SendEmailAsync(IEnumerable<string> toList, string subject, string body, bool isHtml, IEnumerable<string>? ccList = null)
    {
      try
      {
        var msg = new Google.Apis.Gmail.v1.Data.Message();
        var GServices = _mainForm.Service;

        // Encode subject using Base64 for UTF-8 compatibility
        string encodedSubject = $"=?UTF-8?B?{Convert.ToBase64String(Encoding.UTF8.GetBytes(subject))}?=";

        string toHeader = string.Join(", ", toList ?? Enumerable.Empty<string>());
        string ccHeader = ccList != null ? string.Join(", ", ccList) : string.Empty;

        // Dynamically set the content type
        string contentType = isHtml ? "text/html" : "text/plain";

        // Build MIME message with optional CC and BCC
        var mimeBuilder = new StringBuilder();
        mimeBuilder.AppendLine($"To: {toHeader}");
        if (!string.IsNullOrWhiteSpace(ccHeader))
          mimeBuilder.AppendLine($"Cc: {ccHeader}"); 
        mimeBuilder.AppendLine($"Subject: {encodedSubject}");
        mimeBuilder.AppendLine($"Content-Type: {contentType}; charset=utf-8");
        mimeBuilder.AppendLine("MIME-Version: 1.0");
        mimeBuilder.AppendLine();
        mimeBuilder.AppendLine(body);

        string mimeMessage = mimeBuilder.ToString();

        msg.Raw = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(mimeMessage))
                    .Replace('+', '-')
                    .Replace('/', '_')
                    .Replace("=", "");

        await GServices.Users.Messages.Send(msg, "me").ExecuteAsync();
        _mainForm.Log($"📧 Email sent to: {toHeader}" +
                     (ccHeader != "" ? $" | CC: {ccHeader}" : "") +
                     $" | Subject: {subject}");
      }
      catch(Exception ex)
      {
        throw ex;
      }
    }

    public string ExtractDateOfService(List<List<string>> rows)
    {
      for (int i = 0; i < rows.Count; i++)
      {
        string rowText = string.Join(" ", rows[i]).ToLower();

        // Look for the specific header row that signals Date of Service section
        //if (rowText.Contains("report") || rowText.Contains("place") || rowText.Contains("service") && rowText.Contains("zip"))
        //{
        //  // Scan next 10 rows looking for a row with a date in the first cell
        //  for (int j = 1; j <= 30; j++)
        //  {
        //    if (i + j >= rows.Count) break;

        //    var currentRow = rows[i + j];
        //    if (currentRow.Count == 0) continue;

        //    foreach (var cell in currentRow)
        //    {
        //      var trimmedCell = cell.Trim();

        //      // If a cell contains a standard date pattern (MM/dd/yyyy or MM/dd/yy)
        //      if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
        //      {
        //        return trimmedCell;
        //      }

        //      // If a cell looks like "MMddyyyy" (8 digits, no slashes)
        //      if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
        //      {
        //        if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
        //            System.Globalization.CultureInfo.InvariantCulture,
        //            System.Globalization.DateTimeStyles.None, out var parsedDate))
        //        {
        //          return parsedDate.ToString("MM/dd/yyyy");
        //        }
        //      }
        //    }
        //    // Alternatively, if the full row contains a date (anywhere in the row), extract it
        //    var fullRowText = string.Join(" ", currentRow);

        //    var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
        //    if (match.Success)
        //    {
        //      return match.Value;
        //    }

        //    //// Check if the first element is a date                  
        //    //var firstCell = currentRow[0].Trim();
        //    //if (Regex.IsMatch(firstCell, @"^\d{2}/\d{2}/\d{4}$"))
        //    //{
        //    //  return firstCell;
        //    //}

        //    //// Alternatively, scan full row if needed
        //    //var fullRowText = string.Join(" ", currentRow);
        //    //var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
        //    //if (match.Success)
        //    //{
        //    //  return match.Value;
        //    //}
        //  }
        //}

        if (rowText.Contains("report of services"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (SoundsLike(rowText, "report of services"))
        {
          // Scan next 10 rows looking for a row with a date in the first cell
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (rowText.Contains("verification of treatment"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();
              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (SoundsLike(rowText, "verification of treatment"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (rowText.Contains("date of"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (SoundsLike(rowText, "date of"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (rowText.Contains("zip code"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();

              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

        else if (SoundsLike(rowText, "zip code"))
        {
          for (int j = 1; j <= 30; j++)
          {
            if (i + j >= rows.Count) break;

            var currentRow = rows[i + j];
            if (currentRow.Count == 0) continue;

            foreach (var cell in currentRow)
            {
              var trimmedCell = cell.Trim();
              if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
              {
                return trimmedCell;
              }
              if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
              {
                if (DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedDate))
                {
                  return parsedDate.ToString("MM/dd/yyyy");
                }
              }
            }
            var fullRowText = string.Join(" ", currentRow);

            var match = Regex.Match(fullRowText, @"\b\d{2}/\d{2}/\d{4}\b");
            if (match.Success)
            {
              return match.Value;
            }
          }
        }

      }
      return "Not Found";
    }
                              
    public static int LevenshteinDistance(string s, string t)
    {
      if (string.IsNullOrEmpty(s)) return t.Length;
      if (string.IsNullOrEmpty(t)) return s.Length;

      int[,] d = new int[s.Length + 1, t.Length + 1];

      for (int i = 0; i <= s.Length; i++)
        d[i, 0] = i;
      for (int j = 0; j <= t.Length; j++)
        d[0, j] = j;

      for (int i = 1; i <= s.Length; i++)
      {
        for (int j = 1; j <= t.Length; j++)
        {
          int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
          d[i, j] = Math.Min(
              Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
              d[i - 1, j - 1] + cost);
        }
      }
      return d[s.Length, t.Length];
    }

    public static bool SoundsLike(string source, string target, int threshold = 3)
    {
      int distance = LevenshteinDistance(source.ToLower(), target.ToLower());
      return distance <= threshold;
    }

    public string ExtractCharges(List<List<string>> rows)
    {
      foreach (var row in rows)
      {
        string rowText = string.Join(" ", row).ToLower();

        if (rowText.Contains("total charges to date") || rowText.Contains("total charges"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);

          //string candidateRow = string.Join(" ", row).Trim();
          //candidateRow = Regex.Replace(candidateRow, @"(\d+)\s+(\d{1,2})\b", "$1.$2");

          //var match = Regex.Match(candidateRow, @"\$ ?\d{1,3}(,\d{3})*(\.\d{2})?");
          //var match = Regex.Match(candidateRow, @"\$?\s?\d{1,}(?:,\d{3})*(?:\.\d{2})?");
          //var match1 = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          var match = Regex.Match(candidateRow, @"\$?\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }


        else if (SoundsLike(rowText, "total charges to date") || SoundsLike(rowText, "total charges"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);

          var match = Regex.Match(candidateRow, @"\$?\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }


        else if (rowText.Contains("total gharges"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);

          // First, try match with "$"
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (!match.Success)
          {
            // If no "$" found, try without "$"
            match = Regex.Match(candidateRow, @"\b\d{1,}(?:,\d{3})*(?:\.\d{1,2})?\b");
          }

          if (match.Success)
            return match.Value;
        }

        else if (SoundsLike(rowText, "total gharges"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }

        else if (rowText.Contains("total"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }

        else if (SoundsLike(rowText, "total"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }
        else if (rowText.Contains("totals"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);

          // First, try match with "$"
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (!match.Success)
          {
            // If no "$" found, try without "$"
            match = Regex.Match(candidateRow, @"\b\d{1,}(?:,\d{3})*(?:\.\d{1,2})?\b");
          }

          if (match.Success)
            return match.Value;
        }

        else if (SoundsLike(rowText, "totals"))
        {
          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }

        else if (rowText.Contains("$"))
        {
          if (Regex.IsMatch(row.FirstOrDefault() ?? "", @"^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$"))
            continue;

          for (int i = 0; i < row.Count - 1; i++)
          {
            if (Regex.IsMatch(row[i], @"^\d+$") && Regex.IsMatch(row[i + 1], @"^\d{1,2}$"))
            {
              row[i] = row[i] + "." + row[i + 1];
              row.RemoveAt(i + 1);
              break;
            }
          }
          string candidateRow = string.Join(" ", row);
          var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
          if (match.Success)
            return match.Value;
        }
      }
      return "Not Found";
    }

    public (string Provider, string DateOfService, string Charges) ExtractFromGeicoPeer(List<List<string>> rows)
    {
      for (int i = 0; i < rows.Count; i++)
      {
        string rowText = string.Join(" ", rows[i]);

        if (rowText.Contains("Providers:", StringComparison.OrdinalIgnoreCase))
        {
          string provider = "Not Found";
          var providerMatch = Regex.Match(rowText, @"Providers:\s*(.*?)\s*Dates", RegexOptions.IgnoreCase);
          if (providerMatch.Success)
            provider = providerMatch.Groups[1].Value.Trim();

          string date = "Not Found";
          string charges = "Not Found";

          var dateMatch = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{4}\b"); // only match with '/'
          if (dateMatch.Success)
          {
            string rawDate = dateMatch.Value.Trim();
            string[] formats = { "M/d/yyyy", "MM/dd/yyyy" };

            if (DateTime.TryParseExact(rawDate, formats, null, System.Globalization.DateTimeStyles.None, out var parsedDate))
            {
              date = parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            }
          }
          var amountMatch = Regex.Match(rowText, @"\$ ?\d+(?:,\d{3})*(?:\.\d{2})?");
          if (amountMatch.Success)
          {
            string rawAmount = amountMatch.Value.Replace("$", "").Replace(",", "").Trim();

            if (decimal.TryParse(rawAmount, out var parsedAmount))
            {
              charges = $"$ {parsedAmount:N2}";
            }
          }
          if (date == "Not Found" || charges == "Not Found")
          {
            string[] formats = { "M/d/yyyy", "MM/dd/yyyy" };
            for (int j = i + 1; j < Math.Min(i + 5, rows.Count); j++)
            {
              rowText = string.Join(" ", rows[j]);
              if (date == "Not Found")
              {
                var dateMatchNext = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{4}\b");
                if (dateMatchNext.Success)
                {
                  string rawDate = dateMatchNext.Value.Trim();
                  if (DateTime.TryParseExact(rawDate, formats, null, System.Globalization.DateTimeStyles.None, out var parsedDate))
                  {
                    date = parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                  }
                }
              }
              if (charges == "Not Found")
              {
                var amountMatchNext = Regex.Match(rowText, @"\$ ?\d+(?:,\d{3})*(?:\.\d{2})?");
                if (amountMatchNext.Success)
                {
                  string rawAmount = amountMatchNext.Value.Replace("$", "").Replace(",", "").Trim();

                  if (decimal.TryParse(rawAmount, out var parsedAmount))
                  {
                    charges = $"$ {parsedAmount:N2}";
                  }
                }
              }
              if (date != "Not Found" && charges != "Not Found")
                break;
            }
          }
          return (provider, date, charges);
        }
      }
      return ("Not Found", "Not Found", "Not Found");
    }

    public string ExtractCaseNumber(List<List<string>> rows)
    {
      foreach (var row in rows)
      {
        string rowText = string.Join(" ", row).Trim();

        if (rowText.IndexOf("case", StringComparison.OrdinalIgnoreCase) >= 0)
        {
          var match = Regex.Match(rowText, @"Case\s*Number[: ]\s*(\d+)", RegexOptions.IgnoreCase);
          if (match.Success)
          {
            return match.Groups[1].Value; // only the number part
          }
        }
      }
      return "Not Found";
    }

    public string ExtractClientName(List<List<string>> rows)
    {
      foreach (var row in rows)
      {
        string rowText = string.Concat(row).Trim(); 
        //string rowText = string.Join(" ", row).Trim();

        if (rowText.IndexOf("regarding", StringComparison.OrdinalIgnoreCase) >= 0)
        {
          var match = Regex.Match(rowText, @"regarding\s+(.*)", RegexOptions.IgnoreCase);
          if (match.Success)
          {
            string clientName = match.Groups[1].Value.Trim();

            if (clientName.EndsWith("."))
              clientName = clientName.Substring(0, clientName.Length - 1);

            return clientName;
          }
        }
      }
      return "Not Found";
    }

    public string ExtractProvider(List<List<string>> rows)
    {
      foreach (var row in rows)
      {
        string rowText = string.Join(" ", row).Trim();

        if (rowText.StartsWith("Dear", StringComparison.OrdinalIgnoreCase))
        {
          string namePart = rowText.Substring(4).Trim();

          string[] tokens = namePart.Split(new char[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);

          if (tokens.Length > 0)
          {
            return tokens[tokens.Length - 1]; // last word (e.g., Mayer)
          }
        }
      }
      return "Not Found";
    }

    public string ExtractDateOfIncident(List<List<string>> rows)
    {
      foreach (var row in rows)
      {
        string rowText = string.Join(" ", row).Trim();
                                                              
        if (rowText.IndexOf("incident", StringComparison.OrdinalIgnoreCase) >= 0)
        {
          string date = "Not Found";
          var match = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{4}\b");
          if (match.Success)
          {
            string rawDate = match.Value.Trim();

            string[] formats = { "M/d/yyyy", "MM/dd/yyyy" };

            if (DateTime.TryParseExact(rawDate, formats, null, System.Globalization.DateTimeStyles.None, out var parsedDate))
            {
              date = parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            }
            return date;
          }
        }
      }
      return "Not Found";
    }

    public int GetPdfPageCount_iTextSharp(Stream filePath)
    {
      var reader = new PdfReader(filePath);
      int pages = reader.NumberOfPages;
      reader.Close();
      return pages;
    }

    public List<Bitmap> ConvertPdfToImages_2(Stream pdfStream)
    {
      var images = new List<Bitmap>();
      try
      {
        var settings = new MagickReadSettings
        {
          Density = new Density(650, 650) // high resolution
        };

        _mainForm.Log("[PDF] Setting Ghostscript directory...");

        string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript", "bin");
        if (Directory.Exists(ghostscriptPath))
          MagickNET.SetGhostscriptDirectory(ghostscriptPath);

        using (var collection = new MagickImageCollection())
        {
          _mainForm.Log("[PDF] Reading PDF stream...");
          collection.Read(pdfStream, settings);
          _mainForm.Log($"[PDF] PDF loaded. Page count: {collection.Count}");

          int pagesToProcess = Math.Min(3, collection.Count);
          _mainForm.Log($"[PDF] Processing up to {pagesToProcess} pages.");

          for (int i = 0; i < pagesToProcess; i++)
          {
            _mainForm.Log($"[PDF] Processing page {i + 1}...");
            var page = collection[i];
            page.ColorType = ImageMagick.ColorType.Grayscale;
            page.Normalize();

            using (var ms = new MemoryStream())
            {
              page.Write(ms, MagickFormat.Png);
              ms.Position = 0;
              images.Add(new Bitmap(ms));
            }
            _mainForm.Log($"[PDF] Page {i + 1} converted to Bitmap.");
          }
        }
      }
      catch (Exception ex)
      {
        _mainForm.Log($"[ERROR] ConvertPdfToImages_2 failed: {ex.Message}");
        throw;
      }
      _mainForm.Log($"[PDF] Finished conversion. Total images: {images.Count}");
      return images;
    }

    public List<Bitmap> ConvertPdfToImages(Stream pdfStream)
    {
      var images = new List<Bitmap>();
      var settings = new MagickReadSettings
      {
        Density = new Density(500, 500) // high resolution
      };

      using (var collection = new MagickImageCollection())
      {
        collection.Read(pdfStream, settings);
        foreach (var page in collection)
        {
          page.ColorType = ImageMagick.ColorType.Grayscale;
          page.Normalize();

          using (var ms = new MemoryStream())
          {
            page.Write(ms, MagickFormat.Png);
            ms.Position = 0;
            images.Add(new Bitmap(ms));
          }
        }
      }

      return images;
    }


    public List<List<string>> ExtractTableRowsFromImage_new(Bitmap image)
    {
      var resultTable = new List<List<string>>();

      try
      {
        _mainForm.Log("[OCR] Starting table extraction from image...");

        using (var ms = new MemoryStream())
        {
          image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
          ms.Position = 0;
          _mainForm.Log("[OCR] Image saved to memory stream.");

          using (var magickImage = new MagickImage(ms))
          {
            _mainForm.Log("[OCR] Image loaded into MagickImage. Starting preprocessing...");

            magickImage.Deskew(new Percentage(0.3));
            magickImage.Grayscale(PixelIntensityMethod.Average);
            magickImage.AutoLevel();
            magickImage.Enhance();
            magickImage.Sharpen();
            magickImage.Contrast();
            magickImage.AdaptiveSharpen(1.2, 0.5);
            magickImage.Resize(new Percentage(220)); // slightly higher upscale

            _mainForm.Log("[OCR] Preprocessing completed.");

            // Optional debug image (remove in production if not needed)
            // string debugPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "debug_ocr_image.png");
            // magickImage.Write(debugPath);

            using (var processedStream = new MemoryStream())
            {
              magickImage.Write(processedStream, MagickFormat.Png);
              processedStream.Position = 0;
              _mainForm.Log("[OCR] Processed image written to stream for OCR.");


              // Ensure tessdata exists
              string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
              if (!Directory.Exists(tessDataPath))
              {
                string errorMsg = $"[OCR] ❌ tessdata folder not found: {tessDataPath}";
                _mainForm.Log(errorMsg);
                throw new DirectoryNotFoundException(errorMsg);
              }

              using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.LstmOnly))
              {
                _mainForm.Log("[OCR] Tesseract engine initialized.");

                // Tweaks for cleaner recognition
                engine.SetVariable("tessedit_pageseg_mode", "6"); // treat as a block of text
                engine.SetVariable("preserve_interword_spaces", "1");
                engine.SetVariable("tessedit_char_blacklist", "|~`^{}[]<>");

                using (var pix = Pix.LoadFromMemory(processedStream.ToArray()))
                using (var page = engine.Process(pix))
                {
                  string text = page.GetText();
                  _mainForm.Log($"[OCR] Raw text extracted: \n{text}");

                  if (string.IsNullOrWhiteSpace(text))
                  {
                    _mainForm.Log("⚠️ [OCR] No text detected by Tesseract.");
                    return resultTable;
                  }

                  var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                  foreach (var line in lines)
                  {
                    var cleaned = line.Trim();
                    if (!string.IsNullOrWhiteSpace(cleaned))
                    {
                      var columns = System.Text.RegularExpressions.Regex.Split(cleaned, @"\s{2,}|\t+");
                      resultTable.Add(new List<string>(columns));
                    }
                  }
                  _mainForm.Log($"✅ [OCR] Extracted {resultTable.Count} rows from image.");
                }
              }
            }
          }
        }
        _mainForm.Log($"✅ Extracted {resultTable.Count} rows successfully.");
      }
      catch (Exception ex)
      {
        _mainForm.Log("❌ OCR processing failed: " + ex.Message);
      }
      return resultTable;
    }

    public List<List<string>> ExtractTableRowsFromImage(Bitmap image)
    {
      var tableRows = new List<List<string>>();
      try
      {
        string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        _mainForm.Log($"[OCR] Using tessdata path: {tessDataPath}");

        using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.Default))
        {
          _mainForm.Log("[OCR] Tesseract engine initialized.");
          using (var ms = new MemoryStream())
          {
            image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            ms.Position = 0;
            _mainForm.Log("[OCR] Image converted to memory stream.");

            using (var pix = Pix.LoadFromMemory(ms.ToArray()))
            using (var page = engine.Process(pix))
            {
              _mainForm.Log("[OCR] OCR processing started.");
              var tsv = page.GetTsvText(0);
              _mainForm.Log($"[OCR] OCR text extracted, length: {tsv.Length}");

              var lines = tsv.Split('\n');
              _mainForm.Log($"[OCR] TSV lines count: {lines.Length}");

              int currentLineNum = -1;
              List<string> row = null;

              foreach (var line in lines.Skip(1))
              {
                var cols = line.Split('\t');
                if (cols.Length < 12) continue;

                int lineNum;
                if (!int.TryParse(cols[4], out lineNum)) continue;

                string word = cols[11].Trim();

                if (lineNum != currentLineNum)
                {
                  if (row != null) tableRows.Add(row);
                  row = new List<string>();
                  currentLineNum = lineNum;
                }
                if (!string.IsNullOrEmpty(word))
                  row.Add(word);
              }
              if (row != null) tableRows.Add(row);
              _mainForm.Log($"[OCR] Extracted {tableRows.Count} rows from image.");
            }
          }
        }
      }
      catch (Exception ex)
      {
        _mainForm.Log($"[ERROR] ExtractTableRowsFromImage failed: {ex.Message}");
        throw;
      }
      return tableRows;
    }

    public  DateTime CalculateTargetSheetDate(DateTime now)
    {
      var time = now.TimeOfDay;
      var cutoff = new TimeSpan(17, 0, 0); // 5 PM

      switch (now.DayOfWeek)
      {
        case DayOfWeek.Monday:
          return now.Date.AddDays(time < cutoff ? 1 : 2); // Tue / Wed

        case DayOfWeek.Tuesday:
          return now.Date.AddDays(time < cutoff ? 1 : 2); // Wed / Thu

        case DayOfWeek.Wednesday:
          return now.Date.AddDays(time < cutoff ? 1 : 2); // Thu / Fri

        case DayOfWeek.Thursday:
          return time < cutoff
              ? now.Date.AddDays(1) // Friday
              : GetNextWeekday(now, DayOfWeek.Monday); // Monday

        case DayOfWeek.Friday:
          return GetNextWeekday(now, DayOfWeek.Monday); // Always Monday

        case DayOfWeek.Saturday:
          return time < cutoff
              ? GetNextWeekday(now, DayOfWeek.Monday) // before 5PM → Monday
              : GetNextWeekday(now, DayOfWeek.Tuesday); // after 5PM → Tuesday

        case DayOfWeek.Sunday:
          return GetNextWeekday(now, DayOfWeek.Tuesday); // always → Tuesday

        default:
          return now.Date.AddDays(1);
      }
    }

    private DateTime GetNextWeekday(DateTime from, DayOfWeek day)
    {
      int daysToAdd = ((int)day - (int)from.DayOfWeek + 7) % 7;
      if (daysToAdd == 0) daysToAdd = 7;
      return from.Date.AddDays(daysToAdd);
    }

    public async Task ProcessAndUploadFilesAsync( string caseNumber, string CLAIMANTNAME, string Status, string PROVIDER, List<(string fileName, byte[] data)> attachments, Google.Apis.Drive.v3.DriveService Driveservices)
    {
      try
      {
        // --- Get current US Eastern Time ---
        TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
        DateTime usNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZone);
        _mainForm.Log($"⏰ Current US (Eastern) time: {usNow}");

        DateTime targetDate = CalculateTargetSheetDate(usNow);

        string today = targetDate.ToString("MM.dd");

        // --- Now create folder after extracting values ---
        //string today = DateTime.Now.AddDays(1).ToString("MM.dd");
        string folderName = $"{today} ISG {CleanFileName(caseNumber)} {CleanFileName(CLAIMANTNAME)}";
        string basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ISG_Messages");
        string saveFolder = Path.Combine(basePath, folderName);

        try
        {
          // Create folder if it doesn't exist
          if (!Directory.Exists(saveFolder))
          {
            Directory.CreateDirectory(saveFolder);
            _mainForm.Log($"Folder created: {saveFolder}");
          }

          // --- Test write permission ---
          try
          {
            string testFile = Path.Combine(saveFolder, "test.tmp");
            File.WriteAllText(testFile, "test");
            File.Delete(testFile);
            _mainForm.Log("Write permission test passed.");
          }
          catch (Exception ex)
          {
            _mainForm.Log("Permission issue: " + ex.Message);
            throw new UnauthorizedAccessException("Cannot write to folder: " + saveFolder, ex);
          }

          // --- Save all attachments safely ---
          foreach (var (fileName, data) in attachments)
          {
            string safeFileName = CleanFileName(fileName);
            string filePath = Path.Combine(saveFolder, safeFileName);

            try
            {
              // Remove read-only if exists
              if (File.Exists(filePath))
              {
                File.SetAttributes(filePath, FileAttributes.Normal);
                File.Delete(filePath);
              }

              // Write file
              using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
              {
                fs.Write(data, 0, data.Length);
              }

              _mainForm.Log($"Final saved attachment: {filePath}");
            }
            catch (Exception ex)
            {
              _mainForm.Log($"Error saving file '{safeFileName}': {ex.Message}");
            }
          }
        }
        catch (Exception ex)
        {
          _mainForm.Log($"Error in saving attachments: {ex.Message}");
        }

        // === Upload to Google Drive ===

        //string parentFolderId = "0AOr8Zxx2A1Y6Uk9PVA"; // "2025 Test Peers"
        string parentFolderId = AppSettingsHelper.Get("GoogleDrive:ParentFolderId");
        string matchedFolderId = null;
        string matchedFolderName = null;

        try
        {
          // Find subfolders inside parent
          var listRequest = Driveservices.Files.List();
          listRequest.Q = $"mimeType='application/vnd.google-apps.folder' and trashed=false and '{parentFolderId}' in parents";
          listRequest.Fields = "files(id, name, webViewLink)";
          listRequest.SupportsAllDrives = true;
          listRequest.IncludeItemsFromAllDrives = true;

          var folderList = await listRequest.ExecuteAsync();

          if (folderList.Files == null || folderList.Files.Count == 0)
          {
            _mainForm.Log("❌ No folders found inside parent folder on Drive.");
          }
          else
          {
            foreach (var folder in folderList.Files)
            {
              if (!string.IsNullOrEmpty(PROVIDER) &&
                  folder.Name.IndexOf(PROVIDER, StringComparison.OrdinalIgnoreCase) >= 0)
              {
                matchedFolderId = folder.Id;
                matchedFolderName = folder.Name;
                _mainForm.Log($"Found matching provider folder on Drive: {matchedFolderName}");
                break;
              }
            }

            if (matchedFolderId == null)
              _mainForm.Log($"❌ No matching folder found for provider '{PROVIDER}' in Drive folder.");
          }

          // === Upload files into matched Drive folder ===
          if (matchedFolderId != null)
          {
            try
            {
              // Determine folder name based on status
              string baseFolderName = Path.GetFileName(saveFolder);
              string folderNameToCreate = baseFolderName;

              if (Status == "Not Matched")
              {
                folderNameToCreate = $"{baseFolderName}_Not Matched";
              }

              // Create subfolder in provider folder
              var newFolderMetadata = new Google.Apis.Drive.v3.Data.File()
              {
                Name = folderNameToCreate, // Use updated folder name here
                MimeType = "application/vnd.google-apps.folder",
                Parents = new List<string> { matchedFolderId }
              };


              //// Create subfolder in provider folder
              //var newFolderMetadata = new Google.Apis.Drive.v3.Data.File()
              //{
              //  Name = Path.GetFileName(saveFolder), // e.g. "10.04 ISG 1892104 Tiessa O Lewis"
              //  MimeType = "application/vnd.google-apps.folder",
              //  Parents = new List<string> { matchedFolderId }
              //};

              var createFolderRequest = Driveservices.Files.Create(newFolderMetadata);
              createFolderRequest.Fields = "id, name, webViewLink";
              createFolderRequest.SupportsAllDrives = true;

              var createdFolder = await createFolderRequest.ExecuteAsync();
              string createdFolderId = createdFolder.Id;

              _mainForm.Log($"Created subfolder '{createdFolder.Name}' under provider folder '{matchedFolderName}'");

              // Upload all files inside this saveFolder into the new Drive folder
              foreach (var filePath in Directory.GetFiles(saveFolder))
              {
                try
                {
                  var fileName = Path.GetFileName(filePath);
                  var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                  {
                    Name = fileName,
                    Parents = new List<string> { createdFolderId } // Upload into subfolder
                  };

                  using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                  {
                    var uploadRequest = Driveservices.Files.Create(fileMetadata, stream, GetMimeType(filePath));
                    uploadRequest.Fields = "id, name, webViewLink";
                    uploadRequest.SupportsAllDrives = true;

                    var progress = await uploadRequest.UploadAsync();

                    if (progress.Status == Google.Apis.Upload.UploadStatus.Failed)
                    {
                      _mainForm.Log($"❌ Upload failed for '{fileName}': {progress.Exception?.Message}");
                      continue;
                    }

                    var uploadedFile = uploadRequest.ResponseBody;
                    if (uploadedFile != null && !string.IsNullOrEmpty(uploadedFile.Id))
                    {
                      string fileUrl = uploadedFile.WebViewLink ?? $"https://drive.google.com/file/d/{uploadedFile.Id}/view";
                      _mainForm.Log($"Uploaded '{fileName}' → Subfolder '{createdFolder.Name}'");
                      _mainForm.Log($"File URL: {fileUrl}");
                    }
                  }
                }
                catch (Exception ex)
                {
                  _mainForm.Log($"❌ Error uploading file '{filePath}': {ex.Message}");
                }
              }
            }
            catch (Exception ex)
            {
              _mainForm.Log($"❌ Error creating/uploading folder '{saveFolder}': {ex.Message}");
            }
          }
        }
        catch (Exception ex)
        {
          _mainForm.Log($"❌ Google Drive error: {ex.Message}");
        }
      }
      catch (Exception ex)
      {
        _mainForm.Log($"❌ Error in ProcessAndUploadFilesAsync: {ex.Message}");
      }
    }

    public void CalculateAndSendEmail()
    {
      TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
      DateTime usNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZone);
      _mainForm.Log($"⏰ Current US (Eastern) time: {usNow}");

      DateTime targetDate = CalculateTargetSheetDate(usNow);
      string todaySheetName = targetDate.ToString("MM/dd", CultureInfo.InvariantCulture);

      _mainForm.Log($"📄 Target sheet date selected: {todaySheetName}");

      // Check if it's after 5 PM
      bool isAfterFivePM = usNow.TimeOfDay == new TimeSpan(5, 0, 0);

      // Depending on the time, decide which sheet to calculate (today or yesterday's)
      string targetSheetNameToProcess = isAfterFivePM ? todaySheetName : GetPreviousSheetName(todaySheetName);

      // Retrieve data from the selected sheet (team name => record count)
      var teamRecordCounts = GetTeamRecordCounts(targetSheetNameToProcess);

      // Send email
      _mainForm.Log("📧 Sending email with calculated data...");
      SendEmailWithCalculatedData(teamRecordCounts, targetSheetNameToProcess);
      _mainForm.Log("✅ Email sent successfully.");
    }

    private string GetPreviousSheetName(string currentSheetName)
    {
      DateTime currentDate = DateTime.ParseExact(currentSheetName, "MM/dd", CultureInfo.InvariantCulture);
      return currentDate.AddDays(-1).ToString("MM/dd", CultureInfo.InvariantCulture);
    }

    private Dictionary<string, Dictionary<string, int>> GetTeamRecordCounts(string sheetName)
    {
      var result = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

      try
      {
        var sheetsService = _mainForm.SheetsService;

        // ✅ Replace / in sheet name (Google Sheets API can't parse /)
        var safeSheetName = sheetName.Replace("/", "-");
        var range = $"'{sheetName}'!A1:Z500";

        _mainForm.Log($"📄 Reading data from sheet range: {range}");

        var request = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
        var response = request.Execute();
        var values = response.Values;

        if (values == null || values.Count == 0)
        {
          _mainForm.Log($"❌ No data found in sheet '{safeSheetName}'.");
          return result;
        }

        string currentTeam = null;

        for (int i = 0; i < values.Count; i++)
        {
          var row = values[i];
          if (row == null || row.Count == 0)
            continue;

          string firstCell = row[0]?.ToString().Trim();

          // ✅ Detect TEAM NAME rows (like "SARAH")
          if (!string.IsNullOrWhiteSpace(firstCell)
              && firstCell.All(c => !char.IsDigit(c))
              && firstCell.Equals(firstCell.ToUpperInvariant())
              && row.Count < 5)
          {
            currentTeam = firstCell;
            if (!result.ContainsKey(currentTeam))
              result[currentTeam] = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            _mainForm.Log($"📍 Found team header: {currentTeam}");
            continue;
          }

          // ✅ Detect DATA ROWS (start with a number)
          if (int.TryParse(firstCell, out _))
          {
            // SCRIBE TEAM is column E (index 4)
            string scribeTeam = row.ElementAtOrDefault(4)?.ToString().Trim();

            if (!string.IsNullOrEmpty(currentTeam) && !string.IsNullOrEmpty(scribeTeam))
            {
              if (!result[currentTeam].ContainsKey(scribeTeam))
                result[currentTeam][scribeTeam] = 0;

              result[currentTeam][scribeTeam]++;
            }
          }
        }

        // ✅ Log final counts for debug
        _mainForm.Log("✅ Team Record Summary:");
        foreach (var team in result)
        {
          _mainForm.Log($"📋 {team.Key}:");
          foreach (var member in team.Value)
            _mainForm.Log($"  - {member.Key}: {member.Value} records");
        }
      }
      catch (Exception ex)
      {
        _mainForm.Log($"❌ Error reading team counts from '{sheetName}': {ex.Message}");
      }

      return result;
    }


    private async void SendEmailWithCalculatedData(Dictionary<string, Dictionary<string, int>> teamRecordCounts, string targetSheetNameToProcess)
    {
      var sb = new StringBuilder();

      // --- Header message ---
      sb.AppendLine("<p>Hello,</p>");
      sb.AppendLine($"<p>This is to notify you that we have finalized the ISG Peer reviwes for date: <strong>{targetSheetNameToProcess}</strong> summary and the brief details are as below:</p>");
      sb.AppendLine("<br>");
      sb.AppendLine("<h2>📊 Calculated Data Summary</h2>");

      foreach (var team in teamRecordCounts)
      {
        var teamName = team.Key;
        var memberDict = team.Value;

        if (memberDict == null || memberDict.Count == 0)
          continue;

        //// Combine member names into a single line
        //string memberList = string.Join(", ", memberDict.Keys);

        // Sum total records for this team
        int totalRecords = memberDict.Values.Sum();

        sb.AppendLine($"<h3>📋 {teamName}</h3>");
        //sb.AppendLine($"<p><strong>Provider:</strong> {memberList}</p>");
        sb.AppendLine($"<p><strong>Total Records:</strong> {totalRecords}</p>");
        sb.AppendLine("<br>");
      }

      string emailSubject = "✅ Calculated Data Summary Report";
      string emailBody = sb.ToString();
      //string CalculateDataEmail = AppSettingsHelper.Get("CalculateDataEmail");


      _mainForm.Log("📧 Sending formatted HTML email...");

      var toList = AppSettingsHelper.Get("EmailTO")
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

      var ccList = AppSettingsHelper.Get("EmailCC")
                      ?.Split(',', StringSplitOptions.RemoveEmptyEntries)
                      .Select(e => e.Trim());


      await SendEmailAsync( toList, emailSubject, emailBody, isHtml: true, ccList );


      _mainForm.Log("✅ Email sent successfully.");
    }


    private string GetMimeType(string filePath)
    {
      string mimeType = "application/octet-stream";
      string ext = Path.GetExtension(filePath).ToLowerInvariant();

      Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
      if (key != null && key.GetValue("Content Type") != null)
      {
        mimeType = key.GetValue("Content Type").ToString();
      }
      else
      {
        // fallback for common formats
        switch (ext)
        {
          case ".pdf": mimeType = "application/pdf"; break;
          case ".jpg":
          case ".jpeg": mimeType = "image/jpeg"; break;
          case ".png": mimeType = "image/png"; break;
          case ".doc": mimeType = "application/msword"; break;
          case ".docx": mimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"; break;
          case ".xls": mimeType = "application/vnd.ms-excel"; break;
          case ".xlsx": mimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"; break;
        }
      }
      return mimeType;
    }

    public string CleanFileName(string name)
    {
      foreach (char c in Path.GetInvalidFileNameChars())
      {
        name = name.Replace(c, '_');
      }
      return name;
    }
    

  }
}
