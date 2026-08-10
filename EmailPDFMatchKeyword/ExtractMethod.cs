using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Google.Apis.Drive.v3;
using Google.Apis.Gmail.v1;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using ImageMagick;
using iTextSharp.text.pdf;
using Microsoft.Extensions.Configuration;
using Org.BouncyCastle.Utilities.Encoders;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tesseract;
using CellFormat = Google.Apis.Sheets.v4.Data.CellFormat;
using Color = Google.Apis.Sheets.v4.Data.Color;
using ColorType = ImageMagick.ColorType;

namespace EmailPDFMatchKeyword
{
    public class ExtractMethod
    {
        // STATIC processing anchor date (do not change)
        public static readonly DateTime ProcessingStartDate = new DateTime(2026, 8, 6);

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



        public async Task<bool> InsertDataIntoSheetORDataBase( string provider, string caseNumber, string claimantName, DateTime emailReceivedUtc,  string incidentDate, int pages, string Matchstatus, string SCRIBETEAM, string Fullsubject)
		{
            try
            {
                _mainForm.ShowLoader();

                TimeZoneInfo indiaZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");

                // Ensure emailReceivedUtc is treated as UTC
                if (emailReceivedUtc.Kind != DateTimeKind.Utc)
                {
                    emailReceivedUtc = DateTime.SpecifyKind(emailReceivedUtc, DateTimeKind.Utc);
                }

                DateTime emailReceivedIndia = TimeZoneInfo.ConvertTimeFromUtc(emailReceivedUtc, indiaZone);

                // (Optional) Only for logging: current India time
                DateTime indiaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, indiaZone);

                _mainForm.Log($"⏰ Current India (IST) time: {indiaNow}");
                _mainForm.Log($"📧 Email received (UTC):    {emailReceivedUtc:u}");
                _mainForm.Log($"📧 Email received (IST):    {emailReceivedIndia:G}");

                // ---------------------------------------------
                // 2️⃣ Decide sheet date based on EMAIL time
                // ---------------------------------------------
                // Decide sheet date based on EMAIL time in IST
                DateTime targetDate = CalculateTargetSheetDate(emailReceivedIndia);

                _mainForm.Log($"📅 Target sheet date (based on email): {targetDate:MM/dd/yyyy}");

                _mainForm.Log("Start inserting Data in Database...");

                // DB me bhi same targetDate ka string store kar rahe hain
                SqliteHelper.InsertCopyTemplateSheet(provider, caseNumber, claimantName, incidentDate, pages, Matchstatus, SCRIBETEAM, targetDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));

                _mainForm.Log("Insert the data into Database successful.");

                // Sheet name = MM/dd
                string todaySheetName = targetDate.ToString("MM/dd", CultureInfo.InvariantCulture);
                _mainForm.Log($"📄 Target sheet name selected: {todaySheetName}");
                var sheetsService = _mainForm.SheetsService;

                // ---------------------------------------------
                // 3️⃣ Check if sheet exists, otherwise create from TEMPLATE
                // ---------------------------------------------
                var spreadsheet = sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
                var todaySheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == todaySheetName);

                if (todaySheet != null)
                {
                    _mainForm.Log($"✅ Using existing sheet: {todaySheetName}");
                }
                else
                {
                    _mainForm.Log($"❌ Sheet not found. Creating new sheet from template for {todaySheetName}...");

                    try
                    {
                        var templateSheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title.ToUpper().Trim() == "TEMPLATE");
                        if (templateSheet == null) throw new Exception("❌ Template sheet not found.");

                        var copyRequest = new CopySheetToAnotherSpreadsheetRequest
                        {
                            DestinationSpreadsheetId = _spreadsheetId
                        };

                        var response = sheetsService.Spreadsheets.Sheets
                            .CopyTo(copyRequest, _spreadsheetId, (int)templateSheet.Properties.SheetId)
                            .Execute();

                        _mainForm.Log($"Renaming copied sheet to {todaySheetName} and positioning it next to template...");

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

                        string sheetLink = $"https://docs.google.com/spreadsheets/d/{_spreadsheetId}/edit#gid={response.SheetId}";
                        _mainForm.Log($"✅ New sheet created: <a href='{sheetLink}' target='_blank'>{todaySheetName}</a>");

						string targetSheetNameToProcess = GetPreviousWorkingDaySheetName(targetDate);

						// Optional: ye logic ab bhi SAME rahega, bas targetDate email-date based hai
						_mainForm.Log("Proceeding to calculate previous sheet data and send email...");
                        await CalculateAndSendEmailAsync(targetSheetNameToProcess);
                        _mainForm.Log("Sheet Data Calculated & Email sent Successfully");

                        _mainForm.Log("Proceeding to calculate Match & Not Matched Records in previous sheet and send email...");

                        
                        _mainForm.Log($"📊 Processing previous working day sheet: {targetSheetNameToProcess}");

                        var matchSummary = await MatchAndNotMatchRecordCountAsync(targetSheetNameToProcess);
                        await SendEmailWithMatchSummary(matchSummary, targetSheetNameToProcess);

						try
						{
							string pathToUse = "";
							// fallback to Documents/InvoiceAttachments/Logs
							var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileAndCreateSheetLog", "Logs");
							Directory.CreateDirectory(fallbackDir);
							pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
							File.AppendAllText(pathToUse, $"Saved attachment: {sheetLink}");
						}
						catch
						{
							// ignore logging failures to file to avoid crashing the app
						}
						_mainForm.Log("Match & Not Matched Records Data Calculated & Email sent Successfully");
                    }
                    catch (Exception ex)
                    {
						try
						{
							string pathToUse = "";
							// fallback to Documents/InvoiceAttachments/Logs
							var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileAndCreateSheetLog", "Logs");
							Directory.CreateDirectory(fallbackDir);
							pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
							string errorMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
					                             $"Saved attachment Error: {ex.Message}";

							if (ex.InnerException != null)
							{
								errorMessage += $" | Inner: {ex.InnerException.Message}";
							}

							errorMessage += Environment.NewLine;

							File.AppendAllText(pathToUse, errorMessage);
						}
						catch
						{
							// ignore logging failures to file to avoid crashing the app
						}
						_mainForm.Log($"❌ Failed to create new sheet: {ex.Message}");
                        _mainForm.HideLoader();
                        return false;
                    }
                }
                _mainForm.Log($"Loading values from {todaySheetName}...");

				try
				{
					// ---------------------------------------------
					// 4️⃣ Load sheet values
					// ---------------------------------------------
					var range = $"{todaySheetName}!A1:Z5000";
					var getRequest = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
					var values = getRequest.Execute().Values ?? new List<IList<object>>();

					// ---------------------------------------------
					// Constants / settings
					// ---------------------------------------------
					const string NotFoundTitle = "Not Found Provider Records";
					const string NotFoundTitleUpper = "NOT FOUND PROVIDER RECORDS";

					// how many data rows each provider block should conceptually have
					const int ReservedDataRowsPerProvider = 14;

					// header keywords for detection
					string[] headerKeywords = { "NO", "DATE", "PROVIDER", "CASE", "CLAIMANT", "PAGES", "STATUS" };

					// ---------------------------------------------
					// Known provider section titles (for boundaries)
					// ---------------------------------------------
					string[] knownProviders = AppSettingsHelper
						.GetArray("KnownProviders")
						.Select(p => p.ToUpperInvariant())
						.ToArray();

					// Always include the special block as a section boundary
					if (!knownProviders.Contains(NotFoundTitleUpper))
					{
						knownProviders = knownProviders
							.Concat(new[] { NotFoundTitleUpper })
							.ToArray();
					}

					var providerUpper = (provider ?? string.Empty).ToUpperInvariant();

					// 5️⃣ Try to find provider section
					_mainForm.Log($"Searching for provider section for '{provider}'...");
					int providerSectionRow = -1;

					if (!string.IsNullOrWhiteSpace(providerUpper))
					{
						for (int r = 0; r < values.Count; r++)
						{
							string rowText = string.Join(" ", values[r]).ToUpperInvariant();
							if (rowText.Contains(providerUpper))
							{
								providerSectionRow = r; // 0-based index (title row)
								break;
							}
						}
					}

					int headerRow = -1;
					int startDataRow = -1;
					bool isNewSectionCreated = false;
					bool isNotFoundProviderBlock = false;

					// ---------------------------------------------
					// If provider section is NOT found → use "Not Found Provider Records" block
					// ---------------------------------------------
					if (providerSectionRow == -1)
					{
						_mainForm.Log($"⚠️ Provider '{provider}' not found. Using '{NotFoundTitle}' block...");

						// Look for existing "Not Found Provider Records" section
						for (int r = 0; r < values.Count; r++)
						{
							string rowText = string.Join(" ", values[r]).ToUpperInvariant();
							if (rowText.Contains(NotFoundTitleUpper))
							{
								providerSectionRow = r; // title row index
								break;
							}
						}

						// 👉 we are definitely using the Not Found Provider block
						isNotFoundProviderBlock = true;

						// If even that block doesn't exist → create it AFTER last provider block
						if (providerSectionRow == -1)
						{
							_mainForm.Log($"'{NotFoundTitle}' block not found. Creating new block after last provider block...");

							// 1️⃣ Find the LAST provider header row among known providers
							int lastProviderHeaderRow = -1;
							var realProviders = knownProviders
								.Where(p => p != NotFoundTitleUpper) // exclude our special block
								.ToArray();

							for (int r = 0; r < values.Count; r++)
							{
								string rowText = string.Join(" ", values[r]).ToUpperInvariant();

								// If this row belongs to a known provider section (title row)
								if (realProviders.Any(p => rowText.Contains(p)))
								{
									// Search for that provider's header row ("NO", "DATE", etc.)
									for (int r2 = r + 1; r2 < values.Count; r2++)
									{
										int matches = headerKeywords.Count(h =>
											values[r2].Any(v => v?.ToString().ToUpperInvariant().Contains(h) == true));

										if (matches >= 2)
										{
											if (r2 > lastProviderHeaderRow)
												lastProviderHeaderRow = r2; // remember the last (lowest) header row
											break;
										}
									}
								}
							}

							// 2️⃣ Decide where to place "Not Found Provider Records" title row
							int titleRowIndex;
							if (lastProviderHeaderRow != -1)
							{
								// place after [header + ReservedDataRows]
								titleRowIndex = lastProviderHeaderRow + 1 + ReservedDataRowsPerProvider;
							}
							else
							{
								// fallback: append at end
								titleRowIndex = values.Count;
							}

							// Calculate needed indices (0-based)
							providerSectionRow = titleRowIndex;       // title row index
							headerRow = providerSectionRow + 1;       // header row index
							startDataRow = headerRow + 1;             // first data row index

							// 🔍 Ensure sheet has enough rows for title + header
							int currentRowCount = todaySheet.Properties.GridProperties.RowCount ?? values.Count;
							// highest row index we will touch = headerRow
							int maxNeededRowIndex = headerRow;
							int neededRowCount = maxNeededRowIndex + 1;  // convert to 1-based count

							if (neededRowCount > currentRowCount)
							{
								int rowsToAdd = neededRowCount - currentRowCount;

								_mainForm.Log($"Sheet has {currentRowCount} rows, need {neededRowCount}. Inserting {rowsToAdd} more row(s) at bottom...");

								var addRowsRequest = new BatchUpdateSpreadsheetRequest
								{
									Requests = new List<Request>
		                            {
			                            new Request
			                            {
				                            InsertDimension = new InsertDimensionRequest
				                            {
					                            Range = new DimensionRange
					                            {
						                            SheetId = todaySheet.Properties.SheetId,
						                            Dimension = "ROWS",
						                            StartIndex = currentRowCount,              // insert after last existing row (0-based)
                                                    EndIndex = currentRowCount + rowsToAdd
					                            },
					                            InheritFromBefore = true
				                            }
			                            }
		                            }
								};

								sheetsService.Spreadsheets.BatchUpdate(addRowsRequest, _spreadsheetId).Execute();

								// Update local row count (optional, for your own logic)
								currentRowCount += rowsToAdd;
								_mainForm.Log($"✅ Inserted {rowsToAdd} row(s). New row count = {currentRowCount}.");
							}

							// 3️⃣ Write title + header (now we know the grid is big enough)
							var titleRow = new List<object>
                            {
	                            NotFoundTitle
                            };

							var headerRowValues = new List<object>
                            {
	                            "NO.", "INITIALS", "DATE", "PROVIDER", "SCRIBE TEAM", "DOA", "VENDOR", "CASE #", "CLAIMANT NAME", "PAGES", "NOTES [Email Subject]", "DATE SUBMITTED", "TIME SUBMITTED", "YES/NO", "STATUS"
                            };

							string sectionRange = $"{todaySheetName}!A{providerSectionRow + 1}:O{headerRow + 1}";
							var sectionValueRange = new ValueRange
							{
								Values = new List<IList<object>> { titleRow, headerRowValues }
							};

							var sectionUpdateRequest =
								sheetsService.Spreadsheets.Values.Update(sectionValueRange, _spreadsheetId, sectionRange);
							sectionUpdateRequest.ValueInputOption =
								SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
							sectionUpdateRequest.Execute();

							isNewSectionCreated = true;
							_mainForm.Log($"✅ Created '{NotFoundTitle}' block starting at row {providerSectionRow + 1} (header at row {headerRow + 1}).");
						}
					}

					// ---------------------------------------------
					// 6️⃣ Find header row (if not already set when creating block)
					// ---------------------------------------------
					if (!isNewSectionCreated)
					{
						_mainForm.Log("Looking for header row...");

						for (int r = providerSectionRow + 1; r < values.Count; r++)
						{
							int matches = headerKeywords.Count(h =>
								values[r].Any(v => v?.ToString().ToUpperInvariant().Contains(h) == true));

							if (matches >= 2)
							{
								headerRow = r;
								break;
							}
						}

						if (headerRow == -1)
						{
							// Header not found → create it right after title row
							_mainForm.Log($"❌ Header row not found for section starting at row {providerSectionRow + 1}. Creating header row...");

							headerRow = providerSectionRow + 1;
							startDataRow = headerRow + 1;

							var headerRowValues = new List<object>
			                {
				                "NO.", "INITIALS", "DATE", "PROVIDER", "SCRIBE TEAM", "DOA", "VENDOR", "CASE #", "CLAIMANT NAME", "PAGES", "NOTES", "DATE SUBMITTED", "TIME SUBMITTED", "YES/NO", "STATUS"
			                };

							string headerRange = $"{todaySheetName}!A{headerRow + 1}:O{headerRow + 1}";
							var headerValueRange = new ValueRange
							{
								Values = new List<IList<object>> { headerRowValues }
							};

							var headerUpdateRequest =
								sheetsService.Spreadsheets.Values.Update(headerValueRange, _spreadsheetId, headerRange);
							headerUpdateRequest.ValueInputOption =
								SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
							headerUpdateRequest.Execute();

							_mainForm.Log($"✅ Header row created at row {headerRow + 1}.");
						}
						else
						{
							startDataRow = headerRow + 1;
						}
					}

					// -----------------------------------------------
					// 7️⃣ CHECK IF CASE NUMBER ALREADY EXISTS IN THIS SECTION
					// Skip check if it's a brand new "Not Found" block (no data yet)
					// -----------------------------------------------
					_mainForm.Log($"Checking if Case Number '{caseNumber}' already exists in this section...");

					bool caseExists = false;

					if (!isNewSectionCreated)
					{
						for (int r = startDataRow; r < values.Count; r++)
						{
							var row = values[r];

							string rowTextCheck = string.Join(" ", row).ToUpperInvariant();

							// Stop scanning if next provider/section begins
							if (!string.IsNullOrWhiteSpace(rowTextCheck) &&
								knownProviders.Any(p => rowTextCheck.Contains(p)))
							{
								break; // new section reached
							}

							// Case Number column index 7 (H)
							if (row.Count > 7)
							{
								string existingCase = row[7]?.ToString()?.Trim();

								if (!string.IsNullOrWhiteSpace(existingCase) &&
									existingCase.Equals(caseNumber, StringComparison.OrdinalIgnoreCase))
								{
									caseExists = true;
									break;
								}
							}
						}
					}

					if (caseExists)
					{
						_mainForm.Log($"⚠️ Case Number '{caseNumber}' already exists in this section. Skipping insert.");
						_mainForm.HideLoader();
                return true;
					}

					_mainForm.Log("Case number not found. Proceeding to insert new row...");

					// ---------------------------------------------
					// 8️⃣ Find insert position within this section
					// ---------------------------------------------
					_mainForm.Log("Finding insert position within section...");

                    string vendor = "ISG";
					int insertRow = -1;
					int providerRecordCount = 0;
					const int PROVIDER_COL = 3; 
					const int VENDOR_COL = 6;

					bool insertBlankRow = false;
					bool insertColorSeparatorRow = false;

					int lastDataRow = -1;
					string lastProvider = null;
					string lastVendor = null;

					string currentProvider = provider ?? "";
					string currentVendor = vendor ?? "";


					if (isNewSectionCreated)
					{
						// First data row in the new "Not Found Provider Records" block
						insertRow = startDataRow;
						providerRecordCount = 0;
					}
					else
					{
						for (int r = startDataRow; r < values.Count; r++)
						{
							var row = values[r];
							string rowText = string.Join(" ", row).ToUpperInvariant();

							// stop at next provider section
							if (!string.IsNullOrWhiteSpace(rowText) &&
								knownProviders.Any(p => rowText.Contains(p)))
								break;

							bool isEmpty = row.All(c => string.IsNullOrWhiteSpace(c?.ToString()));
							if (isEmpty)
								break;

							providerRecordCount++;

							lastDataRow = r;

							if (row.Count > VENDOR_COL)
							{
								lastProvider = row[PROVIDER_COL]?.ToString();
								lastVendor = row[VENDOR_COL]?.ToString();
							}
						}
						if (lastDataRow != -1)
						{
							// CONDITION 1: Same provider, vendor different
							if (string.Equals(lastProvider, currentProvider, StringComparison.OrdinalIgnoreCase) &&
								!string.Equals(lastVendor, currentVendor, StringComparison.OrdinalIgnoreCase))
							{
								insertBlankRow = true;
								insertRow = lastDataRow + 1;
							}
							// CONDITION 2: Provider changed
							else if (!string.Equals(lastProvider, currentProvider, StringComparison.OrdinalIgnoreCase))
							{
								insertColorSeparatorRow = true;
								insertRow = lastDataRow + 1;
							}
							else
							{
								insertRow = lastDataRow + 1;
							}
						}
						else
						{
							insertRow = startDataRow;
						}

						// If no empty row or new section found, append at end
						if (insertRow == -1)
							insertRow = values.Count;
					}

					// ✅ If section already has 5 records, expand by inserting a blank row
					if (!isNewSectionCreated && providerRecordCount >= 5)
					{
						_mainForm.Log($"Section has {providerRecordCount} records. Expanding by inserting a new blank row...");
						var expandRequest = new BatchUpdateSpreadsheetRequest
						{
							Requests = new List<Request>
			                {
				                new Request
				                {
					                InsertDimension = new InsertDimensionRequest
					                {
						                Range = new DimensionRange
						                {
							                SheetId = todaySheet.Properties.SheetId,
							                Dimension = "ROWS",
							                StartIndex = insertRow,
							                EndIndex = insertRow + 1
						                },
						                InheritFromBefore = true
					                }
				                }
			                }
						};
						sheetsService.Spreadsheets.BatchUpdate(expandRequest, _spreadsheetId).Execute();
					}
					if (insertBlankRow)
					{
						_mainForm.Log("Vendor changed — inserting blank row...");

						var blankRowRequest = new BatchUpdateSpreadsheetRequest
						{
							Requests = new List<Request>
		                    {
			                    new Request
			                    {
				                    InsertDimension = new InsertDimensionRequest
				                    {
					                    Range = new DimensionRange
					                    {
						                    SheetId = todaySheet.Properties.SheetId,
						                    Dimension = "ROWS",
						                    StartIndex = insertRow,
						                    EndIndex = insertRow + 1
					                    },
					                    InheritFromBefore = true
				                    }
			                    }
		                    }
						};

						sheetsService.Spreadsheets.BatchUpdate(blankRowRequest, _spreadsheetId).Execute();

						insertRow++; // 🔴 VERY IMPORTANT
					}

                    if (insertColorSeparatorRow)
                    {
                        _mainForm.Log("Provider changed — inserting colored separator row...");

                        // Insert exactly one separator row.
                        // The row inherits formatting from the previous row.
                        sheetsService.Spreadsheets.BatchUpdate(
                            new BatchUpdateSpreadsheetRequest
                            {
                                Requests = new List<Request>
                                {
                new Request
                {
                    InsertDimension = new InsertDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = todaySheet.Properties.SheetId,
                            Dimension = "ROWS",
                            StartIndex = insertRow,
                            EndIndex = insertRow + 1
                        },

                        // Keep existing table formatting inheritance.
                        InheritFromBefore = true
                    }
                }
                                }
                            },
                            _spreadsheetId
                        ).Execute();

                        // ---------------------------------------------------------
                        // IMPORTANT:
                        // The newly inserted row is the colored divider row.
                        // We want the SAME color as the previous provider table.
                        // Only columns A:O should receive the color.
                        // ---------------------------------------------------------

                        int separatorRowIndex = insertRow;

                        try
                        {
                            // Use the previous provider's last data row as the
                            // source for detecting the existing table color.
                            int sampleRowIndex =
                                lastDataRow != -1
                                    ? lastDataRow
                                    : (headerRow != -1 ? headerRow : separatorRowIndex);

                            // Read formatting from A:O of the previous row.
                            var getReq = sheetsService.Spreadsheets.Get(_spreadsheetId);

                            getReq.Ranges = new List<string>
        {
            $"{todaySheetName}!A{sampleRowIndex + 1}:O{sampleRowIndex + 1}"
        };

                            getReq.IncludeGridData = true;

                            var gridResp = getReq.Execute();

                            var sheetWithGrid = gridResp.Sheets?
                                .FirstOrDefault(
                                    s => (s.Properties?.Title ?? string.Empty)
                                        .Equals(todaySheetName, StringComparison.OrdinalIgnoreCase)
                                );

                            // Default fallback color = black.
                            var dividerColor = new Color
                            {
                                Red = 0,
                                Green = 0,
                                Blue = 0
                            };

                            // ---------------------------------------------------------
                            // Get the existing table border color.
                            // The screenshot's table color is also used as the
                            // divider/background color.
                            // ---------------------------------------------------------

                            if (sheetWithGrid?.Data != null &&
                                sheetWithGrid.Data.Count > 0 &&
                                sheetWithGrid.Data[0].RowData != null &&
                                sheetWithGrid.Data[0].RowData.Count > 0)
                            {
                                var rowData = sheetWithGrid.Data[0].RowData[0];

                                var cell = rowData?.Values?.FirstOrDefault();

                                var effectiveFormat =
                                    cell?.EffectiveFormat ??
                                    cell?.UserEnteredFormat;

                                var sampleBorder =
                                    effectiveFormat?.Borders?.Top ??
                                    effectiveFormat?.Borders?.Bottom ??
                                    effectiveFormat?.Borders?.Left ??
                                    effectiveFormat?.Borders?.Right;

                                if (sampleBorder?.Color != null)
                                {
                                    dividerColor = new Color
                                    {
                                        Red = sampleBorder.Color.Red ?? 0,
                                        Green = sampleBorder.Color.Green ?? 0,
                                        Blue = sampleBorder.Color.Blue ?? 0,
                                        Alpha = sampleBorder.Color.Alpha ?? 1
                                    };
                                }
                            }

                            // ---------------------------------------------------------
                            // Apply the SAME color as background to A:O.
                            // This creates the solid colored divider row exactly
                            // like the existing provider table.
                            // ---------------------------------------------------------

                            var colorRequest = new BatchUpdateSpreadsheetRequest
                            {
                                Requests = new List<Request>
            {
                new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = todaySheet.Properties.SheetId,

                            // Separator row only
                            StartRowIndex = separatorRowIndex,
                            EndRowIndex = separatorRowIndex + 1,

                            // A:O only
                            StartColumnIndex = 0,
                            EndColumnIndex = 15
                        },

                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = dividerColor
                            }
                        },

                        Fields = "userEnteredFormat.backgroundColor"
                    }
                }
            }
                            };

                            sheetsService.Spreadsheets
                                .BatchUpdate(colorRequest, _spreadsheetId)
                                .Execute();

                            _mainForm.Log(
                                $"✅ Colored provider separator applied to A:O. " +
                                $"Provider = '{currentProvider}'"
                            );
                        }
                        catch (Exception ex)
                        {
                            _mainForm.Log(
                                $"⚠️ Failed to apply colored provider separator: {ex.Message}"
                            );
                        }

                        // Move to the row after the separator.
                        // The actual data will be inserted here.
                        insertRow++;
                    }

                    // ---------------------------------------------
                    // 9️⃣ Build new row values
                    // ---------------------------------------------


                    int currentRowCountForData = todaySheet.Properties.GridProperties.RowCount ?? values.Count;
					int neededRowIndexForData = insertRow;        // 0-based
					int neededRowCountForData = neededRowIndexForData + 1; // convert to 1-based

					if (neededRowCountForData > currentRowCountForData)
					{
						int rowsToAdd = neededRowCountForData - currentRowCountForData;

						_mainForm.Log($"Sheet has {currentRowCountForData} rows, need {neededRowCountForData}. Inserting {rowsToAdd} more row(s) at bottom for data row...");

						var addRowsForDataRequest = new BatchUpdateSpreadsheetRequest
						{
							Requests = new List<Request>
		                    {
			                    new Request
			                    {
				                    InsertDimension = new InsertDimensionRequest
				                    {
					                    Range = new DimensionRange
					                    {
						                    SheetId = todaySheet.Properties.SheetId,
						                    Dimension = "ROWS",
						                    StartIndex = currentRowCountForData,          // insert after last existing row (0-based)
                                            EndIndex = currentRowCountForData + rowsToAdd
					                    },
					                    InheritFromBefore = true
				                    }
			                    }
		                    }
						};

						sheetsService.Spreadsheets.BatchUpdate(addRowsForDataRequest, _spreadsheetId).Execute();

						// Update local row count in the sheet object so future checks are correct
						todaySheet.Properties.GridProperties.RowCount = currentRowCountForData + rowsToAdd;
						_mainForm.Log($"✅ Inserted {rowsToAdd} row(s) for data. New row count = {todaySheet.Properties.GridProperties.RowCount}.");
					}


					_mainForm.Log("Building new row for insertion...");

            List<object> newRow;

            // Determine NO. value. If provider changed and we inserted a colored separator row,
            // reset numbering to 1 for the new provider. Otherwise continue sequence normally.
            string noValue;
            if (insertColorSeparatorRow)
            {
                noValue = "1";
            }
            else
            {
                noValue = (insertRow - startDataRow + 1).ToString();
            }

					if (isNotFoundProviderBlock)
					{
						// ✅ Provider NOT found → store ONLY the email subject in NOTES column
                        newRow = new List<object>
                    {
                        noValue,                         // NO.
                            "",                                                                // INITIALS
                            DateTime.Parse(targetDate.ToString()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture), // DATE
                            provider ?? "",                                                                // PROVIDER (unknown)
                            SCRIBETEAM ?? "",                                                                // SCRIBE TEAM
                            "",                                                                // DOA
                            vendor ?? "",                                                             // VENDOR
                            caseNumber ?? "",                                                                // CASE #
                            "",                                                                // CLAIMANT NAME
                            "",                                                                // PAGES
                            Fullsubject ?? "",                                                // NOTES  ⬅️ only this is filled
                            "",                                                                // DATE SUBMITTED
                            "",                                                                // TIME SUBMITTED
                            "",                                                                // YES/NO
                            ""                                                  // STATUS (you can also leave blank if you prefer)
                        };
					}
					else
					{
						// ✅ Provider FOUND → existing full row behavior
                        newRow = new List<object>
                    {
                        noValue,                         // NO.
                            "",                                                                // INITIALS
                            DateTime.Parse(targetDate.ToString()).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture), // DATE
                            provider ?? "",                                                    // PROVIDER
                            SCRIBETEAM ?? "",                                                  // SCRIBE TEAM
                            incidentDate ?? "",                                                // DOA
                            vendor ?? "",                                                             // VENDOR
                            caseNumber ?? "",                                                  // CASE #
                            claimantName ?? "",                                                // CLAIMANT NAME
                            pages > 0 ? pages.ToString() : "",                                 // PAGES
                            "",                                                                // NOTES
                            "",                                                                // DATE SUBMITTED
                            "",                                                                // TIME SUBMITTED
                            "",                                                                // YES/NO
                            Matchstatus ?? ""                                                  // STATUS
                        };
					}


					// ---------------------------------------------
					// 🔟 Insert row
					// ---------------------------------------------
					_mainForm.Log($"Inserting new row at {todaySheetName}!A{insertRow + 1}...");
					string insertRange = $"{todaySheetName}!A{insertRow + 1}";
					var valueRange = new ValueRange { Values = new List<IList<object>> { newRow } };

					var updateRequest =
						sheetsService.Spreadsheets.Values.Update(valueRange, _spreadsheetId, insertRange);
					updateRequest.ValueInputOption =
						SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
					updateRequest.Execute();

					_mainForm.Log($"✅ Row inserted at {todaySheetName}!A{insertRow + 1}. Provider = '{provider}' (section may be '{NotFoundTitle}').");
					try
					{
						string pathToUse = "";
						// fallback to Documents/InvoiceAttachments/Logs
						var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileAndCreateSheetLog", "Logs");
						Directory.CreateDirectory(fallbackDir);
						pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
                        string errorMessage = $"✅ Row inserted at {todaySheetName}!A{insertRow + 1}. Provider = '{provider}' (section may be '{NotFoundTitle}').";



						errorMessage += Environment.NewLine;

						File.AppendAllText(pathToUse, errorMessage);
					}
					catch
					{
						// ignore logging failures to file to avoid crashing the app
					}
                    _mainForm.HideLoader();
                    return true;
				}
				catch (Google.GoogleApiException gEx)
				{
					try
					{
						string pathToUse = "";
						// fallback to Documents/InvoiceAttachments/Logs
						var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileAndCreateSheetLog", "Logs");
						Directory.CreateDirectory(fallbackDir);
						pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
						string errorMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
											 $"Saved attachment Error1: {gEx.Message}";

						if (gEx.InnerException != null)
						{
							errorMessage += $" | Inner: {gEx.InnerException.Message}";
						}

						errorMessage += Environment.NewLine;

						File.AppendAllText(pathToUse, errorMessage);
					}
					catch
					{
						// ignore logging failures to file to avoid crashing the app
					}
					_mainForm.Log($"❌ Google Sheets API Error while reading sheet '{todaySheetName}': {gEx.Message}");
                    _mainForm.HideLoader();
                    return false;
				}

			}
			catch (Exception ex)
            {
				try
				{
					string pathToUse = "";
					// fallback to Documents/InvoiceAttachments/Logs
					var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileAndCreateSheetLog", "Logs");
					Directory.CreateDirectory(fallbackDir);
					pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
					string errorMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] " +
										 $"Saved attachment EPPlus Error: {ex.Message}";

					if (ex.InnerException != null)
					{
						errorMessage += $" | Inner: {ex.InnerException.Message}";
					}

					errorMessage += Environment.NewLine;

					File.AppendAllText(pathToUse, errorMessage);
				}
				catch
				{
					// ignore logging failures to file to avoid crashing the app
				}
                _mainForm.Log($"EPPlus error: {ex.Message}\r\nCheck if the file is a valid Excel format and not open in another program.");
                _mainForm.HideLoader();
                return false;
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
            var threadId = message.ThreadId; // ✅ Get the thread ID

            if (!string.IsNullOrEmpty(subjectHeader))
                _mainForm.Log($"Email Subject: {subjectHeader}");
            else
                _mainForm.Log("Subject header not found.");


            // 3️⃣ Prepare modify request (remove "UNREAD")
            var modifyRequest = new Google.Apis.Gmail.v1.Data.ModifyThreadRequest
            {
                RemoveLabelIds = new[] { "UNREAD" }
            };

            // 4️⃣ Mark the entire thread as read
            await GServices.Users.Threads.Modify(modifyRequest, "me", threadId).ExecuteAsync();

            // Log the IST timestamp for when the thread was marked as read
            try
            {
                TimeZoneInfo indiaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                DateTime indiaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, indiaTimeZone);
                _mainForm.Log($"{indiaNow:dd/MM/yyyy HH:mm:ss} IST - ✅ Entire thread '{subjectHeader}' marked as read.");
            }
            catch (Exception ex)
            {
                // If timezone lookup fails, still log the event with UTC fallback
                _mainForm.Log($"{DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC - ✅ Entire thread '{subjectHeader}' marked as read. (Failed to get IST: {ex.Message})");
            }

            //await GServices.Users.Messages.Modify(mods, "me", messageId).ExecuteAsync();
            //_mainForm.Log($"Message {subjectHeader} marked as read.");

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
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string ExtractDateOfService(List<List<string>> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                string rowText = string.Join(" ", rows[i]).ToLower();

                if (rowText.Contains("report of services rendered") || rowText.Contains("attach additional sheets"))
                {
                    if (rowText.Contains("report of services"))
                    {
                        for (int j = 1; j <= 30; j++)
                        {
                            if (i + j >= rows.Count) break;

                            var currentRow = rows[i + j];
                            if (currentRow.Count == 0) continue;

                            foreach (var cell in currentRow)
                            {
                                var cleanedCell = cell.Trim().Replace("[", "").Replace("]", "").Replace(",", "").Replace("f", "").Replace("'", "").Replace("‘", "");

                                // Try MM/dd/yyyy, M/d/yyyy, MM/dd, M/d
                                if (Regex.IsMatch(cleanedCell, @"^\d{1,2}/\d{1,2}/\d{2,4}$") ||
                                    Regex.IsMatch(cleanedCell, @"^\d{1,2}/\d{1,2}$"))
                                {
                                    if (DateTime.TryParseExact(cleanedCell, new string[] { "MM/dd/yyyy", "M/d/yyyy", "MM/dd", "M/d" },
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                                    {
                                        return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                    }
                                }

                                var trimmedCell = cell.Trim();

                                if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
                                {
                                    return trimmedCell;
                                }
                                if (Regex.IsMatch(trimmedCell, @"^\d{8}$"))
                                {
                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                        }
                                    }
                                }

                                // Try to recover from bad OCR dates like "0972472025"
                                var digitsOnly = Regex.Replace(trimmedCell, @"[^\d]", "");
                                if (digitsOnly.Length == 8 || digitsOnly.Length == 9)
                                {
                                    string cleaned = digitsOnly.Length == 9 ? digitsOnly.Substring(1) : digitsOnly;
                                    if (DateTime.TryParseExact(cleaned, "MMddyyyy",
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        out var recoveredDate))
                                    {
                                        return recoveredDate.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                                    }
                                }
                            }

                            var fullRowText = string.Join(" ", currentRow);

                            var extractedDate = ExtractDateFromLine(fullRowText);

                            var match = Regex.Match(extractedDate, @"\b\d{2}/\d{2}/\d{4}\b");
                            if (match.Success)
                            {
                                return match.Value;
                            }
                            //var match_1 = Regex.Match(extractedDate, @"\b\d{1,2}/\d{1,2}/\d{2,4}\b");
                            //if (match_1.Success && DateTime.TryParse(match.Value, out var finalParsedDate))
                            //{
                            //    return finalParsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                            //}
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
                                    }
                                }

                                // Try to recover from bad OCR dates like "0972472025"
                                var digitsOnly = Regex.Replace(trimmedCell, @"[^\d]", "");
                                if (digitsOnly.Length == 8 || digitsOnly.Length == 9)
                                {
                                    string cleaned = digitsOnly.Length == 9 ? digitsOnly.Substring(1) : digitsOnly;
                                    if (DateTime.TryParseExact(cleaned, "MMddyyyy",
                                        System.Globalization.CultureInfo.InvariantCulture,
                                        System.Globalization.DateTimeStyles.None,
                                        out var recoveredDate))
                                    {
                                        return recoveredDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                        }
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

                                    string[] formats = { "MMddyyyy", "ddMMyyyy", "yyyyMMdd" };
                                    foreach (var format in formats)
                                    {
                                        if (DateTime.TryParseExact(trimmedCell, format,
                                            System.Globalization.CultureInfo.InvariantCulture,
                                            System.Globalization.DateTimeStyles.None, out var parsedDate))
                                        {
                                            return parsedDate.ToString("MM/dd/yyyy");
                                        }
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

            }
            return "Not Found";
        }

        public string ExtractDateFromLine(string line)
        {
            // Normalize line
            line = line.Trim();

            // 1️⃣ Try to find MM/dd/yyyy pattern with a 5-digit year (e.g. 09/24/20725)
            var badDateMatch = Regex.Match(line, @"\b(\d{2})/(\d{2})/(\d{5})\b");
            if (badDateMatch.Success)
            {
                // Extract components
                string month = badDateMatch.Groups[1].Value;
                string day = badDateMatch.Groups[2].Value;
                string badYear = badDateMatch.Groups[3].Value;

                // Fix the year by trimming the first digit (assuming it's an extra 2 or 0)
                string correctedYear = badYear.Substring(1);

                string fixedDate = $"{month}/{day}/{correctedYear}";
                if (DateTime.TryParseExact(fixedDate, "MM/dd/yyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsedFixedDate))
                {
                    return parsedFixedDate.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
                }
            }

            // 2️⃣ Try normal MM/dd/yyyy (in case it's already correct)
            var normalDateMatch = Regex.Match(line, @"\b\d{1,2}/\d{1,2}/\d{4}\b");
            if (normalDateMatch.Success && DateTime.TryParse(normalDateMatch.Value, out var parsedDate))
            {
                return parsedDate.ToString("MM/dd/yyyy");
            }

            // 3️⃣ Try MMddyyyy format (as a fallback)
            var fallbackDigits = Regex.Match(line, @"\b\d{8}\b");
            if (fallbackDigits.Success)
            {
                string dateDigits = fallbackDigits.Value;
                if (DateTime.TryParseExact(dateDigits, "MMddyyyy",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var fallbackDate))
                {
                    return fallbackDate.ToString("MM/dd/yyyy", System.Globalization.CultureInfo.InvariantCulture);
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
            bool result = false;
            bool startProcessing = false;

            // 1️⃣ First pass: apply all keyword and fuzzy logic (except "$" check)
            foreach (var row in rows)
            {
                string rowText = string.Join(" ", row).ToLower();

                // ✅ Check when to start processing
                if (!startProcessing && (rowText.Contains("report of services rendered") || rowText.Contains("attach additional sheets")))
                {
                    startProcessing = true;
                    continue; // skip this row, start processing from the next one
                }

                if (!startProcessing)
                    continue;

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

                    // Fix patterns like "$ 4 500.00" → "$4,500.00"
                    candidateRow = Regex.Replace(candidateRow, @"\$\s*(\d+)\s+(\d{3})\s*(?:\.(\d{2}))?", m =>
                    {
                        var dollars = m.Groups[1].Value;
                        var thousands = m.Groups[2].Value;
                        var cents = m.Groups[3].Success ? "." + m.Groups[3].Value : "";
                        return $"${dollars},{thousands}{cents}";
                    });

                    //string candidateRow = string.Join(" ", row).Trim();
                    //candidateRow = Regex.Replace(candidateRow, @"(\d+)\s+(\d{1,2})\b", "$1.$2");

                    //var match = Regex.Match(candidateRow, @"\$ ?\d{1,3}(,\d{3})*(\.\d{2})?");
                    //var match = Regex.Match(candidateRow, @"\$?\s?\d{1,}(?:,\d{3})*(?:\.\d{2})?");
                    //var match1 = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
                    var match = Regex.Match(candidateRow, @"\$?\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
                    if (match.Success)
                    {
                        result = true;
                        return match.Value;
                    }
                }

                if (rowText.Contains("total charges to"))
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
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
                    {
                        result = true;
                        return match.Value;
                    }
                }
            }

            // 2️⃣ Second pass: look for last row with a valid $ charge
            string lastDollarValue = null;

            if (startProcessing)
            {

                foreach (var row in rows)
                {
                    string rowText = string.Join(" ", row).ToLower();

                    if (Regex.IsMatch(row.FirstOrDefault() ?? "", @"^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$"))
                        continue;

                    MergeAmount(row);
                    string candidateRow = string.Join(" ", row);

                    var match = Regex.Match(candidateRow, @"\$\s?\d{1,}(?:,\d{3})*(?:\.\d{1,2})?");
                    if (match.Success)
                    {
                        lastDollarValue = match.Value;
                    }
                }
            }
            return lastDollarValue ?? "Not Found";
            //return "Not Found";
        }

        public void MergeAmount(List<string> row)
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
        }


        public string ExtractChargesAPI(List<List<string>> rows)
        {
            foreach (var row in rows)
            {
                // Convert entire row to a single lowercase string for comparison
                string rowText = string.Join(" ", row).ToLower().Trim();

                // Check for "total charges to date" with common OCR misspelling tolerance
                if (rowText.Contains("total charges to date") || rowText.Contains("total gharges to date"))
                {
                    // Attempt to fix split values like "3797" followed by "60" -> "3797.60"
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
                    {
                        var result = match.Value.Trim();
                        // Add dollar sign if missing
                        if (!result.StartsWith("$"))
                            result = "$ " + result;

                        return result;
                    }
                }
            }

            return "Not Found";
        }

        public string ExtractDateOfServiceAPI(List<List<string>> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                string rowText = string.Join(" ", rows[i]).ToLower().Trim();

                if (rowText.Contains("report of services rendered"))
                {
                    // Start scanning up to 30 rows after the keyword is found
                    for (int j = 1; j <= 30; j++)
                    {
                        if (i + j >= rows.Count) break;

                        var currentRow = rows[i + j];
                        if (currentRow.Count == 0) continue;

                        foreach (var cell in currentRow)
                        {
                            var trimmedCell = cell.Trim();

                            // Match MM/dd/yyyy or MM/dd/yy
                            if (Regex.IsMatch(trimmedCell, @"\b\d{2}/\d{2}/\d{2,4}\b"))
                            {
                                return trimmedCell;
                            }

                            // Match compact format MMddyyyy (e.g., 09112025)
                            if (Regex.IsMatch(trimmedCell, @"^\d{8}$") &&
                                DateTime.TryParseExact(trimmedCell, "MMddyyyy",
                                    System.Globalization.CultureInfo.InvariantCulture,
                                    System.Globalization.DateTimeStyles.None,
                                    out DateTime parsedDate))
                            {
                                return parsedDate.ToString("MM/dd/yyyy");
                            }
                        }

                        // Check the full row for any embedded date
                        string fullRow = string.Join(" ", currentRow);
                        var rowMatch = Regex.Match(fullRow, @"\b\d{2}/\d{2}/\d{2,4}\b");
                        if (rowMatch.Success)
                        {
                            return rowMatch.Value;
                        }
                    }
                }
            }

            return "Not Found";
        }


        public (string Provider, string DateOfService, string Charges) ExtractFromGeicoPeer(List<List<string>> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                string rowText = string.Join(" ", rows[i]);

                if (rowText.Contains("Providers:", StringComparison.OrdinalIgnoreCase) || rowText.Contains("PRV", StringComparison.OrdinalIgnoreCase))
                {
                    string provider = "Not Found";
                    var providerMatch = Regex.Match(rowText, @"Providers:\s*(.*?)\s*Dates", RegexOptions.IgnoreCase);
                    if (providerMatch.Success)
                        provider = providerMatch.Groups[1].Value.Trim();

                    string date = "Not Found";
                    string charges = "Not Found";

                    var dateMatch = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{2,4}\b"); // only match with '/'
					if (dateMatch.Success)
                    {
                        string rawDate = dateMatch.Value.Trim();
                        string[] formats = { "M/d/yyyy", "MM/dd/yyyy", "MM/dd/yy" };

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
                        string[] formats = { "M/d/yyyy", "MM/dd/yyyy", "MM/dd/yy" };
                        for (int j = i + 1; j < Math.Min(i + 5, rows.Count); j++)
                        {
                            rowText = string.Join(" ", rows[j]);
                            if (date == "Not Found")
                            {
                                var dateMatchNext = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{2,4}\b");
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

				// Normalize spaces
				rowText = Regex.Replace(rowText, @"\s+", " ");

				// ✅ 1) Case Number (primary)
				var caseMatch = Regex.Match(rowText, @"Case\s*Number[:# ]*\s*([A-Za-z0-9\-\/]+)", RegexOptions.IgnoreCase);

				if (caseMatch.Success)
				{
					return caseMatch.Groups[1].Value;
				}

				// ✅ 2) ISG / ISF File Number (fallback)
				var isgMatch = Regex.Match(rowText, @"IS[GF]\s*File\s*#[: ]*\s*([A-Za-z0-9\-\/]+)", RegexOptions.IgnoreCase);

				if (isgMatch.Success)
				{
					return isgMatch.Groups[1].Value;
				}
			}

			return "Not Found";
		}

		public string ExtractClientName(List<List<string>> rows)
		{
			foreach (var row in rows)
			{
				string rowText = string.Join(" ", row).Trim();

				// normalize whitespace a bit
				rowText = Regex.Replace(rowText, @"\s+", " ");

				// 1) Existing logic: "regarding <name>"
				if (rowText.IndexOf("regarding", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					var matchRegarding = Regex.Match(
						rowText,
						@"regarding\s+(.+)$",
						RegexOptions.IgnoreCase);

					if (matchRegarding.Success)
					{
						var clientName = matchRegarding.Groups[1].Value.Trim();

						if (clientName.EndsWith("."))
							clientName = clientName.Substring(0, clientName.Length - 1);

						return clientName;
					}
				}

				// 2) New logic: "Claimant Name: <name>   ISG File # ..."
				if (rowText.IndexOf("claimant name", StringComparison.OrdinalIgnoreCase) >= 0)
				{
					var matchClaimant = Regex.Match(
						rowText,
						@"Claimant\s*Name[: ]\s*(.+?)(?:\s{2,}|ISG\s*File|Insured:|$)",
						RegexOptions.IgnoreCase);

					if (matchClaimant.Success)
					{
						var clientName = matchClaimant.Groups[1].Value.Trim();

						// strip trailing dot if OCR left one
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

			    // ✅ Trigger on either "incident" OR "Date of Injury"
			    bool hasIncident =
				    rowText.IndexOf("incident", StringComparison.OrdinalIgnoreCase) >= 0;
			    bool hasDateOfInjury =
				    rowText.IndexOf("date of injury", StringComparison.OrdinalIgnoreCase) >= 0;

			    if (hasIncident || hasDateOfInjury)
			    {
				    string date = "Not Found";
				    var match = Regex.Match(rowText, @"\b\d{1,2}/\d{1,2}/\d{4}\b");
				    if (match.Success)
				    {
					    string rawDate = match.Value.Trim();

					    string[] formats = { "M/d/yyyy", "MM/dd/yyyy" };

					    if (DateTime.TryParseExact(
							    rawDate,
							    formats,
							    CultureInfo.InvariantCulture,
							    DateTimeStyles.None,
							    out var parsedDate))
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

        public async Task<List<Bitmap>> ConvertPdfToImages_2Async(Stream pdfStream)
        {
            return await Task.Run(() => ConvertPdfToImages_2(pdfStream));
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

                //string ghostscriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ghostscript", "bin");
                //if (Directory.Exists(ghostscriptPath))
                //  MagickNET.SetGhostscriptDirectory(ghostscriptPath);

                //// ✅ Ensure MagickTemp directory exists
                //string magickTempPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MagickTemp");
                //if (!Directory.Exists(magickTempPath))
                //  Directory.CreateDirectory(magickTempPath);

                //MagickNET.SetTempDirectory(magickTempPath);

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

        public async Task<List<Bitmap>> ConvertPdfToImagesAsync(Stream pdfStream)
        {
            return await Task.Run(() => ConvertPdfToImages(pdfStream));
        }


		public List<Bitmap> ConvertPdfToImages(Stream pdfStream)
		{
			var images = new List<Bitmap>();

			// 300 dpi is usually enough; 500 is overkill and slow
			var settings = new MagickReadSettings
			{
				Density = new Density(500, 500)
			};

			using (var collection = new MagickImageCollection())
			{
				collection.Read(pdfStream, settings);

				foreach (var page in collection)
				{
					// Make sure background is white and there is no alpha
					page.Alpha(AlphaOption.Remove);
					page.BackgroundColor = MagickColors.White;

					// Convert to grayscale for better OCR
					page.ColorType = ColorType.Grayscale;

					// Light cleanup
					page.Deskew(new Percentage(40));          // straighten if skewed
					page.ContrastStretch(new Percentage(2));  // improve contrast a bit
					page.Sharpen();                           // sharpen edges

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

		public async Task<List<List<string>>> ExtractTableRowsFromImageAllStateAsync(Bitmap image)
		{
			return await Task.Run(() => ExtractTableRowsFromImageAllState(image));
		}


		public List<List<string>> ExtractTableRowsFromImageAllState(Bitmap image)
		{
			var tableRows = new List<List<string>>();
			var lines = new List<string>();

			string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");

			using (var engine = new TesseractEngine(tessDataPath, "eng", EngineMode.LstmOnly))
			{
				// Optional: whitelist common characters you expect
				engine.SetVariable("tessedit_char_whitelist",
					"0123456789./:$ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz-, ");

				using (var ms = new MemoryStream())
				{
					image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
					ms.Position = 0;

					using (var pix = Pix.LoadFromMemory(ms.ToArray()))
					using (var page = engine.Process(pix, PageSegMode.Auto)) // Auto or SingleColumn
					{
						var text = page.GetText() ?? string.Empty;

						lines = text
							.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
							.Select(l => l.Trim())
							.Where(l => !string.IsNullOrWhiteSpace(l))
							.ToList();
					}
				}
			}

			// 🔹 Convert each line into a "row" (single cell)
			foreach (var line in lines)
			{
				tableRows.Add(new List<string> { line });
			}

			return tableRows;
		}





		public async Task<List<List<string>>> ExtractTableRowsFromImageAsync(Bitmap image)
        {
            return await Task.Run(() => ExtractTableRowsFromImage(image));
        }

        public List<List<string>> ExtractTableRowsFromImage(Bitmap image)
        {
            var tableRows = new List<List<string>>();
            try
            {
                _mainForm.ShowLoader();
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
                _mainForm.HideLoader();
            }
            catch (Exception ex)
            {
                _mainForm.Log($"[ERROR] ExtractTableRowsFromImage failed: {ex.Message}");
                throw;
            }
            return tableRows;
        }

        // Business rule: target sheet date is exactly the EMAIL RECEIVED DATE in IST (date portion only)
        // This must not depend on current system date or any "next business day" heuristics.
        public DateTime CalculateTargetSheetDate(DateTime emailReceivedIst)
        {
            // emailReceivedIst is expected to already be converted to India Standard Time
            return emailReceivedIst.Date;
        }


		public async Task ProcessAndUploadFilesAsync(DateTime emailReceivedUtc, string caseNumber, string CLAIMANTNAME, string Status, string PROVIDER, List<(string fileName, byte[] data)> attachments, Google.Apis.Drive.v3.DriveService Driveservices)
		{
			try
			{
                // --- Use India Standard Time for business date decisions ---
                TimeZoneInfo indiaZone = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                DateTime indiaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, indiaZone);
                _mainForm.Log($"⏰ Current India (IST) time: {indiaNow}");

                DateTime emailReceivedIndia = TimeZoneInfo.ConvertTimeFromUtc(emailReceivedUtc, indiaZone);
                _mainForm.Log($"📧 Email received (UTC):    {emailReceivedUtc:yyyy-MM-dd HH:mm:ss} (UTC)");
                _mainForm.Log($"📧 Email received (IST): {emailReceivedIndia:yyyy-MM-dd HH:mm:ss} (IST)");

                // Use emailReceivedIndia for sheet-date logic
                DateTime targetDate = CalculateTargetSheetDate(emailReceivedIndia);
				string today = targetDate.ToString("MM.dd");


				//DateTime targetDate = CalculateTargetSheetDate(usNow);
				//string today = targetDate.ToString("MM.dd");

				// --- Build local folder path ---
				string folderName = $"{today} ISG {CleanFileName(caseNumber)} {CleanFileName(CLAIMANTNAME)}-{"TBC"}";
				//string basePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ISG_Messages");
				string basePath = GetBaseFolderPath();
				string? matchedFolder = FindDoctorFolder(basePath, PROVIDER);
				string saveFolder;

				// If found → use it
				if (!string.IsNullOrEmpty(matchedFolder))
				{
					saveFolder = Path.Combine(matchedFolder, folderName);

					_mainForm.Log($"✅ Existing doctor folder found: {saveFolder}");
				}
				else
				{
					// Else → create new (old logic)
					saveFolder = Path.Combine(basePath, folderName);

					_mainForm.Log($"🆕 No matching folder. Creating new: {saveFolder}");
				}

				// =========================
				// 1) SAVE ATTACHMENTS LOCALLY
				// =========================
				try
				{
					_mainForm.ShowLoader();

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
						LogPdfStatus(caseNumber, fileName, "InProgress");

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
							LogPdfStatus(caseNumber, fileName, "Completed");
							try
							{
								string pathToUse = "";
								// fallback to Documents/InvoiceAttachments/Logs
								var fallbackDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "FinalEmailReadFileLog", "Logs");
								Directory.CreateDirectory(fallbackDir);
								pathToUse = Path.Combine(fallbackDir, $"Log_{DateTime.Now:yyyyMMdd}.txt");
								File.AppendAllText(pathToUse, $"Saved attachment: {filePath}");
							}
							catch
							{
								// ignore logging failures to file to avoid crashing the app
							}
						}
						catch (Exception ex)
						{
							_mainForm.Log($"Error saving file '{safeFileName}': {ex.Message}");
						}
					}
					_mainForm.HideLoader();
				}
				catch (Exception ex)
				{
					_mainForm.Log($"Error in saving attachments: {ex.Message}");
					_mainForm.HideLoader();
				}

				string parentFolderId = AppSettingsHelper.Get("GoogleDrive:ParentFolderId");
				if (string.IsNullOrWhiteSpace(parentFolderId))
				{
					_mainForm.Log("❌ GoogleDrive:ParentFolderId not configured or empty. Skipping Drive upload.");
					return;
				}
				string matchedFolderId = null;
				string matchedFolderName = null;

				try
				{
					_mainForm.ShowLoader();

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
							if (!string.IsNullOrEmpty(PROVIDER) && folder.Name.IndexOf(PROVIDER, StringComparison.OrdinalIgnoreCase) >= 0)
							{
								matchedFolderId = folder.Id;
								matchedFolderName = folder.Name;
								_mainForm.Log($"Found matching provider folder on Drive: {matchedFolderName}");
								break;
							}
						}

						// 🔑 CHANGE 3: Proper check for not-found
						if (string.IsNullOrEmpty(matchedFolderId))
						{
							_mainForm.Log($"❌ No matching folder found for provider '{PROVIDER}' in Drive parent '{parentFolderId}'.");
						}
						//if (matchedFolderId == null)
						//	_mainForm.Log($"❌ No matching folder found for provider '{PROVIDER}' in Drive folder.");
					}

					_mainForm.HideLoader();

					// =========================
					// 3) UPLOAD FILES INTO MATCHED PROVIDER FOLDER
					// =========================

					//if (matchedFolderId != null)
                    if (!string.IsNullOrEmpty(matchedFolderId))
					{
						try
						{
							_mainForm.ShowLoader();

							// 🔍 NEW PART: CHECK IF CASE FOLDER ALREADY EXISTS UNDER THIS PROVIDER
							if (!string.IsNullOrWhiteSpace(caseNumber))
							{
								var subListReq = Driveservices.Files.List();
								subListReq.Q =
									$"mimeType='application/vnd.google-apps.folder' and trashed=false and '{matchedFolderId}' in parents";
								subListReq.Fields = "files(id, name)";
								subListReq.SupportsAllDrives = true;
								subListReq.IncludeItemsFromAllDrives = true;

								var subListResp = await subListReq.ExecuteAsync();
								var existingFolders = subListResp.Files ?? new List<Google.Apis.Drive.v3.Data.File>();
								Google.Apis.Drive.v3.Data.File existingCaseFolder = default!;

								if (!string.IsNullOrWhiteSpace(CLAIMANTNAME))
								{
									existingCaseFolder = existingFolders.FirstOrDefault(f =>
										!string.IsNullOrEmpty(f.Name) &&
										f.Name.IndexOf(caseNumber, StringComparison.OrdinalIgnoreCase) >= 0 &&
										f.Name.IndexOf(CLAIMANTNAME, StringComparison.OrdinalIgnoreCase) >= 0
									);
								}
								else
								{
									existingCaseFolder = existingFolders.FirstOrDefault(f =>
										!string.IsNullOrEmpty(f.Name) &&
										f.Name.IndexOf(caseNumber, StringComparison.OrdinalIgnoreCase) >= 0
									);
								}

								//var existingCaseFolder = existingFolders
								//	.FirstOrDefault(f =>
								//		!string.IsNullOrEmpty(f.Name) &&
								//		f.Name.IndexOf(caseNumber, StringComparison.OrdinalIgnoreCase) >= 0);

								if (existingCaseFolder != null)
								{
									_mainForm.Log($"📁 Folder already exists for Case #{caseNumber} under provider '{matchedFolderName}'.");
									_mainForm.Log($"Existing folder name: {existingCaseFolder.Name}");
									_mainForm.Log("⏩ Skipping folder creation and file uploads for this case.");

									_mainForm.HideLoader();
									return; // ⛔ STOP: do not create or upload for this case
								}
							}
							else
							{
								_mainForm.Log("⚠ Case number is empty; skipping case-folder duplication check.");
							}

							// If we reach here, no folder for this case exists yet.
							// Determine folder name based on status
							string baseFolderName = Path.GetFileName(saveFolder); // e.g. "10.04 ISG 1892104 Tiessa O Lewis"
							string folderNameToCreate = baseFolderName;

							if (Status == "Not Matched")
							{
								folderNameToCreate = $"{baseFolderName}_Not Matched";
							}

							// Create subfolder in provider folder
							var newFolderMetadata = new Google.Apis.Drive.v3.Data.File()
							{
								Name = folderNameToCreate,
								MimeType = "application/vnd.google-apps.folder",
								Parents = new List<string> { matchedFolderId }
							};

							_mainForm.Log($"Creating Drive subfolder '{folderNameToCreate}' under provider '{matchedFolderName}' (ParentId={matchedFolderId})...");

							var createFolderRequest = Driveservices.Files.Create(newFolderMetadata);
							createFolderRequest.Fields = "id, name, webViewLink";
							createFolderRequest.SupportsAllDrives = true;

							var createdFolder = await createFolderRequest.ExecuteAsync();
							string createdFolderId = createdFolder.Id;

							_mainForm.Log($"✅ Created subfolder '{createdFolder.Name}' under provider folder '{matchedFolderName}'");

							// Upload all files inside this saveFolder into the new Drive folder
							foreach (var filePath in Directory.GetFiles(saveFolder))
							{
								var fileName = Path.GetFileName(filePath);

								// Max 3 attempts per file
								const int maxAttempts = 3;
								int attempt = 0;
								bool success = false;

								while (attempt < maxAttempts && !success)
								{
									attempt++;

									try
									{
										_mainForm.Log($"📤 [{attempt}/{maxAttempts}] Uploading '{fileName}'...");

										var fileMetadata = new Google.Apis.Drive.v3.Data.File
										{
											Name = fileName,
											Parents = new List<string> { createdFolderId } // Upload into subfolder
										};

										using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
										{
											var uploadRequest = Driveservices.Files.Create(fileMetadata, stream, GetMimeType(filePath));
											uploadRequest.Fields = "id, name, webViewLink";
											uploadRequest.SupportsAllDrives = true;

											var progress = await uploadRequest.UploadAsync();

											if (progress.Status == Google.Apis.Upload.UploadStatus.Failed)
											{
												var ex = progress.Exception;
												_mainForm.Log($"❌ Upload failed for '{fileName}' on attempt {attempt}: {ex?.Message}");

												// Agar GoogleApiException hai to status code bhi log karein
												if (ex is Google.GoogleApiException gex)
												{
													_mainForm.Log($"   ↳ HTTP Status: {gex.HttpStatusCode}, Errors: {gex.Error?.Message}");
												}

												// Transient error ho sakta hai → next attempt (thoda wait)
												if (attempt < maxAttempts)
												{
													await Task.Delay(TimeSpan.FromSeconds(3));
													continue;
												}

												// Max attempts complete → hard fail
												_mainForm.Log($"🚫 Giving up on '{fileName}' after {maxAttempts} failed attempts.");
												break;
											}

											// ✅ Success
											var uploadedFile = uploadRequest.ResponseBody;
											if (uploadedFile != null && !string.IsNullOrEmpty(uploadedFile.Id))
											{
												string fileUrl = uploadedFile.WebViewLink ??
																 $"https://drive.google.com/file/d/{uploadedFile.Id}/view";

												_mainForm.Log($"✅ Uploaded '{fileName}' → Subfolder '{createdFolder.Name}'");
												_mainForm.Log($"🔗 File URL: {fileUrl}");
											}

											success = true;
										}
									}
									catch (HttpRequestException httpEx)
									{
										_mainForm.Log($"🌐 HTTP error while uploading '{fileName}' on attempt {attempt}: {httpEx.Message}");

										if (attempt < maxAttempts)
										{
											await Task.Delay(TimeSpan.FromSeconds(3));
											continue;
										}

										_mainForm.Log($"🚫 Giving up on '{fileName}' after {maxAttempts} HTTP failures.");
									}
									catch (IOException ioEx)
									{
										_mainForm.Log($"💾 IO error while uploading '{fileName}': {ioEx.Message} (file locked/missing?)");
										// IO error usually local issue, retry optional – yahan ek hi attempt enough
										break;
									}
									catch (Exception ex)
									{
										_mainForm.Log($"❌ Unexpected error uploading file '{filePath}' on attempt {attempt}: {ex.Message}");

										if (attempt < maxAttempts)
										{
											await Task.Delay(TimeSpan.FromSeconds(2));
											continue;
										}

										_mainForm.Log($"🚫 Giving up on '{fileName}' after {maxAttempts} unexpected failures.");
									}
								}
							}

							_mainForm.HideLoader();

						}
						catch (Exception ex)
						{
							_mainForm.Log($"❌ Error creating/uploading folder '{saveFolder}': {ex.Message}");
							_mainForm.HideLoader();
						}
					}
				}
				catch (Exception ex)
				{
					_mainForm.Log($"❌ Google Drive error: {ex.Message}");
					_mainForm.HideLoader();
				}
			}
			catch (Exception ex)
			{
				_mainForm.Log($"❌ Error in ProcessAndUploadFilesAsync: {ex.Message}");
				_mainForm.HideLoader();
			}
		}


		public async Task<Dictionary<string, (int Matched, int NotMatched)>> MatchAndNotMatchRecordCountAsync(string sheetName)
        {
            var result = new Dictionary<string, (int Matched, int NotMatched)>(StringComparer.OrdinalIgnoreCase)
            {
                ["Mikhail"] = (0, 0),
                ["Amurta"] = (0, 0),
                ["Sarah"] = (0, 0),
                ["Krina"] = (0, 0),
                ["Patrizia"] = (0, 0),
                ["Amanda"] = (0, 0)
            };

            try
            {
                _mainForm.ShowLoader();

                var sheetsService = _mainForm.SheetsService;
                var range = $"'{sheetName}'!A1:Z2000";
                _mainForm.Log($"📄 Reading data from: {range}");

                var request = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();
                var values = response.Values;

                if (values == null || values.Count == 0)
                {
                    _mainForm.Log($"❌ No data found in sheet '{sheetName}'.");
                    return result;
                }

                // Known teams
                var teamNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Mikhail", "Amurta", "Sarah", "Krina", "Patrizia", "Amanda"
        };

                string currentTeam = null;
                bool inDataSection = false;
                int statusColumnIndex = -1;

                for (int i = 0; i < values.Count; i++)
                {
                    var row = values[i];
                    if (row == null || row.Count == 0)
                        continue;

                    string firstCell = row[0]?.ToString().Trim() ?? "";

                    if (string.IsNullOrEmpty(firstCell))
                        continue;

                    // --- Detect team header ---
                    if (teamNames.Contains(firstCell, StringComparer.OrdinalIgnoreCase))
                    {
                        currentTeam = firstCell;
                        inDataSection = false;
                        statusColumnIndex = -1;
                        _mainForm.Log($"📍 Found team: {currentTeam}");
                        continue;
                    }

                    // --- Detect header row ---
                    if (currentTeam != null && !inDataSection)
                    {
                        bool hasCaseHeader = row.Any(c =>
                            c != null && c.ToString().Trim().Equals("CASE #", StringComparison.OrdinalIgnoreCase));

                        if (hasCaseHeader)
                        {
                            inDataSection = true;
                            // Find STATUS column index
                            for (int col = 0; col < row.Count; col++)
                            {
                                string colName = row[col]?.ToString().Trim();
                                if (colName.Equals("STATUS", StringComparison.OrdinalIgnoreCase))
                                {
                                    statusColumnIndex = col;
                                    break;
                                }
                            }
                            continue;
                        }
                    }

                    // --- Count records by STATUS ---
                    if (inDataSection && currentTeam != null && statusColumnIndex >= 0 && statusColumnIndex < row.Count)
                    {
                        string statusValue = row[statusColumnIndex]?.ToString().Trim().ToLowerInvariant();

                        if (statusValue == "matched")
                        {
                            var data = result[currentTeam];
                            data.Matched++;
                            result[currentTeam] = data;
                        }
                        else if (statusValue == "not matched")
                        {
                            var data = result[currentTeam];
                            data.NotMatched++;
                            result[currentTeam] = data;
                        }
                    }
                }

                // ✅ Log summary
                _mainForm.Log("✅ Match/NotMatch Summary");
                foreach (var kv in result)
                {
                    _mainForm.Log($"📋 {kv.Key}\nMatched Records: {kv.Value.Matched}\nNotMatched Records: {kv.Value.NotMatched}\n");
                }

                int totalMatched = result.Values.Sum(v => v.Matched);
                int totalNotMatched = result.Values.Sum(v => v.NotMatched);
                _mainForm.Log($"📊 Overall Total: Matched = {totalMatched}, NotMatched = {totalNotMatched}");
            }
            catch (Exception ex)
            {
                _mainForm.Log($"❌ Error reading Match/NotMatch counts: {ex.Message}");
            }
            finally
            {
                _mainForm.HideLoader();
            }

            return result;
        }

        private async Task SendEmailWithMatchSummary( Dictionary<string, (int Matched, int NotMatched)> teamMatchSummary, string targetSheetNameToProcess)
        {
            _mainForm.ShowLoader();
            var sb = new StringBuilder();

            // --- Header ---
            sb.AppendLine("<p>Hello,</p>");
            sb.AppendLine($"<p>This is to notify you that the Match/NotMatch record summary for the ISG Peer reviews dated <strong>{targetSheetNameToProcess}</strong> is as follows:</p>");
            sb.AppendLine("<br>");
            sb.AppendLine("<h2>📊 Match vs NotMatch Summary</h2>");

            if (teamMatchSummary == null || teamMatchSummary.Count == 0)
            {
                sb.AppendLine("<p><strong>No data found for this date.</strong></p>");
            }
            else
            {
                int totalMatched = 0;
                int totalNotMatched = 0;

                foreach (var team in teamMatchSummary)
                {
                    var teamName = team.Key;
                    var matchedCount = team.Value.Matched;
                    var notMatchedCount = team.Value.NotMatched;

                    sb.AppendLine($"<h3>📋 {teamName}</h3>");
                    sb.AppendLine($"<p>✅ <strong>Matched Records:</strong> {matchedCount}</p>");
                    sb.AppendLine($"<p>❌ <strong>NotMatched Records:</strong> {notMatchedCount}</p>");
                    sb.AppendLine("<br>");

                    totalMatched += matchedCount;
                    totalNotMatched += notMatchedCount;
                }

                sb.AppendLine("<hr>");
                sb.AppendLine($"<h3>📊 <strong>Overall Summary:</strong></h3>");
                sb.AppendLine($"<p>✅ Total Matched Records: <strong>{totalMatched}</strong></p>");
                sb.AppendLine($"<p>❌ Total NotMatched Records: <strong>{totalNotMatched}</strong></p>");
            }

            string emailSubject = "✅ Match vs NotMatch Summary Report";
            string emailBody = sb.ToString();

            _mainForm.Log("📧 Sending Match/NotMatch summary email...");

            var toList = AppSettingsHelper.Get("EmailTO")
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            var ccList = AppSettingsHelper.Get("EmailCC")
                ?.Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(e => e.Trim());

            await SendEmailAsync(toList, emailSubject, emailBody, isHtml: true, ccList);

            _mainForm.HideLoader();
            _mainForm.Log("✅ Match/NotMatch summary email sent successfully.");
        }

        public async Task CalculateAndSendEmailAsync(string targetSheetNameToProcess)
        {
            await Task.Run(() => CalculateAndSendEmail(targetSheetNameToProcess));
        }

        public async Task CalculateAndSendEmail(string targetSheetNameToProcess)
        {
            _mainForm.ShowLoader();
            //TimeZoneInfo easternZone = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
            //DateTime usNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZone);
            //_mainForm.Log($"⏰ Current US (Eastern) time: {usNow}");

            //DateTime targetDate = CalculateTargetSheetDate(usNow);
            //string todaySheetName = targetDate.ToString("MM/dd", CultureInfo.InvariantCulture);

            _mainForm.Log($"📄 Target sheet date selected: {targetSheetNameToProcess}");

            // Step 3: Always process the previous *working day’s* sheet
            //string targetSheetNameToProcess = GetPreviousWorkingDaySheetName(targetDate);
            _mainForm.Log($"📊 Processing previous working day sheet: {targetSheetNameToProcess}");

            //// Check if it's after 5 PM
            //bool isAfterFivePM = usNow.TimeOfDay == new TimeSpan(5, 0, 0);

            //// Depending on the time, decide which sheet to calculate (today or yesterday's)
            //string targetSheetNameToProcess = isAfterFivePM ? todaySheetName : GetPreviousSheetName(todaySheetName);

            // Retrieve data from the selected sheet (team name => record count)
            //var teamRecordCounts = GetTeamRecordCounts(targetSheetNameToProcess);
            var teamRecordCounts = await GetTeamRecordCountsAsync(targetSheetNameToProcess);


            // Send email
            _mainForm.Log("📧 Sending email with calculated data...");
            await SendEmailWithCalculatedData(teamRecordCounts, targetSheetNameToProcess);
            _mainForm.HideLoader();
            _mainForm.Log("✅ Email sent successfully.");
        }

        private string GetPreviousWorkingDaySheetName(DateTime currentTargetDate)
        {
            DateTime previousDate = currentTargetDate.AddDays(-1);

            // Skip weekends
            while (previousDate.DayOfWeek == DayOfWeek.Saturday || previousDate.DayOfWeek == DayOfWeek.Sunday)
            {
                previousDate = previousDate.AddDays(-1);
            }

            return previousDate.ToString("MM/dd", CultureInfo.InvariantCulture);
        }

        public async Task<Dictionary<string, int>> GetTeamRecordCountsAsync(string sheetName)
        {
            return await Task.Run(() => GetTeamRecordCounts(sheetName));
        }
        private async Task<Dictionary<string, int>> GetTeamRecordCounts(string sheetName)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
            {
                ["Mikhail"] = 0,
                ["Amurta"] = 0,
                ["Sarah"] = 0,
                ["Krina"] = 0,
                ["Patrizia"] = 0,
                ["Amanda"] = 0
            };

            try
            {
                _mainForm.ShowLoader();

                var sheetsService = _mainForm.SheetsService;
                var range = $"'{sheetName}'!A1:Z1000";
                _mainForm.Log($"📄 Reading data from: {range}");

                var request = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
                var response = await request.ExecuteAsync();
                var values = response.Values;

                if (values == null || values.Count == 0)
                {
                    _mainForm.Log($"❌ No data found in sheet '{sheetName}'.");
                    return result;
                }

                // Our known team list
                var teamNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Mikhail", "Amurta", "Sarah", "Krina", "Patrizia", "Amanda"
        };

                string currentTeam = null;
                bool inDataSection = false;

                foreach (var row in values)
                {
                    if (row == null || row.Count == 0)
                        continue;

                    string firstCell = row[0]?.ToString().Trim() ?? "";

                    if (string.IsNullOrEmpty(firstCell))
                        continue;

                    // --- Detect team header by name match ---
                    if (teamNames.Contains(firstCell, StringComparer.OrdinalIgnoreCase))
                    {
                        currentTeam = firstCell;
                        inDataSection = false;
                        _mainForm.Log($"📍 Found team: {currentTeam}");
                        continue;
                    }

                    // --- Detect the "CASE #" header row (signals start of data section) ---
                    bool isCaseHeader = row.Any(c =>
                        c != null &&
                        c.ToString().Trim().Equals("CASE #", StringComparison.OrdinalIgnoreCase));

                    if (isCaseHeader)
                    {
                        inDataSection = true;
                        continue;
                    }

                    // --- Count data rows ---
                    if (inDataSection && !string.IsNullOrEmpty(currentTeam))
                    {
                        // Detect valid data by checking if any cell looks like a numeric CASE #
                        bool hasCaseNumber = row.Any(c =>
                            int.TryParse(c?.ToString().Trim() ?? "", out _));

                        if (hasCaseNumber)
                        {
                            result[currentTeam]++;
                        }
                    }
                }

                // --- Log summary ---
                _mainForm.Log("✅ Calculated Data Summary");
                foreach (var kv in result)
                {
                    _mainForm.Log($"📋 {kv.Key}\nTotal Records: {kv.Value}\n");
                }

                int overallTotal = result.Values.Sum();
                _mainForm.Log($"📊 Overall Total Records: {overallTotal}");
            }
            catch (Exception ex)
            {
                _mainForm.Log($"❌ Error reading counts: {ex.Message}");
            }
            finally
            {
                _mainForm.HideLoader();
            }

            return result;
        }

        private async Task SendEmailWithCalculatedData(Dictionary<string, int> teamRecordCounts, string targetSheetNameToProcess)
        {
            _mainForm.ShowLoader();
            var sb = new StringBuilder();

            sb.AppendLine("<p>Hello,</p>");
            sb.AppendLine($"<p>This is to notify you that we have finalized the ISG Peer reviews for date: <strong>{targetSheetNameToProcess}</strong> summary and the brief details are as below:</p>");
            sb.AppendLine("<br>");
            sb.AppendLine("<h2>📊 Calculated Data Summary</h2>");

            if (teamRecordCounts == null || teamRecordCounts.Count == 0)
            {
                sb.AppendLine("<p><strong>No team data found for this date.</strong></p>");
            }
            else
            {
                int grandTotal = 0;

                foreach (var team in teamRecordCounts)
                {
                    sb.AppendLine($"<h3>📋 {team.Key}</h3>");
                    sb.AppendLine($"<p><strong>Total Records:</strong> {team.Value}</p>");
                    sb.AppendLine("<br>");
                    grandTotal += team.Value;
                }

                sb.AppendLine("<hr>");
                sb.AppendLine($"<h3>📊 <strong>Overall Total Records:</strong> {grandTotal}</h3>");
            }

            string emailSubject = "✅ Calculated Data Summary Report";
            string emailBody = sb.ToString();

            _mainForm.Log("📧 Sending formatted HTML email...");

            var toList = AppSettingsHelper.Get("EmailTO")
                    .Split(',', StringSplitOptions.RemoveEmptyEntries)
                    .Select(e => e.Trim());

            var ccList = AppSettingsHelper.Get("EmailCC")
                            ?.Split(',', StringSplitOptions.RemoveEmptyEntries)
                            .Select(e => e.Trim());

            await SendEmailAsync(toList, emailSubject, emailBody, isHtml: true, ccList);

            _mainForm.HideLoader();
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
		private string GetBaseFolderPath()
		{
			// TXT file path
			string txtPath = Path.Combine(Application.StartupPath,"basepath.txt");

			// Default fallback
			string defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),"ISG_Messages");

			// 1st Priority → TXT
			if (File.Exists(txtPath))
			{
				string txtValue = File.ReadAllText(txtPath).Trim();

				if (!string.IsNullOrEmpty(txtValue))
					return txtValue;
			}
			// Last → Default
			return defaultPath;
		}
		private string? FindDoctorFolder(string basePath, string doctorName)
		{
			if (!Directory.Exists(basePath))
				return null;

			// Clean search name
			string cleanDoctor = CleanNameFolder(doctorName);

			foreach (var dir in Directory.GetDirectories(basePath))
			{
				string folderName = Path.GetFileName(dir);

				string cleanFolder = CleanNameFolder(folderName);

				// Partial match
				if (cleanFolder.Contains(cleanDoctor))
				{
					return dir; // Found
				}
			}

			return null; // Not found
		}
		private string CleanNameFolder(string name)
		{
			name = name.ToLower();

			// Remove common titles
			name = Regex.Replace(name, @"\b(dr|md)\b", "", RegexOptions.IgnoreCase);

			// Remove special chars
			name = Regex.Replace(name, @"[^a-z\s]", "");

			// Remove extra spaces
			name = Regex.Replace(name, @"\s+", " ").Trim();

			return name;
		}

		private readonly object _statusLock = new object();

		public void LogPdfStatus(string caseNo, string fileName, string status)
		{
			try
			{
				string baseDir = Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
					"InvoiceAttachments",
					"Logs");

				Directory.CreateDirectory(baseDir);

				string statusFile =
					Path.Combine(baseDir, $"ProcessStatus_{DateTime.Now:yyyyMMdd}.txt");

				string line =
					$"{DateTime.Now:dd/MM/yyyy HH:mm:ss} - ISG {caseNo} - {fileName} - {status}\r\n";

				lock (_statusLock)
				{
					File.AppendAllText(statusFile, line);
				}
			}
			catch
			{
				// silent
			}
		}

	}
}
