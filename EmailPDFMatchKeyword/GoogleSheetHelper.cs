using Google.Apis.Sheets.v4;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailPDFMatchKeyword
{
  public class GoogleSheetHelper
  {
    private readonly SheetsService _sheetsService;
    private readonly string _spreadsheetId;

    public GoogleSheetHelper(SheetsService sheetsService, string spreadsheetId)
    {
      _sheetsService = sheetsService ?? throw new ArgumentNullException(nameof(sheetsService));
      _spreadsheetId = spreadsheetId ?? throw new ArgumentNullException(nameof(spreadsheetId));
    }

    public string SpreadsheetId => _spreadsheetId;
  }
}
