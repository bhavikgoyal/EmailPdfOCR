using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EmailPDFMatchKeyword
{
  public static class SqliteHelper
  {
    private static Dictionary<string, HashSet<string>> ProviderGroups;

    // Static constructor to load config once when the class is first used
    static SqliteHelper()
    {
      LoadProviderGroupsFromConfig();
    }

    private static void LoadProviderGroupsFromConfig()
    {
      var builder = new ConfigurationBuilder()
          .SetBasePath(Directory.GetCurrentDirectory()) // Make sure your appsettings.json is here
          .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

      IConfigurationRoot configuration = builder.Build();

      // Read "ProviderGroups" section into dictionary<string, List<string>>
      var groups = configuration.GetSection("ProviderGroups").Get<Dictionary<string, List<string>>>();

      if (groups == null)
      {
        throw new Exception("Failed to load 'ProviderGroups' from appsettings.json.");
      }

      // Convert to dictionary<string, HashSet<string>> for fast lookup (case-insensitive)
      ProviderGroups = groups.ToDictionary(
          kvp => kvp.Key,
          kvp => new HashSet<string>(kvp.Value ?? new List<string>(), StringComparer.OrdinalIgnoreCase),
          StringComparer.OrdinalIgnoreCase);
    }

    public static void InsertCopyTemplateSheet(string provider, string caseNumber, string claimantName, string incidentDate, int pages, string matchStatus, string scribeTeam, string targetDate)
    {
      try
      {

        string tableName = GetTableNameByProvider(provider);
        Console.WriteLine($"DEBUG: Provider '{provider}' maps to table '{tableName}'");

        if (string.IsNullOrEmpty(tableName))
        {
          Console.WriteLine($"❌ Provider '{provider}' not found in any group. Insert aborted.");
          return;
        }

        string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "EmailPDFMatchKeyword_DB.db");
        string connectionString = $"Data Source={dbPath};";

        using (var connection = new SqliteConnection(connectionString))
        {
          connection.Open();

          // Create table if it does not exist
          string createTableQuery = $@"  CREATE TABLE IF NOT EXISTS [{tableName}] ( No INTEGER PRIMARY KEY AUTOINCREMENT, Initials TEXT, DATE TEXT, PROVIDER TEXT, [SCRIBE TEAM] TEXT, DOA TEXT, VENDOR TEXT, [CASE #] INTEGER, [CLAIMANT NAME] TEXT, PAGES INTEGER, NOTES TEXT, [DATE SUBMITTED] TEXT, [TIME SUBMITTED] TEXT, [YES/NO] TEXT, STATUS TEXT );";

          using (var createCmd = new SqliteCommand(createTableQuery, connection))
          {
            createCmd.ExecuteNonQuery();
          }

          // Adjusted column names to match your example and added brackets for spaces
          string insertQuery = $@" INSERT INTO [{tableName}]  (Initials, DATE, PROVIDER, [SCRIBE TEAM], DOA, VENDOR, [CASE #], [CLAIMANT NAME], PAGES, NOTES, [DATE SUBMITTED], [TIME SUBMITTED], [YES/NO], STATUS) VALUES  (@Initials, @Date, @Provider, @ScribeTeam, @DOA, @Vendor, @CaseNumber, @ClaimantName, @Pages, @Notes, @DateSubmitted, @TimeSubmitted, @YesNo, @Status); ";

          using (var cmd = new SqliteCommand(insertQuery, connection))
          {
            cmd.Parameters.AddWithValue("@Initials", DBNull.Value);
            cmd.Parameters.AddWithValue("@Date", targetDate);
            cmd.Parameters.AddWithValue("@Provider", provider);
            cmd.Parameters.AddWithValue("@ScribeTeam", scribeTeam);
            cmd.Parameters.AddWithValue("@DOA", incidentDate);
            cmd.Parameters.AddWithValue("@Vendor", "ISG");
            cmd.Parameters.AddWithValue("@CaseNumber", caseNumber);
            cmd.Parameters.AddWithValue("@ClaimantName", claimantName);
            cmd.Parameters.AddWithValue("@Pages", pages);
            cmd.Parameters.AddWithValue("@Notes", DBNull.Value);
            cmd.Parameters.AddWithValue("@DateSubmitted", DBNull.Value);
            cmd.Parameters.AddWithValue("@TimeSubmitted", DBNull.Value);
            cmd.Parameters.AddWithValue("@YesNo", DBNull.Value);
            cmd.Parameters.AddWithValue("@Status", matchStatus);

            cmd.ExecuteNonQuery();
          }
        }
        Console.WriteLine($"✅ Data inserted successfully into {tableName} table!");
      }
      catch(Exception ex)
      {
        throw ex;
      }
    }

    // Returns the table name/group that contains the provider
    private static string GetTableNameByProvider(string provider)
    {
      foreach (var group in ProviderGroups)
      {
        if (group.Value.Contains(provider))
          return group.Key;
      }
      return null;
    }
  }
}
