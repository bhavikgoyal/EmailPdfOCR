using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EmailPDFMatchKeyword
{
  public class AppSettingsHelper
  {
    private static IConfigurationRoot _configuration;

    static AppSettingsHelper()
    {
      _configuration = new ConfigurationBuilder()
          .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
          .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
          .Build();
    }

    public static string Get(string key)
    {
      return _configuration[key];
    }

    public static T GetSection<T>(string sectionName)
    {
      return _configuration.GetSection(sectionName).Get<T>();
    }
  }
}
