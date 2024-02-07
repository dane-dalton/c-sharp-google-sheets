using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Numerics;

namespace GoogleSheetsAndCsharp
{
  class Program
  {

    static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
    static readonly string ApplicationName = "Google Sheets API Test";
    static readonly string SpreadsheetId = "1nvlcBzrto0Y2ULvykmT40Mkt-4TP-kFu9jAZyK70-w8";
    static readonly string SpreadsheetHoursId = "1HkbIHtNdF5N9uYSAAbQ6Mj5ZHJuyQchel_YhcVBuQ5s";
    static readonly string sheet = "Schedule";
    static readonly string hoursSheet = "Test";
    static SheetsService service;


    static void Main(string[] args)
    {
      GoogleCredential credential;
      using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
      {
        credential = GoogleCredential.FromStream(stream)
          .CreateScoped(Scopes);
      }

      service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
      {
        HttpClientInitializer = credential,
        ApplicationName = ApplicationName,
      });

      CreateEntries(ReadEntries());
    }

    static Dictionary<string, int[]> ReadEntries()
    {
      Dictionary<string, int[]> classesPerTA = new Dictionary<string, int[]>();
      Dictionary<string, int> classMonthlyValue = new Dictionary<string, int>()
      {
        { "Monday0", 4 },
        { "Tuesday0", 4 },
        { "Friday0", 4 },
        { "Saturday0", 4 },
        { "Monday1", 4 },
        { "Tuesday1", 4 },
        { "Wednesday1", 4 },
        { "Thursday1", 4 },
        { "Saturday1", 4 },
        { "Monday2", 3 },
        { "Tuesday2", 3 },
        { "Wednesday2", 3 },
        { "Thursday2", 3 },
        { "Monday3", 3 },
        { "Tuesday3", 3 },
        { "Wednesday3", 3 },
        { "Thursday3", 3 },
        { "Monday4", 2 },
        { "Tuesday4", 2 },
        { "Wednesday4", 2 },
        { "Thursday4", 2 },
        { "Friday4", 2 },
        { "Monday5", 2 },
        { "Wednesday5", 2 },
        { "Thursday5", 2 },
      };

      var range = $"{sheet}!C2:H38";
      var request = service.Spreadsheets.Values.Get(SpreadsheetId, range);

      var response = request.Execute();
      var values = response.Values;
      if(values != null && values.Count > 0)
      {
        foreach(var row in values)
        {
          string classKey = $"{row[0]}{row[2]}";
          string leadTAName = row[4].ToString();
          string taName = row[5].ToString();

          if (!classesPerTA.ContainsKey(leadTAName))
          {
            classesPerTA.Add(leadTAName, new int[] { classMonthlyValue[classKey], 0 });
          }
          else
          {
            classesPerTA[leadTAName][0] += classMonthlyValue[classKey];
          }

          if (!classesPerTA.ContainsKey(taName))
          {
            classesPerTA.Add(taName, new int[] { 0, classMonthlyValue[classKey] });
          }
          else
          {
            classesPerTA[taName][1] += classMonthlyValue[classKey];
          }

          //Console.WriteLine(row[0]);
          //CreateEntries(row[0]);
        }
      }
      else
      {
        Console.WriteLine("No data found");
      }

      return classesPerTA;
    }

    static void CreateEntries(Dictionary<string, int[]> taHours)
    {
      var range = $"{hoursSheet}!A2:C2";
      var valueRange = new ValueRange();

      foreach(KeyValuePair<string, int[]> element in taHours)
      {
      var objectList = new List<object>() { element.Key, element.Value[0], element.Value[1] };
      valueRange.Values = new List<IList<object>> { objectList };

      var appendRequest = service.Spreadsheets.Values.Append(valueRange, SpreadsheetHoursId, range);
      appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
      var appendResponse = appendRequest.Execute();
      }

    }
  }
}