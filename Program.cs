using ConsoleApp1;
using ExcelDataReader;
using System.Text;
using System.Text.RegularExpressions;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
string filePath = @"D:\Projects\acm.timus_artemiy_lessons\2023\from-table-to-oop-cs\ConsoleApp1\original-table.xlsx";

using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
    IExcelDataReader reader;
    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

    //// reader.IsFirstRowAsColumnNames
    var conf = new ExcelDataSetConfiguration
    {
        ConfigureDataTable = _ => new ExcelDataTableConfiguration
        {
            UseHeaderRow = true
        }
    };

    var dataSet = reader.AsDataSet(conf);

    // Now you can get data from each sheet by its index or its "name"
    var dataTable = dataSet.Tables[0];

    List<Language> languages = new List<Language>();

    for (var i = 0; i < dataTable.Rows.Count; i++)
    {
        var name = (string)dataTable.Rows[i][0];
        var year = dataTable.Rows[i][1].ToString();
        var author = (string)dataTable.Rows[i][2];
        var oop = (string)dataTable.Rows[i][3];
        var activeDev = (string)dataTable.Rows[i][4];
        var lastVersion = (string)dataTable.Rows[i][5];
        var lastVerDate = (string)dataTable.Rows[i][6];

        bool oopBool;
        if (oop == "нет") oopBool = false; else oopBool = true;

        bool activeDevBool;
        if (activeDev == "нет") activeDevBool = false; else activeDevBool = true;

        var lastVerDateSplitted = lastVerDate.Split('.');
        int day = Convert.ToInt32(lastVerDateSplitted[0]); 
        int month = Convert.ToInt32(lastVerDateSplitted[1]);
        int yearFor = Convert.ToInt32(lastVerDateSplitted[2]);
        DateOnly lastVerDateDate = new DateOnly(yearFor, month, day);

        var language = new Language();
        language.Name = name;
        language.Year = Convert.ToInt32(year);
        language.Author = author;
        language.Oop = oopBool;
        language.ActiveDev = activeDevBool;
        language.LastVersion = lastVersion;
        language.LastVerDate = lastVerDateDate;

        languages.Add(language);

    }

    foreach (Language i in languages)
    {
        Console.WriteLine("В {0}-том году придуман язык {1} \nА его автор {2}\n", i.Year, i.Name, i.Author);
    }

}