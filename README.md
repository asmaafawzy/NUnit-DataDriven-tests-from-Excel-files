# NUnit-DataDriven-tests-from-Excel-files

The purpose of Unit Testing is to validate that each unit of the software works as expected, so we're gonna though Nunit which is the most popular unit test framework for .NET and know how to read data from excel file and use this data through Nunit attribute(TestCaseSource). Let's start ;)

```c#
 public class ExcelReader
        {
            public static IEnumerable<TestCaseData> ReadFromExcel(string excelFileName, string excelsheetTabName)
            {
                
                string executableLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string xslLocation = Path.Combine(executableLocation, "data/"+ excelFileName);
        
                string cmdText = "SELECT * FROM [" + excelsheetTabName + "$]";
                if (!File.Exists(xslLocation))
                    throw new Exception(string.Format("File name: {0}", xslLocation), new FileNotFoundException());
                string connectionStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES\";", xslLocation);
                var testCases = new List<TestCaseData>();
                using (var connection = new OleDbConnection(connectionStr))
                {
                    connection.Open();
                    var command = new OleDbCommand(cmdText, connection);
                    var reader = command.ExecuteReader();
                    if (reader == null)
                        throw new Exception(string.Format("No data return from file, file name:{0}", xslLocation));
                    while (reader.Read())
                    {
                        var row = new List<string>();
                        var feildCnt = reader.FieldCount;
                        for (var i = 0; i < feildCnt; i++)
                            row.Add(reader.GetValue(i).ToString());
                        testCases.Add(new TestCaseData(row.ToArray()));
                    }
                }

                if (testCases != null)
                    foreach (TestCaseData testCaseData in testCases)
                        yield return testCaseData;            
            }
        }    
```
