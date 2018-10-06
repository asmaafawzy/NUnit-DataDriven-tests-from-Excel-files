# NUnit-DataDriven-tests-from-Excel-files
How to read Data from excel file and pass it through nunit attribute (TestCaseSource)

```c#
 if (testCases != null)
                    foreach (TestCaseData testCaseData in testCases)
                        yield return testCaseData;      
```
