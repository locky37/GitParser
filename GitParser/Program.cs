
namespace GitParser;

internal class Program
{
    static void Main(string[] args)
    {
        // Использование
        // @"C:\Path\To\Repository", @"C:\Path\To\Report.docx"
        CommitReportGenerator.GenerateReport(@$"{args[0]}", @$"{args[1]}");
    }

}
