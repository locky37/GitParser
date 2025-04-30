using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LibGit2Sharp;

namespace GitParser;

class CommitReportGenerator
{
    public static void GenerateReport(string repoPath, string outputPath)
    {
        // Получение истории коммитов
        using var repo = new Repository(repoPath);
        var commits = repo.Commits.Take(150); // Берем последние 50 коммитов

        // Создание Word-документа
        using var wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = wordDoc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());

        var body = mainPart.Document.Body;

        // Добавление заголовка
        body.AppendChild(new Paragraph(new Run(new Text("Commit History Report")) { RunProperties = new RunProperties(new Bold()) }));

        // Добавление таблицы
        var table = new Table();

        // Заголовок таблицы
        var headerRow = new TableRow();
        headerRow.Append(
            CreateCell("Hash"),
            CreateCell("Author"),
            CreateCell("Date"),
            CreateCell("Message")
        );
        table.Append(headerRow);

        // Данные коммитов
        foreach (var commit in commits)
        {
            var row = new TableRow();
            row.Append(
                CreateCell(commit.Sha.Substring(0, 7)),
                CreateCell(commit.Author.Name),
                CreateCell(commit.Author.When.ToString("yyyy-MM-dd HH:mm:ss")),
                CreateCell(commit.MessageShort)
            );
            table.Append(row);
        }

        body.Append(table);
    }

    private static TableCell CreateCell(string text)
    {
        return new TableCell(new Paragraph(new Run(new Text(text))));
    }
}
