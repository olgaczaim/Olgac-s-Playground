//try and fail for sp db read and export files

using Microsoft.Data.SqlClient;
using Spectre.Console;
using System.Data;
using SimpleImpersonation;
using System.Security.Principal;
using Microsoft.Win32.SafeHandles;
using System.Reflection;
using System.Threading.Tasks;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Text;


AnsiConsole.Write(
    new FigletText("SP DB Explorer")
        .Centered()
        .Color(Color.Green));

AnsiConsole.MarkupLine("[grey]Read-only SharePoint content database explorer[/]");
AnsiConsole.WriteLine();

try
{
    var connectionInfo = PromptForConnectionString();

    var cobaltCoreDllPath = AnsiConsole.Prompt(
        new TextPrompt<string>("Microsoft.CobaltCore.dll path [grey](leave empty to disable CobaltCore export)[/]:")
            .DefaultValue(@"C:\Users\uname\SpContentDbExplorer\bin\Microsoft.CobaltCore.dll"))
        .Trim()
        .Trim('"');

    if (!string.IsNullOrWhiteSpace(cobaltCoreDllPath) && !File.Exists(cobaltCoreDllPath))
    {
        AnsiConsole.MarkupLine(
            $"[yellow]Microsoft.CobaltCore.dll was not found at:[/] [grey]{Markup.Escape(cobaltCoreDllPath)}[/]");
        AnsiConsole.MarkupLine("[yellow]CobaltCore export will be disabled for this run.[/]");
        cobaltCoreDllPath = "";
    }

    var scanner = new SharePointDatabaseScanner(
        connectionInfo.ConnectionString,
        connectionInfo.WindowsCredentials,
        string.IsNullOrWhiteSpace(cobaltCoreDllPath) ? null : cobaltCoreDllPath);

    await scanner.TestConnectionAsync();

    //if (AnsiConsole.Confirm("Inspect Microsoft.CobaltCore.dll before continuing?"))
    //{
    //    var cobaltDllPath = AnsiConsole.Ask<string>(
    //        "Full path to Microsoft.CobaltCore.dll:");
    //
    //    CobaltCoreInspector.Inspect(cobaltDllPath);
    //
    //    AnsiConsole.MarkupLine("[yellow]Paste the output here so we can wire the correct API.[/]");
    //}

    var mode = AnsiConsole.Prompt(
        new SelectionPrompt<string>()
            .Title("How do you want to select the SharePoint content database?")
            .AddChoices(
                "Scan all online databases",
                "Enter database name manually"));

    DatabaseCandidate? selectedDatabase = null;

    if (mode == "Scan all online databases")
    {
        List<DatabaseCandidate> candidates = new();

        await AnsiConsole.Status()
            .StartAsync("Scanning SQL Server databases...", async _ =>
            {
                candidates = await scanner.FindCandidateDatabasesAsync();
            });

        if (candidates.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]No likely SharePoint content databases found.[/]");
            return;
        }

        DisplayCandidates(candidates);

        selectedDatabase = AnsiConsole.Prompt(
            new SelectionPrompt<DatabaseCandidate>()
                .Title("Select a SharePoint content database")
                .PageSize(15)
                .UseConverter(db =>
                    $"{Markup.Escape(db.Name)} | {db.Confidence} | Sites: {db.SiteCollectionCount?.ToString() ?? "?"}")
                .AddChoices(candidates));
    }
    else
    {
        var dbName = AnsiConsole.Prompt(
          new TextPrompt<string>("Enter content database name:")
              .DefaultValue("WSS_Content"));


        await AnsiConsole.Status()
            .StartAsync($"Checking database [yellow]{dbName}[/]...", async _ =>
            {
                selectedDatabase = await scanner.InspectDatabaseAsync(dbName);
            });

        if (selectedDatabase == null)
        {
            AnsiConsole.MarkupLine("[red]Database not found, not accessible, or does not look like a SharePoint content database.[/]");
            return;
        }
    }

    AnsiConsole.MarkupLine($"[green]Selected database:[/] [yellow]{Markup.Escape(selectedDatabase.Name)}[/]");
    AnsiConsole.WriteLine();

    SharePointDatabaseSnapshot snapshot = new();

    await AnsiConsole.Status()
        .StartAsync("Reading site collections and webs...", async _ =>
        {
            snapshot = await scanner.ReadSitesAndWebsAsync(selectedDatabase.Name);
        });

    DisplaySitesAndWebs(selectedDatabase.Name, snapshot);

    AnsiConsole.WriteLine();

    if (snapshot.Webs.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]No webs found, so document libraries cannot be listed yet.[/]");
        return;
    }

    var selectedWeb = AnsiConsole.Prompt(
        new SelectionPrompt<WebInfo>()
            .Title("Select a site/web to list document libraries")
            .PageSize(15)
            .UseConverter(web =>
            {
                var path = string.IsNullOrWhiteSpace(web.FullUrl)
                    ? "/"
                    : "/" + web.FullUrl.Trim('/');

                var title = string.IsNullOrWhiteSpace(web.Title)
                    ? "(no title)"
                    : web.Title;

                return $"{Markup.Escape(title)} | {Markup.Escape(path)}";
            })
            .AddChoices(snapshot.Webs.OrderBy(w => w.FullUrl)));

    List<DocumentLibraryInfo> libraries = new();

    await AnsiConsole.Status()
        .StartAsync("Reading document libraries...", async _ =>
        {
            libraries = await scanner.ReadDocumentLibrariesForWebAsync(
                selectedDatabase.Name,
                selectedWeb.Id);
        });

    DisplayDocumentLibraries(selectedWeb, libraries);

    AnsiConsole.WriteLine();

    if (libraries.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]No document libraries found.[/]");
        return;
    }

    var selectedLibrary = AnsiConsole.Prompt(
        new SelectionPrompt<DocumentLibraryInfo>()
            .Title("Select a document library")
            .PageSize(15)
            .UseConverter(library =>
            {
                var title = string.IsNullOrWhiteSpace(library.Title)
                    ? "(no title)"
                    : library.Title;

                return $"{Markup.Escape(title)} | Template: {library.ServerTemplate} | Items: {library.ItemCount?.ToString() ?? "?"}";
            })
            .AddChoices(libraries.OrderBy(l => l.Title)));

    await BrowseDocumentLibraryAsync(
        scanner,
        selectedDatabase.Name,
        selectedWeb,
        selectedLibrary);

    AnsiConsole.WriteLine();
    AnsiConsole.MarkupLine("[green]Done.[/]");


}
catch (Exception ex)
{
    AnsiConsole.WriteException(ex, ExceptionFormats.ShortenEverything);
}

static (string ConnectionString, UserCredentials? WindowsCredentials) PromptForConnectionString()
{

    var server = AnsiConsole.Prompt(
          new TextPrompt<string>("SQL Server name or instance:")
              .DefaultValue(@"instance\name"));


    var useCurrentWindowsAuth = true;// AnsiConsole.Confirm("Use current Windows Authentication?");

    UserCredentials? windowsCredentials = null;

    if (!useCurrentWindowsAuth)
    {
        AnsiConsole.MarkupLine("[yellow]Enter domain Windows credentials for NTLM/Kerberos authentication.[/]");

        var domain = AnsiConsole.Prompt(
          new TextPrompt<string>("Domain:")
              .DefaultValue(Path.Combine(Environment.CurrentDirectory, "aaaa")));

        var username = AnsiConsole.Prompt(
          new TextPrompt<string>("Username:")
              .DefaultValue(Path.Combine(Environment.CurrentDirectory, "bbb")));

        var password = AnsiConsole.Prompt(
            new TextPrompt<string>("Password:")
                .Secret());

        windowsCredentials = new UserCredentials(domain, username, password);
    }

    var trustServerCertificate = AnsiConsole.Confirm("Trust SQL Server certificate?");

    var builder = new SqlConnectionStringBuilder
    {
        DataSource = server,
        InitialCatalog = "master",
        ApplicationName = "SpContentDbExplorer",
        ConnectTimeout = 15,
        TrustServerCertificate = trustServerCertificate,
        IntegratedSecurity = true,
        Pooling = false
    };

    builder["Encrypt"] = "True";

    return (builder.ConnectionString, windowsCredentials);
}


static void DisplayDocumentLibraries(WebInfo web, List<DocumentLibraryInfo> libraries)
{
    AnsiConsole.WriteLine();

    var webPath = string.IsNullOrWhiteSpace(web.FullUrl)
        ? "/"
        : "/" + web.FullUrl.Trim('/');

    AnsiConsole.MarkupLine(
        $"[bold green]Document libraries for:[/] [yellow]{Markup.Escape(web.Title)}[/] [grey]{Markup.Escape(webPath)}[/]");

    if (libraries.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]No document libraries found for this web.[/]");
        return;
    }

    var table = new Table()
        .Border(TableBorder.Rounded)
        .Title("Document Libraries");

    table.AddColumn("Title");
    table.AddColumn("Template");
    table.AddColumn("Items");
    table.AddColumn("Created");
    table.AddColumn("Modified");
    table.AddColumn("List Id");

    foreach (var library in libraries)
    {
        var modified = library.Modified ?? library.LastItemModifiedDate;

        table.AddRow(
            Markup.Escape(string.IsNullOrWhiteSpace(library.Title) ? "(no title)" : library.Title),
            library.ServerTemplate.ToString(),
            library.ItemCount?.ToString() ?? "",
            library.Created?.ToString("yyyy-MM-dd HH:mm") ?? "",
            modified?.ToString("yyyy-MM-dd HH:mm") ?? "",
            Markup.Escape(library.Id.ToString()));
    }

    AnsiConsole.Write(table);
}

static void DisplayCandidates(List<DatabaseCandidate> candidates)
{
    var table = new Table()
        .Border(TableBorder.Rounded)
        .Title("Candidate SharePoint Content Databases");

    table.AddColumn("Database");
    table.AddColumn("Matched Tables");
    table.AddColumn("Site Collections");
    table.AddColumn("Confidence");

    foreach (var db in candidates)
    {
        table.AddRow(
            Markup.Escape(db.Name),
            db.MatchedTableCount.ToString(),
            db.SiteCollectionCount?.ToString() ?? "?",
            db.Confidence);
    }

    AnsiConsole.Write(table);
    AnsiConsole.WriteLine();
}

static void DisplaySitesAndWebs(string databaseName, SharePointDatabaseSnapshot snapshot)
{
    if (snapshot.SiteCollections.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]No site collections found.[/]");
        return;
    }

    var tree = new Tree($"[bold green]Database:[/] [yellow]{Markup.Escape(databaseName)}[/]");

    foreach (var site in snapshot.SiteCollections.OrderBy(x => x.FullUrl))
    {
        var sitePath = BuildDisplayPath(site.FullUrl, null);

        var siteNode = tree.AddNode(
            $"[bold blue]Site Collection:[/] {Markup.Escape(sitePath)} [grey]({site.Id})[/]");

        var webs = snapshot.Webs
            .Where(w => w.SiteId == site.Id)
            .OrderBy(w => w.FullUrl)
            .ToList();

        foreach (var web in webs)
        {
            var webPath = BuildDisplayPath(site.FullUrl, web.FullUrl);
            var title = string.IsNullOrWhiteSpace(web.Title)
                ? "(no title)"
                : web.Title;

            siteNode.AddNode(
                $"[green]{Markup.Escape(title)}[/] [grey]{Markup.Escape(webPath)}[/]");
        }
    }

    AnsiConsole.Write(tree);
}

static void DisplayFolderContents(List<DocumentItemInfo> items)
{
    if (items.Count == 0)
    {
        AnsiConsole.MarkupLine("[yellow]This folder is empty.[/]");
        return;
    }

    var table = new Table()
        .Border(TableBorder.Rounded)
        .Title("Folder Contents");

    table.AddColumn("Type");
    table.AddColumn("Name");
    table.AddColumn("Ext");
    table.AddColumn("Size");
    table.AddColumn("Modified");
    table.AddColumn("Version");
    table.AddColumn("Id");

    foreach (var item in items
                 .OrderByDescending(x => x.IsFolder)
                 .ThenBy(x => x.LeafName))
    {
        table.AddRow(
     item.IsFolder ? "[blue]Folder[/]" : "[green]File[/]",
     Markup.Escape(item.LeafName),
     Markup.Escape(item.DisplayExtension),
     item.IsFolder ? "" : FormatBytes(item.SizeBytes),
     item.TimeLastModified?.ToString("yyyy-MM-dd HH:mm") ?? "",
     Markup.Escape(item.UIVersionString ?? ""),
     Markup.Escape(item.Id.ToString()));
    }

    AnsiConsole.Write(table);
}

static string FormatBytes(long? bytes)
{
    if (bytes == null)
        return "";

    double value = bytes.Value;
    string[] suffixes = { "B", "KB", "MB", "GB", "TB" };

    var index = 0;

    while (value >= 1024 && index < suffixes.Length - 1)
    {
        value /= 1024;
        index++;
    }

    return $"{value:0.##} {suffixes[index]}";
}

static async Task BrowseDocumentLibraryAsync(
    SharePointDatabaseScanner scanner,
    string databaseName,
    WebInfo web,
    DocumentLibraryInfo library)
{
    AnsiConsole.WriteLine();

    AnsiConsole.MarkupLine(
        $"[bold green]Opening library:[/] [yellow]{Markup.Escape(library.Title)}[/]");

    var rootFolder = await scanner.FindLibraryRootFolderAsync(databaseName, library.Id);

    if (rootFolder == null)
    {
        AnsiConsole.MarkupLine("[red]Could not find the root folder for this document library in AllDocs.[/]");
        return;
    }

    var folderStack = new Stack<DocumentItemInfo>();
    var currentFolder = rootFolder;

    while (true)
    {
        AnsiConsole.WriteLine();

        var currentPath = currentFolder.FullPath;

        AnsiConsole.MarkupLine(
            $"[bold blue]Current folder:[/] [yellow]{Markup.Escape(currentPath)}[/]");

        List<DocumentItemInfo> children = new();

        await AnsiConsole.Status()
            .StartAsync("Reading folder contents...", async _ =>
            {
                children = await scanner.ReadChildItemsAsync(
                    databaseName,
                    library.Id,
                    currentFolder.Id);
            });

        DisplayFolderContents(children);

        var choices = new List<ExplorerChoice>();

        if (folderStack.Count > 0)
        {
            choices.Add(ExplorerChoice.Up());
        }

        choices.Add(ExplorerChoice.Exit());

        foreach (var child in children
                     .OrderByDescending(x => x.IsFolder)
                     .ThenBy(x => x.LeafName))
        {
            choices.Add(ExplorerChoice.FromItem(child));
        }

        var selected = AnsiConsole.Prompt(
            new SelectionPrompt<ExplorerChoice>()
                .Title("Select folder/file")
                .PageSize(20)
                .UseConverter(choice => choice.DisplayText)
                .AddChoices(choices));

        if (selected.Kind == ExplorerChoiceKind.Exit)
        {
            break;
        }

        if (selected.Kind == ExplorerChoiceKind.Up)
        {
            currentFolder = folderStack.Pop();
            continue;
        }

        if (selected.Item == null)
            continue;

        if (selected.Item.IsFolder)
        {
            folderStack.Push(currentFolder);
            currentFolder = selected.Item;
            continue;
        }

        var exportFolder = AnsiConsole.Prompt(
            new TextPrompt<string>("Export folder:")
                .DefaultValue(Path.Combine(Environment.CurrentDirectory, "exports")));

        var showStreamDiagnostics = AnsiConsole.Confirm("Show stream diagnostics before export?", false);

        await scanner.ExportCurrentFileAsync(
            databaseName,
            selected.Item,
            exportFolder,
            showStreamDiagnostics);
    }
}



static string BuildDisplayPath(string? siteFullUrl, string? webFullUrl)
{
    var parts = new List<string>();

    if (!string.IsNullOrWhiteSpace(siteFullUrl))
        parts.Add(siteFullUrl!.Trim('/'));

    if (!string.IsNullOrWhiteSpace(webFullUrl))
        parts.Add(webFullUrl!.Trim('/'));

    var path = string.Join("/", parts);

    return string.IsNullOrWhiteSpace(path)
        ? "/"
        : "/" + path;
}

public sealed class SharePointDatabaseScanner
{

    private async Task DumpShreddedRowsAsync(
        SqlConnection connection,
        string quotedDatabaseName,
        DocumentItemInfo item,
        string exportFolder)
    {
        var rows = await ReadShreddedStreamRowsAsync(
            connection,
            quotedDatabaseName,
            item);

        var dumpFolder = Path.Combine(
            exportFolder,
            MakeSafeFileName(item.LeafName) + "_shred_dump");

        Directory.CreateDirectory(dumpFolder);

        foreach (var row in rows)
        {
            var fileName =
                $"hist_{row.HistVersion}_type_{row.Type}_stream_{row.StreamId}_bsn_{row.BSN}.bin";

            var path = Path.Combine(dumpFolder, fileName);

            await Compat.WriteAllBytesAsync(path, row.Content);
        }

        AnsiConsole.MarkupLine(
            $"[grey]Shred dump written to:[/] [yellow]{Markup.Escape(dumpFolder)}[/]");
    }
    private async Task<List<ShreddedStreamRow>> ReadShreddedStreamRowsAsync(
        SqlConnection connection,
        string quotedDatabaseName,
        DocumentItemInfo item)
    {
        var rows = new List<ShreddedStreamRow>();

        using var command = connection.CreateCommand();

        command.CommandText = $@"
DECLARE @latestHistVersion int;

SELECT @latestHistVersion = MAX(HistVersion)
FROM {quotedDatabaseName}.dbo.DocsToStreams
WHERE SiteId = @siteId
  AND DocId = @docId;

SELECT
    dts.HistVersion,
    dts.Partition,
    dts.BSN,
    dts.StreamId,
    ds.Type,
    ds.Size,
    ds.Content
FROM {quotedDatabaseName}.dbo.DocsToStreams dts
INNER JOIN {quotedDatabaseName}.dbo.DocStreams ds
    ON ds.SiteId = dts.SiteId
   AND ds.DocId = dts.DocId
   AND ds.Partition = dts.Partition
   AND ds.BSN = dts.BSN
WHERE dts.SiteId = @siteId
  AND dts.DocId = @docId
  AND dts.HistVersion = @latestHistVersion
ORDER BY
    dts.StreamId,
    dts.Partition,
    dts.BSN;";

        command.Parameters.Add("@siteId", SqlDbType.UniqueIdentifier).Value = item.SiteId;
        command.Parameters.Add("@docId", SqlDbType.UniqueIdentifier).Value = item.Id;

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            if (reader["Content"] == DBNull.Value)
                continue;

            rows.Add(new ShreddedStreamRow
            {
                HistVersion = Convert.ToInt32(reader["HistVersion"]),
                Partition = Convert.ToInt32(reader["Partition"]),
                BSN = Convert.ToInt64(reader["BSN"]),
                StreamId = Convert.ToInt64(reader["StreamId"]),
                Type = Convert.ToInt32(reader["Type"]),
                Size = Convert.ToInt32(reader["Size"]),
                Content = (byte[])reader["Content"]
            });
        }

        return rows;
    }

    private async Task<ShreddedExportResult> AnalyzeSharePointShreddedStorageAsync(
    SqlConnection connection,
    string quotedDatabaseName,
    DocumentItemInfo item,
    string outputPath)
    {
        var rows = await ReadShreddedStreamRowsAsync(
            connection,
            quotedDatabaseName,
            item);

        if (rows.Count == 0)
        {
            return new ShreddedExportResult
            {
                Success = false,
                Message = "No DocStreams rows were found for this file."
            };
        }

        var type10Rows = rows.Where(x => x.Type == 10).ToList();
        var type11Rows = rows.Where(x => x.Type == 11).ToList();

        var totalRawBytes = rows.Sum(x => (long)x.Content.Length);
        var totalType10Bytes = type10Rows.Sum(x => (long)x.Content.Length);
        var totalType11Bytes = type11Rows.Sum(x => (long)x.Content.Length);

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("Shredded Storage Parser Status");

        table.AddColumn("Property");
        table.AddColumn("Value");

        table.AddRow("Total rows", rows.Count.ToString());
        table.AddRow("Type 10 rows", type10Rows.Count.ToString());
        table.AddRow("Type 11 rows", type11Rows.Count.ToString());
        table.AddRow("Raw total bytes", totalRawBytes.ToString("N0"));
        table.AddRow("Type 10 bytes", totalType10Bytes.ToString("N0"));
        table.AddRow("Type 11 bytes", totalType11Bytes.ToString("N0"));
        table.AddRow("Expected file size", item.SizeBytes?.ToString("N0") ?? "?");

        AnsiConsole.Write(table);

        if (rows.Count == 1 && type11Rows.Count == 0 && item.SizeBytes.HasValue)
        {
            var onlyRow = rows[0];

            if (onlyRow.Content.Length == item.SizeBytes.Value)
            {
                await Compat.WriteAllBytesAsync(outputPath, onlyRow.Content);

                return new ShreddedExportResult
                {
                    Success = true,
                    OutputPath = outputPath,
                    BytesWritten = onlyRow.Content.Length,
                    Message = "Exported single raw stream row."
                };
            }
        }

        var workFolder = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory,
            MakeSafeFileName(item.LeafName) + "_cobalt_work");

        Directory.CreateDirectory(workFolder);

        var databaseMappingSearchPath = await WriteCobaltDatabaseMappingSearchAsync(
            connection,
            quotedDatabaseName,
            item,
            workFolder);

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt database mapping search written to:[/] [yellow]{Markup.Escape(databaseMappingSearchPath)}[/]");

        var formatNeutralResult = TryExportFormatNeutralCobaltPayload(
            rows,
            item,
            outputPath,
            workFolder);

        if (formatNeutralResult.Success)
            return formatNeutralResult;

        var embeddedPdfResult = TryExportEmbeddedPdfPayload(
            rows,
            item,
            outputPath);

        if (embeddedPdfResult.Success)
            return embeddedPdfResult;

        var embeddedPdfMessage = string.Equals(item.DisplayExtension, "pdf", StringComparison.OrdinalIgnoreCase)
            ? embeddedPdfResult.Message
            : "";

        await DumpShreddedRowsAsync(
    connection,
    quotedDatabaseName,
    item,
    Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory);

        var cobaltCoreDllPath = _cobaltCoreDllPath;

        if (string.IsNullOrWhiteSpace(cobaltCoreDllPath))
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = JoinMessages(
                    formatNeutralResult.Message,
                    embeddedPdfMessage,
                    "This file uses SharePoint Cobalt/shredded storage. Microsoft.CobaltCore.dll path was not provided, so the app can only dump raw shreds and recovery diagnostics.")
            };
        }

        var adapter = new MicrosoftCobaltCoreAdapter(cobaltCoreDllPath!);

        var cobaltResult = adapter.TryExportFromSharePointDbRows(
            rows,
            item,
            outputPath,
            workFolder);

        if (!cobaltResult.Success && !string.IsNullOrWhiteSpace(embeddedPdfMessage))
        {
            cobaltResult.Message = JoinMessages(
                formatNeutralResult.Message,
                embeddedPdfMessage,
                cobaltResult.Message);
        }
        else if (!cobaltResult.Success)
        {
            cobaltResult.Message = JoinMessages(
                formatNeutralResult.Message,
                cobaltResult.Message);
        }

        return cobaltResult;

    }

    private static async Task<string> WriteCobaltDatabaseMappingSearchAsync(
        SqlConnection connection,
        string quotedDatabaseName,
        DocumentItemInfo item,
        string workFolder)
    {
        var outputPath = Path.Combine(workFolder, "cobalt_database_mapping_search.csv");
        var candidates = await ReadPotentialCobaltMappingObjectsAsync(connection, quotedDatabaseName);

        using var writer = new StreamWriter(outputPath, false, Encoding.UTF8);
        writer.WriteLine("Schema,Object,ObjectType,Columns,Filter,RowCount,SampleRows,Error");

        foreach (var candidate in candidates)
        {
            try
            {
                var filter = BuildSelectedDocumentFilter(candidate.Columns);
                var rowCount = await CountCandidateMappingRowsAsync(
                    connection,
                    quotedDatabaseName,
                    candidate,
                    filter,
                    item);

                var sampleRows = rowCount > 0
                    ? await ReadCandidateMappingSampleRowsAsync(
                        connection,
                        quotedDatabaseName,
                        candidate,
                        filter,
                        item)
                    : "";

                writer.WriteLine(
                    string.Join(",",
                        Csv(candidate.SchemaName),
                        Csv(candidate.ObjectName),
                        Csv(candidate.ObjectType),
                        Csv(string.Join("|", candidate.Columns)),
                        Csv(filter.Description),
                        rowCount,
                        Csv(sampleRows),
                        ""));
            }
            catch (Exception ex)
            {
                writer.WriteLine(
                    string.Join(",",
                        Csv(candidate.SchemaName),
                        Csv(candidate.ObjectName),
                        Csv(candidate.ObjectType),
                        Csv(string.Join("|", candidate.Columns)),
                        "",
                        "",
                        "",
                        Csv(GetUsefulExceptionMessage(ex))));
            }
        }

        return outputPath;
    }

    private static async Task<List<CobaltMappingObjectCandidate>> ReadPotentialCobaltMappingObjectsAsync(
        SqlConnection connection,
        string quotedDatabaseName)
    {
        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT
    s.name AS SchemaName,
    o.name AS ObjectName,
    o.type_desc AS ObjectType,
    c.name AS ColumnName
FROM {quotedDatabaseName}.sys.objects o
INNER JOIN {quotedDatabaseName}.sys.schemas s
    ON s.schema_id = o.schema_id
INNER JOIN {quotedDatabaseName}.sys.columns c
    ON c.object_id = o.object_id
WHERE o.type IN ('U', 'V')
  AND
  (
      c.name IN ('SiteId', 'DocId', 'StreamId', 'BSN', 'Partition', 'HistVersion', 'Size', 'Type', 'Content', 'RbsId')
      OR c.name LIKE '%Stream%'
      OR c.name LIKE '%Bsn%'
      OR c.name LIKE '%Blob%'
      OR c.name LIKE '%Ordinal%'
      OR c.name LIKE '%Sequence%'
      OR c.name LIKE '%Offset%'
      OR c.name LIKE '%Index%'
      OR c.name LIKE '%Map%'
      OR o.name LIKE '%Stream%'
      OR o.name LIKE '%Blob%'
      OR o.name LIKE '%Doc%'
      OR o.name LIKE '%Map%'
  )
ORDER BY s.name, o.name, c.column_id;";

        var lookup = new Dictionary<(string Schema, string Object), CobaltMappingObjectCandidate>();

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var schemaName = reader["SchemaName"]?.ToString() ?? "";
            var objectName = reader["ObjectName"]?.ToString() ?? "";
            var objectType = reader["ObjectType"]?.ToString() ?? "";
            var columnName = reader["ColumnName"]?.ToString() ?? "";

            var key = (schemaName, objectName);

            if (!lookup.TryGetValue(key, out var candidate))
            {
                candidate = new CobaltMappingObjectCandidate(schemaName, objectName, objectType);
                lookup[key] = candidate;
            }

            if (!candidate.Columns.Contains(columnName, StringComparer.OrdinalIgnoreCase))
                candidate.Columns.Add(columnName);
        }

        return lookup.Values
            .Where(x =>
                x.Columns.Contains("DocId", StringComparer.OrdinalIgnoreCase) ||
                x.Columns.Contains("StreamId", StringComparer.OrdinalIgnoreCase) ||
                x.Columns.Contains("BSN", StringComparer.OrdinalIgnoreCase) ||
                x.ObjectName.IndexOf("Map", StringComparison.OrdinalIgnoreCase) >= 0 ||
                x.ObjectName.IndexOf("Blob", StringComparison.OrdinalIgnoreCase) >= 0)
            .OrderByDescending(x => x.Columns.Contains("DocId", StringComparer.OrdinalIgnoreCase))
            .ThenByDescending(x => x.Columns.Contains("StreamId", StringComparer.OrdinalIgnoreCase))
            .ThenBy(x => x.SchemaName)
            .ThenBy(x => x.ObjectName)
            .ToList();
    }

    private static CobaltMappingFilter BuildSelectedDocumentFilter(List<string> columns)
    {
        var predicates = new List<string>();

        if (columns.Contains("SiteId", StringComparer.OrdinalIgnoreCase))
            predicates.Add("[SiteId] = @siteId");

        if (columns.Contains("DocId", StringComparer.OrdinalIgnoreCase))
            predicates.Add("[DocId] = @docId");

        if (predicates.Count == 0)
            return new CobaltMappingFilter("", "none");

        return new CobaltMappingFilter(
            "WHERE " + string.Join(" AND ", predicates),
            string.Join(" AND ", predicates));
    }

    private static async Task<long> CountCandidateMappingRowsAsync(
        SqlConnection connection,
        string quotedDatabaseName,
        CobaltMappingObjectCandidate candidate,
        CobaltMappingFilter filter,
        DocumentItemInfo item)
    {
        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT COUNT_BIG(*)
FROM {quotedDatabaseName}.{QuoteSqlIdentifier(candidate.SchemaName)}.{QuoteSqlIdentifier(candidate.ObjectName)}
{filter.Sql};";

        AddDocumentFilterParameters(command, filter, item);

        return Convert.ToInt64(await command.ExecuteScalarAsync());
    }

    private static async Task<string> ReadCandidateMappingSampleRowsAsync(
        SqlConnection connection,
        string quotedDatabaseName,
        CobaltMappingObjectCandidate candidate,
        CobaltMappingFilter filter,
        DocumentItemInfo item)
    {
        var usefulColumns = candidate.Columns
            .Where(IsUsefulMappingSearchColumn)
            .Take(16)
            .ToList();

        if (usefulColumns.Count == 0)
            usefulColumns = candidate.Columns.Take(8).ToList();

        var selectList = string.Join(", ", usefulColumns.Select(QuoteSqlIdentifier));

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT TOP (10) {selectList}
FROM {quotedDatabaseName}.{QuoteSqlIdentifier(candidate.SchemaName)}.{QuoteSqlIdentifier(candidate.ObjectName)}
{filter.Sql}
{BuildMappingSearchOrderBy(usefulColumns)};";

        AddDocumentFilterParameters(command, filter, item);

        using var reader = await command.ExecuteReaderAsync();
        var rows = new List<string>();

        while (await reader.ReadAsync())
        {
            var values = new List<string>();

            foreach (var column in usefulColumns)
            {
                var value = reader[column];

                values.Add(value == DBNull.Value
                    ? column + "="
                    : column + "=" + Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture));
            }

            rows.Add(string.Join(";", values));
        }

        return string.Join(" || ", rows);
    }

    private static bool IsUsefulMappingSearchColumn(string columnName)
    {
        return columnName.Equals("SiteId", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("DocId", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("StreamId", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("BSN", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("Partition", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("HistVersion", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("Type", StringComparison.OrdinalIgnoreCase) ||
               columnName.Equals("Size", StringComparison.OrdinalIgnoreCase) ||
               columnName.IndexOf("Ordinal", StringComparison.OrdinalIgnoreCase) >= 0 ||
               columnName.IndexOf("Sequence", StringComparison.OrdinalIgnoreCase) >= 0 ||
               columnName.IndexOf("Offset", StringComparison.OrdinalIgnoreCase) >= 0 ||
               columnName.IndexOf("Index", StringComparison.OrdinalIgnoreCase) >= 0 ||
               columnName.IndexOf("Blob", StringComparison.OrdinalIgnoreCase) >= 0 ||
               columnName.IndexOf("Map", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string BuildMappingSearchOrderBy(List<string> columns)
    {
        var orderColumns = new[]
            {
                "HistVersion",
                "StreamId",
                "Partition",
                "BSN",
                "Type"
            }
            .Where(x => columns.Contains(x, StringComparer.OrdinalIgnoreCase))
            .Select(QuoteSqlIdentifier)
            .ToList();

        return orderColumns.Count == 0
            ? ""
            : "ORDER BY " + string.Join(", ", orderColumns);
    }

    private static void AddDocumentFilterParameters(
        SqlCommand command,
        CobaltMappingFilter filter,
        DocumentItemInfo item)
    {
        if (filter.Sql.IndexOf("@siteId", StringComparison.OrdinalIgnoreCase) >= 0)
            command.Parameters.Add("@siteId", SqlDbType.UniqueIdentifier).Value = item.SiteId;

        if (filter.Sql.IndexOf("@docId", StringComparison.OrdinalIgnoreCase) >= 0)
            command.Parameters.Add("@docId", SqlDbType.UniqueIdentifier).Value = item.Id;
    }

    private static ShreddedExportResult TryExportFormatNeutralCobaltPayload(
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string outputPath,
        string workFolder)
    {
        var manifestPath = Path.Combine(workFolder, "cobalt_recovery_manifest.csv");
        var candidateFolder = Path.Combine(workFolder, "format_neutral_candidates");

        WriteFormatNeutralCobaltRecoveryManifest(rows, manifestPath);
        Directory.CreateDirectory(candidateFolder);

        var candidates = BuildFormatNeutralCobaltCandidates(rows, candidateFolder);
        var reportPath = Path.Combine(candidateFolder, "candidate_report.csv");
        WriteFormatNeutralCandidateReport(candidates, item, reportPath);

        var exactCandidate = item.SizeBytes.HasValue
            ? candidates
                .Where(x => x.Length == item.SizeBytes.Value)
                .Where(x => IsFormatNeutralCandidateValidForSelectedFileType(x.Path, item))
                .OrderByDescending(x => !string.IsNullOrWhiteSpace(DetectKnownFileSignature(x.Path)))
                .ThenBy(x => x.Name)
                .FirstOrDefault()
            : null;

        if (exactCandidate != null)
        {
            File.Copy(exactCandidate.Path, outputPath, overwrite: true);

            return new ShreddedExportResult
            {
                Success = true,
                OutputPath = outputPath,
                BytesWritten = exactCandidate.Length,
                Message =
                    "Exported by strict format-neutral Cobalt payload reconstruction. " +
                    $"Candidate strategy: {exactCandidate.Name}."
            };
        }

        var type10RowCount = rows.Count(x => x.Type == 10);

        if (item.SizeBytes.HasValue &&
            type10RowCount <= 4 &&
            string.Equals(item.DisplayExtension, "pdf", StringComparison.OrdinalIgnoreCase))
        {
            var nearPdfCandidate = candidates
                .Where(x => Math.Abs(x.Length - item.SizeBytes.Value) <= 256)
                .Where(x => IsFormatNeutralCandidateValidForSelectedFileType(x.Path, item))
                .OrderBy(x => Math.Abs(x.Length - item.SizeBytes.Value))
                .ThenByDescending(x => string.Equals(DetectKnownFileSignature(x.Path), "PDF", StringComparison.OrdinalIgnoreCase))
                .ThenBy(x => x.Name)
                .FirstOrDefault();

            if (nearPdfCandidate != null)
            {
                File.Copy(nearPdfCandidate.Path, outputPath, overwrite: true);

                return new ShreddedExportResult
                {
                    Success = true,
                    OutputPath = outputPath,
                    BytesWritten = nearPdfCandidate.Length,
                    Message =
                        "Exported by validated format-neutral Cobalt PDF payload reconstruction. " +
                        $"Candidate strategy: {nearPdfCandidate.Name}. " +
                        $"Size delta from SharePoint metadata: {nearPdfCandidate.Length - item.SizeBytes.Value:N0} bytes."
                };
            }
        }

        return new ShreddedExportResult
        {
            Success = false,
            OutputPath = outputPath,
            Message =
                "Format-neutral Cobalt recovery is required for this file." + Environment.NewLine +
                "The exporter must reconstruct the original byte stream from Cobalt/FSSHTTPB fragments before any file-type validation is attempted." + Environment.NewLine +
                $"A format-neutral recovery manifest was written to: {manifestPath}." + Environment.NewLine +
                $"Strict payload candidates were written to: {candidateFolder}." + Environment.NewLine +
                $"Candidate size report was written to: {reportPath}."
        };
    }

    private static string JoinMessages(params string[] messages)
    {
        return string.Join(
            Environment.NewLine,
            messages
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim()));
    }

    private static void WriteFormatNeutralCandidateReport(
        List<FormatNeutralCobaltCandidate> candidates,
        DocumentItemInfo item,
        string reportPath)
    {
        var expectedSize = item.SizeBytes;

        using var writer = new StreamWriter(reportPath, false, Encoding.UTF8);

        writer.WriteLine("Name,Length,ExpectedSize,DeltaFromExpected,DetectedSignature,ValidForSelectedFileType,Path");

        foreach (var candidate in candidates.OrderBy(x => Math.Abs((expectedSize ?? x.Length) - x.Length)))
        {
            writer.WriteLine(
                string.Join(",",
                    CsvValue(candidate.Name),
                    candidate.Length,
                    expectedSize?.ToString() ?? "",
                    expectedSize.HasValue ? (candidate.Length - expectedSize.Value).ToString() : "",
                    CsvValue(DetectKnownFileSignature(candidate.Path)),
                    IsFormatNeutralCandidateValidForSelectedFileType(candidate.Path, item) ? "1" : "0",
                    CsvValue(candidate.Path)));
        }
    }

    private static bool IsFormatNeutralCandidateValidForSelectedFileType(
        string path,
        DocumentItemInfo item)
    {
        if (!File.Exists(path))
            return false;

        var extension = item.DisplayExtension.Trim('.').ToLowerInvariant();

        if (extension == "pdf")
        {
            var bytes = File.ReadAllBytes(path);

            return StartsWithBytes(bytes, Encoding.ASCII.GetBytes("%PDF-")) &&
                   LastIndexOfBytes(bytes, Encoding.ASCII.GetBytes("%%EOF")) >= 0 &&
                   LastIndexOfBytes(bytes, Encoding.ASCII.GetBytes("startxref")) >= 0;
        }

        var signature = KnownFileSignatures.FirstOrDefault(x =>
            (extension == "docx" || extension == "xlsx" || extension == "pptx") && x.Name == "ZIP Office Open XML" ||
            extension == "jpg" && x.Name == "JPEG" ||
            extension == "jpeg" && x.Name == "JPEG" ||
            extension == "png" && x.Name == "PNG" ||
            extension == "gif" && x.Name == "GIF" ||
            extension == "doc" && x.Name == "OLE Compound" ||
            extension == "xls" && x.Name == "OLE Compound" ||
            extension == "ppt" && x.Name == "OLE Compound" ||
            extension == "rtf" && x.Name == "RTF");

        if (signature == null)
            return true;

        var max = (int)Math.Min(new FileInfo(path).Length, signature.Bytes.Length);
        var buffer = new byte[max];

        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            fs.Read(buffer, 0, buffer.Length);
        }

        return StartsWithBytes(buffer, signature.Bytes);
    }

    private static bool StartsWithBytes(byte[] data, byte[] prefix)
    {
        if (data.Length < prefix.Length)
            return false;

        for (var i = 0; i < prefix.Length; i++)
        {
            if (data[i] != prefix[i])
                return false;
        }

        return true;
    }

    private static string DetectKnownFileSignature(string path)
    {
        if (!File.Exists(path))
            return "";

        var max = (int)Math.Min(new FileInfo(path).Length, 512);
        var buffer = new byte[max];

        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            fs.Read(buffer, 0, buffer.Length);
        }

        foreach (var signature in KnownFileSignatures)
        {
            if (IndexOfBytes(buffer, signature.Bytes) >= 0)
                return signature.Name;
        }

        return "";
    }

    private static List<FormatNeutralCobaltCandidate> BuildFormatNeutralCobaltCandidates(
        List<ShreddedStreamRow> rows,
        string candidateFolder)
    {
        var candidates = new List<FormatNeutralCobaltCandidate>();

        var type10Rows = rows
            .Where(x => x.Type == 10)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.Partition)
            .ThenBy(x => x.BSN)
            .ToList();

        if (type10Rows.Count == 0)
            return candidates;

        AddFormatNeutralCandidate(
            candidates,
            "type10_payload_stream_order.bin",
            candidateFolder,
            type10Rows,
            firstRowStartOffset: 96);

        AddFormatNeutralCandidate(
            candidates,
            "type10_payload_stream_order_strip_inner108.bin",
            candidateFolder,
            type10Rows,
            firstRowStartOffset: 96 + 108);

        foreach (var signature in KnownFileSignatures)
        {
            AddFormatNeutralCandidate(
                candidates,
                "type10_payload_from_" + MakeSafeFileName(signature.Name.ToLowerInvariant().Replace(" ", "_")) + ".bin",
                candidateFolder,
                type10Rows,
                firstRowStartOffset: 96,
                rotateFromSignature: signature);

            AddFormatNeutralCandidate(
                candidates,
                "type10_payload_from_" + MakeSafeFileName(signature.Name.ToLowerInvariant().Replace(" ", "_")) + "_strip_inner108.bin",
                candidateFolder,
                type10Rows,
                firstRowStartOffset: 96 + 108,
                rotateFromSignature: signature);
        }

        foreach (var row in type10Rows)
        {
            var payloadStart = Math.Min(96, row.Content.Length);
            var payloadEnd = GetCobaltFramedPayloadEnd(
                row.Content,
                new byte[] { 0x15, 0xA4, 0x01 },
                payloadStart);

            if (payloadEnd <= payloadStart)
                continue;

            var singlePath = Path.Combine(
                candidateFolder,
                $"single_type_{row.Type}_stream_{row.StreamId}_bsn_{row.BSN}.bin");

            using (var output = new FileStream(singlePath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                output.Write(row.Content, payloadStart, payloadEnd - payloadStart);
            }

            candidates.Add(new FormatNeutralCobaltCandidate(
                Path.GetFileName(singlePath),
                singlePath,
                new FileInfo(singlePath).Length));
        }

        return candidates;
    }

    private static void AddFormatNeutralCandidate(
        List<FormatNeutralCobaltCandidate> candidates,
        string fileName,
        string candidateFolder,
        List<ShreddedStreamRow> orderedRows,
        int firstRowStartOffset,
        KnownFileSignature? rotateFromSignature = null)
    {
        var rows = orderedRows;
        var signatureMatch = rotateFromSignature == null
            ? null
            : FindFirstKnownFileSignatureRow(orderedRows, rotateFromSignature);

        if (signatureMatch != null)
        {
            rows = orderedRows
                .Where(x => x.StreamId >= signatureMatch.Row.StreamId)
                .Concat(orderedRows.Where(x => x.StreamId < signatureMatch.Row.StreamId))
                .ToList();
        }

        var path = Path.Combine(candidateFolder, fileName);

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var row in rows)
            {
                var payloadStart = Math.Min(firstRowStartOffset, row.Content.Length);

                if (signatureMatch != null && ReferenceEquals(row, signatureMatch.Row))
                    payloadStart = signatureMatch.Offset;

                var payloadEnd = GetCobaltFramedPayloadEnd(
                    row.Content,
                    new byte[] { 0x15, 0xA4, 0x01 },
                    payloadStart);

                if (payloadEnd <= payloadStart)
                    continue;

                output.Write(row.Content, payloadStart, payloadEnd - payloadStart);
            }
        }

        candidates.Add(new FormatNeutralCobaltCandidate(
            fileName,
            path,
            new FileInfo(path).Length));
    }

    private static KnownSignatureMatch? FindFirstKnownFileSignatureRow(
        List<ShreddedStreamRow> rows,
        KnownFileSignature signature)
    {
        return rows
            .Select(row => new
            {
                Row = row,
                Offset = IndexOfBytes(row.Content, signature.Bytes)
            })
            .Where(x => x.Offset >= 0)
            .OrderBy(x => x.Row.StreamId)
            .ThenBy(x => x.Offset)
            .Select(x => new KnownSignatureMatch(x.Row, x.Offset))
            .FirstOrDefault();
    }

    private static void WriteFormatNeutralCobaltRecoveryManifest(
        List<ShreddedStreamRow> rows,
        string manifestPath)
    {
        using var writer = new StreamWriter(manifestPath, false, Encoding.UTF8);

        writer.WriteLine(
            "HistVersion,Type,StreamId,Partition,BSN,ContentLength," +
            "PdfOffset,ZipOffset,JpegOffset,PngOffset,GifOffset,OleOffset,RtfOffset," +
            "First32Hex,Last32Hex");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.Partition)
                     .ThenBy(x => x.BSN))
        {
            writer.WriteLine(
                string.Join(",",
                    row.HistVersion,
                    row.Type,
                    row.StreamId,
                    row.Partition,
                    row.BSN,
                    row.Content.Length,
                    IndexOfBytes(row.Content, Encoding.ASCII.GetBytes("%PDF-")),
                    IndexOfBytes(row.Content, new byte[] { 0x50, 0x4B, 0x03, 0x04 }),
                    IndexOfBytes(row.Content, new byte[] { 0xFF, 0xD8, 0xFF }),
                    IndexOfBytes(row.Content, new byte[] { 0x89, 0x50, 0x4E, 0x47 }),
                    IndexOfBytes(row.Content, Encoding.ASCII.GetBytes("GIF8")),
                    IndexOfBytes(row.Content, new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }),
                    IndexOfBytes(row.Content, Encoding.ASCII.GetBytes(@"{\rtf")),
                    CsvValue(ToHexString(row.Content.Take(32).ToArray())),
                    CsvValue(ToHexString(row.Content.Length <= 32
                        ? row.Content
                        : row.Content.Skip(row.Content.Length - 32).Take(32).ToArray()))));
        }
    }

    private sealed class FormatNeutralCobaltCandidate
    {
        public FormatNeutralCobaltCandidate(string name, string path, long length)
        {
            Name = name;
            Path = path;
            Length = length;
        }

        public string Name { get; }

        public string Path { get; }

        public long Length { get; }
    }

    private sealed class KnownSignatureMatch
    {
        public KnownSignatureMatch(ShreddedStreamRow row, int offset)
        {
            Row = row;
            Offset = offset;
        }

        public ShreddedStreamRow Row { get; }

        public int Offset { get; }
    }

    private sealed class KnownFileSignature
    {
        public KnownFileSignature(string name, byte[] bytes)
        {
            Name = name;
            Bytes = bytes;
        }

        public string Name { get; }

        public byte[] Bytes { get; }
    }

    private static readonly KnownFileSignature[] KnownFileSignatures =
    {
        new("PDF", Encoding.ASCII.GetBytes("%PDF-")),
        new("ZIP Office Open XML", new byte[] { 0x50, 0x4B, 0x03, 0x04 }),
        new("JPEG", new byte[] { 0xFF, 0xD8, 0xFF }),
        new("PNG", new byte[] { 0x89, 0x50, 0x4E, 0x47 }),
        new("GIF", Encoding.ASCII.GetBytes("GIF8")),
        new("OLE Compound", new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }),
        new("RTF", Encoding.ASCII.GetBytes(@"{\rtf"))
    };

    private static string ToHexString(byte[] data)
    {
        return BitConverter.ToString(data).Replace("-", "");
    }

    private static ShreddedExportResult TryExportEmbeddedPdfPayload(
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string outputPath)
    {
        if (!string.Equals(item.DisplayExtension, "pdf", StringComparison.OrdinalIgnoreCase))
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Embedded PDF recovery was skipped because the item is not a PDF."
            };
        }

        var type10Rows = rows
            .Where(x => x.Type == 10 && x.Content.Length > 0)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ToList();

        if (type10Rows.Count == 0)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Embedded PDF recovery found no Type 10 rows."
            };
        }

        var pdfHeader = System.Text.Encoding.ASCII.GetBytes("%PDF-");
        var pdfEof = System.Text.Encoding.ASCII.GetBytes("%%EOF");
        var cobaltRowSuffix = new byte[] { 0x15, 0xA4, 0x01 };

        var startMatch = type10Rows
            .Select(row => new
            {
                Row = row,
                Offset = IndexOfBytes(row.Content, pdfHeader)
            })
            .FirstOrDefault(x => x.Offset >= 0);

        var eofMatch = type10Rows
            .Select(row => new
            {
                Row = row,
                Offset = LastIndexOfBytes(row.Content, pdfEof)
            })
            .LastOrDefault(x => x.Offset >= 0);

        if (startMatch == null || eofMatch == null)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Embedded PDF recovery could not find both %PDF- and %%EOF markers."
            };
        }

        const int cobaltRowPrefixLength = 96;

        var candidates = new List<PdfRecoveryCandidate>
        {
            BuildPdfRecoveryCandidate(
                "type10_stripped_stop_at_eof",
                type10Rows,
                startMatch.Row,
                startMatch.Offset,
                eofMatch.Row,
                eofMatch.Offset,
                cobaltRowPrefixLength,
                cobaltRowSuffix,
                stripCobaltRowSuffix: true,
                stopAtEof: true),

            BuildPdfRecoveryCandidate(
                "type10_stripped_full_eof_row",
                type10Rows,
                startMatch.Row,
                startMatch.Offset,
                eofMatch.Row,
                eofMatch.Offset,
                cobaltRowPrefixLength,
                cobaltRowSuffix,
                stripCobaltRowSuffix: true,
                stopAtEof: false),

            BuildPdfRecoveryCandidate(
                "type10_prefix_only_full_eof_row",
                type10Rows,
                startMatch.Row,
                startMatch.Offset,
                eofMatch.Row,
                eofMatch.Offset,
                cobaltRowPrefixLength,
                cobaltRowSuffix,
                stripCobaltRowSuffix: false,
                stopAtEof: false)
        };

        var allContentRows = rows
            .Where(x => x.Content.Length > 0)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ToList();

        var allStartMatch = allContentRows
            .Select(row => new
            {
                Row = row,
                Offset = IndexOfBytes(row.Content, pdfHeader)
            })
            .FirstOrDefault(x => x.Offset >= 0);

        var allEofMatch = allContentRows
            .Select(row => new
            {
                Row = row,
                Offset = LastIndexOfBytes(row.Content, pdfEof)
            })
            .LastOrDefault(x => x.Offset >= 0);

        if (allStartMatch != null && allEofMatch != null)
        {
            candidates.Add(BuildPdfRecoveryCandidate(
                "all_types_stripped_full_eof_row",
                allContentRows,
                allStartMatch.Row,
                allStartMatch.Offset,
                allEofMatch.Row,
                allEofMatch.Offset,
                cobaltRowPrefixLength,
                cobaltRowSuffix,
                stripCobaltRowSuffix: true,
                stopAtEof: false));

            candidates.Add(BuildPdfRecoveryCandidate(
                "all_types_prefix_only_full_eof_row",
                allContentRows,
                allStartMatch.Row,
                allStartMatch.Offset,
                allEofMatch.Row,
                allEofMatch.Offset,
                cobaltRowPrefixLength,
                cobaltRowSuffix,
                stripCobaltRowSuffix: false,
                stopAtEof: false));
        }

        var markerValidCandidates = candidates
            .Where(x => LooksLikePdfBytes(x.Bytes, pdfHeader, pdfEof))
            .ToList();

        if (markerValidCandidates.Count == 0)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Embedded PDF recovery produced candidate files, but PDF marker validation failed."
            };
        }

        var validCandidates = markerValidCandidates
            .Where(x => LooksLikeConsistentPdfXref(x.Bytes))
            .ToList();

        if (validCandidates.Count == 0)
        {
            var diagnosticFolder = WritePdfRecoveryCandidates(
                outputPath,
                markerValidCandidates,
                markerValidCandidates[0],
                item.SizeBytes ?? 0);

            foreach (var candidate in markerValidCandidates)
            {
                WritePdfObjectMapDiagnostic(
                    Path.Combine(diagnosticFolder, candidate.Name + "_object_map.csv"),
                    candidate.Bytes);

                WritePdfImageStreamGapReport(
                    Path.Combine(diagnosticFolder, candidate.Name + "_image_stream_gaps.csv"),
                    candidate.Bytes);

                WritePdfImageStreamFragmentAssignmentReport(
                    Path.Combine(diagnosticFolder, candidate.Name + "_image_stream_fragment_assignments.csv"),
                    candidate.Bytes);

                TryWriteObjectRepackedPdfCandidate(
                    Path.Combine(diagnosticFolder, candidate.Name + "_object_repacked_length_normalized.pdf"),
                    candidate.Bytes);
            }

            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message =
                    "Embedded PDF recovery found PDF markers, but the PDF cross-reference table does not match the recovered byte layout. " +
                    "This indicates the file content is fragmented/interleaved in Cobalt shredded storage and cannot be recovered by raw row concatenation. " +
                    "Marker-only candidates, object maps, image stream gap reports, and image stream fragment assignment reports were written to: " +
                    diagnosticFolder
            };
        }

        var selectedCandidate = SelectBestPdfRecoveryCandidate(validCandidates, item.SizeBytes);

        File.WriteAllBytes(outputPath, selectedCandidate.Bytes);

        if (!LooksLikePdfFile(outputPath))
        {
            try
            {
                File.Delete(outputPath);
            }
            catch
            {
            }

            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Embedded PDF recovery produced a file, but PDF marker validation failed."
            };
        }

        var fileInfo = new FileInfo(outputPath);

        if (item.SizeBytes.HasValue && fileInfo.Length != item.SizeBytes.Value)
        {
            var diagnosticFolder = WritePdfRecoveryCandidates(
                outputPath,
                validCandidates,
                selectedCandidate,
                item.SizeBytes.Value);

            return new ShreddedExportResult
            {
                Success = true,
                OutputPath = outputPath,
                BytesWritten = fileInfo.Length,
                Message =
                    "Exported by experimental embedded PDF payload recovery. " +
                    "PDF header/trailer markers validated, but SharePoint's stored file size did not match the recovered PDF payload. " +
                    $"Selected candidate: {selectedCandidate.Name}. " +
                    $"SharePoint size: {item.SizeBytes.Value:N0} bytes; recovered PDF: {fileInfo.Length:N0} bytes. " +
                    $"Alternate recovery candidates were written to: {diagnosticFolder}"
            };
        }

        return new ShreddedExportResult
        {
            Success = true,
            OutputPath = outputPath,
            BytesWritten = fileInfo.Length,
            Message =
                $"Exported by experimental embedded PDF payload recovery ({selectedCandidate.Name}). " +
                "This is a fallback for non-Office files stored in SharePoint shredded storage."
        };
    }

    private static PdfRecoveryCandidate BuildPdfRecoveryCandidate(
        string name,
        List<ShreddedStreamRow> rows,
        ShreddedStreamRow startRow,
        int startOffset,
        ShreddedStreamRow eofRow,
        int eofOffset,
        int rowPrefixLength,
        byte[] rowSuffixMarker,
        bool stripCobaltRowSuffix,
        bool stopAtEof)
    {
        var orderedRows = rows
            .Where(x => x.StreamId >= startRow.StreamId && !ReferenceEquals(x, eofRow))
            .Concat(rows.Where(x => x.StreamId < startRow.StreamId))
            .Concat(new[] { eofRow })
            .ToList();

        using var output = new MemoryStream();

        foreach (var row in orderedRows)
        {
            var start = ReferenceEquals(row, startRow)
                ? startOffset
                : Math.Min(rowPrefixLength, row.Content.Length);

            var framedEnd = stripCobaltRowSuffix
                ? GetCobaltFramedPayloadEnd(row.Content, rowSuffixMarker, start)
                : row.Content.Length;

            var end = ReferenceEquals(row, eofRow) && stopAtEof
                ? Math.Min(framedEnd, ConsumePdfEofWhitespace(row.Content, eofOffset, 5, framedEnd))
                : framedEnd;

            if (end <= start)
                continue;

            output.Write(row.Content, start, end - start);
        }

        return new PdfRecoveryCandidate(name, output.ToArray());
    }

    private static PdfRecoveryCandidate SelectBestPdfRecoveryCandidate(
        List<PdfRecoveryCandidate> candidates,
        long? expectedSize)
    {
        if (expectedSize.HasValue)
        {
            var exact = candidates.FirstOrDefault(x => x.Bytes.LongLength == expectedSize.Value);

            if (exact != null)
                return exact;

            return candidates
                .OrderBy(x => Math.Abs(x.Bytes.LongLength - expectedSize.Value))
                .ThenBy(x => x.Name)
                .First();
        }

        return candidates
            .OrderBy(x => x.Name)
            .First();
    }

    private static string WritePdfRecoveryCandidates(
        string outputPath,
        List<PdfRecoveryCandidate> candidates,
        PdfRecoveryCandidate selectedCandidate,
        long expectedSize)
    {
        var folder = Path.Combine(
            Path.GetDirectoryName(outputPath) ?? Environment.CurrentDirectory,
            Path.GetFileName(outputPath) + "_pdf_recovery_candidates");

        Directory.CreateDirectory(folder);

        var reportPath = Path.Combine(folder, "candidate_report.csv");

        using (var writer = new StreamWriter(reportPath, false, System.Text.Encoding.UTF8))
        {
            writer.WriteLine("Selected,Name,Bytes,ExpectedBytes,Delta");

            foreach (var candidate in candidates.OrderBy(x => Math.Abs(x.Bytes.LongLength - expectedSize)))
            {
                var candidatePath = Path.Combine(folder, candidate.Name + ".pdf");

                File.WriteAllBytes(candidatePath, candidate.Bytes);

                writer.WriteLine(
                    string.Join(",",
                        ReferenceEquals(candidate, selectedCandidate) ? "1" : "0",
                        CsvValue(candidate.Name),
                        candidate.Bytes.LongLength,
                        expectedSize,
                        candidate.Bytes.LongLength - expectedSize));
            }
        }

        return folder;
    }

    private static void WritePdfObjectMapDiagnostic(string path, byte[] bytes)
    {
        using var writer = new StreamWriter(path, false, System.Text.Encoding.UTF8);

        writer.WriteLine("Kind,ObjectNumber,Offset,XrefOffset,MatchesXref,Snippet");

        var objectHeaderPattern = new System.Text.RegularExpressions.Regex(
            @"(?m)(\d+)\s+0\s+obj",
            System.Text.RegularExpressions.RegexOptions.Compiled);

        var text = System.Text.Encoding.ASCII.GetString(bytes);
        var actualOffsets = new Dictionary<int, List<int>>();

        foreach (System.Text.RegularExpressions.Match match in objectHeaderPattern.Matches(text))
        {
            var objectNumber = int.Parse(match.Groups[1].Value);

            if (!actualOffsets.TryGetValue(objectNumber, out var offsets))
            {
                offsets = new List<int>();
                actualOffsets[objectNumber] = offsets;
            }

            offsets.Add(match.Index);

            writer.WriteLine(
                string.Join(",",
                    "actual",
                    objectNumber,
                    match.Index,
                    "",
                    "",
                    CsvValue(GetAsciiSnippet(bytes, match.Index, 80))));
        }

        foreach (var entry in ReadPdfXrefEntries(bytes))
        {
            var matches = actualOffsets.TryGetValue(entry.ObjectNumber, out var offsets) &&
                          offsets.Contains(entry.Offset);

            writer.WriteLine(
                string.Join(",",
                    "xref",
                    entry.ObjectNumber,
                    "",
                    entry.Offset,
                    matches ? "1" : "0",
                    CsvValue(GetAsciiSnippet(bytes, entry.Offset, 80))));
        }
    }

    private static bool TryWriteObjectRepackedPdfCandidate(string path, byte[] bytes)
    {
        var text = System.Text.Encoding.ASCII.GetString(bytes);
        var matches = System.Text.RegularExpressions.Regex.Matches(
            text,
            @"(?m)(\d+)\s+0\s+obj");

        if (matches.Count == 0)
            return false;

        var positions = matches
            .Cast<System.Text.RegularExpressions.Match>()
            .Select(match => new PdfObjectPosition(
                int.Parse(match.Groups[1].Value),
                match.Index))
            .OrderBy(x => x.Offset)
            .ToList();

        var objects = new Dictionary<int, byte[]>();

        for (var i = 0; i < positions.Count; i++)
        {
            var objectNumber = positions[i].ObjectNumber;
            var start = positions[i].Offset;
            var end = bytes.Length;

            if (i + 1 < positions.Count)
            {
                end = positions[i + 1].Offset;
            }
            else
            {
                var xref = text.IndexOf("xref", start, StringComparison.Ordinal);

                if (xref > start)
                    end = xref;
            }

            if (end <= start)
                continue;

            var chunk = new byte[end - start];
            Array.Copy(bytes, start, chunk, 0, chunk.Length);

            if (!objects.TryGetValue(objectNumber, out var existing) ||
                chunk.Length < existing.Length)
            {
                objects[objectNumber] = chunk;
            }
        }

        if (!objects.ContainsKey(1) || !objects.ContainsKey(2))
            return false;

        RewriteIndirectPdfStreamLengths(objects);

        using var output = new MemoryStream();
        var header = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%Object repacked diagnostic candidate\n");
        output.Write(header, 0, header.Length);

        var offsets = new Dictionary<int, int>();

        foreach (var objectNumber in objects.Keys.OrderBy(x => x))
        {
            offsets[objectNumber] = (int)output.Position;

            var chunk = objects[objectNumber];
            output.Write(chunk, 0, chunk.Length);

            if (chunk.Length == 0 || chunk[chunk.Length - 1] != 0x0A)
                output.WriteByte(0x0A);
        }

        var xrefOffset = (int)output.Position;
        var maxObjectNumber = objects.Keys.Max();
        var xrefText = new StringBuilder();

        xrefText.AppendLine("xref");
        xrefText.AppendLine($"0 {maxObjectNumber + 1}");
        xrefText.AppendLine("0000000000 65535 f ");

        for (var objectNumber = 1; objectNumber <= maxObjectNumber; objectNumber++)
        {
            if (offsets.TryGetValue(objectNumber, out var offset))
                xrefText.AppendLine($"{offset:0000000000} 00000 n ");
            else
                xrefText.AppendLine("0000000000 65535 f ");
        }

        xrefText.AppendLine("trailer");
        xrefText.AppendLine("<<");
        xrefText.AppendLine($"/Size {maxObjectNumber + 1}");
        xrefText.AppendLine("/Root 2 0 R");
        xrefText.AppendLine(">>");
        xrefText.AppendLine("startxref");
        xrefText.AppendLine(xrefOffset.ToString());
        xrefText.AppendLine("%%EOF");

        var xrefBytes = System.Text.Encoding.ASCII.GetBytes(xrefText.ToString());
        output.Write(xrefBytes, 0, xrefBytes.Length);

        File.WriteAllBytes(path, output.ToArray());
        return true;
    }

    private static void WritePdfImageStreamGapReport(string path, byte[] bytes)
    {
        var text = System.Text.Encoding.ASCII.GetString(bytes);

        using var writer = new StreamWriter(path, false, System.Text.Encoding.UTF8);

        writer.WriteLine(
            "ImageObject,LengthObject,DeclaredLength,ObservedStreamLength,Delta,ObjectOffset,StreamOffset,EndStreamOffset,Snippet");

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
        {
            var objectNumber = int.Parse(match.Groups[1].Value);
            var objectStart = match.Index;
            var nextObjectStart = FindNextPdfObjectOffset(text, objectStart + match.Length);
            var objectEnd = nextObjectStart >= 0 ? nextObjectStart : text.Length;
            var objectText = text.Substring(objectStart, objectEnd - objectStart);

            if (objectText.IndexOf("/Subtype /Image", StringComparison.Ordinal) < 0)
                continue;

            var lengthMatch = System.Text.RegularExpressions.Regex.Match(
                objectText,
                @"/Length\s+(\d+)\s+0\s+R");

            if (!lengthMatch.Success)
                continue;

            var lengthObjectNumber = int.Parse(lengthMatch.Groups[1].Value);
            var declaredLength = ReadSimplePdfLengthObject(text, lengthObjectNumber);

            var streamOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
            var endStreamOffset = text.IndexOf("endstream", objectStart, StringComparison.Ordinal);

            var observedLength = streamOffset >= 0 && endStreamOffset > streamOffset
                ? endStreamOffset - (streamOffset + "stream".Length + 1)
                : -1;

            writer.WriteLine(
                string.Join(",",
                    objectNumber,
                    lengthObjectNumber,
                    declaredLength?.ToString() ?? "",
                    observedLength,
                    declaredLength.HasValue && observedLength >= 0 ? (observedLength - declaredLength.Value).ToString() : "",
                    objectStart,
                    streamOffset,
                    endStreamOffset,
                    CsvValue(GetAsciiSnippet(bytes, objectStart, 140))));
        }
    }

    private static void WritePdfImageStreamFragmentAssignmentReport(string path, byte[] bytes)
    {
        var text = System.Text.Encoding.ASCII.GetString(bytes);
        var segmentStarts = InferLikelyCobaltSegmentStarts(bytes);

        using var writer = new StreamWriter(path, false, System.Text.Encoding.UTF8);

        writer.WriteLine(
            "ImageObject,LengthObject,DeclaredLength,StreamDataStart,ExpectedStreamEnd,ObservedEndStreamOffset," +
            "SegmentIndex,SegmentStart,SegmentEnd,SegmentOverlapStart,SegmentOverlapEnd,SegmentOverlapLength,SegmentFirstObject,Snippet");

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
        {
            var objectNumber = int.Parse(match.Groups[1].Value);
            var objectStart = match.Index;
            var nextObjectStart = FindNextPdfObjectOffset(text, objectStart + match.Length);
            var objectEnd = nextObjectStart >= 0 ? nextObjectStart : text.Length;
            var objectText = text.Substring(objectStart, objectEnd - objectStart);

            if (objectText.IndexOf("/Subtype /Image", StringComparison.Ordinal) < 0)
                continue;

            var lengthMatch = System.Text.RegularExpressions.Regex.Match(
                objectText,
                @"/Length\s+(\d+)\s+0\s+R");

            if (!lengthMatch.Success)
                continue;

            var lengthObjectNumber = int.Parse(lengthMatch.Groups[1].Value);
            var declaredLength = ReadSimplePdfLengthObject(text, lengthObjectNumber);

            if (!declaredLength.HasValue)
                continue;

            var streamKeywordOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
            var observedEndStreamOffset = text.IndexOf("endstream", objectStart, StringComparison.Ordinal);

            if (streamKeywordOffset < 0)
                continue;

            var streamDataStart = streamKeywordOffset + "stream".Length;

            if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0D)
                streamDataStart++;

            if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0A)
                streamDataStart++;

            var expectedStreamEnd = Math.Min(bytes.Length, streamDataStart + declaredLength.Value);

            foreach (var segment in EnumerateSegmentOverlaps(segmentStarts, bytes.Length, streamDataStart, expectedStreamEnd))
            {
                writer.WriteLine(
                    string.Join(",",
                        objectNumber,
                        lengthObjectNumber,
                        declaredLength.Value,
                        streamDataStart,
                        expectedStreamEnd,
                        observedEndStreamOffset,
                        segment.Index,
                        segment.Start,
                        segment.End,
                        segment.OverlapStart,
                        segment.OverlapEnd,
                        segment.OverlapEnd - segment.OverlapStart,
                        CsvValue(GetFirstObjectNumberInRange(text, segment.Start, segment.End)?.ToString() ?? ""),
                        CsvValue(GetAsciiSnippet(bytes, segment.OverlapStart, 80))));
            }
        }
    }

    private static List<int> InferLikelyCobaltSegmentStarts(byte[] bytes)
    {
        var text = System.Text.Encoding.ASCII.GetString(bytes);
        var starts = new SortedSet<int> { 0 };

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
        {
            starts.Add(match.Index);
        }

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)endobj\s+(\d+)\s+0\s+obj"))
        {
            var objectTextOffset = match.Index + match.Value.LastIndexOf(match.Groups[1].Value, StringComparison.Ordinal);
            starts.Add(objectTextOffset);
        }

        return starts.ToList();
    }

    private static IEnumerable<SegmentOverlap> EnumerateSegmentOverlaps(
        List<int> segmentStarts,
        int totalLength,
        int rangeStart,
        int rangeEnd)
    {
        for (var i = 0; i < segmentStarts.Count; i++)
        {
            var segmentStart = segmentStarts[i];
            var segmentEnd = i + 1 < segmentStarts.Count
                ? segmentStarts[i + 1]
                : totalLength;

            var overlapStart = Math.Max(segmentStart, rangeStart);
            var overlapEnd = Math.Min(segmentEnd, rangeEnd);

            if (overlapEnd <= overlapStart)
                continue;

            yield return new SegmentOverlap(
                i,
                segmentStart,
                segmentEnd,
                overlapStart,
                overlapEnd);
        }
    }

    private static int? GetFirstObjectNumberInRange(string text, int start, int end)
    {
        if (start < 0 || start >= text.Length || end <= start)
            return null;

        var match = System.Text.RegularExpressions.Regex.Match(
            text.Substring(start, Math.Min(end, text.Length) - start),
            @"(?m)(\d+)\s+0\s+obj");

        return match.Success
            ? int.Parse(match.Groups[1].Value)
            : null;
    }

    private static int FindNextPdfObjectOffset(string text, int startOffset)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            text.Substring(startOffset),
            @"(?m)(\d+)\s+0\s+obj");

        return match.Success
            ? startOffset + match.Index
            : -1;
    }

    private static int? ReadSimplePdfLengthObject(string text, int objectNumber)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            text,
            $@"(?m){objectNumber}\s+0\s+obj\s*(\d+)\s*endobj");

        return match.Success
            ? int.Parse(match.Groups[1].Value)
            : null;
    }

    private static void RewriteIndirectPdfStreamLengths(Dictionary<int, byte[]> objects)
    {
        var lengthUpdates = new Dictionary<int, int>();

        foreach (var pair in objects.ToList())
        {
            var text = System.Text.Encoding.ASCII.GetString(pair.Value);
            var lengthMatch = System.Text.RegularExpressions.Regex.Match(
                text,
                @"/Length\s+(\d+)\s+0\s+R");

            if (!lengthMatch.Success)
                continue;

            var streamOffset = text.IndexOf("stream", StringComparison.Ordinal);
            var endStreamOffset = text.IndexOf("endstream", StringComparison.Ordinal);

            if (streamOffset < 0 || endStreamOffset <= streamOffset)
                continue;

            var actualLength = endStreamOffset - (streamOffset + "stream".Length + 1);

            if (actualLength < 0)
                continue;

            var lengthObjectNumber = int.Parse(lengthMatch.Groups[1].Value);
            lengthUpdates[lengthObjectNumber] = actualLength;
        }

        foreach (var update in lengthUpdates)
        {
            if (!objects.ContainsKey(update.Key))
                continue;

            objects[update.Key] = System.Text.Encoding.ASCII.GetBytes(
                $"{update.Key} 0 obj\n{update.Value}\nendobj\n");
        }
    }

    private static List<PdfXrefEntry> ReadPdfXrefEntries(byte[] bytes)
    {
        var entries = new List<PdfXrefEntry>();
        var startXrefMarker = System.Text.Encoding.ASCII.GetBytes("startxref");
        var startXrefOffset = LastIndexOfBytes(bytes, startXrefMarker);

        if (startXrefOffset < 0)
            return entries;

        var xrefValue = TryReadPdfIntegerAfter(bytes, startXrefOffset + startXrefMarker.Length);

        if (!xrefValue.HasValue ||
            xrefValue.Value < 0 ||
            xrefValue.Value >= bytes.Length)
        {
            return entries;
        }

        var cursor = xrefValue.Value + "xref".Length;

        if (!TryReadPdfLine(bytes, ref cursor, out _))
            return entries;

        while (TryReadPdfLine(bytes, ref cursor, out var line))
        {
            line = line.Trim();

            if (line.StartsWith("trailer", StringComparison.Ordinal))
                break;

            var subsection = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (subsection.Length != 2 ||
                !int.TryParse(subsection[0], out var firstObjectNumber) ||
                !int.TryParse(subsection[1], out var count))
            {
                break;
            }

            for (var i = 0; i < count; i++)
            {
                if (!TryReadPdfLine(bytes, ref cursor, out var entry) ||
                    entry.Length < 18)
                {
                    break;
                }

                if (entry[17] != 'n')
                    continue;

                if (int.TryParse(entry.Substring(0, 10), out var offset))
                {
                    entries.Add(new PdfXrefEntry(firstObjectNumber + i, offset));
                }
            }
        }

        return entries;
    }

    private static string GetAsciiSnippet(byte[] bytes, int offset, int length)
    {
        if (offset < 0 || offset >= bytes.Length)
            return "";

        var count = Math.Min(length, bytes.Length - offset);
        var chars = new char[count];

        for (var i = 0; i < count; i++)
        {
            var value = bytes[offset + i];

            chars[i] = value >= 32 && value <= 126
                ? (char)value
                : '.';
        }

        return new string(chars);
    }

    private static string CsvValue(object? value)
    {
        var text = value?.ToString() ?? "";

        if (text.Contains("\""))
            text = text.Replace("\"", "\"\"");

        if (text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n"))
            return "\"" + text + "\"";

        return text;
    }

    private static bool LooksLikePdfBytes(
        byte[] bytes,
        byte[] pdfHeader,
        byte[] pdfEof)
    {
        if (bytes.Length < 10)
            return false;

        return IndexOfBytes(bytes, pdfHeader) == 0 &&
               LastIndexOfBytes(bytes, pdfEof) >= 0;
    }

    private static bool LooksLikeConsistentPdfXref(byte[] bytes)
    {
        var startXrefMarker = System.Text.Encoding.ASCII.GetBytes("startxref");
        var xrefMarker = System.Text.Encoding.ASCII.GetBytes("xref");

        var startXrefOffset = LastIndexOfBytes(bytes, startXrefMarker);

        if (startXrefOffset < 0)
            return false;

        var xrefValue = TryReadPdfIntegerAfter(bytes, startXrefOffset + startXrefMarker.Length);

        if (!xrefValue.HasValue ||
            xrefValue.Value < 0 ||
            xrefValue.Value + xrefMarker.Length >= bytes.Length)
        {
            return false;
        }

        for (var i = 0; i < xrefMarker.Length; i++)
        {
            if (bytes[xrefValue.Value + i] != xrefMarker[i])
                return false;
        }

        return PdfXrefObjectOffsetsLookValid(bytes, xrefValue.Value);
    }

    private static int? TryReadPdfIntegerAfter(byte[] bytes, int offset)
    {
        while (offset < bytes.Length && IsPdfWhitespace(bytes[offset]))
            offset++;

        if (offset >= bytes.Length || bytes[offset] < (byte)'0' || bytes[offset] > (byte)'9')
            return null;

        long value = 0;

        while (offset < bytes.Length && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9')
        {
            value = (value * 10) + (bytes[offset] - (byte)'0');

            if (value > int.MaxValue)
                return null;

            offset++;
        }

        return (int)value;
    }

    private static bool PdfXrefObjectOffsetsLookValid(byte[] bytes, int xrefOffset)
    {
        var cursor = xrefOffset + "xref".Length;

        if (!TryReadPdfLine(bytes, ref cursor, out _))
            return false;

        var checkedEntries = 0;
        var validEntries = 0;

        while (TryReadPdfLine(bytes, ref cursor, out var line))
        {
            line = line.Trim();

            if (line.StartsWith("trailer", StringComparison.Ordinal))
                break;

            var subsection = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (subsection.Length != 2 ||
                !int.TryParse(subsection[0], out _) ||
                !int.TryParse(subsection[1], out var count))
            {
                return false;
            }

            for (var i = 0; i < count; i++)
            {
                if (!TryReadPdfLine(bytes, ref cursor, out var entry))
                    return false;

                if (entry.Length < 18)
                    return false;

                if (entry[17] != 'n')
                    continue;

                if (!int.TryParse(entry.Substring(0, 10), out var objectOffset))
                    return false;

                checkedEntries++;

                if (LooksLikePdfObjectHeaderAt(bytes, objectOffset))
                    validEntries++;
            }
        }

        return checkedEntries > 0 && validEntries == checkedEntries;
    }

    private static bool TryReadPdfLine(byte[] bytes, ref int offset, out string line)
    {
        line = "";

        if (offset >= bytes.Length)
            return false;

        var start = offset;

        while (offset < bytes.Length && bytes[offset] != 0x0A && bytes[offset] != 0x0D)
            offset++;

        line = System.Text.Encoding.ASCII.GetString(bytes, start, offset - start);

        while (offset < bytes.Length && (bytes[offset] == 0x0A || bytes[offset] == 0x0D))
            offset++;

        return true;
    }

    private static bool LooksLikePdfObjectHeaderAt(byte[] bytes, int offset)
    {
        if (offset < 0 || offset >= bytes.Length)
            return false;

        while (offset < bytes.Length && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9')
            offset++;

        if (offset >= bytes.Length || !IsPdfWhitespace(bytes[offset]))
            return false;

        while (offset < bytes.Length && IsPdfWhitespace(bytes[offset]))
            offset++;

        while (offset < bytes.Length && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9')
            offset++;

        if (offset >= bytes.Length || !IsPdfWhitespace(bytes[offset]))
            return false;

        while (offset < bytes.Length && IsPdfWhitespace(bytes[offset]))
            offset++;

        var obj = System.Text.Encoding.ASCII.GetBytes("obj");

        if (offset + obj.Length > bytes.Length)
            return false;

        for (var i = 0; i < obj.Length; i++)
        {
            if (bytes[offset + i] != obj[i])
                return false;
        }

        return true;
    }

    private static bool IsPdfWhitespace(byte value)
    {
        return value == 0x00 ||
               value == 0x09 ||
               value == 0x0A ||
               value == 0x0C ||
               value == 0x0D ||
               value == 0x20;
    }

    private sealed class PdfRecoveryCandidate
    {
        public PdfRecoveryCandidate(string name, byte[] bytes)
        {
            Name = name;
            Bytes = bytes;
        }

        public string Name { get; }

        public byte[] Bytes { get; }
    }

    private sealed class PdfXrefEntry
    {
        public PdfXrefEntry(int objectNumber, int offset)
        {
            ObjectNumber = objectNumber;
            Offset = offset;
        }

        public int ObjectNumber { get; }

        public int Offset { get; }
    }

    private sealed class PdfObjectPosition
    {
        public PdfObjectPosition(int objectNumber, int offset)
        {
            ObjectNumber = objectNumber;
            Offset = offset;
        }

        public int ObjectNumber { get; }

        public int Offset { get; }
    }

    private sealed class SegmentOverlap
    {
        public SegmentOverlap(int index, int start, int end, int overlapStart, int overlapEnd)
        {
            Index = index;
            Start = start;
            End = end;
            OverlapStart = overlapStart;
            OverlapEnd = overlapEnd;
        }

        public int Index { get; }

        public int Start { get; }

        public int End { get; }

        public int OverlapStart { get; }

        public int OverlapEnd { get; }
    }

    private static int ConsumePdfEofWhitespace(
        byte[] data,
        int eofOffset,
        int eofLength,
        int maxEnd)
    {
        var end = Math.Min(eofOffset + eofLength, maxEnd);

        while (end < maxEnd)
        {
            var value = data[end];

            if (value != 0x00 &&
                value != 0x09 &&
                value != 0x0A &&
                value != 0x0C &&
                value != 0x0D &&
                value != 0x20)
            {
                break;
            }

            end++;
        }

        return end;
    }

    private static int GetCobaltFramedPayloadEnd(
        byte[] data,
        byte[] suffixMarker,
        int payloadStart)
    {
        const int cobaltRowSuffixLength = 28;

        var suffixOffset = data.Length - cobaltRowSuffixLength;

        if (suffixOffset <= payloadStart ||
            suffixOffset < 0 ||
            suffixOffset + suffixMarker.Length > data.Length)
        {
            return data.Length;
        }

        for (var i = 0; i < suffixMarker.Length; i++)
        {
            if (data[suffixOffset + i] != suffixMarker[i])
                return data.Length;
        }

        return suffixOffset;
    }


    private Task RunWithImpersonationAsync(Func<Task> action)
    {
        if (_windowsCredentials == null)
            return action();

        return Task.Run(() =>
        {
            using (SafeAccessTokenHandle userHandle =
                _windowsCredentials.LogonUser(LogonType.NewCredentials))
            {
                using (var context = WindowsIdentity.Impersonate(userHandle.DangerousGetHandle()))
                {
                    action().GetAwaiter().GetResult();
                }
            }
        });
    }

    private Task<T> RunWithImpersonationAsync<T>(Func<Task<T>> action)
    {
        if (_windowsCredentials == null)
            return action();

        return Task.Run(() =>
        {
            using (SafeAccessTokenHandle userHandle =
                _windowsCredentials.LogonUser(LogonType.NewCredentials))
            {
                using (var context = WindowsIdentity.Impersonate(userHandle.DangerousGetHandle()))
                {
                    return action().GetAwaiter().GetResult();
                }
            }
        });
    }

    private static Task<bool> LooksLikePlausibleExportAsync(
    string outputPath,
    DocumentItemInfo item,
    long exportedBytes)
    {
        if (!File.Exists(outputPath))
            return Task.FromResult(false);

        var fileInfo = new FileInfo(outputPath);

        if (fileInfo.Length == 0)
            return Task.FromResult(false);

        if (exportedBytes <= 0)
            return Task.FromResult(false);

        if (item.SizeBytes.HasValue)
        {
            return Task.FromResult(fileInfo.Length == item.SizeBytes.Value);
        }

        // If SharePoint size is unknown, we can only confirm that bytes were written.
        return Task.FromResult(true);
    }

    private static bool LooksLikePdfFile(string path)
    {
        if (!File.Exists(path))
            return false;

        var bytes = File.ReadAllBytes(path);

        if (bytes.Length < 10)
            return false;

        var pdfHeader = System.Text.Encoding.ASCII.GetBytes("%PDF-");
        var pdfEof = System.Text.Encoding.ASCII.GetBytes("%%EOF");

        return IndexOfBytes(bytes, pdfHeader) == 0 &&
               LastIndexOfBytes(bytes, pdfEof) >= 0;
    }

    private static int IndexOfBytes(byte[] data, byte[] pattern)
    {
        return IndexOfBytes(data, pattern, 0);
    }

    private static int IndexOfBytes(byte[] data, byte[] pattern, int startOffset)
    {
        if (pattern.Length == 0 || data.Length < pattern.Length)
            return -1;

        startOffset = Math.Max(0, startOffset);

        for (var i = startOffset; i <= data.Length - pattern.Length; i++)
        {
            var match = true;

            for (var j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
                return i;
        }

        return -1;
    }

    private static int LastIndexOfBytes(byte[] data, byte[] pattern)
    {
        if (pattern.Length == 0 || data.Length < pattern.Length)
            return -1;

        for (var i = data.Length - pattern.Length; i >= 0; i--)
        {
            var match = true;

            for (var j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
                return i;
        }

        return -1;
    }

    public async Task DiagnoseSelectedFileStreamHeadersAsync(
    string databaseName,
    DocumentItemInfo item)
    {
        await RunWithImpersonationAsync(async () =>
   {
       using var connection = new SqlConnection(_masterConnectionString);
       await connection.OpenAsync();

       var db = QuoteDatabaseName(databaseName);

       using var command = connection.CreateCommand();

       command.CommandText = $@"
DECLARE @latestHistVersion int;

SELECT @latestHistVersion = MAX(HistVersion)
FROM {db}.dbo.DocsToStreams
WHERE SiteId = @siteId
  AND DocId = @docId;

SELECT TOP (200)
    dts.HistVersion,
    dts.Partition,
    dts.BSN,
    dts.StreamId,
    ds.Type,
    ds.Size,
    DATALENGTH(ds.Content) AS ContentLength,
    CONVERT(varchar(64), SUBSTRING(ds.Content, 1, 16), 2) AS First16Hex,
    CHARINDEX(0x255044462D, ds.Content) AS PdfHeaderOffset
FROM {db}.dbo.DocsToStreams dts
INNER JOIN {db}.dbo.DocStreams ds
    ON ds.SiteId = dts.SiteId
   AND ds.DocId = dts.DocId
   AND ds.Partition = dts.Partition
   AND ds.BSN = dts.BSN
WHERE dts.SiteId = @siteId
  AND dts.DocId = @docId
  AND dts.HistVersion = @latestHistVersion
ORDER BY
    CASE WHEN CHARINDEX(0x255044462D, ds.Content) > 0 THEN 0 ELSE 1 END,
    ds.Type,
    dts.StreamId,
    dts.Partition,
    dts.BSN;";

       command.Parameters.Add("@siteId", SqlDbType.UniqueIdentifier).Value = item.SiteId;
       command.Parameters.Add("@docId", SqlDbType.UniqueIdentifier).Value = item.Id;

       using var reader = await command.ExecuteReaderAsync();

       var table = new Table()
           .Border(TableBorder.Rounded)
           .Title("Stream Header Diagnostic");

       table.AddColumn("Hist");
       table.AddColumn("Type");
       table.AddColumn("StreamId");
       table.AddColumn("Partition");
       table.AddColumn("BSN");
       table.AddColumn("Size");
       table.AddColumn("ContentLen");
       table.AddColumn("First16Hex");
       table.AddColumn("PDF Offset");

       while (await reader.ReadAsync())
       {
           table.AddRow(
               reader["HistVersion"]?.ToString() ?? "",
               reader["Type"]?.ToString() ?? "",
               reader["StreamId"]?.ToString() ?? "",
               reader["Partition"]?.ToString() ?? "",
               reader["BSN"]?.ToString() ?? "",
               reader["Size"]?.ToString() ?? "",
               reader["ContentLength"]?.ToString() ?? "",
               Markup.Escape(reader["First16Hex"]?.ToString() ?? ""),
               reader["PdfHeaderOffset"]?.ToString() ?? "");
       }

       AnsiConsole.Write(table);
   });
    }

    public async Task DiagnoseSelectedFileStreamsAsync(
        string databaseName,
        DocumentItemInfo item)
    {
        await RunWithImpersonationAsync(async () =>
    {
        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT
    dts.HistVersion,
    dts.Partition,
    dts.BSN,
    dts.StreamId,
    ds.Size,
    DATALENGTH(ds.Content) AS ContentLength,
    ds.Type,
    CASE 
        WHEN ds.RbsId IS NULL THEN 0
        WHEN DATALENGTH(ds.RbsId) = 0 THEN 0
        ELSE 1
    END AS HasRbsId
FROM {db}.dbo.DocsToStreams dts
INNER JOIN {db}.dbo.DocStreams ds
    ON ds.SiteId = dts.SiteId
   AND ds.DocId = dts.DocId
   AND ds.Partition = dts.Partition
   AND ds.BSN = dts.BSN
WHERE dts.SiteId = @siteId
  AND dts.DocId = @docId
ORDER BY
    dts.HistVersion,
    dts.StreamId,
    dts.Partition,
    dts.BSN;";

        command.Parameters.Add("@siteId", SqlDbType.UniqueIdentifier).Value = item.SiteId;
        command.Parameters.Add("@docId", SqlDbType.UniqueIdentifier).Value = item.Id;

        using var reader = await command.ExecuteReaderAsync();

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("Selected File Stream Rows");

        table.AddColumn("HistVersion");
        table.AddColumn("Partition");
        table.AddColumn("BSN");
        table.AddColumn("StreamId");
        table.AddColumn("Size");
        table.AddColumn("ContentLength");
        table.AddColumn("Type");
        table.AddColumn("RBS");

        while (await reader.ReadAsync())
        {
            table.AddRow(
                reader["HistVersion"]?.ToString() ?? "",
                reader["Partition"]?.ToString() ?? "",
                reader["BSN"]?.ToString() ?? "",
                reader["StreamId"]?.ToString() ?? "",
                reader["Size"]?.ToString() ?? "",
                reader["ContentLength"]?.ToString() ?? "",
                reader["Type"]?.ToString() ?? "",
                reader["HasRbsId"]?.ToString() ?? "");
        }

        AnsiConsole.Write(table);
    });
    }

    private static async Task PrintExportValidationAsync(
    string outputPath,
    DocumentItemInfo item,
    long exportedBytes)
    {
        if (!File.Exists(outputPath))
        {
            AnsiConsole.MarkupLine("[red]Output file was not created.[/]");
            return;
        }

        var fileInfo = new FileInfo(outputPath);

        var firstBytes = new byte[Math.Min(16, (int)Math.Min(fileInfo.Length, 16))];

        using (var fs = new FileStream(
            outputPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 4096,
            useAsync: true))
        {
            await fs.ReadAsync(firstBytes, 0, firstBytes.Length);
        }

        var hex = BitConverter.ToString(firstBytes);

        AnsiConsole.MarkupLine($"[grey]First bytes HEX:[/] {Markup.Escape(hex)}");
        AnsiConsole.MarkupLine($"[grey]Output file size:[/] {fileInfo.Length:N0} bytes");

        if (item.SizeBytes.HasValue)
        {
            AnsiConsole.MarkupLine($"[grey]SharePoint file size:[/] {item.SizeBytes.Value:N0} bytes");

            if (fileInfo.Length == item.SizeBytes.Value)
            {
                AnsiConsole.MarkupLine("[green]Size check passed.[/]");
            }
            else
            {
                AnsiConsole.MarkupLine(
                    $"[red]Size check failed.[/] Expected {item.SizeBytes.Value:N0}, exported {fileInfo.Length:N0}.");
            }
        }
    }

    private async Task<long> TryExportShreddedFileAsync(
    SqlConnection connection,
    string quotedDatabaseName,
    DocumentItemInfo item,
    string outputPath,
    bool useExactInternalVersion,
    int? streamTypeFilter = null)
    {
        using var command = connection.CreateCommand();

        string versionFilterSql;

        if (useExactInternalVersion && item.InternalVersion.HasValue)
        {
            versionFilterSql = "AND dts.HistVersion = @histVersion";
            command.Parameters.Add("@histVersion", SqlDbType.Int).Value = item.InternalVersion.Value;
        }
        else
        {
            versionFilterSql = $@"
AND dts.HistVersion =
(
    SELECT MAX(dts2.HistVersion)
    FROM {quotedDatabaseName}.dbo.DocsToStreams dts2
    WHERE dts2.SiteId = @siteId
      AND dts2.DocId = @docId
)";
        }

        command.CommandText = $@"
SELECT
    ds.RbsId,
    ds.Size,
    dts.HistVersion,
    dts.Partition,
    dts.BSN,
    dts.StreamId,
    ds.Type,
    ds.Content
FROM {quotedDatabaseName}.dbo.DocsToStreams dts
INNER JOIN {quotedDatabaseName}.dbo.DocStreams ds
    ON ds.SiteId = dts.SiteId
   AND ds.DocId = dts.DocId
   AND ds.Partition = dts.Partition
   AND ds.BSN = dts.BSN
WHERE dts.SiteId = @siteId
  AND dts.DocId = @docId
  {versionFilterSql}
  AND (@streamTypeFilter IS NULL OR ds.Type = @streamTypeFilter)
ORDER BY
    dts.StreamId,
    dts.Partition,
    dts.BSN;";

        command.Parameters.Add("@siteId", SqlDbType.UniqueIdentifier).Value = item.SiteId;
        command.Parameters.Add("@docId", SqlDbType.UniqueIdentifier).Value = item.Id;
        command.Parameters.Add("@streamTypeFilter", SqlDbType.Int).Value =
    streamTypeFilter.HasValue ? (object)streamTypeFilter.Value : DBNull.Value;

        using var reader = await command.ExecuteReaderAsync(CommandBehavior.SequentialAccess);

        long totalBytesWritten = 0;
        var wroteAnyRow = false;

        using var output = new FileStream(
            outputPath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 1024 * 128,
            useAsync: true);

        while (await reader.ReadAsync())
        {
            var rbsOrdinal = reader.GetOrdinal("RbsId");

            if (!await reader.IsDBNullAsync(rbsOrdinal))
            {
                var rbsValue = reader.GetValue(rbsOrdinal);

                if (rbsValue is byte[] rbsBytes && rbsBytes.Length > 0)
                {
                    throw new InvalidOperationException(
                        "This file uses RBS/external BLOB storage. This exporter cannot read external BLOB content yet.");
                }
            }

            var contentOrdinal = reader.GetOrdinal("Content");

            if (await reader.IsDBNullAsync(contentOrdinal))
                continue;

            wroteAnyRow = true;

            const int bufferSize = 1024 * 128;
            var buffer = new byte[bufferSize];
            long offset = 0;

            while (true)
            {
                var bytesRead = reader.GetBytes(
                    contentOrdinal,
                    offset,
                    buffer,
                    0,
                    buffer.Length);

                if (bytesRead == 0)
                    break;

                await output.WriteAsync(buffer, 0, (int)bytesRead);

                offset += bytesRead;
                totalBytesWritten += bytesRead;
            }
        }

        if (!wroteAnyRow)
        {
            output.Close();

            if (File.Exists(outputPath))
                File.Delete(outputPath);

            return 0;
        }

        return totalBytesWritten;
    }
    public async Task DiagnoseStreamStorageAsync(string databaseName)
    {
        await RunWithImpersonationAsync(async () =>
    {
        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT 
    s.name AS SchemaName,
    o.name AS ObjectName,
    o.type_desc AS ObjectType,
    c.name AS ColumnName,
    ty.name AS DataType,
    c.max_length
FROM {db}.sys.objects o
INNER JOIN {db}.sys.schemas s
    ON s.schema_id = o.schema_id
INNER JOIN {db}.sys.columns c
    ON c.object_id = o.object_id
INNER JOIN {db}.sys.types ty
    ON ty.user_type_id = c.user_type_id
WHERE 
    o.type IN ('U', 'V')
    AND
    (
        o.name LIKE '%Doc%'
        OR o.name LIKE '%Stream%'
        OR c.name LIKE '%Content%'
        OR c.name LIKE '%Rbs%'
        OR c.name LIKE '%Stream%'
        OR ty.name IN ('varbinary', 'image')
    )
ORDER BY 
    o.name,
    c.column_id;";

        using var reader = await command.ExecuteReaderAsync();

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("Possible Document Stream Storage Objects");

        table.AddColumn("Schema");
        table.AddColumn("Object");
        table.AddColumn("Type");
        table.AddColumn("Column");
        table.AddColumn("Data Type");
        table.AddColumn("Max Length");

        while (await reader.ReadAsync())
        {
            table.AddRow(
                Markup.Escape(reader["SchemaName"]?.ToString() ?? ""),
                Markup.Escape(reader["ObjectName"]?.ToString() ?? ""),
                Markup.Escape(reader["ObjectType"]?.ToString() ?? ""),
                Markup.Escape(reader["ColumnName"]?.ToString() ?? ""),
                Markup.Escape(reader["DataType"]?.ToString() ?? ""),
                Markup.Escape(reader["max_length"]?.ToString() ?? ""));
        }

        AnsiConsole.Write(table);
    });
    }

    private readonly string _masterConnectionString;
    private readonly UserCredentials? _windowsCredentials;
    private readonly string? _cobaltCoreDllPath;

    private static readonly string[] ImportantSharePointTables =
    {
        "AllSites",
        "Webs",
        "AllLists",
        "AllDocs",
        "AllUserData",
        "UserInfo",
        "Versions"
    };

    public SharePointDatabaseScanner(
    string masterConnectionString,
    UserCredentials? windowsCredentials,
    string? cobaltCoreDllPath)
    {
        _masterConnectionString = masterConnectionString;
        _windowsCredentials = windowsCredentials;
        _cobaltCoreDllPath = cobaltCoreDllPath;
    }

    public async Task TestConnectionAsync()
    {
        await RunWithImpersonationAsync(async () =>
        {
            using var connection = new SqlConnection(_masterConnectionString);
            await connection.OpenAsync();

            using var command = connection.CreateCommand();
            command.CommandText = @"
SELECT 
    @@SERVERNAME AS ServerName,
    SERVERPROPERTY('MachineName') AS MachineName,
    SERVERPROPERTY('InstanceName') AS InstanceName,
    SERVERPROPERTY('ServerName') AS FullServerName,
    SUSER_SNAME() AS SqlLogin,
    ORIGINAL_LOGIN() AS OriginalLogin;";

            using var reader = await command.ExecuteReaderAsync();

            if (await reader.ReadAsync())
            {
                AnsiConsole.MarkupLine($"[green]Connected to SQL Server:[/] {Markup.Escape(reader["FullServerName"]?.ToString() ?? "?")}");
                AnsiConsole.MarkupLine($"[grey]Machine:[/] {Markup.Escape(reader["MachineName"]?.ToString() ?? "?")}");
                AnsiConsole.MarkupLine($"[grey]Instance:[/] {Markup.Escape(reader["InstanceName"]?.ToString() ?? "Default")}");
                AnsiConsole.MarkupLine($"[grey]SQL login:[/] {Markup.Escape(reader["SqlLogin"]?.ToString() ?? "?")}");
            }

            AnsiConsole.WriteLine();
        });
    }

    public async Task<List<DatabaseCandidate>> FindCandidateDatabasesAsync()
    {
        return await RunWithImpersonationAsync(async () =>
        {
            var databaseNames = await GetOnlineDatabaseNamesAsync();
            var result = new List<DatabaseCandidate>();

            foreach (var databaseName in databaseNames)
            {
                try
                {
                    var candidate = await InspectDatabaseAsync(databaseName, verbose: false);

                    if (candidate != null)
                        result.Add(candidate);
                }
                catch
                {
                }
            }

            return result
                .OrderByDescending(x => x.IsLikelySharePointContentDatabase)
                .ThenByDescending(x => x.MatchedTableCount)
                .ThenBy(x => x.Name)
                .ToList();
        });
    }
    public async Task<DatabaseCandidate?> InspectDatabaseAsync(string databaseName, bool verbose = true)
    {
        return await RunWithImpersonationAsync(async () =>
   {
       if (!await DatabaseExistsAsync(databaseName))
           return null;

       var existingTables = await GetExistingImportantTablesAsync(databaseName, verbose);

       var hasAllSites = existingTables.Contains("AllSites");
       var hasWebs = existingTables.Contains("Webs");
       var hasAllLists = existingTables.Contains("AllLists");
       var hasAllDocs = existingTables.Contains("AllDocs");
       var hasAllUserData = existingTables.Contains("AllUserData");

       var isCandidate =
   hasAllSites &&
   hasAllLists &&
   hasAllDocs &&
   hasAllUserData;

       if (verbose)
       {
           AnsiConsole.WriteLine();
           AnsiConsole.MarkupLine("[bold yellow]Detection result:[/]");

           AnsiConsole.MarkupLine($"AllSites:    {(hasAllSites ? "[green]Yes[/]" : "[red]No[/]")}");
           AnsiConsole.MarkupLine($"Webs:        {(hasWebs ? "[green]Yes[/]" : "[red]No[/]")}");
           AnsiConsole.MarkupLine($"AllLists:    {(hasAllLists ? "[green]Yes[/]" : "[red]No[/]")}");
           AnsiConsole.MarkupLine($"AllDocs:     {(hasAllDocs ? "[green]Yes[/]" : "[red]No[/]")}");
           AnsiConsole.MarkupLine($"AllUserData: {(hasAllUserData ? "[green]Yes[/]" : "[red]No[/]")}");
           AnsiConsole.WriteLine();
       }

       if (!isCandidate)
           return null;

       long? siteCollectionCount = null;

       if (hasAllSites)
           siteCollectionCount = await GetSiteCollectionCountAsync(databaseName);

       return new DatabaseCandidate
       {
           Name = databaseName,
           MatchedTableCount = existingTables.Count,
           HasAllSites = hasAllSites,
           HasWebs = hasWebs,
           HasAllLists = hasAllLists,
           HasAllDocs = hasAllDocs,
           HasAllUserData = hasAllUserData,
           SiteCollectionCount = siteCollectionCount
       };
   });
    }

    public async Task<SharePointDatabaseSnapshot> ReadSitesAndWebsAsync(string databaseName)
    {
        return await RunWithImpersonationAsync(async () =>
        {
            var snapshot = new SharePointDatabaseSnapshot();

            using var connection = new SqlConnection(_masterConnectionString);
            await connection.OpenAsync();

            var db = QuoteDatabaseName(databaseName);

            using (var command = connection.CreateCommand())
            {
                command.CommandText = $@"
SELECT
    Id,
    FullUrl,
    RootWebId,
    TimeCreated
FROM {db}.dbo.AllSites
WHERE ISNULL(Deleted, 0) = 0
ORDER BY FullUrl;";

                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    snapshot.SiteCollections.Add(new SiteCollectionInfo
                    {
                        Id = reader.GetGuid(reader.GetOrdinal("Id")),
                        FullUrl = reader["FullUrl"] as string ?? string.Empty,
                        RootWebId = reader.GetGuid(reader.GetOrdinal("RootWebId")),
                        TimeCreated = reader["TimeCreated"] == DBNull.Value
                            ? null
                            : Convert.ToDateTime(reader["TimeCreated"])
                    });
                }
            }

            var hasWebsObject = await ObjectExistsAsync(connection, databaseName, "Webs");

            if (hasWebsObject)
            {
                using var command = connection.CreateCommand();

                command.CommandText = $@"
SELECT
    Id,
    SiteId,
    FullUrl,
    Title,
    ParentWebId
FROM {db}.dbo.Webs
ORDER BY SiteId, FullUrl;";

                using var reader = await command.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    snapshot.Webs.Add(new WebInfo
                    {
                        Id = reader.GetGuid(reader.GetOrdinal("Id")),
                        SiteId = reader.GetGuid(reader.GetOrdinal("SiteId")),
                        FullUrl = reader["FullUrl"] as string ?? string.Empty,
                        Title = reader["Title"] as string ?? string.Empty,
                        ParentWebId = reader["ParentWebId"] == DBNull.Value
                            ? null
                            : reader.GetGuid(reader.GetOrdinal("ParentWebId"))
                    });
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]Webs object was not found. Showing site collections only for now.[/]");
            }

            return snapshot;
        });
    }

    private async Task<List<string>> GetOnlineDatabaseNamesAsync()
    {
        var result = new List<string>();

        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT name
FROM sys.databases
WHERE state_desc = 'ONLINE'
  AND database_id > 4
ORDER BY name;";

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(reader.GetString(0));
        }

        return result;
    }

    private async Task<bool> DatabaseExistsAsync(string databaseName)
    {
        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        using var command = connection.CreateCommand();
        command.CommandText = @"
SELECT COUNT_BIG(*)
FROM sys.databases
WHERE name = @name
  AND state_desc = 'ONLINE';";

        command.Parameters.Add("@name", SqlDbType.NVarChar, 128).Value = databaseName;

        var count = Convert.ToInt64(await command.ExecuteScalarAsync());

        return count > 0;
    }

    private async Task<HashSet<string>> GetExistingImportantTablesAsync(string databaseName, bool verbose = true)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();
        command.CommandText = $@"
SELECT 
    s.name AS SchemaName,
    o.name AS ObjectName,
    o.type_desc AS ObjectType
FROM {db}.sys.objects o
INNER JOIN {db}.sys.schemas s
    ON s.schema_id = o.schema_id
WHERE o.name IN ({BuildSqlStringList(ImportantSharePointTables)})
  AND o.type IN ('U', 'V')
ORDER BY o.name;";

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            var schemaName = reader["SchemaName"]?.ToString() ?? "";
            var objectName = reader["ObjectName"]?.ToString() ?? "";
            var objectType = reader["ObjectType"]?.ToString() ?? "";

            result.Add(objectName);

            if (verbose)
            {
                AnsiConsole.MarkupLine(
                    $"[grey]Found object:[/] {Markup.Escape(schemaName)}.{Markup.Escape(objectName)} [grey]({Markup.Escape(objectType)})[/]");
            }
        }

        return result;
    }

    private async Task<long> GetSiteCollectionCountAsync(string databaseName)
    {
        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();
        command.CommandText = $@"
SELECT COUNT_BIG(*)
FROM {db}.dbo.AllSites
WHERE ISNULL(Deleted, 0) = 0;";

        return Convert.ToInt64(await command.ExecuteScalarAsync());
    }

    private static string QuoteDatabaseName(string databaseName)
    {
        return "[" + databaseName.Replace("]", "]]") + "]";
    }

    private static string QuoteSqlIdentifier(string value)
    {
        return "[" + value.Replace("]", "]]") + "]";
    }

    private static string BuildSqlStringList(IEnumerable<string> values)
    {
        return string.Join(", ", values.Select(v => "'" + v.Replace("'", "''") + "'"));
    }

    private static string Csv(string? value)
    {
        value ??= "";

        if (value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) < 0)
            return value;

        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    private static string GetUsefulExceptionMessage(Exception ex)
    {
        var parts = new List<string>();
        var current = ex;

        while (current != null)
        {
            parts.Add(current.GetType().Name + ": " + current.Message);
            current = current.InnerException;
        }

        return string.Join(" -> ", parts);
    }

    private static async Task<bool> ObjectExistsAsync(
    SqlConnection connection,
    string databaseName,
    string objectName)
    {
        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT COUNT_BIG(*)
FROM {db}.sys.objects o
INNER JOIN {db}.sys.schemas s
    ON s.schema_id = o.schema_id
WHERE s.name = 'dbo'
  AND o.name = @objectName
  AND o.type IN ('U', 'V');";

        command.Parameters.Add("@objectName", SqlDbType.NVarChar, 128).Value = objectName;

        var count = Convert.ToInt64(await command.ExecuteScalarAsync());

        return count > 0;
    }

    public async Task<List<DocumentLibraryInfo>> ReadDocumentLibrariesForWebAsync(
    string databaseName,
    Guid webId)
    {
        return await RunWithImpersonationAsync(async () =>
    {
        var result = new List<DocumentLibraryInfo>();

        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);

        var columns = await GetColumnNamesAsync(connection, databaseName, "AllLists");

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT
    tp_ID AS Id,
    tp_WebId AS WebId,
    tp_Title AS Title,
    tp_BaseType AS BaseType,
    tp_ServerTemplate AS ServerTemplate,
    {SelectNullableInt(columns, "tp_ItemCount", "ItemCount")},
    {SelectNullableDateTime(columns, "tp_Created", "Created")},
    {SelectNullableDateTime(columns, "tp_Modified", "Modified")},
    {SelectNullableDateTime(columns, "tp_LastItemModifiedDate", "LastItemModifiedDate")}
FROM {db}.dbo.AllLists
WHERE tp_WebId = @webId
  AND tp_BaseType = 1
ORDER BY tp_Title;";

        command.Parameters.Add("@webId", SqlDbType.UniqueIdentifier).Value = webId;

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(new DocumentLibraryInfo
            {
                Id = reader.GetGuid(reader.GetOrdinal("Id")),
                WebId = reader.GetGuid(reader.GetOrdinal("WebId")),
                Title = reader["Title"] as string ?? string.Empty,
                BaseType = Convert.ToInt32(reader["BaseType"]),
                ServerTemplate = Convert.ToInt32(reader["ServerTemplate"]),
                ItemCount = GetNullableInt(reader, "ItemCount"),
                Created = GetNullableDateTime(reader, "Created"),
                Modified = GetNullableDateTime(reader, "Modified"),
                LastItemModifiedDate = GetNullableDateTime(reader, "LastItemModifiedDate")
            });
        }

        return result;
    });
    }

    private static async Task<HashSet<string>> GetColumnNamesAsync(
        SqlConnection connection,
        string databaseName,
        string objectName)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var db = QuoteDatabaseName(databaseName);

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT c.name
FROM {db}.sys.columns c
INNER JOIN {db}.sys.objects o
    ON o.object_id = c.object_id
INNER JOIN {db}.sys.schemas s
    ON s.schema_id = o.schema_id
WHERE s.name = 'dbo'
  AND o.name = @objectName
  AND o.type IN ('U', 'V')
ORDER BY c.column_id;";

        command.Parameters.Add("@objectName", SqlDbType.NVarChar, 128).Value = objectName;

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(reader.GetString(0));
        }

        return result;
    }

    private static string SelectNullableInt(
        HashSet<string> columns,
        string columnName,
        string alias)
    {
        return columns.Contains(columnName)
            ? $"[{columnName}] AS [{alias}]"
            : $"CAST(NULL AS int) AS [{alias}]";
    }

    private static string SelectNullableDateTime(
        HashSet<string> columns,
        string columnName,
        string alias)
    {
        return columns.Contains(columnName)
            ? $"[{columnName}] AS [{alias}]"
            : $"CAST(NULL AS datetime) AS [{alias}]";
    }

    private static int? GetNullableInt(IDataRecord reader, string columnName)
    {
        var value = reader[columnName];

        return value == DBNull.Value
            ? null
            : Convert.ToInt32(value);
    }

    private static DateTime? GetNullableDateTime(IDataRecord reader, string columnName)
    {
        var value = reader[columnName];

        return value == DBNull.Value
            ? null
            : Convert.ToDateTime(value);
    }

    public async Task<DocumentItemInfo?> FindLibraryRootFolderAsync(
    string databaseName,
    Guid listId)
    {
        return await RunWithImpersonationAsync(async () =>
    {
        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);
        var allDocsColumns = await GetColumnNamesAsync(connection, databaseName, "AllDocs");

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT TOP (1)
    Id,
    SiteId,
    ListId,
    ParentId,
    DirName,
    LeafName,
    Type,
    {SelectNullableLong(allDocsColumns, "Size", "SizeBytes")},
    {SelectNullableDateTime(allDocsColumns, "TimeCreated", "TimeCreated")},
    {SelectNullableDateTime(allDocsColumns, "TimeLastModified", "TimeLastModified")},
    {SelectNullableInt(allDocsColumns, "InternalVersion", "InternalVersion")},
    {SelectNullableString(allDocsColumns, "UIVersionString", "UIVersionString")},
    {SelectNullableString(allDocsColumns, "Extension", "Extension")},
{SelectNullableString(allDocsColumns, "ExtensionForFile", "ExtensionForFile")},
{SelectNullableString(allDocsColumns, "ProgId", "ProgId")}
FROM {db}.dbo.AllDocs
WHERE ListId = @listId
  AND Type <> 0
  {BuildCurrentVersionFilter(allDocsColumns)}
  {BuildNotDeletedFilter(allDocsColumns)}
ORDER BY
    CASE WHEN ParentId = '00000000-0000-0000-0000-000000000000' THEN 0 ELSE 1 END,
    LEN(DirName),
    DirName,
    LeafName;";

        command.Parameters.Add("@listId", SqlDbType.UniqueIdentifier).Value = listId;

        using var reader = await command.ExecuteReaderAsync();

        if (!await reader.ReadAsync())
            return null;

        return ReadDocumentItem(reader);
    });
    }

    public async Task<List<DocumentItemInfo>> ReadChildItemsAsync(
        string databaseName,
        Guid listId,
        Guid parentId)
    {
        return await RunWithImpersonationAsync(async () =>
    {
        var result = new List<DocumentItemInfo>();

        using var connection = new SqlConnection(_masterConnectionString);
        await connection.OpenAsync();

        var db = QuoteDatabaseName(databaseName);
        var allDocsColumns = await GetColumnNamesAsync(connection, databaseName, "AllDocs");

        using var command = connection.CreateCommand();

        command.CommandText = $@"
SELECT
    Id,
    SiteId,
    ListId,
    ParentId,
    DirName,
    LeafName,
    Type,
    {SelectNullableLong(allDocsColumns, "Size", "SizeBytes")},
    {SelectNullableDateTime(allDocsColumns, "TimeCreated", "TimeCreated")},
    {SelectNullableDateTime(allDocsColumns, "TimeLastModified", "TimeLastModified")},
    {SelectNullableInt(allDocsColumns, "InternalVersion", "InternalVersion")},
    {SelectNullableString(allDocsColumns, "UIVersionString", "UIVersionString")},
    {SelectNullableString(allDocsColumns, "Extension", "Extension")},
{SelectNullableString(allDocsColumns, "ExtensionForFile", "ExtensionForFile")},
{SelectNullableString(allDocsColumns, "ProgId", "ProgId")}
FROM {db}.dbo.AllDocs
WHERE ListId = @listId
  AND ParentId = @parentId
  {BuildCurrentVersionFilter(allDocsColumns)}
  {BuildNotDeletedFilter(allDocsColumns)}
ORDER BY
    CASE WHEN Type = 0 THEN 1 ELSE 0 END,
    LeafName;";

        command.Parameters.Add("@listId", SqlDbType.UniqueIdentifier).Value = listId;
        command.Parameters.Add("@parentId", SqlDbType.UniqueIdentifier).Value = parentId;

        using var reader = await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            result.Add(ReadDocumentItem(reader));
        }

        return result;
    });
    }

    public async Task ExportCurrentFileAsync(
    string databaseName,
    DocumentItemInfo item,
    string exportFolder,
    bool showStreamDiagnostics = false)
    {
        await RunWithImpersonationAsync(async () =>
        {
            if (item.IsFolder)
                throw new InvalidOperationException("Selected item is a folder, not a file.");

            Directory.CreateDirectory(exportFolder);

            using var connection = new SqlConnection(_masterConnectionString);
            await connection.OpenAsync();

            var db = QuoteDatabaseName(databaseName);

            var docsToStreamsColumns = await GetColumnNamesAsync(connection, databaseName, "DocsToStreams");
            var docStreamsColumns = await GetColumnNamesAsync(connection, databaseName, "DocStreams");

            if (!docsToStreamsColumns.Contains("DocId") ||
                !docsToStreamsColumns.Contains("SiteId") ||
                !docsToStreamsColumns.Contains("Partition") ||
                !docsToStreamsColumns.Contains("BSN"))
            {
                throw new InvalidOperationException("DocsToStreams table does not contain expected columns.");
            }

            if (!docStreamsColumns.Contains("DocId") ||
                !docStreamsColumns.Contains("SiteId") ||
                !docStreamsColumns.Contains("Partition") ||
                !docStreamsColumns.Contains("BSN") ||
                !docStreamsColumns.Contains("Content"))
            {
                throw new InvalidOperationException("DocStreams table does not contain expected columns.");
            }

            var safeName = MakeSafeFileName(item.LeafName);
            var outputPath = GetAvailableOutputPath(exportFolder, safeName);

            AnsiConsole.MarkupLine("[grey]Using SharePoint content DB shredded-storage parser.[/]");

            if (showStreamDiagnostics)
            {
                await DiagnoseSelectedFileStreamsAsync(databaseName, item);
            }

            var result = await AnalyzeSharePointShreddedStorageAsync(
                connection,
                db,
                item,
                outputPath);

            if (!result.Success)
            {
                AnsiConsole.MarkupLine("[red]Export failed.[/]");
                AnsiConsole.MarkupLine($"[yellow]{Markup.Escape(result.Message)}[/]");
                return;
            }

            AnsiConsole.MarkupLine(
                $"[green]Exported:[/] [yellow]{Markup.Escape(result.OutputPath ?? outputPath)}[/]");

            AnsiConsole.MarkupLine(
                $"[green]Written bytes:[/] {result.BytesWritten:N0}");

            if (!string.IsNullOrWhiteSpace(result.Message))
            {
                AnsiConsole.MarkupLine(
                    $"[grey]{Markup.Escape(result.Message)}[/]");
            }

            var finalOutputPath = result.OutputPath ?? outputPath;

            await PrintExportValidationAsync(finalOutputPath, item, result.BytesWritten);

            var looksValid = await LooksLikePlausibleExportAsync(finalOutputPath, item, result.BytesWritten);

            if (!looksValid)
            {
                AnsiConsole.MarkupLine(
                    "[red]Export finished, but output does not match SharePoint file size.[/]");

                return;
            }

            var finalFileInfo = new FileInfo(finalOutputPath);

            if (item.SizeBytes.HasValue && finalFileInfo.Length != item.SizeBytes.Value)
            {
                AnsiConsole.MarkupLine(
                    "[yellow]Export produced a file-like payload, but the size mismatch remains. Do not treat this as a complete disaster-recovery export.[/]");

                return;
            }

            AnsiConsole.MarkupLine("[green]Export completed successfully.[/]");
        });
    }

    private static DocumentItemInfo ReadDocumentItem(IDataRecord reader)
    {
        return new DocumentItemInfo
        {
            Id = reader.GetGuid(reader.GetOrdinal("Id")),
            SiteId = reader.GetGuid(reader.GetOrdinal("SiteId")),
            ListId = reader["ListId"] == DBNull.Value
                ? Guid.Empty
                : reader.GetGuid(reader.GetOrdinal("ListId")),
            ParentId = reader["ParentId"] == DBNull.Value
                ? null
                : reader.GetGuid(reader.GetOrdinal("ParentId")),
            DirName = reader["DirName"] as string ?? string.Empty,
            LeafName = reader["LeafName"] as string ?? string.Empty,
            Type = Convert.ToInt32(reader["Type"]),
            SizeBytes = GetNullableLong(reader, "SizeBytes"),
            TimeCreated = GetNullableDateTime(reader, "TimeCreated"),
            TimeLastModified = GetNullableDateTime(reader, "TimeLastModified"),
            InternalVersion = GetNullableInt(reader, "InternalVersion"),
            UIVersionString = reader["UIVersionString"] == DBNull.Value
                ? null
                : reader["UIVersionString"]?.ToString(),
            Extension = reader["Extension"] == DBNull.Value
                ? null
                : reader["Extension"]?.ToString(),
            ExtensionForFile = reader["ExtensionForFile"] == DBNull.Value
                ? null
                : reader["ExtensionForFile"]?.ToString(),
            ProgId = reader["ProgId"] == DBNull.Value
                ? null
                : reader["ProgId"]?.ToString()
        };
    }

    private static string BuildCurrentVersionFilter(HashSet<string> columns)
    {
        return columns.Contains("IsCurrentVersion")
            ? "AND IsCurrentVersion = 1"
            : "";
    }

    private static string BuildNotDeletedFilter(HashSet<string> columns)
    {
        return columns.Contains("DeleteTransactionId")
            ? "AND DeleteTransactionId = 0x"
            : "";
    }

    private static string SelectNullableString(
        HashSet<string> columns,
        string columnName,
        string alias)
    {
        return columns.Contains(columnName)
            ? $"[{columnName}] AS [{alias}]"
            : $"CAST(NULL AS nvarchar(max)) AS [{alias}]";
    }

    private static string SelectNullableLong(
        HashSet<string> columns,
        string columnName,
        string alias)
    {
        return columns.Contains(columnName)
            ? $"CAST([{columnName}] AS bigint) AS [{alias}]"
            : $"CAST(NULL AS bigint) AS [{alias}]";
    }

    private static long? GetNullableLong(IDataRecord reader, string columnName)
    {
        var value = reader[columnName];

        return value == DBNull.Value
            ? null
            : Convert.ToInt64(value);
    }

    private static string MakeSafeFileName(string fileName)
    {
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(invalidChar, '_');
        }

        return string.IsNullOrWhiteSpace(fileName)
            ? "exported_file.bin"
            : fileName;
    }

    private static string GetAvailableOutputPath(string folder, string fileName)
    {
        var path = Path.Combine(folder, fileName);

        if (!File.Exists(path))
            return path;

        var nameWithoutExtension = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);

        for (var i = 1; i < int.MaxValue; i++)
        {
            var candidate = Path.Combine(folder, $"{nameWithoutExtension} ({i}){extension}");

            if (!File.Exists(candidate))
                return candidate;
        }

        throw new IOException("Could not find an available export file name.");
    }

    private sealed class CobaltMappingObjectCandidate
    {
        public CobaltMappingObjectCandidate(string schemaName, string objectName, string objectType)
        {
            SchemaName = schemaName;
            ObjectName = objectName;
            ObjectType = objectType;
        }

        public string SchemaName { get; }

        public string ObjectName { get; }

        public string ObjectType { get; }

        public List<string> Columns { get; } = new();
    }

    private sealed class CobaltMappingFilter
    {
        public CobaltMappingFilter(string sql, string description)
        {
            Sql = sql;
            Description = description;
        }

        public string Sql { get; }

        public string Description { get; }
    }
}

public sealed class DocumentItemInfo
{
    public Guid Id { get; set; }

    public Guid SiteId { get; set; }

    public Guid ListId { get; set; }

    public Guid? ParentId { get; set; }

    public string DirName { get; set; } = string.Empty;

    public string LeafName { get; set; } = string.Empty;

    public int Type { get; set; }

    public long? SizeBytes { get; set; }

    public DateTime? TimeCreated { get; set; }

    public DateTime? TimeLastModified { get; set; }

    public int? InternalVersion { get; set; }

    public string? UIVersionString { get; set; }

    public string? Extension { get; set; }

    public string? ExtensionForFile { get; set; }

    public string? ProgId { get; set; }

    public bool IsFolder => Type != 0;

    public string DisplayExtension
    {
        get
        {
            if (IsFolder)
                return "";

            if (!string.IsNullOrWhiteSpace(ExtensionForFile))
                return ExtensionForFile!.Trim('.');

            if (!string.IsNullOrWhiteSpace(Extension))
                return Extension!.Trim('.');

            var ext = Path.GetExtension(LeafName);

            return string.IsNullOrWhiteSpace(ext)
                ? ""
                : ext.Trim('.');
        }
    }

    public string FullPath
    {
        get
        {
            if (string.IsNullOrWhiteSpace(DirName))
                return "/" + LeafName.Trim('/');

            if (string.IsNullOrWhiteSpace(LeafName))
                return "/" + DirName.Trim('/');

            return "/" + DirName.Trim('/') + "/" + LeafName.Trim('/');
        }
    }
}

public enum ExplorerChoiceKind
{
    Item,
    Up,
    Exit
}

public sealed class ExplorerChoice
{
    public ExplorerChoiceKind Kind { get; set; }

    public DocumentItemInfo? Item { get; set; }

    public string DisplayText
    {
        get
        {
            if (Kind == ExplorerChoiceKind.Up)
                return ".. Back to parent folder";

            if (Kind == ExplorerChoiceKind.Exit)
                return "Exit library browser";

            if (Item == null)
                return "";

            var icon = Item.IsFolder ? "📁" : "📄";

            var name = string.IsNullOrWhiteSpace(Item.LeafName)
                ? "(root)"
                : Item.LeafName;

            var size = Item.IsFolder ? "" : $" | {FormatBytesStatic(Item.SizeBytes)}";

            return $"{icon} {Markup.Escape(name)}{size}";
        }
    }

    public static ExplorerChoice FromItem(DocumentItemInfo item)
    {
        return new ExplorerChoice
        {
            Kind = ExplorerChoiceKind.Item,
            Item = item
        };
    }

    public static ExplorerChoice Up()
    {
        return new ExplorerChoice
        {
            Kind = ExplorerChoiceKind.Up
        };
    }

    public static ExplorerChoice Exit()
    {
        return new ExplorerChoice
        {
            Kind = ExplorerChoiceKind.Exit
        };
    }

    private static string FormatBytesStatic(long? bytes)
    {
        if (bytes == null)
            return "";

        double value = bytes.Value;
        string[] suffixes = { "B", "KB", "MB", "GB", "TB" };

        var index = 0;

        while (value >= 1024 && index < suffixes.Length - 1)
        {
            value /= 1024;
            index++;
        }

        return $"{value:0.##} {suffixes[index]}";
    }
}
public sealed class DatabaseCandidate
{
    public string Name { get; set; } = string.Empty;

    public int MatchedTableCount { get; set; }

    public bool HasAllSites { get; set; }

    public bool HasWebs { get; set; }

    public bool HasAllLists { get; set; }

    public bool HasAllDocs { get; set; }

    public bool HasAllUserData { get; set; }

    public long? SiteCollectionCount { get; set; }

    public bool IsLikelySharePointContentDatabase =>
    HasAllSites &&
    HasAllLists &&
    HasAllDocs &&
    HasAllUserData;

    public string Confidence =>
        IsLikelySharePointContentDatabase
            ? "[green]Likely SharePoint Content DB[/]"
            : "[yellow]Possible SharePoint Content DB[/]";
}

public sealed class ShreddedExportResult
{
    public bool Success { get; set; }

    public string Message { get; set; } = string.Empty;

    public long BytesWritten { get; set; }

    public string? OutputPath { get; set; }
}

public sealed class ShreddedStreamRow
{
    public int HistVersion { get; set; }

    public int Partition { get; set; }

    public long BSN { get; set; }

    public long StreamId { get; set; }

    public int Type { get; set; }

    public int Size { get; set; }

    public byte[] Content { get; set; } = Array.Empty<byte>();
}
public sealed class SharePointDatabaseSnapshot
{
    public List<SiteCollectionInfo> SiteCollections { get; } = new();

    public List<WebInfo> Webs { get; } = new();
}

public sealed class SiteCollectionInfo
{
    public Guid Id { get; set; }

    public string FullUrl { get; set; } = string.Empty;

    public Guid RootWebId { get; set; }

    public DateTime? TimeCreated { get; set; }
}

public sealed class WebInfo
{
    public Guid Id { get; set; }

    public Guid SiteId { get; set; }

    public string FullUrl { get; set; } = string.Empty;

    public string Title { get; set; } = string.Empty;

    public Guid? ParentWebId { get; set; }
}

public sealed class DocumentLibraryInfo
{
    public Guid Id { get; set; }

    public Guid WebId { get; set; }

    public string Title { get; set; } = string.Empty;

    public int BaseType { get; set; }

    public int ServerTemplate { get; set; }

    public int? ItemCount { get; set; }

    public DateTime? Created { get; set; }

    public DateTime? Modified { get; set; }

    public DateTime? LastItemModifiedDate { get; set; }
}


public static class CobaltCoreInspector
{
    public static void Inspect(string cobaltCoreDllPath)
    {
        if (!File.Exists(cobaltCoreDllPath))
        {
            AnsiConsole.MarkupLine("[red]DLL not found.[/]");
            return;
        }

        var folder = Path.GetDirectoryName(cobaltCoreDllPath)!;

        AppDomain.CurrentDomain.AssemblyResolve += (_, args) =>
        {
            var name = new AssemblyName(args.Name).Name + ".dll";
            var candidate = Path.Combine(folder, name);

            return File.Exists(candidate)
                ? Assembly.LoadFrom(candidate)
                : null;
        };

        var asm = Assembly.LoadFrom(cobaltCoreDllPath);

        AnsiConsole.MarkupLine($"[green]Loaded:[/] {Markup.Escape(asm.FullName ?? cobaltCoreDllPath)}");

        var interestingWords = new[]
        {
            "Cobalt",
            "Cell",
            "Storage",
            "Partition",
            "Object",
            "Blob",
            "Data",
            "File",
            "Protocol",
            "Request",
            "Response",
            "Stream",
            "Shred"
        };

        var types = asm.GetTypes()
            .Where(t => interestingWords.Any(w =>
                t.FullName?.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0))
            .OrderBy(t => t.FullName)
            .ToList();

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("Microsoft.CobaltCore Interesting Types");

        table.AddColumn("Type");
        table.AddColumn("Kind");
        table.AddColumn("Public");
        table.AddColumn("Constructors");
        table.AddColumn("Methods");

        foreach (var type in types.Take(200))
        {
            var constructors = string.Join("\n",
                type.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Take(5)
                    .Select(FormatConstructor));

            var methods = string.Join("\n",
                type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)
                    .Where(m => !m.IsSpecialName)
                    .Take(10)
                    .Select(FormatMethod));

            table.AddRow(
                Markup.Escape(type.FullName ?? type.Name),
                type.IsClass ? "class" : type.IsInterface ? "interface" : type.IsEnum ? "enum" : "other",
                type.IsPublic || type.IsNestedPublic ? "yes" : "no",
                Markup.Escape(constructors),
                Markup.Escape(methods));
        }

        AnsiConsole.Write(table);

        DumpPossibleEntryPoints(types);
    }

    private static void DumpPossibleEntryPoints(List<Type> types)
    {
        var keywords = new[]
        {
            "Deserialize",
            "Serialize",
            "Execute",
            "GetBytes",
            "PutChanges",
            "GetContent",
            "GetStream",
            "Write",
            "Read",
            "Save",
            "Load",
            "Flush",
            "Partition",
            "Cell"
        };

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("Possible Cobalt Entry Points");

        table.AddColumn("Type");
        table.AddColumn("Method");

        foreach (var type in types)
        {
            foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)
                         .Where(m => !m.IsSpecialName)
                         .Where(m => keywords.Any(k => m.Name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                         .Take(20))
            {
                table.AddRow(
                    Markup.Escape(type.FullName ?? type.Name),
                    Markup.Escape(FormatMethod(method)));
            }
        }

        AnsiConsole.Write(table);
    }

    private static string FormatConstructor(ConstructorInfo ctor)
    {
        var args = string.Join(", ",
            ctor.GetParameters()
                .Select(p => $"{PrettyTypeName(p.ParameterType)} {p.Name}"));

        return $"{ctor.Name}({args})";
    }

    private static string FormatMethod(MethodInfo method)
    {
        var args = string.Join(", ",
            method.GetParameters()
                .Select(p => $"{PrettyTypeName(p.ParameterType)} {p.Name}"));

        return $"{PrettyTypeName(method.ReturnType)} {method.Name}({args})";
    }

    private static string PrettyTypeName(Type type)
    {
        if (!type.IsGenericType)
            return type.Name;

        var name = type.Name;

        var tick = name.IndexOf('`');

        if (tick >= 0)
            name = name.Substring(0, tick);

        return name + "<" + string.Join(", ", type.GetGenericArguments().Select(PrettyTypeName)) + ">";
    }
}



public sealed class MicrosoftCobaltCoreAdapter
{

    private object CreateStoreUpdateOptions(
     CobaltCoreBinding binding,
     object? transactionContext)
    {
        if (binding.StoreUpdateOptionsType == null)
            throw new InvalidOperationException("StoreUpdateOptions type was not found.");

        var options = CreateObjectWithoutDefaultConstructor(binding.StoreUpdateOptionsType);

        if (transactionContext != null)
            SetPropertyIfExists(options, "TransactionContext", transactionContext);

        return options;
    }

    private object CreateHostBlobStoreIdAllocator(
    CobaltCoreBinding binding,
    object hostBlobStore,
    object? transactionContext)
    {
        if (binding.HostBlobStoreIdsType == null)
            throw new InvalidOperationException("HostBlobStoreIds type was not found.");

        var createIdAllocator = hostBlobStore.GetType()
            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(m =>
            {
                if (!m.Name.Equals("CreateIdAllocator", StringComparison.OrdinalIgnoreCase))
                    return false;

                var p = m.GetParameters();

                return p.Length == 2 &&
                       p[1].ParameterType == binding.HostBlobStoreIdsType;
            });

        if (createIdAllocator == null)
            throw new InvalidOperationException("HostBlobStore.CreateIdAllocator(TransactionContext, HostBlobStoreIds) was not found.");

        var allocatorStart = CreateHostBlobStoreIds(binding);

        var result = createIdAllocator.Invoke(hostBlobStore, new object?[]
        {
        transactionContext,
        allocatorStart
        });

        if (result == null)
            throw new InvalidOperationException("CreateIdAllocator returned null.");

        return result;
    }

    private object CreateHostBlobStoreUpdate(
    CobaltCoreBinding binding,
    object hostBlobStore,
    object storeUpdateOptions,
    object idAllocator)
    {
        if (binding.StoreUpdateOptionsType == null)
            throw new InvalidOperationException("StoreUpdateOptions type was not found.");

        var createUpdate = hostBlobStore.GetType()
            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(m =>
            {
                if (!m.Name.Equals("CreateUpdate", StringComparison.OrdinalIgnoreCase))
                    return false;

                var p = m.GetParameters();

                return p.Length == 2 &&
                       p[0].ParameterType == binding.StoreUpdateOptionsType;
            });

        if (createUpdate == null)
            throw new InvalidOperationException("HostBlobStore.CreateUpdate(StoreUpdateOptions, IHostBlobStoreIdAllocator) was not found.");

        var result = createUpdate.Invoke(hostBlobStore, new object[]
        {
        storeUpdateOptions,
        idAllocator
        });

        if (result == null)
            throw new InvalidOperationException("CreateUpdate returned null.");

        return result;
    }

    private void CommitHostBlobStoreUpdate(object update)
    {
        var commit = update.GetType()
            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(m =>
                m.Name.Equals("Commit", StringComparison.OrdinalIgnoreCase) &&
                m.GetParameters().Length == 0);

        if (commit == null)
            throw new InvalidOperationException("HostBlobStore.Update.Commit() was not found.");

        AnsiConsole.MarkupLine("[grey]Stage:[/] [yellow]CommitHostBlobStoreUpdate[/]");

        commit.Invoke(update, Array.Empty<object>());
    }

    private static object? CreatePreferredEnumValue(Type? enumType, string[] preferredNames)
    {
        if (enumType == null)
            return null;

        if (!enumType.IsEnum)
            return CreateDefaultValue(enumType);

        var values = Enum.GetValues(enumType).Cast<object>().ToList();

        if (values.Count == 0)
            return Activator.CreateInstance(enumType);

        foreach (var preferredName in preferredNames)
        {
            var match = values.FirstOrDefault(v =>
                v.ToString()!.IndexOf(preferredName, StringComparison.OrdinalIgnoreCase) >= 0);

            if (match != null)
                return match;
        }

        return values[0];
    }
    private static void DumpEnumValues(Type? type, string title)
    {
        if (type == null || !type.IsEnum)
            return;

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title(title);

        table.AddColumn("Name");
        table.AddColumn("Value");

        foreach (var value in Enum.GetValues(type))
        {
            table.AddRow(
                Markup.Escape(value.ToString() ?? ""),
                Convert.ToInt64(value).ToString());
        }

        AnsiConsole.Write(table);
    }

    private object CreateHostBlobStoreIds(CobaltCoreBinding binding)
    {
        if (binding.HostBlobStoreIdsType == null)
            throw new InvalidOperationException("HostBlobStoreIds type was not found.");

        var ids = CreateObjectWithoutDefaultConstructor(binding.HostBlobStoreIdsType);

        SetPropertyIfExists(ids, "StoreBsn", (ulong)1);
        SetPropertyIfExists(ids, "MultiSessionStoreBsn", (ulong)1);

        return ids;
    }

    private static T Stage<T>(string name, Func<T> action)
    {
        try
        {
            AnsiConsole.MarkupLine($"[grey]Stage:[/] [yellow]{Markup.Escape(name)}[/]");
            return action();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Stage failed: " + name + " -> " + GetUsefulExceptionMessage(ex),
                ex);
        }
    }

    private static void Stage(string name, Action action)
    {
        try
        {
            AnsiConsole.MarkupLine($"[grey]Stage:[/] [yellow]{Markup.Escape(name)}[/]");
            action();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                "Stage failed: " + name + " -> " + GetUsefulExceptionMessage(ex),
                ex);
        }
    }

    private ShreddedExportResult? AnalyzeFsshttpbRows(
    List<ShreddedStreamRow> rows,
    DocumentItemInfo item,
    string outputPath,
    string workFolder)
    {
        var scanCsvPath = Path.Combine(workFolder, "fsshttpb_row_headers_scan.csv");
        var sequentialCsvPath = Path.Combine(workFolder, "fsshttpb_sequential_parse.csv");
        var objectGraphCsvPath = Path.Combine(workFolder, "fsshttpb_object_graph.csv");
        var hexCsvPath = Path.Combine(workFolder, "fsshttpb_row_hex_summary.csv");
        var fragmentMapCsvPath = Path.Combine(workFolder, "cobalt_fragment_map.csv");
        var envelopeCsvPath = Path.Combine(workFolder, "cobalt_row_envelope_fields.csv");
        var type11ReferenceMapCsvPath = Path.Combine(workFolder, "cobalt_type11_reference_map.csv");
        var type11RecordScanCsvPath = Path.Combine(workFolder, "cobalt_type11_record_scan.csv");
        var type11SequenceCsvPath = Path.Combine(workFolder, "cobalt_type11_sequence_scan.csv");
        var type11PairedReferenceCsvPath = Path.Combine(workFolder, "cobalt_type11_paired_references.csv");
        var fileHostBlobMapProbeCsvPath = Path.Combine(workFolder, "cobalt_file_host_blob_map_probe.csv");
        var fileHostBlobMapDetailsCsvPath = Path.Combine(workFolder, "cobalt_file_host_blob_map_details.csv");
        var cobaltHeaderFooterProbeCsvPath = Path.Combine(workFolder, "cobalt_header_footer_probe.csv");
        var cobaltItemGroupPayloadCsvPath = Path.Combine(workFolder, "cobalt_item_group_payloads.csv");
        var cobaltType10DescriptorCsvPath = Path.Combine(workFolder, "cobalt_type10_descriptor_map.csv");
        var cobaltItemGroupCandidateFolder = Path.Combine(workFolder, "cobalt_item_group_candidates");
        var cobaltItemGroupCandidateReportPath = Path.Combine(cobaltItemGroupCandidateFolder, "candidate_report.csv");
        var cobaltType11StreamOrdersCsvPath = Path.Combine(workFolder, "cobalt_type11_stream_orders.csv");

        WriteFsshttpbHexSummary(rows, hexCsvPath);
        WriteFsshttpbSequentialParse(rows, sequentialCsvPath);
        WriteFsshttpbObjectGraph(rows, objectGraphCsvPath);
        WriteCobaltRowEnvelopeFields(rows, envelopeCsvPath);
        WriteCobaltFragmentMap(rows, fragmentMapCsvPath);
        WriteCobaltType11ReferenceMap(rows, type11ReferenceMapCsvPath);
        WriteCobaltType11RecordScan(rows, type11RecordScanCsvPath);
        WriteCobaltType11SequenceScan(rows, type11SequenceCsvPath);
        WriteCobaltType11PairedReferences(rows, type11PairedReferenceCsvPath);
        WriteCobaltFileHostBlobMapProbe(rows, fileHostBlobMapProbeCsvPath, fileHostBlobMapDetailsCsvPath, workFolder);
        WriteCobaltHeaderFooterProbe(rows, cobaltHeaderFooterProbeCsvPath);
        WriteCobaltItemGroupPayloadSummary(rows, cobaltHeaderFooterProbeCsvPath, cobaltItemGroupPayloadCsvPath);
        WriteCobaltType10DescriptorMap(rows, cobaltHeaderFooterProbeCsvPath, cobaltType10DescriptorCsvPath);
        var strictItemGroupCandidate = WriteCobaltItemGroupPayloadCandidates(
            rows,
            item,
            cobaltHeaderFooterProbeCsvPath,
            cobaltItemGroupCandidateFolder,
            cobaltItemGroupCandidateReportPath,
            cobaltType11StreamOrdersCsvPath,
            item.SizeBytes);

        AnsiConsole.MarkupLine(
            $"[grey]FSSHTTPB hex summary written to:[/] [yellow]{Markup.Escape(hexCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]FSSHTTPB sequential parse written to:[/] [yellow]{Markup.Escape(sequentialCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]FSSHTTPB object graph written to:[/] [yellow]{Markup.Escape(objectGraphCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt fragment map written to:[/] [yellow]{Markup.Escape(fragmentMapCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt row envelope fields written to:[/] [yellow]{Markup.Escape(envelopeCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 11 reference map written to:[/] [yellow]{Markup.Escape(type11ReferenceMapCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 11 record scan written to:[/] [yellow]{Markup.Escape(type11RecordScanCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 11 sequence scan written to:[/] [yellow]{Markup.Escape(type11SequenceCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 11 paired references written to:[/] [yellow]{Markup.Escape(type11PairedReferenceCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt FileHostBlobMap probe written to:[/] [yellow]{Markup.Escape(fileHostBlobMapProbeCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt FileHostBlobMap details written to:[/] [yellow]{Markup.Escape(fileHostBlobMapDetailsCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt header/footer probe written to:[/] [yellow]{Markup.Escape(cobaltHeaderFooterProbeCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt item group payload summary written to:[/] [yellow]{Markup.Escape(cobaltItemGroupPayloadCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 10 descriptor map written to:[/] [yellow]{Markup.Escape(cobaltType10DescriptorCsvPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt item group candidates written to:[/] [yellow]{Markup.Escape(cobaltItemGroupCandidateFolder)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt item group candidate report written to:[/] [yellow]{Markup.Escape(cobaltItemGroupCandidateReportPath)}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Cobalt Type 11 stream orders written to:[/] [yellow]{Markup.Escape(cobaltType11StreamOrdersCsvPath)}[/]");

        if (strictItemGroupCandidate == null)
            return null;

        File.Copy(strictItemGroupCandidate.Path, outputPath, overwrite: true);

        return new ShreddedExportResult
        {
            Success = true,
            OutputPath = outputPath,
            BytesWritten = strictItemGroupCandidate.Length,
            Message =
                "Exported by strict format-neutral Cobalt item-group payload reconstruction. " +
                $"Candidate strategy: {strictItemGroupCandidate.Name}."
        };
    }

    private static CobaltItemGroupCandidate? WriteCobaltItemGroupPayloadCandidates(
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string headerFooterProbeCsvPath,
        string candidateFolder,
        string reportPath,
        string type11StreamOrdersCsvPath,
        long? expectedSize)
    {
        Directory.CreateDirectory(candidateFolder);

        var payloads = ReadCobaltItemGroupPayloads(rows, headerFooterProbeCsvPath)
            .Where(x => x.Type == 10 && x.Payload.Length > 100)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step)
            .ToList();

        var candidates = new List<CobaltItemGroupCandidate>();

        if (payloads.Count > 0)
        {
            AddCobaltItemGroupCandidateOrderedByPayloadOrdinal(
                candidates,
                candidateFolder,
                payloads,
                "type10_payload_ordinal_strip79_tail26.bin",
                leadingStrip: 79,
                trailingStrip: 26);

            foreach (var signature in CobaltItemGroupKnownSignatures)
            {
                AddCobaltItemGroupCandidateOrderedByPayloadOrdinalFromSignature(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_payload_ordinal_from_{MakeSafeCandidateToken(signature.Name)}_signature_tail26.bin",
                    signature.Bytes,
                    trailingStrip: 26);

                if (signature.Name == "PDF" && expectedSize.HasValue)
                {
                    AddPdfOrdinalVariableStartCandidate(
                        candidates,
                        candidateFolder,
                        payloads,
                        "type10_payload_ordinal_pdf_variable_start_tail26.bin",
                        signature.Bytes,
                        expectedSize.Value,
                        trailingStrip: 26);

                    AddPdfOrdinalDynamicBoundaryCandidate(
                        candidates,
                        candidateFolder,
                        payloads,
                        "type10_payload_ordinal_pdf_dynamic_boundary_exact.bin",
                        signature.Bytes,
                        expectedSize.Value,
                        trailingStrip: 26);

                    AddPdfDescriptorOrderDynamicBoundaryCandidate(
                        candidates,
                        candidateFolder,
                        payloads,
                        "type10_descriptor_order_pdf_dynamic_boundary_exact.bin",
                        signature.Bytes,
                        expectedSize.Value,
                        trailingStrip: 26);
                }
            }

            foreach (var strip in new[] { 0, 8, 16, 23, 32, 64, 79, 96, 104, 105, 106, 107, 108, 128 })
            {
                AddCobaltItemGroupCandidate(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_strip_{strip}.bin",
                    strip);
            }

            foreach (var signature in CobaltItemGroupKnownSignatures)
            {
                AddCobaltItemGroupCandidateFromSignature(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}.bin",
                    signature.Bytes);

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following79.bin",
                    signature.Bytes,
                    followingStrip: 79,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following79_trim_tail26_each.bin",
                    signature.Bytes,
                    followingStrip: 79,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "",
                    tailTrimEachPayload: 26);

                AddCobaltItemGroupCandidateFromSignatureWithDescriptorOrdering(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_descriptor_order_from_{MakeSafeCandidateToken(signature.Name)}_strip_following79_trim_tail26_each.bin",
                    signature.Bytes,
                    followingStrip: 79,
                    tailTrimEachPayload: 26);

                if (signature.Name == "PDF")
                {
                    AddPdfTerminalXrefLastCobaltItemGroupCandidate(
                        candidates,
                        candidateFolder,
                        payloads,
                        "type10_from_pdf_strip_following79_trim_tail26_each_terminal_xref_last.bin",
                        signature.Bytes,
                        followingStrip: 79,
                        tailTrimEachPayload: 26);
                }

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following96.bin",
                    signature.Bytes,
                    followingStrip: 96,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following105.bin",
                    signature.Bytes,
                    followingStrip: 105,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following105_tail26.bin",
                    signature.Bytes,
                    followingStrip: 105,
                    extraBytesToRemove: 26,
                    extraRemovalMode: "tail");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following105_first26_extra1.bin",
                    signature.Bytes,
                    followingStrip: 105,
                    extraBytesToRemove: 26,
                    extraRemovalMode: "first-following");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    payloads,
                    $"type10_item_groups_from_{MakeSafeCandidateToken(signature.Name)}_strip_following105_last26_extra1.bin",
                    signature.Bytes,
                    followingStrip: 105,
                    extraBytesToRemove: 26,
                    extraRemovalMode: "last-following");
            }

            AddType11GuidedCobaltItemGroupCandidates(
                candidates,
                rows,
                candidateFolder,
                payloads,
                type11StreamOrdersCsvPath);

            if (string.Equals(item.DisplayExtension, "pdf", StringComparison.OrdinalIgnoreCase))
            {
                TryRunPdfDiagnosticCandidateWriters(candidates, candidateFolder);
            }
        }

        using (var writer = new StreamWriter(reportPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Name,Length,ExpectedSize,DeltaFromExpected,DetectedSignature,ValidForSelectedFileType,PayloadCount,Path");

            foreach (var candidate in candidates
                         .OrderBy(x => Math.Abs((expectedSize ?? x.Length) - x.Length))
                         .ThenBy(x => x.Name))
            {
                writer.WriteLine(
                    string.Join(",",
                        Csv(candidate.Name),
                        candidate.Length,
                        expectedSize?.ToString() ?? "",
                        expectedSize.HasValue ? (candidate.Length - expectedSize.Value).ToString() : "",
                        Csv(DetectKnownSignatureInFile(candidate.Path)),
                        IsCandidateValidForSelectedFileType(candidate.Path, item) ? "1" : "0",
                        candidate.PayloadCount,
                        Csv(candidate.Path)));
            }
        }

        if (!expectedSize.HasValue)
            return null;

        var exactCandidate = candidates
            .Where(x => x.Length == expectedSize.Value)
            .Where(x => IsCandidateValidForSelectedFileType(x.Path, item))
            .OrderBy(x => GetCandidateStructuralScore(x.Path, item))
            .ThenByDescending(x => !string.IsNullOrWhiteSpace(DetectKnownSignatureInFile(x.Path)))
            .ThenBy(x => x.Name)
            .FirstOrDefault();

        if (exactCandidate != null)
            return exactCandidate;

        if (string.Equals(item.DisplayExtension, "pdf", StringComparison.OrdinalIgnoreCase))
        {
            return candidates
                .Where(x => Math.Abs(x.Length - expectedSize.Value) <= 64)
                .Where(x => IsCandidateValidForSelectedFileType(x.Path, item))
                .Where(x => !IsDiagnosticCandidateName(x.Name))
                .OrderBy(x => GetCandidateStructuralScore(x.Path, item))
                .ThenBy(x => Math.Abs(x.Length - expectedSize.Value))
                .ThenByDescending(x => !string.IsNullOrWhiteSpace(DetectKnownSignatureInFile(x.Path)))
                .ThenBy(x => x.Name)
                .FirstOrDefault();
        }

        return null;
    }

    private static void TryRunPdfDiagnosticCandidateWriters(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder)
    {
        var errors = new List<string>();

        TryRunPdfDiagnosticStep(
            "xref_repaired",
            () => AddPdfXrefRepairedDiagnosticCandidates(candidates, candidateFolder),
            errors);

        TryRunPdfDiagnosticStep(
            "object_repacked",
            () => AddPdfObjectRepackedDiagnosticCandidates(candidates, candidateFolder),
            errors);

        TryRunPdfDiagnosticStep(
            "declared_length_repacked",
            () => AddPdfDeclaredLengthRepackedDiagnosticCandidates(candidates, candidateFolder),
            errors);

        TryRunPdfDiagnosticStep(
            "image_flate_report",
            () => WritePdfImageFlateDiagnosticReport(candidates, candidateFolder),
            errors);

        if (errors.Count == 0)
            return;

        var path = Path.Combine(candidateFolder, "pdf_diagnostic_errors.csv");

        using var writer = new StreamWriter(path, false, Encoding.UTF8);

        writer.WriteLine("Step,Error");

        foreach (var error in errors)
            writer.WriteLine(error);
    }

    private static void TryRunPdfDiagnosticStep(
        string step,
        Action action,
        List<string> errors)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            errors.Add(
                string.Join(",",
                    Csv(step),
                    Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static List<CobaltItemGroupPayload> ReadCobaltItemGroupPayloads(
        List<ShreddedStreamRow> rows,
        string headerFooterProbeCsvPath)
    {
        if (!File.Exists(headerFooterProbeCsvPath))
            return new List<CobaltItemGroupPayload>();

        var rowLookup = rows.ToDictionary(
            x => (x.Type, x.StreamId, x.BSN),
            x => x);

        var result = new List<CobaltItemGroupPayload>();

        foreach (var probe in File.ReadLines(headerFooterProbeCsvPath)
                     .Skip(1)
                     .Select(ParseHeaderFooterProbeCsvLine)
                     .Where(x => x != null)
                     .Select(x => x!)
                     .Where(x =>
                         x.HeaderFooterType == "ItemGroupHeader" &&
                         x.Status == "1" &&
                         x.Variant == "full"))
        {
            if (!rowLookup.TryGetValue((probe.Type, probe.StreamId, probe.BSN), out var row))
                continue;

            if (probe.Length <= 0 ||
                probe.PositionAfter < 0 ||
                probe.PositionAfter + probe.Length > row.Content.Length)
            {
                continue;
            }

            var payload = new byte[probe.Length];
            Buffer.BlockCopy(row.Content, probe.PositionAfter, payload, 0, payload.Length);

            result.Add(new CobaltItemGroupPayload(
                probe.Type,
                probe.StreamId,
                probe.BSN,
                probe.Step,
                probe.Id,
                probe.PositionAfter,
                payload));
        }

        return result;
    }

    private static void AddCobaltItemGroupCandidate(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        int stripEachPayload)
    {
        var path = Path.Combine(candidateFolder, fileName);

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var payload in payloads)
            {
                var start = Math.Min(stripEachPayload, payload.Payload.Length);
                output.Write(payload.Payload, start, payload.Payload.Length - start);
            }
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            payloads.Count));
    }

    private static void AddCobaltItemGroupCandidateOrderedByPayloadOrdinal(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        int leadingStrip,
        int trailingStrip)
    {
        var orderedPayloads = payloads
            .Where(x => x.Payload.Length > leadingStrip + trailingStrip + 2)
            .Select(x => new
            {
                Payload = x,
                Ordinal = BitConverter.ToUInt16(x.Payload, 0)
            })
            .OrderBy(x => x.Ordinal)
            .ThenBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .Select(x => x.Payload)
            .ToList();

        if (orderedPayloads.Count == 0)
            return;

        var path = Path.Combine(candidateFolder, fileName);

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var payload in orderedPayloads)
            {
                var start = Math.Min(leadingStrip, payload.Payload.Length);
                var length = Math.Max(0, payload.Payload.Length - start - trailingStrip);

                output.Write(payload.Payload, start, length);
            }
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            orderedPayloads.Count));
    }

    private static void AddCobaltItemGroupCandidateOrderedByPayloadOrdinalFromSignature(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        int trailingStrip)
    {
        var orderedPayloads = payloads
            .Where(x => x.Payload.Length > signature.Length + trailingStrip + 2)
            .Select(x => new
            {
                Payload = x,
                Ordinal = BitConverter.ToUInt16(x.Payload, 0),
                SignatureOffset = IndexOfBytes(x.Payload, signature)
            })
            .OrderBy(x => x.Ordinal)
            .ThenBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .ToList();

        var match = orderedPayloads.FirstOrDefault(x => x.SignatureOffset >= 0);

        if (match == null)
            return;

        var rotatedPayloads = orderedPayloads
            .Where(x => x.Ordinal >= match.Ordinal)
            .Concat(orderedPayloads.Where(x => x.Ordinal < match.Ordinal))
            .ToList();

        var path = Path.Combine(candidateFolder, fileName);
        var first = true;

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var item in rotatedPayloads)
            {
                var start = first
                    ? match.SignatureOffset
                    : match.SignatureOffset;

                first = false;

                start = Math.Min(start, item.Payload.Payload.Length);

                var length = Math.Max(0, item.Payload.Payload.Length - start - trailingStrip);

                output.Write(item.Payload.Payload, start, length);
            }
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            rotatedPayloads.Count));
    }

    private static void AddPdfOrdinalVariableStartCandidate(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        long expectedSize,
        int trailingStrip)
    {
        var orderedPayloads = payloads
            .Where(x => x.Payload.Length > signature.Length + trailingStrip + 2)
            .Select(x => new
            {
                Payload = x,
                Ordinal = BitConverter.ToUInt16(x.Payload, 0),
                SignatureOffset = IndexOfBytes(x.Payload, signature)
            })
            .OrderBy(x => x.Ordinal)
            .ThenBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .ToList();

        var match = orderedPayloads.FirstOrDefault(x => x.SignatureOffset >= 0);

        if (match == null)
            return;

        var rotatedPayloads = orderedPayloads
            .Where(x => x.Ordinal >= match.Ordinal)
            .Concat(orderedPayloads.Where(x => x.Ordinal < match.Ordinal))
            .Select(x => x.Payload)
            .ToList();

        // Exhaustive nearby-boundary search is intentionally limited. Large files usually
        // use a stable start offset; small files can vary by one or two bytes per fragment.
        if (rotatedPayloads.Count == 0 || rotatedPayloads.Count > 12)
            return;

        var choices = rotatedPayloads
            .Select(_ => new[]
                {
                    match.SignatureOffset - 1,
                    match.SignatureOffset,
                    match.SignatureOffset + 1
                }
                .Where(x => x >= 0)
                .Distinct()
                .ToArray())
            .ToList();

        var offsets = new int[rotatedPayloads.Count];
        byte[]? bestBytes = null;

        SearchPdfOrdinalVariableStarts(
            rotatedPayloads,
            choices,
            trailingStrip,
            expectedSize,
            offsets,
            index: 0,
            ref bestBytes);

        if (bestBytes == null)
            return;

        var path = Path.Combine(candidateFolder, fileName);
        File.WriteAllBytes(path, bestBytes);

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            rotatedPayloads.Count));
    }

    private static void SearchPdfOrdinalVariableStarts(
        List<CobaltItemGroupPayload> payloads,
        List<int[]> choices,
        int trailingStrip,
        long expectedSize,
        int[] offsets,
        int index,
        ref byte[]? bestBytes)
    {
        if (bestBytes != null)
            return;

        if (index == payloads.Count)
        {
            using var output = new MemoryStream();

            for (var i = 0; i < payloads.Count; i++)
            {
                var payload = payloads[i].Payload;
                var start = Math.Min(offsets[i], payload.Length);
                var length = Math.Max(0, payload.Length - start - trailingStrip);

                output.Write(payload, start, length);
            }

            if (output.Length != expectedSize)
                return;

            var bytes = output.ToArray();

            if (LooksLikePdfCandidateBytes(bytes))
                bestBytes = bytes;

            return;
        }

        foreach (var choice in choices[index])
        {
            offsets[index] = choice;
            SearchPdfOrdinalVariableStarts(
                payloads,
                choices,
                trailingStrip,
                expectedSize,
                offsets,
                index + 1,
                ref bestBytes);

            if (bestBytes != null)
                return;
        }
    }

    private static void AddCobaltItemGroupCandidateFromSignatureWithDescriptorOrdering(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        int followingStrip,
        int tailTrimEachPayload)
    {
        var orderedByDescriptor = payloads
            .Where(x => x.Payload.Length > 100)
            .OrderBy(x => GetType10PayloadLogicalOffset(x))
            .ThenBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step)
            .ToList();

        if (orderedByDescriptor.Count == 0)
            return;

        AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
            candidates,
            candidateFolder,
            orderedByDescriptor,
            fileName,
            signature,
            followingStrip,
            extraBytesToRemove: 0,
            extraRemovalMode: "",
            tailTrimEachPayload: tailTrimEachPayload);
    }

    private static void AddPdfTerminalXrefLastCobaltItemGroupCandidate(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        int followingStrip,
        int tailTrimEachPayload)
    {
        var match = payloads
            .Select(x => new
            {
                Payload = x,
                Offset = IndexOfBytes(x.Payload, signature)
            })
            .Where(x => x.Offset >= 0)
            .OrderBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .FirstOrDefault();

        if (match == null)
            return;

        var orderedPayloads = payloads
            .Where(x =>
                x.StreamId > match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN > match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step >= match.Payload.Step))
            .Concat(payloads.Where(x =>
                x.StreamId < match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN < match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step < match.Payload.Step)))
            .ToList();

        var terminalPayload = orderedPayloads
            .Where(x => ContainsAscii(x.Payload, "xref") && ContainsAscii(x.Payload, "%%EOF"))
            .OrderByDescending(x => x.StreamId)
            .ThenByDescending(x => x.BSN)
            .ThenByDescending(x => x.Step)
            .FirstOrDefault();

        if (terminalPayload == null)
            return;

        orderedPayloads.Remove(terminalPayload);
        orderedPayloads.Add(terminalPayload);

        var path = Path.Combine(candidateFolder, fileName);
        var first = true;

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var payload in orderedPayloads)
            {
                var start = first && payload == match.Payload
                    ? match.Offset
                    : Math.Min(followingStrip, payload.Payload.Length);

                first = false;

                var length = payload.Payload.Length - start;

                if (tailTrimEachPayload > 0)
                    length = Math.Max(0, length - tailTrimEachPayload);

                output.Write(payload.Payload, start, length);
            }
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            orderedPayloads.Count));
    }

    private static bool ContainsAscii(byte[] data, string value)
    {
        return IndexOfBytes(data, Encoding.ASCII.GetBytes(value)) >= 0;
    }

    private static void AddCobaltItemGroupCandidateFromSignature(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature)
    {
        var match = payloads
            .Select(x => new
            {
                Payload = x,
                Offset = IndexOfBytes(x.Payload, signature)
            })
            .Where(x => x.Offset >= 0)
            .OrderBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .FirstOrDefault();

        if (match == null)
            return;

        var orderedPayloads = payloads
            .Where(x =>
                x.StreamId > match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN > match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step >= match.Payload.Step))
            .Concat(payloads.Where(x =>
                x.StreamId < match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN < match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step < match.Payload.Step)))
            .ToList();

        var path = Path.Combine(candidateFolder, fileName);
        var first = true;

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            foreach (var payload in orderedPayloads)
            {
                var start = first ? match.Offset : 0;
                first = false;

                output.Write(payload.Payload, start, payload.Payload.Length - start);
            }
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            payloads.Count));
    }

    private static void AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        int followingStrip,
        int extraBytesToRemove,
        string extraRemovalMode,
        int tailTrimEachPayload = 0)
    {
        var match = payloads
            .Select(x => new
            {
                Payload = x,
                Offset = IndexOfBytes(x.Payload, signature)
            })
            .Where(x => x.Offset >= 0)
            .OrderBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .FirstOrDefault();

        if (match == null)
            return;

        var orderedPayloads = payloads
            .Where(x =>
                x.StreamId > match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN > match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step >= match.Payload.Step))
            .Concat(payloads.Where(x =>
                x.StreamId < match.Payload.StreamId ||
                (x.StreamId == match.Payload.StreamId && x.BSN < match.Payload.BSN) ||
                (x.StreamId == match.Payload.StreamId && x.BSN == match.Payload.BSN && x.Step < match.Payload.Step)))
            .ToList();

        var path = Path.Combine(candidateFolder, fileName);
        var first = true;

        using (var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            for (var i = 0; i < orderedPayloads.Count; i++)
            {
                var payload = orderedPayloads[i];
                var start = first
                    ? match.Offset
                    : Math.Min(followingStrip, payload.Payload.Length);

                first = false;

                if (extraRemovalMode == "first-following" && i > 0 && i <= extraBytesToRemove)
                    start = Math.Min(start + 1, payload.Payload.Length);

                if (extraRemovalMode == "last-following" && i > 0 && i > orderedPayloads.Count - extraBytesToRemove)
                    start = Math.Min(start + 1, payload.Payload.Length);

                var length = payload.Payload.Length - start;

                if (tailTrimEachPayload > 0)
                    length = Math.Max(0, length - tailTrimEachPayload);

                output.Write(payload.Payload, start, length);
            }

            if (extraRemovalMode == "tail" && extraBytesToRemove > 0 && output.Length > extraBytesToRemove)
                output.SetLength(output.Length - extraBytesToRemove);
        }

        candidates.Add(new CobaltItemGroupCandidate(
            fileName,
            path,
            new FileInfo(path).Length,
            payloads.Count));
    }

    private static void AddType11GuidedCobaltItemGroupCandidates(
        List<CobaltItemGroupCandidate> candidates,
        List<ShreddedStreamRow> rows,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string type11StreamOrdersCsvPath)
    {
        var dataPayloads = payloads
            .Where(x => x.Type == 10 && x.Payload.Length > 100)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step)
            .ToList();

        if (dataPayloads.Count == 0)
            return;

        var orders = InferType11StreamOrders(rows, dataPayloads)
            .OrderByDescending(x => x.StreamIds.Distinct().Count())
            .ThenByDescending(x => x.StreamIds.Count)
            .ThenBy(x => x.SourceStreamId)
            .ThenBy(x => x.PayloadOffset)
            .Take(12)
            .ToList();

        using (var writer = new StreamWriter(type11StreamOrdersCsvPath, false, Encoding.UTF8))
        {
            writer.WriteLine("Name,SourceType,SourceStreamId,SourceBSN,SourceStep,PayloadOffset,Count,UniqueCount,StreamIds");

            foreach (var order in orders)
            {
                writer.WriteLine(
                    string.Join(",",
                        Csv(order.Name),
                        order.SourceType,
                        order.SourceStreamId,
                        order.SourceBSN,
                        order.SourceStep,
                        order.PayloadOffset,
                        order.StreamIds.Count,
                        order.StreamIds.Distinct().Count(),
                        Csv(string.Join("|", order.StreamIds))));
            }
        }

        foreach (var order in orders.Take(8))
        {
            var orderedPayloads = OrderDataPayloadsByType11StreamOrder(dataPayloads, order.StreamIds);

            foreach (var signature in CobaltItemGroupKnownSignatures)
            {
                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    orderedPayloads,
                    $"type11_{order.Name}_from_{MakeSafeCandidateToken(signature.Name)}_strip_following79.bin",
                    signature.Bytes,
                    followingStrip: 79,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    orderedPayloads,
                    $"type11_{order.Name}_from_{MakeSafeCandidateToken(signature.Name)}_strip_following79_trim_tail26_each.bin",
                    signature.Bytes,
                    followingStrip: 79,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "",
                    tailTrimEachPayload: 26);

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    orderedPayloads,
                    $"type11_{order.Name}_from_{MakeSafeCandidateToken(signature.Name)}_strip_following96.bin",
                    signature.Bytes,
                    followingStrip: 96,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");

                AddCobaltItemGroupCandidateFromSignatureWithFollowingStrips(
                    candidates,
                    candidateFolder,
                    orderedPayloads,
                    $"type11_{order.Name}_from_{MakeSafeCandidateToken(signature.Name)}_strip_following105.bin",
                    signature.Bytes,
                    followingStrip: 105,
                    extraBytesToRemove: 0,
                    extraRemovalMode: "");
            }
        }
    }

    private static void AddPdfOrdinalDynamicBoundaryCandidate(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        long expectedSize,
        int trailingStrip)
    {
        var orderedPayloads = payloads
            .Where(x => x.Payload.Length > signature.Length + trailingStrip + 2)
            .Select(x => new CobaltOrdinalPayload(
                x,
                BitConverter.ToUInt16(x.Payload, 0),
                IndexOfBytes(x.Payload, signature)))
            .OrderBy(x => x.Ordinal)
            .ThenBy(x => x.Payload.StreamId)
            .ThenBy(x => x.Payload.BSN)
            .ThenBy(x => x.Payload.Step)
            .ToList();

        var match = orderedPayloads.FirstOrDefault(x => x.SignatureOffset >= 0);

        if (match == null)
            return;

        var rotatedPayloads = orderedPayloads
            .Where(x => x.Ordinal >= match.Ordinal)
            .Concat(orderedPayloads.Where(x => x.Ordinal < match.Ordinal))
            .Select(x => x.Payload)
            .ToList();

        if (rotatedPayloads.Count == 0)
            return;

        var baseStart = match.SignatureOffset;
        var choiceSets = rotatedPayloads
            .Select(x => BuildPdfBoundaryChoices(x.Payload, baseStart, trailingStrip))
            .ToList();

        if (choiceSets.Any(x => x.Count == 0))
            return;

        long baseLength = 0;

        foreach (var payload in rotatedPayloads)
        {
            var length = Math.Max(0, payload.Payload.Length - baseStart - trailingStrip);
            baseLength += length;
        }

        var targetReduction = baseLength - expectedSize;

        if (targetReduction < -1024 || targetReduction > 1024)
            return;

        var states = new Dictionary<long, List<PdfBoundarySearchState>>
        {
            [0] = new List<PdfBoundarySearchState>
            {
                new PdfBoundarySearchState(null, null, 0, 0)
            }
        };

        const int maxStatesPerReduction = 48;
        const int reductionWindow = 1024;

        foreach (var choices in choiceSets)
        {
            var next = new Dictionary<long, List<PdfBoundarySearchState>>();

            foreach (var stateBucket in states)
            {
                foreach (var state in stateBucket.Value)
                {
                    foreach (var choice in choices)
                    {
                        var reduction = state.Reduction + choice.Reduction;

                        if (Math.Abs(reduction - targetReduction) > reductionWindow)
                            continue;

                        var nextState = new PdfBoundarySearchState(
                            state,
                            choice,
                            reduction,
                            state.Score + choice.Score);

                        if (!next.TryGetValue(reduction, out var bucket))
                        {
                            bucket = new List<PdfBoundarySearchState>();
                            next[reduction] = bucket;
                        }

                        bucket.Add(nextState);
                    }
                }
            }

            foreach (var key in next.Keys.ToList())
            {
                next[key] = next[key]
                    .OrderBy(x => x.Score)
                    .ThenBy(x => x.NonBaseChoiceCount)
                    .Take(maxStatesPerReduction)
                    .ToList();
            }

            states = next;

            if (states.Count == 0)
                return;
        }

        if (!states.TryGetValue(targetReduction, out var finalStates))
            return;

        byte[]? bestBytes = null;
        long bestStructuralScore = long.MaxValue;
        PdfBoundaryChoice[]? bestChoices = null;
        string bestValidationError = "";
        byte[]? bestRejectedBytes = null;

        foreach (var state in finalStates
                     .OrderBy(x => x.Score)
                     .ThenBy(x => x.NonBaseChoiceCount)
                     .Take(128))
        {
            var choices = UnwindPdfBoundaryChoices(state, rotatedPayloads.Count);
            var bytes = BuildCobaltPayloadBytes(rotatedPayloads, choices);

            if (bytes.Length != expectedSize)
                continue;

            if (!TryValidatePdfCandidateBytes(bytes, out var validationError))
            {
                if (TryRepairPdfStartXref(bytes, out var repairedBytes) &&
                    repairedBytes.Length == expectedSize &&
                    TryValidatePdfCandidateBytes(repairedBytes, out _))
                {
                    var structuralScore = GetPdfXrefConsistencyScore(repairedBytes);

                    if (structuralScore < bestStructuralScore)
                    {
                        bestBytes = repairedBytes;
                        bestStructuralScore = structuralScore;
                        bestChoices = choices;
                    }
                }
                else if (bestRejectedBytes == null)
                {
                    bestRejectedBytes = bytes;
                    bestChoices = choices;
                    bestValidationError = validationError;
                }

                continue;
            }

            var score = GetPdfXrefConsistencyScore(bytes);

            if (score < bestStructuralScore)
            {
                bestBytes = bytes;
                bestStructuralScore = score;
                bestChoices = choices;
            }
        }

        if (bestBytes != null)
        {
            var path = Path.Combine(candidateFolder, fileName);
            File.WriteAllBytes(path, bestBytes);

            candidates.Add(new CobaltItemGroupCandidate(
                fileName,
                path,
                new FileInfo(path).Length,
                rotatedPayloads.Count));
            return;
        }

        if (bestRejectedBytes != null && bestChoices != null)
        {
            WritePdfDynamicBoundaryRejectedCandidateReport(
                candidateFolder,
                fileName,
                rotatedPayloads,
                bestChoices,
                bestRejectedBytes.Length,
                bestValidationError,
                bestRejectedBytes);
        }
    }

    private static void AddPdfDescriptorOrderDynamicBoundaryCandidate(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder,
        List<CobaltItemGroupPayload> payloads,
        string fileName,
        byte[] signature,
        long expectedSize,
        int trailingStrip)
    {
        var orderedPayloads = payloads
            .Where(x => x.Type == 10 && x.Payload.Length > signature.Length + trailingStrip + 2)
            .OrderBy(x => GetType10PayloadLogicalOffset(x))
            .ThenBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step)
            .ToList();

        var match = orderedPayloads
            .Select((payload, index) => new
            {
                Payload = payload,
                Index = index,
                SignatureOffset = IndexOfBytes(payload.Payload, signature)
            })
            .FirstOrDefault(x => x.SignatureOffset >= 0);

        if (match == null)
            return;

        var rotatedPayloads = orderedPayloads
            .Skip(match.Index)
            .Concat(orderedPayloads.Take(match.Index))
            .ToList();

        var baseStart = match.SignatureOffset;
        var choiceSets = rotatedPayloads
            .Select(x => BuildPdfBoundaryChoices(x.Payload, baseStart, trailingStrip))
            .ToList();

        if (choiceSets.Any(x => x.Count == 0))
            return;

        long baseLength = 0;

        foreach (var payload in rotatedPayloads)
        {
            var length = Math.Max(0, payload.Payload.Length - baseStart - trailingStrip);
            baseLength += length;
        }

        var targetReduction = baseLength - expectedSize;

        if (targetReduction < -1024 || targetReduction > 1024)
            return;

        var states = new Dictionary<long, List<PdfBoundarySearchState>>
        {
            [0] = new List<PdfBoundarySearchState>
            {
                new PdfBoundarySearchState(null, null, 0, 0)
            }
        };

        const int maxStatesPerReduction = 48;
        const int reductionWindow = 1024;

        foreach (var choices in choiceSets)
        {
            var next = new Dictionary<long, List<PdfBoundarySearchState>>();

            foreach (var stateBucket in states)
            {
                foreach (var state in stateBucket.Value)
                {
                    foreach (var choice in choices)
                    {
                        var reduction = state.Reduction + choice.Reduction;

                        if (Math.Abs(reduction - targetReduction) > reductionWindow)
                            continue;

                        var nextState = new PdfBoundarySearchState(
                            state,
                            choice,
                            reduction,
                            state.Score + choice.Score);

                        if (!next.TryGetValue(reduction, out var bucket))
                        {
                            bucket = new List<PdfBoundarySearchState>();
                            next[reduction] = bucket;
                        }

                        bucket.Add(nextState);
                    }
                }
            }

            foreach (var key in next.Keys.ToList())
            {
                next[key] = next[key]
                    .OrderBy(x => x.Score)
                    .ThenBy(x => x.NonBaseChoiceCount)
                    .Take(maxStatesPerReduction)
                    .ToList();
            }

            states = next;

            if (states.Count == 0)
                return;
        }

        if (!states.TryGetValue(targetReduction, out var finalStates))
            return;

        byte[]? bestBytes = null;
        long bestStructuralScore = long.MaxValue;
        PdfBoundaryChoice[]? bestChoices = null;
        string bestValidationError = "";
        byte[]? bestRejectedBytes = null;

        foreach (var state in finalStates
                     .OrderBy(x => x.Score)
                     .ThenBy(x => x.NonBaseChoiceCount)
                     .Take(128))
        {
            var choices = UnwindPdfBoundaryChoices(state, rotatedPayloads.Count);
            var bytes = BuildCobaltPayloadBytes(rotatedPayloads, choices);

            if (bytes.Length != expectedSize)
                continue;

            if (!TryValidatePdfCandidateBytes(bytes, out var validationError))
            {
                if (TryRepairPdfStartXref(bytes, out var repairedBytes) &&
                    repairedBytes.Length == expectedSize &&
                    TryValidatePdfCandidateBytes(repairedBytes, out _))
                {
                    var structuralScore = GetPdfXrefConsistencyScore(repairedBytes);

                    if (structuralScore < bestStructuralScore)
                    {
                        bestBytes = repairedBytes;
                        bestStructuralScore = structuralScore;
                        bestChoices = choices;
                    }
                }
                else if (bestRejectedBytes == null)
                {
                    bestRejectedBytes = bytes;
                    bestChoices = choices;
                    bestValidationError = validationError;
                }

                continue;
            }

            var score = GetPdfXrefConsistencyScore(bytes);

            if (score < bestStructuralScore)
            {
                bestBytes = bytes;
                bestStructuralScore = score;
                bestChoices = choices;
            }
        }

        if (bestBytes != null)
        {
            var path = Path.Combine(candidateFolder, fileName);
            File.WriteAllBytes(path, bestBytes);

            candidates.Add(new CobaltItemGroupCandidate(
                fileName,
                path,
                new FileInfo(path).Length,
                rotatedPayloads.Count));
            return;
        }

        if (bestRejectedBytes != null && bestChoices != null)
        {
            WritePdfDynamicBoundaryRejectedCandidateReport(
                candidateFolder,
                fileName,
                rotatedPayloads,
                bestChoices,
                bestRejectedBytes.Length,
                bestValidationError,
                bestRejectedBytes);
        }
    }

    private static void WritePdfDynamicBoundaryRejectedCandidateReport(
        string candidateFolder,
        string fileName,
        List<CobaltItemGroupPayload> payloads,
        PdfBoundaryChoice[] choices,
        long length,
        string validationError,
        byte[] rejectedBytes)
    {
        var path = Path.Combine(
            candidateFolder,
            Path.GetFileNameWithoutExtension(fileName) + "_rejected_report.csv");

        if (File.Exists(path))
            return;

        var binPath = Path.Combine(
            candidateFolder,
            Path.GetFileNameWithoutExtension(fileName) + "_rejected.bin");

        if (!File.Exists(binPath))
            File.WriteAllBytes(binPath, rejectedBytes);

        using var writer = new StreamWriter(path, false, Encoding.UTF8);

        writer.WriteLine("Length,ValidationError,PayloadIndex,StreamId,BSN,Step,PayloadLength,Start,Tail,Reduction");

        for (var i = 0; i < payloads.Count; i++)
        {
            writer.WriteLine(
                    string.Join(",",
                        length,
                    Csv(validationError),
                        i,
                        payloads[i].StreamId,
                    payloads[i].BSN,
                    payloads[i].Step,
                    payloads[i].Payload.Length,
                    choices[i].Start,
                    choices[i].TrailingStrip,
                    choices[i].Reduction));
        }
    }

    private static List<PdfBoundaryChoice> BuildPdfBoundaryChoices(
        byte[] payload,
        int baseStart,
        int baseTrailingStrip)
    {
        var starts = new HashSet<int>();

        for (var delta = -4; delta <= 4; delta++)
            starts.Add(baseStart + delta);

        var markerOffset = FindLastCobaltInnerBoundaryMarker(payload);

        if (markerOffset >= 0)
        {
            for (var delta = 7; delta <= 10; delta++)
                starts.Add(markerOffset + delta);
        }

        var choices = new List<PdfBoundaryChoice>();

        foreach (var start in starts
                     .Where(x => x >= 0 && x < payload.Length)
                     .Distinct()
                     .OrderBy(x => Math.Abs(x - baseStart))
                     .ThenBy(x => x))
        {
            foreach (var tail in new[] { baseTrailingStrip - 1, baseTrailingStrip, baseTrailingStrip + 1 })
            {
                if (tail < 0)
                    continue;

                if (payload.Length - start - tail <= 0)
                    continue;

                var reduction = (start - baseStart) + (tail - baseTrailingStrip);
                var nonBase = start == baseStart && tail == baseTrailingStrip ? 0 : 1;
                var markerBonus = markerOffset >= 0 && start >= markerOffset + 7 && start <= markerOffset + 10
                    ? -1
                    : 0;

                var score =
                    nonBase * 50 +
                    Math.Abs(start - baseStart) * 20 +
                    Math.Abs(tail - baseTrailingStrip) * 5 +
                    markerBonus;

                choices.Add(new PdfBoundaryChoice(start, tail, reduction, score, nonBase));
            }
        }

        return choices
            .GroupBy(x => (x.Start, x.TrailingStrip))
            .Select(x => x.OrderBy(y => y.Score).First())
            .OrderBy(x => x.Score)
            .ToList();
    }

    private static int FindLastCobaltInnerBoundaryMarker(byte[] payload)
    {
        var end = Math.Min(payload.Length - 2, 112);
        var result = -1;

        for (var i = 48; i <= end; i++)
        {
            if (payload[i] == 0x20 && payload[i + 1] == 0x03)
                result = i;
        }

        return result;
    }

    private static PdfBoundaryChoice[] UnwindPdfBoundaryChoices(
        PdfBoundarySearchState state,
        int count)
    {
        var result = new PdfBoundaryChoice[count];
        var index = count - 1;
        var current = state;

        while (current.Choice != null && index >= 0)
        {
            result[index] = current.Choice;
            current = current.Previous!;
            index--;
        }

        return result;
    }

    private static byte[] BuildCobaltPayloadBytes(
        List<CobaltItemGroupPayload> payloads,
        PdfBoundaryChoice[] choices)
    {
        using var output = new MemoryStream();

        for (var i = 0; i < payloads.Count; i++)
        {
            var payload = payloads[i].Payload;
            var choice = choices[i];
            var start = Math.Min(choice.Start, payload.Length);
            var length = Math.Max(0, payload.Length - start - choice.TrailingStrip);

            output.Write(payload, start, length);
        }

        return output.ToArray();
    }

    private static List<CobaltType11StreamOrder> InferType11StreamOrders(
        List<ShreddedStreamRow> rows,
        List<CobaltItemGroupPayload> dataPayloads)
    {
        var knownStreamIds = dataPayloads
            .Select(x => (int)x.StreamId)
            .Distinct()
            .OrderBy(x => x)
            .ToHashSet();

        var orders = new List<CobaltType11StreamOrder>();

        foreach (var row in rows
                     .Where(x => x.Type == 11 && x.Content.Length >= 32)
                     .OrderBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            var matches = new List<(int Offset, int StreamId)>();

            for (var i = 0; i + 3 < row.Content.Length; i++)
            {
                if (row.Content[i + 2] != 0x00 ||
                    row.Content[i + 3] != 0x20)
                {
                    continue;
                }

                var value = BitConverter.ToUInt16(row.Content, i);

                if (knownStreamIds.Contains(value))
                    matches.Add((i, value));
            }

            if (matches.Count < 4)
                continue;

            var firstSeenOrder = matches
                .GroupBy(x => x.StreamId)
                .Select(g => g.First())
                .OrderBy(x => x.Offset)
                .Select(x => x.StreamId)
                .ToList();

            if (firstSeenOrder.Count >= 4)
            {
                orders.Add(new CobaltType11StreamOrder(
                    "first_seen_" + row.StreamId + "_" + row.BSN,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    0,
                    firstSeenOrder.Count > 0 ? matches.First(x => x.StreamId == firstSeenOrder[0]).Offset : 0,
                    firstSeenOrder));
            }

            var collapsedOffsetOrder = new List<int>();

            foreach (var match in matches.OrderBy(x => x.Offset))
            {
                if (collapsedOffsetOrder.Count == 0 ||
                    collapsedOffsetOrder[collapsedOffsetOrder.Count - 1] != match.StreamId)
                {
                    collapsedOffsetOrder.Add(match.StreamId);
                }
            }

            if (collapsedOffsetOrder.Count >= 4)
            {
                orders.Add(new CobaltType11StreamOrder(
                    "offset_collapsed_" + row.StreamId + "_" + row.BSN,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    0,
                    matches.Min(x => x.Offset),
                    collapsedOffsetOrder));
            }

            foreach (var run in BuildType11StreamRuns(matches))
            {
                if (run.StreamIds.Count < 4)
                    continue;

                    orders.Add(new CobaltType11StreamOrder(
                        "run_" + orders.Count.ToString("000"),
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        0,
                        run.StartOffset,
                        run.StreamIds));
            }
        }

        return orders
            .GroupBy(x => string.Join("|", x.StreamIds))
            .Select(g => g.First())
            .ToList();
    }

    private static List<(int StartOffset, List<int> StreamIds)> BuildType11StreamRuns(
        List<(int Offset, int StreamId)> matches)
    {
        var runs = new List<(int StartOffset, List<int> StreamIds)>();

        for (var i = 0; i < matches.Count; i++)
        {
            foreach (var spacing in new[] { 4, 7, 65, 72, 89, 96 })
            {
                var values = new List<int> { matches[i].StreamId };
                var expectedOffset = matches[i].Offset + spacing;

                for (var j = i + 1; j < matches.Count; j++)
                {
                    if (matches[j].Offset != expectedOffset)
                        continue;

                    values.Add(matches[j].StreamId);
                    expectedOffset += spacing;
                }

                if (values.Count >= 4)
                    runs.Add((matches[i].Offset, values));
            }
        }

        return runs;
    }

    private static List<CobaltItemGroupPayload> OrderDataPayloadsByType11StreamOrder(
        List<CobaltItemGroupPayload> dataPayloads,
        List<int> streamOrder)
    {
        var remaining = new List<CobaltItemGroupPayload>(dataPayloads);
        var ordered = new List<CobaltItemGroupPayload>();

        foreach (var streamId in streamOrder)
        {
            var match = remaining
                .Where(x => x.StreamId == streamId)
                .OrderBy(x => x.Step)
                .ThenBy(x => x.BSN)
                .FirstOrDefault();

            if (match == null)
                continue;

            ordered.Add(match);
            remaining.Remove(match);
        }

        ordered.AddRange(remaining
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step));

        return ordered;
    }

    private sealed class CobaltType11StreamOrder
    {
        public CobaltType11StreamOrder(
            string name,
            int sourceType,
            long sourceStreamId,
            long sourceBsn,
            int sourceStep,
            int payloadOffset,
            List<int> streamIds)
        {
            Name = name;
            SourceType = sourceType;
            SourceStreamId = sourceStreamId;
            SourceBSN = sourceBsn;
            SourceStep = sourceStep;
            PayloadOffset = payloadOffset;
            StreamIds = streamIds;
        }

        public string Name { get; }
        public int SourceType { get; }
        public long SourceStreamId { get; }
        public long SourceBSN { get; }
        public int SourceStep { get; }
        public int PayloadOffset { get; }
        public List<int> StreamIds { get; }
    }

    private static void AddPdfXrefRepairedDiagnosticCandidates(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder)
    {
        foreach (var candidate in candidates.ToList())
        {
            if (IsDiagnosticCandidateName(candidate.Name))
                continue;

            var repaired = TryCreatePdfXrefRepairedDiagnosticCandidate(
                candidate.Path,
                BuildDiagnosticCandidatePath(candidateFolder, candidate.Name, "xref_repaired"));

            if (repaired == null)
                continue;

            candidates.Add(new CobaltItemGroupCandidate(
                Path.GetFileName(repaired),
                repaired,
                new FileInfo(repaired).Length,
                candidate.PayloadCount));
        }
    }

    private static void AddPdfObjectRepackedDiagnosticCandidates(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder)
    {
        foreach (var candidate in candidates.ToList())
        {
            if (IsDiagnosticCandidateName(candidate.Name))
                continue;

            var repacked = TryCreatePdfObjectRepackedDiagnosticCandidate(
                candidate.Path,
                BuildDiagnosticCandidatePath(candidateFolder, candidate.Name, "object_repacked"));

            if (repacked == null)
                continue;

            candidates.Add(new CobaltItemGroupCandidate(
                Path.GetFileName(repacked),
                repacked,
                new FileInfo(repacked).Length,
                candidate.PayloadCount));
        }
    }

    private static void AddPdfDeclaredLengthRepackedDiagnosticCandidates(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder)
    {
        var reportPath = Path.Combine(candidateFolder, "declared_length_diagnostic_report.csv");

        using var report = new StreamWriter(reportPath, false, Encoding.UTF8);
        report.WriteLine("Candidate,Status,Message,OutputPath");

        foreach (var candidate in candidates.ToList())
        {
            if (IsDiagnosticCandidateName(candidate.Name))
                continue;

            var repacked = TryCreatePdfDeclaredLengthRepackedDiagnosticCandidate(
                candidate.Path,
                BuildDiagnosticCandidatePath(candidateFolder, candidate.Name, "declared_length"),
                out var status,
                out var message);

            report.WriteLine(
                string.Join(",",
                    Csv(candidate.Name),
                    Csv(status),
                    Csv(message),
                    Csv(repacked ?? "")));

            if (repacked == null)
                continue;

            candidates.Add(new CobaltItemGroupCandidate(
                Path.GetFileName(repacked),
                repacked,
                new FileInfo(repacked).Length,
                candidate.PayloadCount));
        }
    }

    private static void WritePdfImageFlateDiagnosticReport(
        List<CobaltItemGroupCandidate> candidates,
        string candidateFolder)
    {
        var reportPath = Path.Combine(candidateFolder, "pdf_image_flate_diagnostic_report.csv");

        using var writer = new StreamWriter(reportPath, false, Encoding.UTF8);
        writer.WriteLine("Candidate,ImageObject,LengthObject,DeclaredLength,StreamDataStart,BytesAvailable,Status,InflatedBytes,Error");

        foreach (var candidate in candidates
                     .Where(x => x.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                     .Where(x => x.Name.IndexOf("declared_length", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                 x.Name.IndexOf("object_repacked", StringComparison.OrdinalIgnoreCase) >= 0))
        {
            WritePdfImageFlateDiagnosticsForCandidate(writer, candidate);
        }
    }

    private static void WritePdfImageFlateDiagnosticsForCandidate(
        StreamWriter writer,
        CobaltItemGroupCandidate candidate)
    {
        try
        {
            var bytes = File.ReadAllBytes(candidate.Path);
            var text = Encoding.ASCII.GetString(bytes);
            var simpleLengthObjects = ReadSimplePdfLengthObjects(text);

            foreach (System.Text.RegularExpressions.Match match in
                     System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
            {
                var objectNumber = int.Parse(match.Groups[1].Value);
                var objectStart = match.Index;
                var streamKeywordOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
                var endObjOffset = text.IndexOf("endobj", objectStart, StringComparison.Ordinal);

                if (streamKeywordOffset < 0 ||
                    endObjOffset < 0 ||
                    streamKeywordOffset > endObjOffset)
                {
                    continue;
                }

                var objectHeader = text.Substring(objectStart, streamKeywordOffset - objectStart);

                if (objectHeader.IndexOf("/Subtype /Image", StringComparison.Ordinal) < 0 ||
                    objectHeader.IndexOf("/FlateDecode", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                var lengthMatch = System.Text.RegularExpressions.Regex.Match(
                    objectHeader,
                    @"/Length\s+(\d+)\s+0\s+R");

                var lengthObjectNumber = lengthMatch.Success
                    ? int.Parse(lengthMatch.Groups[1].Value)
                    : 0;

                var declaredLength = TryReadReferencedPdfStreamLength(
                    text,
                    objectStart,
                    streamKeywordOffset,
                    simpleLengthObjects);

                var streamDataStart = streamKeywordOffset + "stream".Length;

                if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0D)
                    streamDataStart++;

                if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0A)
                    streamDataStart++;

                var available = declaredLength.HasValue
                    ? Math.Min(declaredLength.Value, Math.Max(0, bytes.Length - streamDataStart))
                    : Math.Max(0, bytes.Length - streamDataStart);

                var status = "not_tested";
                long inflatedBytes = 0;
                var error = "";

                if (available > 2 &&
                    streamDataStart + available <= bytes.Length)
                {
                    TryInflateZlibPayload(
                        bytes,
                        streamDataStart,
                        available,
                        out status,
                        out inflatedBytes,
                        out error);
                }

                writer.WriteLine(
                    string.Join(",",
                        Csv(candidate.Name),
                        objectNumber,
                        lengthObjectNumber == 0 ? "" : lengthObjectNumber.ToString(),
                        declaredLength?.ToString() ?? "",
                        streamDataStart,
                        available,
                        Csv(status),
                        inflatedBytes,
                        Csv(error)));
            }
        }
        catch (Exception ex)
        {
            writer.WriteLine(
                string.Join(",",
                    Csv(candidate.Name),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "error",
                    "",
                    Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static void TryInflateZlibPayload(
        byte[] bytes,
        int offset,
        int length,
        out string status,
        out long inflatedBytes,
        out string error)
    {
        inflatedBytes = 0;
        error = "";

        try
        {
            using var input = new MemoryStream(bytes, offset + 2, length - 2, writable: false);
            using var deflate = new System.IO.Compression.DeflateStream(
                input,
                System.IO.Compression.CompressionMode.Decompress);

            var buffer = new byte[64 * 1024];

            while (true)
            {
                var read = deflate.Read(buffer, 0, buffer.Length);

                if (read == 0)
                    break;

                inflatedBytes += read;

                if (inflatedBytes > 1024L * 1024L * 1024L)
                    break;
            }

            status = "ok";
        }
        catch (Exception ex)
        {
            status = "failed";
            error = GetUsefulExceptionMessage(ex);
        }
    }

    private static byte[] TryInflateZlibPayloadToBytes(
        byte[] bytes,
        int offset,
        int length,
        out string status,
        out string error)
    {
        error = "";

        try
        {
            using var input = new MemoryStream(bytes, offset + 2, length - 2, writable: false);
            using var deflate = new System.IO.Compression.DeflateStream(
                input,
                System.IO.Compression.CompressionMode.Decompress);
            using var output = new MemoryStream();

            deflate.CopyTo(output);

            status = "ok";
            return output.ToArray();
        }
        catch (Exception ex)
        {
            status = "failed";
            error = GetUsefulExceptionMessage(ex);
            return Array.Empty<byte>();
        }
    }

    private static bool IsDiagnosticCandidateName(string candidateName)
    {
        return candidateName.IndexOf("_diagnostic", StringComparison.OrdinalIgnoreCase) >= 0 ||
               candidateName.IndexOf("_xref_repaired", StringComparison.OrdinalIgnoreCase) >= 0 ||
               candidateName.IndexOf("_object_repacked", StringComparison.OrdinalIgnoreCase) >= 0 ||
               candidateName.IndexOf("_declared_length", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string BuildDiagnosticCandidatePath(
        string candidateFolder,
        string sourceCandidateName,
        string diagnosticKind)
    {
        var stem = Path.GetFileNameWithoutExtension(sourceCandidateName);
        var safeStem = MakeSafeCandidateToken(stem);

        if (safeStem.Length > 70)
            safeStem = safeStem.Substring(0, 70);

        return Path.Combine(
            candidateFolder,
            safeStem + "_" + diagnosticKind + "_diagnostic.pdf");
    }

    private static string? TryCreatePdfXrefRepairedDiagnosticCandidate(
        string sourcePath,
        string repairedPath)
    {
        if (!File.Exists(sourcePath))
            return null;

        var bytes = File.ReadAllBytes(sourcePath);
        var pdfHeader = Encoding.ASCII.GetBytes("%PDF-");

        if (!StartsWithBytes(bytes, pdfHeader))
            return null;

        var startXrefMarker = Encoding.ASCII.GetBytes("startxref");
        var eofMarker = Encoding.ASCII.GetBytes("%%EOF");
        var startXrefOffset = LastIndexOfBytes(bytes, startXrefMarker);

        if (startXrefOffset < 0)
            return null;

        var eofOffset = IndexOfBytes(bytes, eofMarker, startXrefOffset);

        if (eofOffset < 0)
            return null;

        var realXrefOffset = FindLikelyPdfXrefTableOffset(bytes);

        if (realXrefOffset < 0)
            return null;

        var declaredXrefOffset = TryReadPositiveIntegerAfter(bytes, startXrefOffset + startXrefMarker.Length);

        if (declaredXrefOffset.HasValue &&
            declaredXrefOffset.Value == realXrefOffset &&
            eofOffset + eofMarker.Length >= bytes.Length)
        {
            return null;
        }

        using (var output = new FileStream(repairedPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            output.Write(bytes, 0, startXrefOffset);

            var tail = Encoding.ASCII.GetBytes(
                "startxref" + Environment.NewLine +
                realXrefOffset + Environment.NewLine +
                "%%EOF" + Environment.NewLine);

            output.Write(tail, 0, tail.Length);
        }

        return repairedPath;
    }

    private static string? TryCreatePdfObjectRepackedDiagnosticCandidate(
        string sourcePath,
        string repackedPath)
    {
        if (!File.Exists(sourcePath))
            return null;

        var bytes = File.ReadAllBytes(sourcePath);

        if (!StartsWithBytes(bytes, Encoding.ASCII.GetBytes("%PDF-")))
            return null;

        var text = Encoding.ASCII.GetString(bytes);
        var matches = System.Text.RegularExpressions.Regex.Matches(
            text,
            @"(?m)(\d+)\s+0\s+obj");

        if (matches.Count == 0)
            return null;

        var positions = matches
            .Cast<System.Text.RegularExpressions.Match>()
            .Select(match => new CobaltPdfObjectPosition(
                int.Parse(match.Groups[1].Value),
                match.Index))
            .OrderBy(x => x.Offset)
            .ToList();

        var objects = new Dictionary<int, byte[]>();

        for (var i = 0; i < positions.Count; i++)
        {
            var objectNumber = positions[i].ObjectNumber;
            var start = positions[i].Offset;
            var end = bytes.Length;

            if (i + 1 < positions.Count)
            {
                end = positions[i + 1].Offset;
            }
            else
            {
                var xref = text.IndexOf("xref", start, StringComparison.Ordinal);

                if (xref > start)
                    end = xref;
            }

            if (end <= start)
                continue;

            var chunk = new byte[end - start];
            Array.Copy(bytes, start, chunk, 0, chunk.Length);

            if (!objects.TryGetValue(objectNumber, out var existing) ||
                chunk.Length < existing.Length)
            {
                objects[objectNumber] = chunk;
            }
        }

        if (!objects.ContainsKey(1) || !objects.ContainsKey(2))
            return null;

        RewriteCobaltPdfIndirectStreamLengths(objects);

        using var output = new MemoryStream();
        var header = Encoding.ASCII.GetBytes("%PDF-1.4\n%Object repacked diagnostic candidate\n");
        output.Write(header, 0, header.Length);

        var offsets = new Dictionary<int, int>();

        foreach (var objectNumber in objects.Keys.OrderBy(x => x))
        {
            offsets[objectNumber] = (int)output.Position;

            var chunk = objects[objectNumber];
            output.Write(chunk, 0, chunk.Length);

            if (chunk.Length == 0 || chunk[chunk.Length - 1] != 0x0A)
                output.WriteByte(0x0A);
        }

        var xrefOffset = (int)output.Position;
        var maxObjectNumber = objects.Keys.Max();
        var xrefText = new StringBuilder();

        xrefText.AppendLine("xref");
        xrefText.AppendLine($"0 {maxObjectNumber + 1}");
        xrefText.AppendLine("0000000000 65535 f ");

        for (var objectNumber = 1; objectNumber <= maxObjectNumber; objectNumber++)
        {
            if (offsets.TryGetValue(objectNumber, out var offset))
                xrefText.AppendLine($"{offset:0000000000} 00000 n ");
            else
                xrefText.AppendLine("0000000000 65535 f ");
        }

        xrefText.AppendLine("trailer");
        xrefText.AppendLine("<<");
        xrefText.AppendLine($"/Size {maxObjectNumber + 1}");
        xrefText.AppendLine("/Root 2 0 R");
        xrefText.AppendLine(">>");
        xrefText.AppendLine("startxref");
        xrefText.AppendLine(xrefOffset.ToString());
        xrefText.AppendLine("%%EOF");

        var xrefBytes = Encoding.ASCII.GetBytes(xrefText.ToString());
        output.Write(xrefBytes, 0, xrefBytes.Length);

        File.WriteAllBytes(repackedPath, output.ToArray());

        return repackedPath;
    }

    private static string? TryCreatePdfDeclaredLengthRepackedDiagnosticCandidate(
        string sourcePath,
        string repackedPath,
        out string status,
        out string message)
    {
        try
        {
            if (!File.Exists(sourcePath))
            {
                status = "missing_source";
                message = "Source candidate does not exist.";
                return null;
            }

            var bytes = File.ReadAllBytes(sourcePath);

            if (!StartsWithBytes(bytes, Encoding.ASCII.GetBytes("%PDF-")))
            {
                status = "not_pdf";
                message = "Candidate does not start with %PDF-.";
                return null;
            }

            var text = Encoding.ASCII.GetString(bytes);
            var simpleLengthObjects = ReadSimplePdfLengthObjects(text);
            var parsedObjects = ParsePdfObjectsByObjectMarkersWithDeclaredStreamLengths(bytes, text, simpleLengthObjects);

            if (!parsedObjects.ContainsKey(1) || !parsedObjects.ContainsKey(2))
            {
                status = "missing_catalog_objects";
                message = $"Parsed {parsedObjects.Count} objects, but object 1 or 2 was not found.";
                return null;
            }

            WritePdfObjectsWithFreshXref(parsedObjects, repackedPath, "Declared length repacked diagnostic candidate");

            status = "created";
            message = $"Parsed {parsedObjects.Count} objects.";
            return repackedPath;
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException)
        {
            status = "error";
            message = GetUsefulExceptionMessage(ex);
            return null;
        }
    }

    private static Dictionary<int, int> ReadSimplePdfLengthObjects(string text)
    {
        var result = new Dictionary<int, int>();

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(
                     text,
                     @"(?m)(\d+)\s+0\s+obj\s+(\d+)\s+endobj"))
        {
            result[int.Parse(match.Groups[1].Value)] = int.Parse(match.Groups[2].Value);
        }

        return result;
    }

    private static Dictionary<int, byte[]> ParsePdfObjectsRespectingDeclaredStreamLengths(
        byte[] bytes,
        string text,
        Dictionary<int, int> simpleLengthObjects)
    {
        var objects = new Dictionary<int, byte[]>();
        var cursor = 0;

        while (cursor >= 0 && cursor < text.Length)
        {
            var match = System.Text.RegularExpressions.Regex.Match(
                text.Substring(cursor),
                @"(?m)(\d+)\s+0\s+obj");

            if (!match.Success)
                break;

            var objectNumber = int.Parse(match.Groups[1].Value);
            var objectStart = cursor + match.Index;

            if (objectNumber == 0)
            {
                cursor = objectStart + match.Length;
                continue;
            }

            var streamKeywordOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
            var firstEndObjOffset = text.IndexOf("endobj", objectStart, StringComparison.Ordinal);
            int objectEnd;

            if (streamKeywordOffset > objectStart &&
                firstEndObjOffset > streamKeywordOffset)
            {
                var declaredLength = TryReadReferencedPdfStreamLength(
                    text,
                    objectStart,
                    streamKeywordOffset,
                    simpleLengthObjects);

                if (declaredLength.HasValue)
                {
                    var streamDataStart = streamKeywordOffset + "stream".Length;

                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0D)
                        streamDataStart++;

                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0A)
                        streamDataStart++;

                    var expectedStreamEnd = Math.Min(bytes.Length, streamDataStart + declaredLength.Value);
                    var endObjAfterDeclaredStream = text.IndexOf("endobj", expectedStreamEnd, StringComparison.Ordinal);

                    objectEnd = endObjAfterDeclaredStream >= 0
                        ? endObjAfterDeclaredStream + "endobj".Length
                        : expectedStreamEnd;
                }
                else
                {
                    objectEnd = firstEndObjOffset + "endobj".Length;
                }
            }
            else
            {
                if (firstEndObjOffset < 0)
                    break;

                objectEnd = firstEndObjOffset + "endobj".Length;
            }

            if (objectEnd <= objectStart || objectStart >= bytes.Length)
                break;

            objectEnd = Math.Min(objectEnd, bytes.Length);
            var chunk = new byte[objectEnd - objectStart];
            Array.Copy(bytes, objectStart, chunk, 0, chunk.Length);

            if (!objects.ContainsKey(objectNumber))
                objects[objectNumber] = chunk;

            cursor = objectEnd;
        }

        return objects;
    }

    private static Dictionary<int, byte[]> ParsePdfObjectsByObjectMarkersWithDeclaredStreamLengths(
        byte[] bytes,
        string text,
        Dictionary<int, int> simpleLengthObjects)
    {
        var markers = System.Text.RegularExpressions.Regex.Matches(
                text,
                @"(?m)(\d+)\s+0\s+obj")
            .Cast<System.Text.RegularExpressions.Match>()
            .Select(match => new CobaltPdfObjectPosition(
                int.Parse(match.Groups[1].Value),
                match.Index))
            .Where(x => x.ObjectNumber > 0)
            .OrderBy(x => x.Offset)
            .ToList();

        var objects = new Dictionary<int, byte[]>();

        for (var i = 0; i < markers.Count; i++)
        {
            var marker = markers[i];
            var objectStart = marker.Offset;
            var nextMarkerOffset = i + 1 < markers.Count
                ? markers[i + 1].Offset
                : bytes.Length;

            var firstEndObjOffset = text.IndexOf("endobj", objectStart, StringComparison.Ordinal);

            if (firstEndObjOffset < 0)
                continue;

            var streamKeywordOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
            int objectEnd;

            if (streamKeywordOffset > objectStart &&
                streamKeywordOffset < firstEndObjOffset)
            {
                var declaredLength = TryReadReferencedPdfStreamLength(
                    text,
                    objectStart,
                    streamKeywordOffset,
                    simpleLengthObjects);

                if (declaredLength.HasValue)
                {
                    var streamDataStart = streamKeywordOffset + "stream".Length;

                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0D)
                        streamDataStart++;

                    if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0A)
                        streamDataStart++;

                    var expectedStreamEnd = Math.Min(bytes.Length, streamDataStart + declaredLength.Value);
                    var declaredEndStreamOffset = text.IndexOf("endstream", expectedStreamEnd, StringComparison.Ordinal);
                    var declaredEndObjOffset = declaredEndStreamOffset >= 0
                        ? text.IndexOf("endobj", declaredEndStreamOffset, StringComparison.Ordinal)
                        : -1;

                    objectEnd = declaredEndObjOffset >= 0
                        ? declaredEndObjOffset + "endobj".Length
                        : Math.Min(expectedStreamEnd, nextMarkerOffset);
                }
                else
                {
                    objectEnd = firstEndObjOffset + "endobj".Length;
                }
            }
            else
            {
                objectEnd = firstEndObjOffset + "endobj".Length;
            }

            if (objectEnd <= objectStart || objectStart >= bytes.Length)
                continue;

            objectEnd = Math.Min(objectEnd, bytes.Length);
            var chunk = new byte[objectEnd - objectStart];
            Array.Copy(bytes, objectStart, chunk, 0, chunk.Length);

            // For duplicate object markers inside binary streams, later real metadata/xref-tail
            // objects are generally more trustworthy than early false positives.
            objects[marker.ObjectNumber] = chunk;
        }

        return objects;
    }

    private static int? TryReadReferencedPdfStreamLength(
        string text,
        int objectStart,
        int streamKeywordOffset,
        Dictionary<int, int> simpleLengthObjects)
    {
        var objectHeader = text.Substring(objectStart, streamKeywordOffset - objectStart);
        var lengthMatch = System.Text.RegularExpressions.Regex.Match(
            objectHeader,
            @"/Length\s+(\d+)\s+0\s+R");

        if (lengthMatch.Success &&
            simpleLengthObjects.TryGetValue(int.Parse(lengthMatch.Groups[1].Value), out var indirectLength))
        {
            return indirectLength;
        }

        lengthMatch = System.Text.RegularExpressions.Regex.Match(
            objectHeader,
            @"/Length\s+(\d+)(?!\s+0\s+R)");

        if (lengthMatch.Success)
            return int.Parse(lengthMatch.Groups[1].Value);

        return null;
    }

    private static void WritePdfObjectsWithFreshXref(
        Dictionary<int, byte[]> objects,
        string outputPath,
        string comment)
    {
        using var output = new MemoryStream();
        var header = Encoding.ASCII.GetBytes("%PDF-1.4\n%" + comment + "\n");
        output.Write(header, 0, header.Length);

        var offsets = new Dictionary<int, int>();

        foreach (var objectNumber in objects.Keys.OrderBy(x => x))
        {
            offsets[objectNumber] = (int)output.Position;

            var chunk = objects[objectNumber];
            output.Write(chunk, 0, chunk.Length);

            if (chunk.Length == 0 || chunk[chunk.Length - 1] != 0x0A)
                output.WriteByte(0x0A);
        }

        var xrefOffset = (int)output.Position;
        var maxObjectNumber = objects.Keys.Max();
        var xrefText = new StringBuilder();

        xrefText.AppendLine("xref");
        xrefText.AppendLine($"0 {maxObjectNumber + 1}");
        xrefText.AppendLine("0000000000 65535 f ");

        for (var objectNumber = 1; objectNumber <= maxObjectNumber; objectNumber++)
        {
            if (offsets.TryGetValue(objectNumber, out var offset))
                xrefText.AppendLine($"{offset:0000000000} 00000 n ");
            else
                xrefText.AppendLine("0000000000 65535 f ");
        }

        xrefText.AppendLine("trailer");
        xrefText.AppendLine("<<");
        xrefText.AppendLine($"/Size {maxObjectNumber + 1}");
        xrefText.AppendLine("/Root 2 0 R");
        xrefText.AppendLine(">>");
        xrefText.AppendLine("startxref");
        xrefText.AppendLine(xrefOffset.ToString());
        xrefText.AppendLine("%%EOF");

        var xrefBytes = Encoding.ASCII.GetBytes(xrefText.ToString());
        output.Write(xrefBytes, 0, xrefBytes.Length);

        File.WriteAllBytes(outputPath, output.ToArray());
    }

    private static void RewriteCobaltPdfIndirectStreamLengths(Dictionary<int, byte[]> objects)
    {
        foreach (var objectNumber in objects.Keys.ToList())
        {
            var bytes = objects[objectNumber];
            var text = Encoding.ASCII.GetString(bytes);
            var lengthMatch = System.Text.RegularExpressions.Regex.Match(
                text,
                @"/Length\s+(\d+)\s+0\s+R");

            if (!lengthMatch.Success)
                continue;

            var lengthObjectNumber = int.Parse(lengthMatch.Groups[1].Value);
            var streamOffset = text.IndexOf("stream", StringComparison.Ordinal);
            var endStreamOffset = text.IndexOf("endstream", StringComparison.Ordinal);

            if (streamOffset < 0 || endStreamOffset <= streamOffset)
                continue;

            var actualLength = endStreamOffset - (streamOffset + "stream".Length + 1);

            if (actualLength < 0)
                continue;

            objects[lengthObjectNumber] = Encoding.ASCII.GetBytes(
                $"{lengthObjectNumber} 0 obj\n{actualLength}\nendobj\n");
        }
    }

    private static int FindLikelyPdfXrefTableOffset(byte[] bytes)
    {
        var xrefMarker = Encoding.ASCII.GetBytes("xref");
        var trailerMarker = Encoding.ASCII.GetBytes("trailer");
        var searchOffset = 0;

        while (searchOffset >= 0 && searchOffset < bytes.Length)
        {
            var xrefOffset = IndexOfBytes(bytes, xrefMarker, searchOffset);

            if (xrefOffset < 0)
                return -1;

            var afterXref = xrefOffset + xrefMarker.Length;

            if (afterXref < bytes.Length &&
                IsAsciiWhitespace(bytes[afterXref]) &&
                IndexOfBytes(bytes, trailerMarker, afterXref) > xrefOffset)
            {
                return xrefOffset;
            }

            searchOffset = xrefOffset + 1;
        }

        return -1;
    }

    private static string DetectKnownSignatureInFile(string path)
    {
        if (!File.Exists(path))
            return "";

        var max = (int)Math.Min(new FileInfo(path).Length, 4096);
        var buffer = new byte[max];

        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            fs.Read(buffer, 0, buffer.Length);
        }

        foreach (var signature in CobaltItemGroupKnownSignatures)
        {
            if (IndexOfBytes(buffer, signature.Bytes) >= 0)
                return signature.Name;
        }

        return "";
    }

    private static bool IsCandidateValidForSelectedFileType(string path, DocumentItemInfo item)
    {
        if (!File.Exists(path))
            return false;

        var extension = item.DisplayExtension.Trim('.').ToLowerInvariant();

        if (extension == "pdf")
            return LooksLikePdfCandidateFile(path);

        var signature = CobaltItemGroupKnownSignatures.FirstOrDefault(x =>
            (extension == "docx" || extension == "xlsx" || extension == "pptx") && x.Name == "ZIP Office Open XML" ||
            extension == "jpg" && x.Name == "JPEG" ||
            extension == "jpeg" && x.Name == "JPEG" ||
            extension == "png" && x.Name == "PNG" ||
            extension == "gif" && x.Name == "GIF" ||
            extension == "doc" && x.Name == "OLE Compound" ||
            extension == "xls" && x.Name == "OLE Compound" ||
            extension == "ppt" && x.Name == "OLE Compound" ||
            extension == "rtf" && x.Name == "RTF");

        if (signature == null)
            return true;

        var bytes = ReadFilePrefix(path, signature.Bytes.Length);

        return StartsWithBytes(bytes, signature.Bytes);
    }

    private static long GetCandidateStructuralScore(string path, DocumentItemInfo item)
    {
        if (!File.Exists(path))
            return long.MaxValue;

        var extension = item.DisplayExtension.Trim('.').ToLowerInvariant();

        if (extension != "pdf")
            return 0;

        try
        {
            return GetPdfXrefConsistencyScore(File.ReadAllBytes(path));
        }
        catch
        {
            return long.MaxValue / 2;
        }
    }

    private static long GetPdfXrefConsistencyScore(byte[] bytes)
    {
        var text = Encoding.ASCII.GetString(bytes);
        var startXrefOffset = text.LastIndexOf("startxref", StringComparison.Ordinal);

        if (startXrefOffset < 0)
            return long.MaxValue / 2;

        var declaredOffset = TryReadPositiveIntegerAfter(
            bytes,
            startXrefOffset + "startxref".Length);

        if (!declaredOffset.HasValue)
            return long.MaxValue / 2;

        var xrefStreamOffset = FindLastPdfXrefStreamObjectOffset(bytes);

        if (xrefStreamOffset < 0)
            return Math.Abs(declaredOffset.Value - FindLikelyPdfXrefTableOffset(bytes));

        var score = Math.Abs(declaredOffset.Value - (long)xrefStreamOffset);
        var streamScore = TryGetPdfXrefStreamOffsetMismatchScore(bytes, text, xrefStreamOffset);

        if (streamScore.HasValue)
            score += streamScore.Value;

        return score;
    }

    private static long? TryGetPdfXrefStreamOffsetMismatchScore(
        byte[] bytes,
        string text,
        int xrefStreamOffset)
    {
        var streamOffset = text.IndexOf("stream", xrefStreamOffset, StringComparison.Ordinal);
        var endStreamOffset = text.IndexOf("endstream", xrefStreamOffset, StringComparison.Ordinal);

        if (streamOffset < 0 ||
            endStreamOffset <= streamOffset)
        {
            return null;
        }

        var objectEndOffset = text.IndexOf("endobj", xrefStreamOffset, StringComparison.Ordinal);

        if (objectEndOffset < 0 ||
            streamOffset > objectEndOffset)
        {
            return null;
        }

        var headerText = text.Substring(xrefStreamOffset, streamOffset - xrefStreamOffset);
        var sizeMatch = System.Text.RegularExpressions.Regex.Match(headerText, @"/Size\s+(\d+)");
        var wMatch = System.Text.RegularExpressions.Regex.Match(headerText, @"/W\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s*\]");

        if (!sizeMatch.Success || !wMatch.Success)
            return null;

        var size = int.Parse(sizeMatch.Groups[1].Value);
        var fieldWidths = new[]
        {
            int.Parse(wMatch.Groups[1].Value),
            int.Parse(wMatch.Groups[2].Value),
            int.Parse(wMatch.Groups[3].Value)
        };

        var columns = fieldWidths.Sum();
        var dataStart = streamOffset + "stream".Length;

        if (dataStart < bytes.Length && bytes[dataStart] == 0x0D)
            dataStart++;

        if (dataStart < bytes.Length && bytes[dataStart] == 0x0A)
            dataStart++;

        var compressedLength = Math.Max(0, endStreamOffset - dataStart);
        var inflated = TryInflateZlibPayloadToBytes(
            bytes,
            dataStart,
            compressedLength,
            out var status,
            out _);

        if (status != "ok")
            return null;

        var decoded = DecodePdfPredictorRows(inflated, columns, size);

        if (decoded == null)
            return null;

        var score = 0L;
        var inspected = 0;

        for (var objectNumber = 1; objectNumber < size; objectNumber++)
        {
            var rowOffset = objectNumber * columns;

            if (rowOffset + columns > decoded.Length)
                break;

            var type = ReadBigEndianField(decoded, rowOffset, fieldWidths[0], defaultValue: 1);
            var offset = ReadBigEndianField(decoded, rowOffset + fieldWidths[0], fieldWidths[1], defaultValue: 0);

            if (type != 1)
                continue;

            var actual = text.IndexOf(objectNumber + " 0 obj", StringComparison.Ordinal);

            if (actual < 0)
                continue;

            inspected++;
            score += Math.Min(Math.Abs(actual - (long)offset), 1_000_000);
        }

        return inspected == 0
            ? null
            : score;
    }

    private static byte[]? DecodePdfPredictorRows(byte[] inflated, int columns, int expectedRows)
    {
        if (columns <= 0)
            return null;

        if (inflated.Length == columns * expectedRows)
            return inflated;

        var rowLength = columns + 1;

        if (inflated.Length < rowLength || inflated.Length % rowLength != 0)
            return null;

        var rows = inflated.Length / rowLength;
        var decoded = new byte[rows * columns];
        var previous = new byte[columns];

        for (var rowIndex = 0; rowIndex < rows; rowIndex++)
        {
            var filter = inflated[rowIndex * rowLength];
            var row = new byte[columns];
            Buffer.BlockCopy(inflated, rowIndex * rowLength + 1, row, 0, columns);

            for (var i = 0; i < columns; i++)
            {
                var left = i > 0 ? row[i - 1] : 0;
                var up = previous[i];
                var upperLeft = i > 0 ? previous[i - 1] : 0;
                var value = row[i];

                switch (filter)
                {
                    case 0:
                        row[i] = value;
                        break;

                    case 1:
                        row[i] = (byte)(value + left);
                        break;

                    case 2:
                        row[i] = (byte)(value + up);
                        break;

                    case 3:
                        row[i] = (byte)(value + ((left + up) / 2));
                        break;

                    case 4:
                        row[i] = (byte)(value + PaethPredictor(left, up, upperLeft));
                        break;

                    default:
                        return null;
                }

                decoded[rowIndex * columns + i] = row[i];
            }

            previous = row;
        }

        return decoded;
    }

    private static int PaethPredictor(int left, int up, int upperLeft)
    {
        var p = left + up - upperLeft;
        var pa = Math.Abs(p - left);
        var pb = Math.Abs(p - up);
        var pc = Math.Abs(p - upperLeft);

        if (pa <= pb && pa <= pc)
            return left;

        return pb <= pc ? up : upperLeft;
    }

    private static long ReadBigEndianField(
        byte[] bytes,
        int offset,
        int length,
        long defaultValue)
    {
        if (length == 0)
            return defaultValue;

        long value = 0;

        for (var i = 0; i < length; i++)
            value = (value << 8) | bytes[offset + i];

        return value;
    }

    private static bool LooksLikePdfCandidateFile(string path)
    {
        return LooksLikePdfCandidateBytes(File.ReadAllBytes(path));
    }

    private static bool LooksLikePdfCandidateBytes(byte[] bytes)
    {
        return TryValidatePdfCandidateBytes(bytes, out _);
    }

    private static bool TryValidatePdfCandidateBytes(byte[] bytes, out string error)
    {
        if (!StartsWithBytes(bytes, Encoding.ASCII.GetBytes("%PDF-")))
        {
            error = "missing_pdf_header";
            return false;
        }

        if (LastIndexOfBytes(bytes, Encoding.ASCII.GetBytes("%%EOF")) < 0)
        {
            error = "missing_eof_marker";
            return false;
        }

        var startXrefMarker = Encoding.ASCII.GetBytes("startxref");
        var xrefMarker = Encoding.ASCII.GetBytes("xref");
        var startXrefOffset = LastIndexOfBytes(bytes, startXrefMarker);

        if (startXrefOffset < 0)
        {
            error = "missing_startxref";
            return false;
        }

        var xrefOffset = TryReadPositiveIntegerAfter(bytes, startXrefOffset + startXrefMarker.Length);

        if (!xrefOffset.HasValue ||
            xrefOffset.Value < 0 ||
            xrefOffset.Value + xrefMarker.Length >= bytes.Length)
        {
            error = "invalid_startxref_value";
            return false;
        }

        var actualXrefOffset = xrefOffset.Value;

        while (actualXrefOffset < bytes.Length &&
               IsAsciiWhitespace(bytes[actualXrefOffset]))
        {
            actualXrefOffset++;
        }

        if (actualXrefOffset + xrefMarker.Length >= bytes.Length ||
            !StartsWithBytes(bytes.Skip(actualXrefOffset).Take(xrefMarker.Length).ToArray(), xrefMarker))
        {
            var xrefStreamObjectOffset = FindPdfXrefStreamObjectOffsetAtOrNear(bytes, actualXrefOffset, 2);

            if (xrefStreamObjectOffset < 0)
            {
                var nearbyXrefStreamOffset = FindNearbyPdfXrefStreamOffset(bytes, actualXrefOffset, 64);

                error = nearbyXrefStreamOffset >= 0
                    ? $"startxref_target_near_xref_stream_delta_{nearbyXrefStreamOffset - actualXrefOffset}"
                    : "startxref_target_is_not_xref_table_or_xref_stream";

                return false;
            }

            if (xrefStreamObjectOffset != actualXrefOffset)
            {
                error = $"startxref_points_inside_xref_stream_object_delta_{xrefStreamObjectOffset - actualXrefOffset}";
                return false;
            }
        }

        if (!AllPdfFlateStreamsInflate(bytes))
        {
            error = "flate_stream_validation_failed";
            return false;
        }

        error = "";
        return true;
    }

    private static bool LooksLikePdfXrefStreamAtOffset(byte[] bytes, int offset)
    {
        if (offset < 0 || offset >= bytes.Length)
            return false;

        var text = Encoding.ASCII.GetString(bytes);

        var match = System.Text.RegularExpressions.Regex.Match(
            text.Substring(offset, Math.Min(256, text.Length - offset)),
            @"^\s*(\d+)\s+(\d+)\s+obj");

        if (!match.Success)
            return false;

        var objectStart = offset + match.Index;
        var streamOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
        var endObjOffset = text.IndexOf("endobj", objectStart, StringComparison.Ordinal);

        if (streamOffset < 0 ||
            endObjOffset < 0 ||
            streamOffset > endObjOffset)
        {
            return false;
        }

        var nextObjectOffset = FindNextPdfObjectMarkerOffset(text, objectStart + match.Length);

        if (nextObjectOffset >= 0 && nextObjectOffset < streamOffset)
            return false;

        var headerText = text.Substring(
            objectStart,
            Math.Min(streamOffset - objectStart, 2048));

        if (headerText.IndexOf("/Type/XRef", StringComparison.Ordinal) < 0 &&
            headerText.IndexOf("/Type /XRef", StringComparison.Ordinal) < 0)
        {
            return false;
        }

        return headerText.IndexOf("/Root", StringComparison.Ordinal) >= 0 ||
               headerText.IndexOf("/Prev", StringComparison.Ordinal) >= 0 ||
               headerText.IndexOf("/Size", StringComparison.Ordinal) >= 0;
    }

    private static int FindNextPdfObjectMarkerOffset(string text, int startOffset)
    {
        if (startOffset < 0 || startOffset >= text.Length)
            return -1;

        var match = System.Text.RegularExpressions.Regex.Match(
            text.Substring(startOffset),
            @"\d+\s+\d+\s+obj");

        return match.Success
            ? startOffset + match.Index
            : -1;
    }

    private static int FindNearbyPdfXrefStreamOffset(
        byte[] bytes,
        int declaredOffset,
        int radius)
    {
        var start = Math.Max(0, declaredOffset - radius);
        var end = Math.Min(bytes.Length - 1, declaredOffset + radius);

        for (var offset = start; offset <= end; offset++)
        {
            if (LooksLikePdfXrefStreamAtOffset(bytes, offset))
                return offset;
        }

        return -1;
    }

    private static int FindPdfXrefStreamObjectOffsetAtOrNear(
        byte[] bytes,
        int declaredOffset,
        int radius)
    {
        var text = Encoding.ASCII.GetString(bytes);
        var start = Math.Max(0, declaredOffset - radius);
        var end = Math.Min(bytes.Length - 1, declaredOffset + radius);

        for (var offset = start; offset <= end; offset++)
        {
            var window = text.Substring(offset, Math.Min(64, text.Length - offset));

            if (!System.Text.RegularExpressions.Regex.IsMatch(window, @"^\d+\s+\d+\s+obj"))
                continue;

            if (LooksLikePdfXrefStreamAtOffset(bytes, offset))
                return offset;
        }

        return -1;
    }

    private static bool TryRepairPdfStartXref(byte[] bytes, out byte[] repairedBytes)
    {
        repairedBytes = Array.Empty<byte>();

        var text = Encoding.ASCII.GetString(bytes);
        var startXrefOffset = text.LastIndexOf("startxref", StringComparison.Ordinal);

        if (startXrefOffset < 0)
            return false;

        var numberStart = startXrefOffset + "startxref".Length;

        while (numberStart < text.Length && IsAsciiWhitespace((byte)text[numberStart]))
            numberStart++;

        var numberEnd = numberStart;

        while (numberEnd < text.Length && text[numberEnd] >= '0' && text[numberEnd] <= '9')
            numberEnd++;

        if (numberEnd <= numberStart)
            return false;

        var oldNumberText = text.Substring(numberStart, numberEnd - numberStart);
        var actualOffset = FindLastPdfXrefStreamObjectOffset(bytes);

        if (actualOffset < 0)
        {
            actualOffset = FindLikelyPdfXrefTableOffset(bytes);

            if (actualOffset < 0)
                return false;
        }

        var newNumberText = actualOffset.ToString();

        if (newNumberText.Length != oldNumberText.Length)
            return false;

        repairedBytes = (byte[])bytes.Clone();
        var newNumberBytes = Encoding.ASCII.GetBytes(newNumberText);
        Buffer.BlockCopy(newNumberBytes, 0, repairedBytes, numberStart, newNumberBytes.Length);

        return true;
    }

    private static int FindLastPdfXrefStreamObjectOffset(byte[] bytes)
    {
        var text = Encoding.ASCII.GetString(bytes);
        var searchOffset = text.Length;

        while (searchOffset > 0)
        {
            var typeOffset = text.LastIndexOf("/Type/XRef", searchOffset - 1, StringComparison.Ordinal);

            if (typeOffset < 0)
                typeOffset = text.LastIndexOf("/Type /XRef", searchOffset - 1, StringComparison.Ordinal);

            if (typeOffset < 0)
                return -1;

            var objectStart = FindPdfObjectStartBefore(text, typeOffset);

            if (objectStart >= 0 && LooksLikePdfXrefStreamAtOffset(bytes, objectStart))
                return objectStart;

            searchOffset = typeOffset;
        }

        return -1;
    }

    private static int FindPdfObjectStartBefore(string text, int offset)
    {
        var searchStart = Math.Max(0, offset - 2048);
        var slice = text.Substring(searchStart, offset - searchStart);
        var matches = System.Text.RegularExpressions.Regex.Matches(
            slice,
            @"(?m)(\d+)\s+(\d+)\s+obj");

        if (matches.Count == 0)
            return -1;

        return searchStart + matches[matches.Count - 1].Index;
    }

    private static bool AllPdfFlateStreamsInflate(byte[] bytes)
    {
        var text = Encoding.ASCII.GetString(bytes);

        // In encrypted PDFs, object streams can be encrypted before/around filter
        // processing. Header, EOF, startxref, and xref-stream validation are the
        // strict structural checks we can safely apply without decrypting content.
        if (text.IndexOf("/Encrypt", StringComparison.Ordinal) >= 0)
            return true;

        var simpleLengthObjects = ReadSimplePdfLengthObjects(text);

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
        {
            var objectStart = match.Index;
            var streamKeywordOffset = text.IndexOf("stream", objectStart, StringComparison.Ordinal);
            var endObjOffset = text.IndexOf("endobj", objectStart, StringComparison.Ordinal);

            if (streamKeywordOffset < 0 ||
                endObjOffset < 0 ||
                streamKeywordOffset > endObjOffset)
            {
                continue;
            }

            var objectHeader = text.Substring(objectStart, streamKeywordOffset - objectStart);

            if (objectHeader.IndexOf("/FlateDecode", StringComparison.Ordinal) < 0)
                continue;

            var streamDataStart = streamKeywordOffset + "stream".Length;

            if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0D)
                streamDataStart++;

            if (streamDataStart < bytes.Length && bytes[streamDataStart] == 0x0A)
                streamDataStart++;

            var declaredLength = TryReadReferencedPdfStreamLength(
                text,
                objectStart,
                streamKeywordOffset,
                simpleLengthObjects);

            var streamDataLength = declaredLength;

            if (!streamDataLength.HasValue)
            {
                var endStreamOffset = text.IndexOf("endstream", streamDataStart, StringComparison.Ordinal);

                if (endStreamOffset < streamDataStart)
                    return false;

                streamDataLength = endStreamOffset - streamDataStart;
            }

            if (streamDataLength.Value <= 2 ||
                streamDataStart < 0 ||
                streamDataStart + streamDataLength.Value > bytes.Length)
            {
                return false;
            }

            var inflated = TryInflateZlibPayloadToBytes(
                bytes,
                streamDataStart,
                streamDataLength.Value,
                out var status,
                out _);

            if (!string.Equals(status, "ok", StringComparison.OrdinalIgnoreCase))
                return false;

            if (objectHeader.IndexOf("/DCTDecode", StringComparison.Ordinal) >= 0 &&
                !LooksLikeCompleteJpegPayload(inflated))
            {
                return false;
            }
        }

        return true;
    }

    private static bool LooksLikeCompleteJpegPayload(byte[] bytes)
    {
        if (bytes.Length < 4)
            return false;

        if (bytes[0] != 0xFF ||
            bytes[1] != 0xD8 ||
            bytes[2] != 0xFF)
        {
            return false;
        }

        for (var i = bytes.Length - 2; i >= Math.Max(0, bytes.Length - 32); i--)
        {
            if (bytes[i] == 0xFF && bytes[i + 1] == 0xD9)
                return true;
        }

        return false;
    }

    private static byte[] ReadFilePrefix(string path, int length)
    {
        var buffer = new byte[length];

        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            var read = fs.Read(buffer, 0, buffer.Length);

            if (read == buffer.Length)
                return buffer;

            return buffer.Take(read).ToArray();
        }
    }

    private static bool StartsWithBytes(byte[] data, byte[] prefix)
    {
        if (data.Length < prefix.Length)
            return false;

        for (var i = 0; i < prefix.Length; i++)
        {
            if (data[i] != prefix[i])
                return false;
        }

        return true;
    }

    private static int? TryReadPositiveIntegerAfter(byte[] bytes, int offset)
    {
        while (offset < bytes.Length && IsAsciiWhitespace(bytes[offset]))
            offset++;

        long value = 0;
        var hasDigit = false;

        while (offset < bytes.Length && bytes[offset] >= (byte)'0' && bytes[offset] <= (byte)'9')
        {
            hasDigit = true;
            value = (value * 10) + bytes[offset] - (byte)'0';

            if (value > int.MaxValue)
                return null;

            offset++;
        }

        return hasDigit ? (int)value : null;
    }

    private static bool IsAsciiWhitespace(byte value)
    {
        return value == 0 ||
               value == 9 ||
               value == 10 ||
               value == 12 ||
               value == 13 ||
               value == 32;
    }

    private static string MakeSafeCandidateToken(string value)
    {
        var builder = new StringBuilder();

        foreach (var ch in value.ToLowerInvariant())
        {
            builder.Append(char.IsLetterOrDigit(ch) ? ch : '_');
        }

        return builder.ToString().Trim('_');
    }

    private sealed class CobaltItemGroupPayload
    {
        public CobaltItemGroupPayload(
            int type,
            long streamId,
            long bsn,
            int step,
            int groupId,
            int payloadStart,
            byte[] payload)
        {
            Type = type;
            StreamId = streamId;
            BSN = bsn;
            Step = step;
            GroupId = groupId;
            PayloadStart = payloadStart;
            Payload = payload;
        }

        public int Type { get; }
        public long StreamId { get; }
        public long BSN { get; }
        public int Step { get; }
        public int GroupId { get; }
        public int PayloadStart { get; }
        public byte[] Payload { get; }
    }

    private sealed class CobaltOrdinalPayload
    {
        public CobaltOrdinalPayload(
            CobaltItemGroupPayload payload,
            int ordinal,
            int signatureOffset)
        {
            Payload = payload;
            Ordinal = ordinal;
            SignatureOffset = signatureOffset;
        }

        public CobaltItemGroupPayload Payload { get; }
        public int Ordinal { get; }
        public int SignatureOffset { get; }
    }

    private sealed class PdfBoundaryChoice
    {
        public PdfBoundaryChoice(
            int start,
            int trailingStrip,
            int reduction,
            int score,
            int nonBaseChoiceCount)
        {
            Start = start;
            TrailingStrip = trailingStrip;
            Reduction = reduction;
            Score = score;
            NonBaseChoiceCount = nonBaseChoiceCount;
        }

        public int Start { get; }
        public int TrailingStrip { get; }
        public int Reduction { get; }
        public int Score { get; }
        public int NonBaseChoiceCount { get; }
    }

    private sealed class PdfBoundarySearchState
    {
        public PdfBoundarySearchState(
            PdfBoundarySearchState? previous,
            PdfBoundaryChoice? choice,
            long reduction,
            int score)
        {
            Previous = previous;
            Choice = choice;
            Reduction = reduction;
            Score = score;
            NonBaseChoiceCount =
                (previous?.NonBaseChoiceCount ?? 0) +
                (choice?.NonBaseChoiceCount ?? 0);
        }

        public PdfBoundarySearchState? Previous { get; }
        public PdfBoundaryChoice? Choice { get; }
        public long Reduction { get; }
        public int Score { get; }
        public int NonBaseChoiceCount { get; }
    }

    private sealed class CobaltItemGroupCandidate
    {
        public CobaltItemGroupCandidate(string name, string path, long length, int payloadCount)
        {
            Name = name;
            Path = path;
            Length = length;
            PayloadCount = payloadCount;
        }

        public string Name { get; }
        public string Path { get; }
        public long Length { get; }
        public int PayloadCount { get; }
    }

    private sealed class CobaltItemGroupKnownSignature
    {
        public CobaltItemGroupKnownSignature(string name, byte[] bytes)
        {
            Name = name;
            Bytes = bytes;
        }

        public string Name { get; }
        public byte[] Bytes { get; }
    }

    private sealed class CobaltPdfObjectPosition
    {
        public CobaltPdfObjectPosition(int objectNumber, int offset)
        {
            ObjectNumber = objectNumber;
            Offset = offset;
        }

        public int ObjectNumber { get; }
        public int Offset { get; }
    }

    private static readonly CobaltItemGroupKnownSignature[] CobaltItemGroupKnownSignatures =
    {
        new("PDF", Encoding.ASCII.GetBytes("%PDF-")),
        new("ZIP Office Open XML", new byte[] { 0x50, 0x4B, 0x03, 0x04 }),
        new("JPEG", new byte[] { 0xFF, 0xD8, 0xFF }),
        new("PNG", new byte[] { 0x89, 0x50, 0x4E, 0x47 }),
        new("GIF", Encoding.ASCII.GetBytes("GIF8")),
        new("OLE Compound", new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }),
        new("RTF", Encoding.ASCII.GetBytes(@"{\rtf"))
    };

    private static void WriteCobaltType10DescriptorMap(
        List<ShreddedStreamRow> rows,
        string headerFooterProbeCsvPath,
        string csvPath)
    {
        var payloads = ReadCobaltItemGroupPayloads(rows, headerFooterProbeCsvPath)
            .Where(x => x.Type == 10)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ThenBy(x => x.Step)
            .ToList();

        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine(
            "StreamId,BSN,Step,PayloadStart,PayloadLength,Kind,U16_0,U16_1,U16_2,U16_3,DescriptorLogicalOffset,InferredLogicalOffset,First32Hex,AsciiPreview");

        foreach (var payload in payloads)
        {
            var u16 = ReadFirstUInt16Values(payload.Payload, 4);
            var descriptorLogicalOffset = payload.Payload.Length == 8 && u16.Count >= 4
                ? (long?)u16[3] * 65536L + u16[1]
                : null;

            writer.WriteLine(
                string.Join(",",
                    payload.StreamId,
                    payload.BSN,
                    payload.Step,
                    payload.PayloadStart,
                    payload.Payload.Length,
                    Csv(payload.Payload.Length == 8 ? "Descriptor" : "Data"),
                    u16.Count > 0 ? u16[0].ToString() : "",
                    u16.Count > 1 ? u16[1].ToString() : "",
                    u16.Count > 2 ? u16[2].ToString() : "",
                    u16.Count > 3 ? u16[3].ToString() : "",
                    descriptorLogicalOffset?.ToString() ?? "",
                    GetType10PayloadLogicalOffset(payload),
                    Csv(ToHex(payload.Payload.Take(Math.Min(32, payload.Payload.Length)).ToArray())),
                    Csv(GetAsciiPreview(payload.Payload, 0, Math.Min(payload.Payload.Length, 96)))));
        }
    }

    private static long GetType10PayloadLogicalOffset(CobaltItemGroupPayload payload)
    {
        // The 8-byte Type 10 descriptors show this logical stream coordinate:
        // stream 298 -> 4*65536 + 43520, stream 320 -> 5*65536 + 512, stream 322 -> 5*65536 + 2560.
        // That fits ((StreamId - 256) * 1024 + 512) + 4*65536. Step is only a tie-breaker within a row.
        return ((payload.StreamId - 256) * 1024L) + 512 + (4 * 65536L) + payload.Step;
    }

    private static List<ushort> ReadFirstUInt16Values(byte[] data, int maxValues)
    {
        var result = new List<ushort>();

        for (var i = 0; i + 1 < data.Length && result.Count < maxValues; i += 2)
            result.Add(BitConverter.ToUInt16(data, i));

        return result;
    }

    private static void WriteCobaltItemGroupPayloadSummary(
        List<ShreddedStreamRow> rows,
        string headerFooterProbeCsvPath,
        string csvPath)
    {
        if (!File.Exists(headerFooterProbeCsvPath))
            return;

        var probeRows = File.ReadLines(headerFooterProbeCsvPath)
            .Skip(1)
            .Select(ParseHeaderFooterProbeCsvLine)
            .Where(x => x != null)
            .Select(x => x!)
            .Where(x =>
                x.HeaderFooterType == "ItemGroupHeader" &&
                x.Status == "1" &&
                x.Variant == "full")
            .ToList();

        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine("Type,StreamId,BSN,Step,GroupId,PayloadStart,PayloadLength,First64Hex,First32U16,First16U32,AsciiPreview");

        foreach (var probe in probeRows)
        {
            var row = rows.FirstOrDefault(x =>
                x.Type == probe.Type &&
                x.StreamId == probe.StreamId &&
                x.BSN == probe.BSN);

            if (row == null)
                continue;

            if (probe.Length <= 0 ||
                probe.PositionAfter < 0 ||
                probe.PositionAfter + probe.Length > row.Content.Length)
            {
                continue;
            }

            var payload = row.Content
                .Skip(probe.PositionAfter)
                .Take(probe.Length)
                .ToArray();

            var first64 = payload.Take(Math.Min(64, payload.Length)).ToArray();
            var firstU16 = new List<string>();
            var firstU32 = new List<string>();

            for (var i = 0; i + 1 < payload.Length && firstU16.Count < 32; i += 2)
                firstU16.Add(BitConverter.ToUInt16(payload, i).ToString());

            for (var i = 0; i + 3 < payload.Length && firstU32.Count < 16; i += 4)
                firstU32.Add(BitConverter.ToUInt32(payload, i).ToString());

            writer.WriteLine(
                string.Join(",",
                    probe.Type,
                    probe.StreamId,
                    probe.BSN,
                    probe.Step,
                    probe.Id,
                    probe.PositionAfter,
                    probe.Length,
                    Csv(ToHex(first64)),
                    Csv(string.Join("|", firstU16)),
                    Csv(string.Join("|", firstU32)),
                    Csv(GetAsciiPreview(payload, 0, Math.Min(payload.Length, 96)))));
        }
    }

    private static HeaderFooterProbeRow? ParseHeaderFooterProbeCsvLine(string line)
    {
        var fields = SplitSimpleCsvLine(line);

        if (fields.Count < 11)
            return null;

        if (!int.TryParse(fields[0], out var type) ||
            !long.TryParse(fields[1], out var streamId) ||
            !long.TryParse(fields[2], out var bsn) ||
            !int.TryParse(fields[4], out var step) ||
            !int.TryParse(fields[6], out var positionAfter) ||
            !int.TryParse(fields[8], out var id) ||
            !int.TryParse(fields[9], out var length))
        {
            return null;
        }

        return new HeaderFooterProbeRow
        {
            Type = type,
            StreamId = streamId,
            BSN = bsn,
            Variant = fields[3],
            Step = step,
            PositionAfter = positionAfter,
            HeaderFooterType = fields[7],
            Id = id,
            Length = length,
            Status = fields[10]
        };
    }

    private static List<string> SplitSimpleCsvLine(string line)
    {
        var result = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;

        for (var i = 0; i < line.Length; i++)
        {
            var ch = line[i];

            if (ch == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }

                continue;
            }

            if (ch == ',' && !inQuotes)
            {
                result.Add(current.ToString());
                current.Clear();
                continue;
            }

            current.Append(ch);
        }

        result.Add(current.ToString());
        return result;
    }

    private sealed class HeaderFooterProbeRow
    {
        public int Type { get; set; }
        public long StreamId { get; set; }
        public long BSN { get; set; }
        public string Variant { get; set; } = "";
        public int Step { get; set; }
        public int PositionAfter { get; set; }
        public string HeaderFooterType { get; set; } = "";
        public int Id { get; set; }
        public int Length { get; set; }
        public string Status { get; set; } = "";
    }

    private void WriteCobaltHeaderFooterOffsetScan(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine("Type,StreamId,BSN,Variant,Offset,BytesConsumed,HeaderFooterType,Id,Length,RemainingAfterHeader,Plausible,Error");

        try
        {
            var binding = BindCobaltHeaderFooterReader(writer, GetAllCobaltTypes());

            if (binding == null)
                return;

            foreach (var row in rows
                         .Where(x => x.Type == 10 || x.Type == 11)
                         .OrderBy(x => x.StreamId)
                         .ThenBy(x => x.BSN))
            {
                foreach (var variant in BuildCobaltRowProbeVariants(row))
                {
                    ScanHeaderFooterOffsetsForVariant(
                        writer,
                        binding.Value.ReaderCtor,
                        binding.Value.ReadMethod,
                        row,
                        variant.Name,
                        variant.Content);
                }
            }
        }
        catch (Exception ex)
        {
            writer.WriteLine(string.Join(",", "", "", "", "fatal", "", "", "", "", "", "", "0",
                Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static CobaltHeaderFooterReaderBinding? BindCobaltHeaderFooterReader(
        StreamWriter writer,
        List<Type> allTypes)
    {
        var readerType = allTypes.FirstOrDefault(t =>
            string.Equals(t.FullName, "Cobalt.Base.IO.CobaltBinaryReader", StringComparison.OrdinalIgnoreCase));
        var headerFootersType = allTypes.FirstOrDefault(t =>
            string.Equals(t.FullName, "Cobalt.Serialization.SerializationHeaderFooters", StringComparison.OrdinalIgnoreCase));

        if (readerType == null || headerFootersType == null)
        {
            writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                Csv("CobaltBinaryReader or SerializationHeaderFooters type was not found.")));
            return null;
        }

        var readerCtor = readerType.GetConstructor(new[] { typeof(Stream), typeof(bool), typeof(int) });

        if (readerCtor == null)
        {
            writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                Csv("CobaltBinaryReader(Stream, Boolean, Int32) constructor was not found.")));
            return null;
        }

        var readMethod = headerFootersType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
            .FirstOrDefault(m =>
            {
                if (!string.Equals(m.Name, "Read", StringComparison.OrdinalIgnoreCase))
                    return false;

                var parameters = m.GetParameters();

                return parameters.Length == 4 &&
                       string.Equals(parameters[0].ParameterType.FullName, "Cobalt.Base.IO.CobaltBinaryReader", StringComparison.OrdinalIgnoreCase);
            });

        if (readMethod == null)
        {
            writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                Csv("SerializationHeaderFooters.Read(...) was not found.")));
            return null;
        }

        return new CobaltHeaderFooterReaderBinding(readerCtor, readMethod);
    }

    private void ScanHeaderFooterOffsetsForVariant(
        StreamWriter writer,
        ConstructorInfo readerCtor,
        MethodInfo readMethod,
        ShreddedStreamRow row,
        string variantName,
        byte[] content)
    {
        for (var offset = 0; offset < content.Length; offset++)
        {
            using var stream = new MemoryStream(content, offset, content.Length - offset, writable: false);
            using var reader = readerCtor.Invoke(new object[] { stream, true, 4096 }) as IDisposable;

            if (reader == null)
                continue;

            var readerObject = (object)reader;
            var logicalPositionProperty = readerObject.GetType().GetProperty(
                "LogicalPosition",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            var seekForwardMethod = readerObject.GetType().GetMethod(
                "SeekForward",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                binder: null,
                types: new[] { typeof(long) },
                modifiers: null);
            var positionBefore = ReadLongProperty(logicalPositionProperty, readerObject, stream.Position);
            var args = new object?[] { readerObject, null, 0, (ulong)0 };

            try
            {
                readMethod.Invoke(null, args);
            }
            catch
            {
                continue;
            }

            var positionAfter = ReadLongProperty(logicalPositionProperty, readerObject, stream.Position);
            var bytesConsumed = positionAfter - positionBefore;
            var headerFooterType = args[1]?.ToString() ?? "";
            var id = args[2] is int idValue ? idValue : 0;
            var length = args[3] is ulong lengthValue ? lengthValue : 0;
            var remainingAfterHeader = content.Length - offset - bytesConsumed;
            var plausible =
                bytesConsumed > 0 &&
                bytesConsumed <= 4 &&
                remainingAfterHeader >= 0 &&
                length <= (ulong)Math.Max(0, remainingAfterHeader) &&
                IsPlausibleHeaderFooter(headerFooterType, id, length) &&
                IsUsefulHeaderFooterOffset(offset, headerFooterType, id, length);

            if (!plausible)
                continue;

            writer.WriteLine(
                string.Join(",",
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    Csv(variantName),
                    offset,
                    bytesConsumed,
                    Csv(headerFooterType),
                    id,
                    length,
                    remainingAfterHeader,
                    "1",
                    ""));
        }
    }

    private static bool IsPlausibleHeaderFooter(string headerFooterType, int id, ulong length)
    {
        if (string.Equals(headerFooterType, "ElementFooter", StringComparison.OrdinalIgnoreCase))
            return false;

        if (string.Equals(headerFooterType, "ElementHeader", StringComparison.OrdinalIgnoreCase))
            return id >= 0 && id <= 1024 && length == 0;

        if (string.Equals(headerFooterType, "ItemGroupHeader", StringComparison.OrdinalIgnoreCase))
            return id >= 0 && id <= 1024;

        return false;
    }

    private static bool IsUsefulHeaderFooterOffset(int offset, string headerFooterType, int id, ulong length)
    {
        if (offset == 0 || offset == 1 || offset == 4 || offset == 8 || offset == 16 || offset == 32 || offset == 64 || offset == 96)
            return true;

        if (string.Equals(headerFooterType, "ItemGroupHeader", StringComparison.OrdinalIgnoreCase))
            return length >= 32;

        return false;
    }

    private readonly struct CobaltHeaderFooterReaderBinding
    {
        public CobaltHeaderFooterReaderBinding(ConstructorInfo readerCtor, MethodInfo readMethod)
        {
            ReaderCtor = readerCtor;
            ReadMethod = readMethod;
        }

        public ConstructorInfo ReaderCtor { get; }

        public MethodInfo ReadMethod { get; }
    }

    private void WriteCobaltHeaderFooterProbe(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine("Type,StreamId,BSN,Variant,Step,PositionBefore,PositionAfter,HeaderFooterType,Id,Length,Status,Error");

        try
        {
            var allTypes = GetAllCobaltTypes();
            var readerType = allTypes.FirstOrDefault(t =>
                string.Equals(t.FullName, "Cobalt.Base.IO.CobaltBinaryReader", StringComparison.OrdinalIgnoreCase));
            var headerFootersType = allTypes.FirstOrDefault(t =>
                string.Equals(t.FullName, "Cobalt.Serialization.SerializationHeaderFooters", StringComparison.OrdinalIgnoreCase));

            if (readerType == null || headerFootersType == null)
            {
                writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                    Csv("CobaltBinaryReader or SerializationHeaderFooters type was not found.")));
                return;
            }

            var readerCtor = readerType.GetConstructor(new[] { typeof(Stream), typeof(bool), typeof(int) });

            if (readerCtor == null)
            {
                writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                    Csv("CobaltBinaryReader(Stream, Boolean, Int32) constructor was not found.")));
                return;
            }

            var readMethod = headerFootersType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                .FirstOrDefault(m =>
                {
                    if (!string.Equals(m.Name, "Read", StringComparison.OrdinalIgnoreCase))
                        return false;

                    var parameters = m.GetParameters();

                    return parameters.Length == 4 &&
                           string.Equals(parameters[0].ParameterType.FullName, "Cobalt.Base.IO.CobaltBinaryReader", StringComparison.OrdinalIgnoreCase);
                });

            if (readMethod == null)
            {
                writer.WriteLine(string.Join(",", "", "", "", "binding", "", "", "", "", "", "", "0",
                    Csv("SerializationHeaderFooters.Read(...) was not found.")));
                return;
            }

            foreach (var row in rows
                         .Where(x => x.Type == 10 || x.Type == 11)
                         .OrderBy(x => x.StreamId)
                         .ThenBy(x => x.BSN))
            {
                foreach (var variant in BuildCobaltRowProbeVariants(row))
                {
                    ProbeHeaderFooterVariant(
                        writer,
                        readerCtor,
                        readMethod,
                        row,
                        variant.Name,
                        variant.Content);
                }
            }
        }
        catch (Exception ex)
        {
            writer.WriteLine(string.Join(",", "", "", "", "fatal", "", "", "", "", "", "", "0",
                Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static void ProbeHeaderFooterVariant(
        StreamWriter writer,
        ConstructorInfo readerCtor,
        MethodInfo readMethod,
        ShreddedStreamRow row,
        string variantName,
        byte[] content)
    {
        try
        {
            using var stream = new MemoryStream(content, writable: false);
            using var reader = readerCtor.Invoke(new object[] { stream, true, 4096 }) as IDisposable;

            if (reader == null)
                throw new InvalidOperationException("Could not create CobaltBinaryReader.");

            var readerObject = (object)reader;
            var logicalPositionProperty = readerObject.GetType().GetProperty(
                "LogicalPosition",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
            var seekForwardMethod = readerObject.GetType().GetMethod(
                "SeekForward",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                binder: null,
                types: new[] { typeof(long) },
                modifiers: null);

            for (var step = 0; step < 500; step++)
            {
                var positionBefore = ReadLongProperty(logicalPositionProperty, readerObject, stream.Position);
                var args = new object?[] { readerObject, null, 0, (ulong)0 };

                try
                {
                    readMethod.Invoke(null, args);
                }
                catch (Exception ex)
                {
                    writer.WriteLine(
                        string.Join(",",
                            row.Type,
                            row.StreamId,
                            row.BSN,
                            Csv(variantName),
                            step,
                            positionBefore,
                            ReadLongProperty(logicalPositionProperty, readerObject, stream.Position),
                            "",
                            "",
                            "",
                            "0",
                            Csv(GetUsefulExceptionMessage(ex))));

                    break;
                }

                var positionAfter = ReadLongProperty(logicalPositionProperty, readerObject, stream.Position);
                var headerFooterType = args[1]?.ToString() ?? "";
                var id = args[2]?.ToString() ?? "";
                var length = args[3]?.ToString() ?? "";
                var lengthValue = args[3] is ulong u ? u : 0;
                var remainingAfterHeader = content.Length - positionAfter;
                var lengthFits = lengthValue <= (ulong)Math.Max(0, remainingAfterHeader);

                writer.WriteLine(
                    string.Join(",",
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        Csv(variantName),
                        step,
                        positionBefore,
                        positionAfter,
                        Csv(headerFooterType),
                        Csv(id),
                        Csv(length),
                        lengthFits ? "1" : "0",
                        ""));

                if (positionAfter <= positionBefore)
                    break;

                if (lengthValue > 0 &&
                    lengthValue < int.MaxValue &&
                    lengthFits)
                {
                    if (seekForwardMethod != null)
                    {
                        seekForwardMethod.Invoke(readerObject, new object[] { (long)lengthValue });
                    }
                    else
                    {
                        stream.Position = positionAfter + (long)lengthValue;
                    }
                }

                var currentPosition = ReadLongProperty(logicalPositionProperty, readerObject, stream.Position);

                if (currentPosition >= content.Length)
                    break;
            }
        }
        catch (Exception ex)
        {
            writer.WriteLine(
                string.Join(",",
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    Csv(variantName),
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "0",
                    Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static long ReadLongProperty(PropertyInfo? property, object target, long fallback)
    {
        if (property == null)
            return fallback;

        try
        {
            return Convert.ToInt64(property.GetValue(target));
        }
        catch
        {
            return fallback;
        }
    }

    private void WriteCobaltFileHostBlobMapProbe(
        List<ShreddedStreamRow> rows,
        string csvPath,
        string detailsCsvPath,
        string workFolder)
    {
        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);
        using var detailsWriter = new StreamWriter(detailsCsvPath, false, Encoding.UTF8);

        writer.WriteLine("Type11StreamId,Type11BSN,Variant,BsnArgument,Success,FileBsn,MapBlobType,MapBlobLength,MapBlobSummary,MapBlobDumpPath,Error");
        detailsWriter.WriteLine("Type11StreamId,Type11BSN,Variant,BsnArgument,ObjectPath,MemberKind,Name,Type,Value");

        try
        {
            var allTypes = GetAllCobaltTypes();
            var mapType = allTypes.FirstOrDefault(t =>
                string.Equals(t.FullName, "Fda.SerializationElements.FileHostBlobMapElement", StringComparison.OrdinalIgnoreCase));

            if (mapType == null)
            {
                WriteFileHostBlobMapProbeError(writer, "binding", "FileHostBlobMapElement type was not found.");
                return;
            }

            var atomType = allTypes.FirstOrDefault(t =>
                string.Equals(t.FullName, "Cobalt.Base.IO.Atom", StringComparison.OrdinalIgnoreCase));

            if (atomType == null)
            {
                WriteFileHostBlobMapProbeError(writer, "binding", "Cobalt.Base.IO.Atom type was not found.");
                return;
            }

            var createAtom = atomType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                .FirstOrDefault(m =>
                    string.Equals(m.Name, "CreateFromArray", StringComparison.OrdinalIgnoreCase) &&
                    atomType.IsAssignableFrom(m.ReturnType) &&
                    m.GetParameters().Length == 1 &&
                    m.GetParameters()[0].ParameterType == typeof(byte[]));

            if (createAtom == null)
            {
                WriteFileHostBlobMapProbeError(writer, "binding", "Atom.CreateFromArray(byte[]) was not found.");
                return;
            }

            var readMethod = mapType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                .FirstOrDefault(m =>
                {
                    if (!string.Equals(m.Name, "Read", StringComparison.OrdinalIgnoreCase))
                        return false;

                    var parameters = m.GetParameters();

                    return parameters.Length == 2 &&
                           string.Equals(parameters[0].ParameterType.FullName, "Cobalt.Base.IO.Atom", StringComparison.OrdinalIgnoreCase) &&
                           parameters[1].ParameterType == typeof(ulong);
                });

            if (readMethod == null)
            {
                WriteFileHostBlobMapProbeError(writer, "binding", "FileHostBlobMapElement.Read(Atom, UInt64) was not found.");
                return;
            }

            var type10Bsns = rows
                .Where(x => x.Type == 10 && x.BSN >= 0)
                .Select(x => (ulong)x.BSN)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            foreach (var type11Row in rows
                         .Where(x => x.Type == 11)
                         .OrderBy(x => x.StreamId)
                         .ThenBy(x => x.BSN))
            {
                var variants = BuildCobaltRowProbeVariants(type11Row);
                var bsnArguments = new List<ulong> { 0, (ulong)Math.Max(0, type11Row.BSN) };

                bsnArguments.AddRange(type10Bsns);

                foreach (var variant in variants)
                {
                    foreach (var bsnArgument in bsnArguments.Distinct())
                    {
                        TryProbeFileHostBlobMapRow(
                            writer,
                            detailsWriter,
                            mapType,
                            createAtom,
                            readMethod,
                            type11Row,
                            variant.Name,
                            variant.Content,
                            bsnArgument,
                            workFolder);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            WriteFileHostBlobMapProbeError(writer, "fatal", GetUsefulExceptionMessage(ex));
        }
    }

    private static List<(string Name, byte[] Content)> BuildCobaltRowProbeVariants(ShreddedStreamRow row)
    {
        var variants = new List<(string Name, byte[] Content)>
        {
            ("full", row.Content)
        };

        foreach (var payloadStart in new[] { 0, 1, 4, 8, 16, 32, 64, 96 })
        {
            if (payloadStart >= row.Content.Length)
                continue;

            var payloadEnd = GetCobaltPayloadEnd(row.Content, payloadStart).PayloadEnd;

            if (payloadEnd > payloadStart)
            {
                variants.Add((
                    "slice_" + payloadStart + "_to_payload_end",
                    row.Content.Skip(payloadStart).Take(payloadEnd - payloadStart).ToArray()));
            }

            variants.Add((
                "slice_" + payloadStart + "_to_end",
                row.Content.Skip(payloadStart).ToArray()));
        }

        return variants
            .Where(x => x.Content.Length > 0)
            .GroupBy(x => Convert.ToBase64String(System.Security.Cryptography.SHA256.Create().ComputeHash(x.Content)))
            .Select(g => g.First())
            .ToList();
    }

    private static void TryProbeFileHostBlobMapRow(
        StreamWriter writer,
        StreamWriter detailsWriter,
        Type mapType,
        MethodInfo createAtom,
        MethodInfo readMethod,
        ShreddedStreamRow type11Row,
        string variantName,
        byte[] content,
        ulong bsnArgument,
        string workFolder)
    {
        try
        {
            var atom = createAtom.Invoke(null, new object[] { content });

            if (atom == null)
                throw new InvalidOperationException("Atom.CreateFromArray returned null.");

            var map = readMethod.Invoke(null, new[] { atom, (object)bsnArgument })
                      ?? throw new InvalidOperationException("FileHostBlobMapElement.Read returned null.");

            var fileBsn = ReadPropertyValue(map, "FileBsn");
            var mapBlob = ReadPropertyValue(map, "MapBlob");
            var mapBlobType = mapBlob?.GetType().FullName ?? "";
            var mapBlobLength = ReadLengthValue(mapBlob);
            var mapBlobSummary = BuildObjectSummary(mapBlob);
            var dumpPath = "";

            if (mapBlobLength.HasValue && mapBlobLength.Value > 0)
            {
                dumpPath = DumpCobaltAtomLikeObject(
                    mapBlob!,
                    workFolder,
                    $"file_host_blob_map_stream_{type11Row.StreamId}_bsn_{type11Row.BSN}_{MakeSafeProbeFileName(variantName)}_arg_{bsnArgument}.bin");
            }

            WriteProbeObjectDetails(
                detailsWriter,
                type11Row,
                variantName,
                bsnArgument,
                "FileHostBlobMapElement",
                map,
                maxMembers: 80);

            if (mapBlob != null)
            {
                WriteProbeObjectDetails(
                    detailsWriter,
                    type11Row,
                    variantName,
                    bsnArgument,
                    "FileHostBlobMapElement.MapBlob",
                    mapBlob,
                    maxMembers: 80);
            }

            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    Csv(variantName),
                    bsnArgument,
                    "1",
                    Csv(fileBsn),
                    Csv(mapBlobType),
                    mapBlobLength?.ToString() ?? "",
                    Csv(mapBlobSummary),
                    Csv(dumpPath),
                    ""));
        }
        catch (Exception ex)
        {
            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    Csv(variantName),
                    bsnArgument,
                    "0",
                    "",
                    "",
                    "",
                    "",
                    "",
                    Csv(GetUsefulExceptionMessage(ex))));
        }
    }

    private static void WriteFileHostBlobMapProbeError(StreamWriter writer, string stage, string error)
    {
        writer.WriteLine(
            string.Join(",",
                "",
                "",
                Csv(stage),
                "",
                "0",
                "",
                "",
                "",
                "",
                "",
                Csv(error)));
    }

    private static void WriteProbeObjectDetails(
        StreamWriter writer,
        ShreddedStreamRow type11Row,
        string variantName,
        ulong bsnArgument,
        string objectPath,
        object target,
        int maxMembers)
    {
        var written = 0;
        var type = target.GetType();

        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            if (!property.CanRead)
                continue;

            if (written >= maxMembers)
                return;

            object? value;

            try
            {
                value = property.GetValue(target);
            }
            catch (Exception ex)
            {
                value = "[read failed: " + GetUsefulExceptionMessage(ex) + "]";
            }

            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    Csv(variantName),
                    bsnArgument,
                    Csv(objectPath),
                    "Property",
                    Csv(property.Name),
                    Csv(property.PropertyType.FullName ?? property.PropertyType.Name),
                    Csv(BuildObjectSummary(value))));

            written++;
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            if (written >= maxMembers)
                return;

            object? value;

            try
            {
                value = field.GetValue(target);
            }
            catch (Exception ex)
            {
                value = "[read failed: " + GetUsefulExceptionMessage(ex) + "]";
            }

            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    Csv(variantName),
                    bsnArgument,
                    Csv(objectPath),
                    "Field",
                    Csv(field.Name),
                    Csv(field.FieldType.FullName ?? field.FieldType.Name),
                    Csv(BuildObjectSummary(value))));

            written++;
        }
    }

    private static string BuildObjectSummary(object? value)
    {
        if (value == null)
            return "";

        if (value is string text)
            return text;

        if (value is byte[] bytes)
            return "byte[" + bytes.Length + "] " + ToHex(bytes.Take(Math.Min(bytes.Length, 32)).ToArray());

        var type = value.GetType();

        if (type.IsPrimitive || type.IsEnum || value is decimal || value is Guid)
            return value.ToString() ?? "";

        var length = ReadLengthValue(value);
        var suffix = length.HasValue
            ? " Length=" + length.Value
            : "";

        return (type.FullName ?? type.Name) + suffix;
    }

    private static object? ReadPropertyValue(object? target, string propertyName)
    {
        if (target == null)
            return null;

        var property = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        return property == null
            ? null
            : property.GetValue(target);
    }

    private static long? ReadLengthValue(object? target)
    {
        if (target == null)
            return null;

        foreach (var propertyName in new[] { "Length", "Count", "Size" })
        {
            var value = ReadPropertyValue(target, propertyName);

            if (value == null)
                continue;

            try
            {
                return Convert.ToInt64(value);
            }
            catch
            {
            }
        }

        return null;
    }

    private static string DumpCobaltAtomLikeObject(
        object atomLikeObject,
        string workFolder,
        string fileName)
    {
        var outputPath = Path.Combine(workFolder, MakeSafeProbeFileName(fileName));

        var writeToFile = atomLikeObject.GetType().GetMethod(
            "WriteToFile",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: new[] { typeof(string) },
            modifiers: null);

        if (writeToFile != null)
        {
            writeToFile.Invoke(atomLikeObject, new object[] { outputPath });
            return outputPath;
        }

        var copyTo = atomLikeObject.GetType().GetMethod(
            "CopyTo",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: new[] { typeof(Stream) },
            modifiers: null);

        if (copyTo != null)
        {
            using (var output = File.Create(outputPath))
            {
                copyTo.Invoke(atomLikeObject, new object[] { output });
            }

            return outputPath;
        }

        return "";
    }

    private static string MakeSafeProbeFileName(string fileName)
    {
        foreach (var invalidChar in Path.GetInvalidFileNameChars())
        {
            fileName = fileName.Replace(invalidChar, '_');
        }

        return string.IsNullOrWhiteSpace(fileName)
            ? "cobalt_probe.bin"
            : fileName;
    }

    private static void WriteCobaltType11PairedReferences(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        var knownStreamIds = rows
            .Where(x => x.Type == 10)
            .Where(x => x.StreamId >= 0 && x.StreamId <= ushort.MaxValue)
            .Select(x => (int)x.StreamId)
            .ToHashSet();

        var knownBsns = rows
            .Where(x => x.Type == 10)
            .Where(x => x.BSN >= 0 && x.BSN <= ushort.MaxValue)
            .Select(x => (int)x.BSN)
            .ToHashSet();

        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine(
            "Type11StreamId,Type11BSN,WindowStart,ReferenceKind,FirstOffset,FirstValue,SecondOffset,SecondValue,Distance,WindowHex,Context");

        foreach (var type11Row in rows
                     .Where(x => x.Type == 11)
                     .OrderBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            WriteCobaltType11PairedReferencesForSet(
                writer,
                type11Row,
                knownStreamIds,
                "Type10StreamId");

            WriteCobaltType11PairedReferencesForSet(
                writer,
                type11Row,
                knownBsns,
                "Type10BSN");
        }
    }

    private static void WriteCobaltType11PairedReferencesForSet(
        StreamWriter writer,
        ShreddedStreamRow type11Row,
        HashSet<int> knownValues,
        string referenceKind)
    {
        var hits = new List<(int Offset, int Value)>();

        for (var offset = 0; offset + 1 < type11Row.Content.Length; offset++)
        {
            var value = BitConverter.ToUInt16(type11Row.Content, offset);

            if (knownValues.Contains(value))
                hits.Add((offset, value));
        }

        for (var i = 0; i + 1 < hits.Count; i++)
        {
            var first = hits[i];
            var second = hits[i + 1];
            var distance = second.Offset - first.Offset;

            if (distance <= 0 || distance > 96)
                continue;

            var windowStart = Math.Max(0, first.Offset - 32);
            var windowEnd = Math.Min(type11Row.Content.Length, second.Offset + 34);
            var window = type11Row.Content.Skip(windowStart).Take(windowEnd - windowStart).ToArray();

            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    windowStart,
                    Csv(referenceKind),
                    first.Offset,
                    first.Value,
                    second.Offset,
                    second.Value,
                    distance,
                    Csv(ToHex(window)),
                    Csv(GetAsciiPreview(type11Row.Content, windowStart, window.Length))));
        }
    }

    private static void WriteCobaltType11SequenceScan(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        var knownStreamIds = rows
            .Where(x => x.Type == 10)
            .Where(x => x.StreamId >= 0 && x.StreamId <= ushort.MaxValue)
            .Select(x => (int)x.StreamId)
            .ToHashSet();

        var knownBsns = rows
            .Where(x => x.Type == 10)
            .Where(x => x.BSN >= 0 && x.BSN <= ushort.MaxValue)
            .Select(x => (int)x.BSN)
            .ToHashSet();

        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine("Type11StreamId,Type11BSN,Offset,SequenceKind,Values,KnownValues,StrideBytes,Context");

        foreach (var type11Row in rows
                     .Where(x => x.Type == 11)
                     .OrderBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            WriteCobaltType11SequencesForSet(
                writer,
                type11Row,
                knownStreamIds,
                "Type10StreamId");

            WriteCobaltType11SequencesForSet(
                writer,
                type11Row,
                knownBsns,
                "Type10BSN");
        }
    }

    private static void WriteCobaltType11SequencesForSet(
        StreamWriter writer,
        ShreddedStreamRow type11Row,
        HashSet<int> knownValues,
        string sequenceKind)
    {
        for (var offset = 0; offset + 16 <= type11Row.Content.Length; offset += 2)
        {
            var values = new List<int>();
            var cursor = offset;

            while (cursor + 3 < type11Row.Content.Length)
            {
                var value = BitConverter.ToUInt16(type11Row.Content, cursor);
                var marker = BitConverter.ToUInt16(type11Row.Content, cursor + 2);

                if (marker != 0x2000)
                    break;

                values.Add(value);
                cursor += 4;
            }

            if (values.Count < 3)
                continue;

            var known = values
                .Where(knownValues.Contains)
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            if (known.Count < 2)
                continue;

            writer.WriteLine(
                string.Join(",",
                    type11Row.StreamId,
                    type11Row.BSN,
                    offset,
                    Csv(sequenceKind),
                    Csv(string.Join("|", values)),
                    Csv(string.Join("|", known)),
                    4,
                    Csv(GetAsciiPreview(type11Row.Content, Math.Max(0, offset - 24), Math.Min(160, type11Row.Content.Length - Math.Max(0, offset - 24))))));
        }
    }

    private static void WriteCobaltType11RecordScan(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        var knownStreamIds = rows
            .Where(x => x.Type == 10)
            .Where(x => x.StreamId >= 0 && x.StreamId <= ushort.MaxValue)
            .Select(x => (int)x.StreamId)
            .ToHashSet();

        var knownBsns = rows
            .Where(x => x.Type == 10)
            .Where(x => x.BSN >= 0 && x.BSN <= ushort.MaxValue)
            .Select(x => (int)x.BSN)
            .ToHashSet();

        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine(
            "Type11StreamId,Type11BSN,RecordOffset,U16_0,U16_2,U16_4,U16_6,U16_8,U16_10,U16_12,U16_14," +
            "U32_0,U32_4,U32_8,U32_12,U32_16,U32_20,U32_24,U32_28," +
            "KnownStreamIds,KnownBSNs,Context");

        foreach (var type11Row in rows
                     .Where(x => x.Type == 11)
                     .OrderBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            for (var offset = 0; offset + 32 <= type11Row.Content.Length; offset++)
            {
                var u16Values = new List<int>();

                for (var i = 0; i < 16; i += 2)
                    u16Values.Add(BitConverter.ToUInt16(type11Row.Content, offset + i));

                var matchingStreamIds = u16Values
                    .Where(knownStreamIds.Contains)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                var matchingBsns = u16Values
                    .Where(knownBsns.Contains)
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                if (matchingStreamIds.Count == 0 && matchingBsns.Count == 0)
                    continue;

                var u32Values = new List<uint>();

                for (var i = 0; i < 32; i += 4)
                    u32Values.Add(BitConverter.ToUInt32(type11Row.Content, offset + i));

                writer.WriteLine(
                    string.Join(",",
                        type11Row.StreamId,
                        type11Row.BSN,
                        offset,
                        u16Values[0],
                        u16Values[1],
                        u16Values[2],
                        u16Values[3],
                        u16Values[4],
                        u16Values[5],
                        u16Values[6],
                        u16Values[7],
                        u32Values[0],
                        u32Values[1],
                        u32Values[2],
                        u32Values[3],
                        u32Values[4],
                        u32Values[5],
                        u32Values[6],
                        u32Values[7],
                        Csv(string.Join("|", matchingStreamIds)),
                        Csv(string.Join("|", matchingBsns)),
                        Csv(GetAsciiPreview(type11Row.Content, Math.Max(0, offset - 24), 96))));
            }
        }
    }

    private static void WriteCobaltType11ReferenceMap(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        var dataRows = rows
            .Where(x => x.Type == 10)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ToList();

        using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine(
            "Type11StreamId,Type11BSN,ReferenceKind,ReferenceValue,U16Offsets,U32Offsets,Context");

        foreach (var type11Row in rows
                     .Where(x => x.Type == 11)
                     .OrderBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            foreach (var dataRow in dataRows)
            {
                WriteCobaltType11ReferenceRow(
                    writer,
                    type11Row,
                    "Type10StreamId",
                    dataRow.StreamId);

                WriteCobaltType11ReferenceRow(
                    writer,
                    type11Row,
                    "Type10BSN",
                    dataRow.BSN);
            }
        }
    }

    private static void WriteCobaltType11ReferenceRow(
        StreamWriter writer,
        ShreddedStreamRow type11Row,
        string referenceKind,
        long referenceValue)
    {
        if (referenceValue < 0 || referenceValue > int.MaxValue)
            return;

        var u16Offsets = FindLittleEndianUInt16Offsets(type11Row.Content, (int)referenceValue);
        var u32Offsets = FindLittleEndianUInt32Offsets(type11Row.Content, (int)referenceValue);

        if (u16Offsets.Count == 0 && u32Offsets.Count == 0)
            return;

        var firstOffset =
            u16Offsets.Count > 0
                ? u16Offsets[0]
                : u32Offsets[0];

        writer.WriteLine(
            string.Join(",",
                type11Row.StreamId,
                type11Row.BSN,
                Csv(referenceKind),
                referenceValue,
                Csv(string.Join("|", u16Offsets)),
                Csv(string.Join("|", u32Offsets)),
                Csv(GetAsciiPreview(type11Row.Content, Math.Max(0, firstOffset - 32), 128))));
    }

    private static List<int> FindLittleEndianUInt16Offsets(byte[] data, int value)
    {
        var result = new List<int>();

        if (value < 0 || value > ushort.MaxValue)
            return result;

        var b0 = (byte)(value & 0xFF);
        var b1 = (byte)((value >> 8) & 0xFF);

        for (var i = 0; i + 1 < data.Length; i++)
        {
            if (data[i] == b0 && data[i + 1] == b1)
                result.Add(i);
        }

        return result;
    }

    private static List<int> FindLittleEndianUInt32Offsets(byte[] data, int value)
    {
        var result = new List<int>();
        var b0 = (byte)(value & 0xFF);
        var b1 = (byte)((value >> 8) & 0xFF);
        var b2 = (byte)((value >> 16) & 0xFF);
        var b3 = (byte)((value >> 24) & 0xFF);

        for (var i = 0; i + 3 < data.Length; i++)
        {
            if (data[i] == b0 &&
                data[i + 1] == b1 &&
                data[i + 2] == b2 &&
                data[i + 3] == b3)
            {
                result.Add(i);
            }
        }

        return result;
    }

    private static void WriteCobaltRowEnvelopeFields(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine(
            "HistVersion,Type,StreamId,BSN,ContentLength,PayloadStart,PayloadEnd,PayloadLength,TrailerDetected," +
            "U16_0,U16_2,U16_4,U16_6,U16_8,U16_10,U16_12,U16_14,U16_16,U16_18,U16_20,U16_22," +
            "U32_6,U32_8,U32_14,U32_16,U32_18,U32_38,U32_42,U32_58,U32_62,U32_74,U32_78,U32_90,U32_92," +
            "PdfOffset,ObjectNumbers,FirstPayload64Hex,TrailerHex");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            var payloadStart = row.Content.Length > 96 ? 96 : 0;
            var payloadEndInfo = GetCobaltPayloadEnd(row.Content, payloadStart);
            var payloadEnd = payloadEndInfo.PayloadEnd;
            var payloadLength = Math.Max(0, payloadEnd - payloadStart);
            var objectNumbers = FindPdfStructuralMarkers(row.Content, payloadStart, payloadEnd)
                .Where(x => x.Kind == "obj" && x.ObjectNumber.HasValue)
                .Select(x => x.ObjectNumber!.Value.ToString())
                .ToList();

            writer.WriteLine(
                string.Join(",",
                    row.HistVersion,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    row.Content.Length,
                    payloadStart,
                    payloadEnd,
                    payloadLength,
                    payloadEndInfo.TrailerDetected ? "1" : "0",
                    ReadUInt16OrEmpty(row.Content, 0),
                    ReadUInt16OrEmpty(row.Content, 2),
                    ReadUInt16OrEmpty(row.Content, 4),
                    ReadUInt16OrEmpty(row.Content, 6),
                    ReadUInt16OrEmpty(row.Content, 8),
                    ReadUInt16OrEmpty(row.Content, 10),
                    ReadUInt16OrEmpty(row.Content, 12),
                    ReadUInt16OrEmpty(row.Content, 14),
                    ReadUInt16OrEmpty(row.Content, 16),
                    ReadUInt16OrEmpty(row.Content, 18),
                    ReadUInt16OrEmpty(row.Content, 20),
                    ReadUInt16OrEmpty(row.Content, 22),
                    ReadUInt32OrEmpty(row.Content, 6),
                    ReadUInt32OrEmpty(row.Content, 8),
                    ReadUInt32OrEmpty(row.Content, 14),
                    ReadUInt32OrEmpty(row.Content, 16),
                    ReadUInt32OrEmpty(row.Content, 18),
                    ReadUInt32OrEmpty(row.Content, 38),
                    ReadUInt32OrEmpty(row.Content, 42),
                    ReadUInt32OrEmpty(row.Content, 58),
                    ReadUInt32OrEmpty(row.Content, 62),
                    ReadUInt32OrEmpty(row.Content, 74),
                    ReadUInt32OrEmpty(row.Content, 78),
                    ReadUInt32OrEmpty(row.Content, 90),
                    ReadUInt32OrEmpty(row.Content, 92),
                    IndexOfBytes(row.Content, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }),
                    Csv(string.Join(" ", objectNumbers)),
                    Csv(ToHex(row.Content.Skip(payloadStart).Take(Math.Min(64, payloadLength)).ToArray())),
                    Csv(payloadEnd < row.Content.Length ? ToHex(row.Content.Skip(payloadEnd).Take(Math.Min(28, row.Content.Length - payloadEnd)).ToArray()) : "")));
        }
    }

    private static void WriteCobaltFragmentMap(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine(
            "HistVersion,Type,StreamId,BSN,ContentLength," +
            "PayloadStart,PayloadEnd,PayloadLength,TrailerDetected," +
            "PrefixFirst16Hex,TrailerHex," +
            "U32_6,U32_8,U32_14,U32_16,U32_18,U32_38,U32_42,U32_58,U32_62,U32_74,U32_78,U32_90,U32_92," +
            "PdfOffset,XrefOffset,TrailerOffset,StartXrefOffset,EofOffset," +
            "MarkerKind,ObjectNumber,RawOffset,PayloadOffset,NextRawOffset,SpanLength,Snippet");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            const int payloadStartCandidate = 96;

            var payloadStart = row.Content.Length > payloadStartCandidate
                ? payloadStartCandidate
                : 0;

            var payloadEndInfo = GetCobaltPayloadEnd(row.Content, payloadStart);
            var payloadEnd = payloadEndInfo.PayloadEnd;
            var payloadLength = Math.Max(0, payloadEnd - payloadStart);

            var markers = FindPdfStructuralMarkers(row.Content, payloadStart, payloadEnd)
                .OrderBy(x => x.RawOffset)
                .ThenBy(x => x.Kind)
                .ToList();

            if (markers.Count == 0)
            {
                WriteCobaltFragmentMapRow(
                    writer,
                    row,
                    payloadStart,
                    payloadEnd,
                    payloadLength,
                    payloadEndInfo.TrailerDetected,
                    "row",
                    null,
                    -1,
                    -1,
                    -1,
                    -1,
                    "");

                continue;
            }

            for (var i = 0; i < markers.Count; i++)
            {
                var marker = markers[i];
                var nextRawOffset = i + 1 < markers.Count
                    ? markers[i + 1].RawOffset
                    : payloadEnd;

                WriteCobaltFragmentMapRow(
                    writer,
                    row,
                    payloadStart,
                    payloadEnd,
                    payloadLength,
                    payloadEndInfo.TrailerDetected,
                    marker.Kind,
                    marker.ObjectNumber,
                    marker.RawOffset,
                    marker.RawOffset - payloadStart,
                    nextRawOffset,
                    Math.Max(0, nextRawOffset - marker.RawOffset),
                    GetAsciiPreview(row.Content, marker.RawOffset, Math.Min(96, Math.Max(0, nextRawOffset - marker.RawOffset))));
            }
        }
    }

    private static void WriteCobaltFragmentMapRow(
        StreamWriter writer,
        ShreddedStreamRow row,
        int payloadStart,
        int payloadEnd,
        int payloadLength,
        bool trailerDetected,
        string markerKind,
        int? objectNumber,
        int rawOffset,
        int payloadOffset,
        int nextRawOffset,
        int spanLength,
        string snippet)
    {
        var content = row.Content;

        writer.WriteLine(
            string.Join(",",
                row.HistVersion,
                row.Type,
                row.StreamId,
                row.BSN,
                content.Length,
                payloadStart,
                payloadEnd,
                payloadLength,
                trailerDetected ? "1" : "0",
                Csv(ToHex(content.Take(Math.Min(16, content.Length)).ToArray())),
                Csv(payloadEnd < content.Length ? ToHex(content.Skip(payloadEnd).Take(Math.Min(28, content.Length - payloadEnd)).ToArray()) : ""),
                ReadUInt32OrEmpty(content, 6),
                ReadUInt32OrEmpty(content, 8),
                ReadUInt32OrEmpty(content, 14),
                ReadUInt32OrEmpty(content, 16),
                ReadUInt32OrEmpty(content, 18),
                ReadUInt32OrEmpty(content, 38),
                ReadUInt32OrEmpty(content, 42),
                ReadUInt32OrEmpty(content, 58),
                ReadUInt32OrEmpty(content, 62),
                ReadUInt32OrEmpty(content, 74),
                ReadUInt32OrEmpty(content, 78),
                ReadUInt32OrEmpty(content, 90),
                ReadUInt32OrEmpty(content, 92),
                IndexOfBytes(content, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }),
                IndexOfBytes(content, System.Text.Encoding.ASCII.GetBytes("xref")),
                IndexOfBytes(content, System.Text.Encoding.ASCII.GetBytes("trailer")),
                IndexOfBytes(content, System.Text.Encoding.ASCII.GetBytes("startxref")),
                LastIndexOfBytes(content, System.Text.Encoding.ASCII.GetBytes("%%EOF")),
                Csv(markerKind),
                objectNumber?.ToString() ?? "",
                rawOffset,
                payloadOffset,
                nextRawOffset,
                spanLength,
                Csv(snippet)));
    }

    private static CobaltPayloadEndInfo GetCobaltPayloadEnd(byte[] content, int payloadStart)
    {
        const int cobaltTrailerLength = 28;
        var suffixMarker = new byte[] { 0x15, 0xA4, 0x01 };
        var suffixOffset = content.Length - cobaltTrailerLength;

        if (suffixOffset <= payloadStart ||
            suffixOffset < 0 ||
            suffixOffset + suffixMarker.Length > content.Length)
        {
            return new CobaltPayloadEndInfo(content.Length, false);
        }

        for (var i = 0; i < suffixMarker.Length; i++)
        {
            if (content[suffixOffset + i] != suffixMarker[i])
                return new CobaltPayloadEndInfo(content.Length, false);
        }

        return new CobaltPayloadEndInfo(suffixOffset, true);
    }

    private static List<PdfStructuralMarker> FindPdfStructuralMarkers(
        byte[] content,
        int payloadStart,
        int payloadEnd)
    {
        var markers = new List<PdfStructuralMarker>();
        var length = Math.Max(0, payloadEnd - payloadStart);

        if (length == 0)
            return markers;

        var text = System.Text.Encoding.ASCII.GetString(content, payloadStart, length);

        foreach (System.Text.RegularExpressions.Match match in
                 System.Text.RegularExpressions.Regex.Matches(text, @"(?m)(\d+)\s+0\s+obj"))
        {
            markers.Add(new PdfStructuralMarker(
                "obj",
                int.Parse(match.Groups[1].Value),
                payloadStart + match.Index));
        }

        foreach (var keyword in new[] { "stream", "endstream", "endobj", "xref", "trailer", "startxref", "%%EOF", "%PDF-" })
        {
            var offset = 0;

            while (offset < text.Length)
            {
                var index = text.IndexOf(keyword, offset, StringComparison.Ordinal);

                if (index < 0)
                    break;

                markers.Add(new PdfStructuralMarker(keyword, null, payloadStart + index));
                offset = index + keyword.Length;
            }
        }

        return markers;
    }

    private static string ReadUInt32OrEmpty(byte[] content, int offset)
    {
        return offset >= 0 && offset + sizeof(uint) <= content.Length
            ? BitConverter.ToUInt32(content, offset).ToString()
            : "";
    }

    private static string ReadUInt16OrEmpty(byte[] content, int offset)
    {
        return offset >= 0 && offset + sizeof(ushort) <= content.Length
            ? BitConverter.ToUInt16(content, offset).ToString()
            : "";
    }

    private static string GetAsciiPreview(byte[] content, int offset, int length)
    {
        if (offset < 0 || offset >= content.Length || length <= 0)
            return "";

        var count = Math.Min(length, content.Length - offset);
        var chars = new char[count];

        for (var i = 0; i < count; i++)
        {
            var value = content[offset + i];

            chars[i] = value >= 32 && value <= 126
                ? (char)value
                : '.';
        }

        return new string(chars);
    }

    private sealed class CobaltPayloadEndInfo
    {
        public CobaltPayloadEndInfo(int payloadEnd, bool trailerDetected)
        {
            PayloadEnd = payloadEnd;
            TrailerDetected = trailerDetected;
        }

        public int PayloadEnd { get; }

        public bool TrailerDetected { get; }
    }

    private sealed class PdfStructuralMarker
    {
        public PdfStructuralMarker(string kind, int? objectNumber, int rawOffset)
        {
            Kind = kind;
            ObjectNumber = objectNumber;
            RawOffset = rawOffset;
        }

        public string Kind { get; }

        public int? ObjectNumber { get; }

        public int RawOffset { get; }
    }

    private static void WriteFsshttpbHexSummary(
    List<ShreddedStreamRow> rows,
    string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine("HistVersion,Type,StreamId,BSN,ContentLength,First64Hex,Last32Hex,PdfOffset,ZipOffset");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            var first64 = row.Content.Take(64).ToArray();

            var last32 = row.Content.Length <= 32
                ? row.Content
                : row.Content.Skip(row.Content.Length - 32).Take(32).ToArray();

            var pdfOffset = IndexOfBytes(row.Content, new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D }); // %PDF-
            var zipOffset = IndexOfBytes(row.Content, new byte[] { 0x50, 0x4B, 0x03, 0x04 });       // PK..

            writer.WriteLine(
                string.Join(",",
                    row.HistVersion,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    row.Content.Length,
                    Csv(ToHex(first64)),
                    Csv(ToHex(last32)),
                    pdfOffset,
                    zipOffset));
        }
    }

    private static void WriteFsshttpbSequentialParse(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, System.Text.Encoding.UTF8);

        writer.WriteLine("HistVersion,Type,StreamId,BSN,ContentLength,StartOffset,Offset,Depth,HeaderKind,HeaderHex,ObjectType,ObjectName,Compound,Length,DataPreviewHex,StackAfter");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            var startOffset = FindLikelyCobaltStartOffset(row.Content);

            if (startOffset < 0)
            {
                writer.WriteLine(
                    string.Join(",",
                        row.HistVersion,
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        row.Content.Length,
                        -1,
                        -1,
                        0,
                        Csv("NoKnownStart"),
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        ""));

                continue;
            }

            ParseSequentialHeadersForRow(row, startOffset, writer);
        }
    }

    private static void WriteFsshttpbObjectGraph(
        List<ShreddedStreamRow> rows,
        string csvPath)
    {
        using var writer = new StreamWriter(csvPath, false, Encoding.UTF8);

        writer.WriteLine(
            "HistVersion,Type,StreamId,BSN,ContentLength,StartOffset,Offset,Depth,HeaderKind,ObjectType,ObjectName," +
            "Compound,Length,DataStart,DataEnd,ParentPath,DecodedKind,DecodedGuid,DecodedValue,ChunkStart,ChunkLength,PreviewHex");

        foreach (var row in rows
                     .OrderBy(x => x.Type)
                     .ThenBy(x => x.StreamId)
                     .ThenBy(x => x.BSN))
        {
            var startOffset = FindLikelyCobaltStartOffset(row.Content);

            if (startOffset < 0)
            {
                writer.WriteLine(
                    string.Join(",",
                        row.HistVersion,
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        row.Content.Length,
                        -1,
                        -1,
                        0,
                        Csv("NoKnownStart"),
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        "",
                        ""));

                continue;
            }

            ParseFsshttpbObjectGraphForRow(row, startOffset, writer);
        }
    }

    private static void ParseFsshttpbObjectGraphForRow(
        ShreddedStreamRow row,
        int startOffset,
        StreamWriter writer)
    {
        var data = row.Content;
        var offset = startOffset;
        var stack = new Stack<string>();
        var maxObjects = 1000;
        var count = 0;

        while (offset < data.Length && count < maxObjects)
        {
            count++;

            var decoded = TryDecodeSequentialHeader(data, offset);

            if (decoded == null)
            {
                offset++;
                continue;
            }

            var dataStart = offset + decoded.HeaderLength;
            var dataLength = Math.Max(0, Math.Min(decoded.Length, data.Length - dataStart));
            var dataEnd = dataStart + dataLength;
            var decodedPayload = DecodeFsshttpbObjectPayload(decoded, data, dataStart, dataLength);
            var parentPath = string.Join(">", stack.Reverse());

            writer.WriteLine(
                string.Join(",",
                    row.HistVersion,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    row.Content.Length,
                    startOffset,
                    offset,
                    stack.Count,
                    Csv(decoded.Kind),
                    Csv(decoded.ObjectType),
                    Csv(decoded.ObjectName),
                    decoded.Compound ? "1" : "0",
                    decoded.Length,
                    decoded.IsEnd ? "" : dataStart.ToString(),
                    decoded.IsEnd ? "" : dataEnd.ToString(),
                    Csv(parentPath),
                    Csv(decodedPayload.Kind),
                    Csv(decodedPayload.GuidText),
                    Csv(decodedPayload.ValueText),
                    decodedPayload.ChunkStart?.ToString() ?? "",
                    decodedPayload.ChunkLength?.ToString() ?? "",
                    Csv(dataLength > 0
                        ? ToHex(data.Skip(dataStart).Take(Math.Min(dataLength, 32)).ToArray())
                        : "")));

            if (decoded.IsEnd)
            {
                if (stack.Count > 0)
                    stack.Pop();

                offset += decoded.HeaderLength;
                continue;
            }

            offset += decoded.HeaderLength;

            if (decoded.Length > 0)
                offset += decoded.Length;

            if (decoded.Compound)
                stack.Push(decoded.ObjectName);
        }
    }

    private static FsshttpbPayloadDecodeResult DecodeFsshttpbObjectPayload(
        SequentialHeaderInfo header,
        byte[] data,
        int dataStart,
        int dataLength)
    {
        if (header.IsEnd || dataLength <= 0)
            return FsshttpbPayloadDecodeResult.Empty;

        if (string.Equals(header.ObjectName, "SpecializedKnowledge", StringComparison.OrdinalIgnoreCase) &&
            dataLength >= 16)
        {
            var guidBytes = new byte[16];
            Buffer.BlockCopy(data, dataStart, guidBytes, 0, guidBytes.Length);
            var guid = new Guid(guidBytes);

            return new FsshttpbPayloadDecodeResult
            {
                Kind = "SpecializedKnowledge",
                GuidText = guid.ToString("D"),
                ValueText = GetSpecializedKnowledgeName(guid)
            };
        }

        if (string.Equals(header.ObjectName, "FragmentKnowledgeEntry", StringComparison.OrdinalIgnoreCase))
        {
            if (TryDecodeFragmentKnowledgeEntry(data, dataStart, dataLength, out var entry))
                return entry;

            return new FsshttpbPayloadDecodeResult
            {
                Kind = "FragmentKnowledgeEntry",
                ValueText = "decode_failed"
            };
        }

        if (string.Equals(header.ObjectName, "DataElementFragment", StringComparison.OrdinalIgnoreCase))
        {
            if (TryDecodeDataElementFragment(data, dataStart, dataLength, out var fragment))
                return fragment;

            return new FsshttpbPayloadDecodeResult
            {
                Kind = "DataElementFragment",
                ValueText = "decode_failed"
            };
        }

        return FsshttpbPayloadDecodeResult.Empty;
    }

    private static bool TryDecodeFragmentKnowledgeEntry(
        byte[] data,
        int offset,
        int length,
        out FsshttpbPayloadDecodeResult result)
    {
        result = FsshttpbPayloadDecodeResult.Empty;
        var cursor = offset;
        var end = offset + length;

        if (!TryReadExtendedGuid(data, ref cursor, end, out var extendedGuid))
            return false;

        if (!TryReadCompactUInt64(data, cursor, out var dataElementSize, out var sizeBytes))
            return false;

        cursor += sizeBytes;

        if (!TryReadFileChunkReference(data, ref cursor, end, out var chunkStart, out var chunkLength))
            return false;

        result = new FsshttpbPayloadDecodeResult
        {
            Kind = "FragmentKnowledgeEntry",
            GuidText = extendedGuid,
            ValueText = "DataElementSize=" + dataElementSize,
            ChunkStart = (long)chunkStart,
            ChunkLength = (long)chunkLength
        };

        return true;
    }

    private static bool TryDecodeDataElementFragment(
        byte[] data,
        int offset,
        int length,
        out FsshttpbPayloadDecodeResult result)
    {
        result = FsshttpbPayloadDecodeResult.Empty;
        var cursor = offset;
        var end = offset + length;

        if (!TryReadExtendedGuid(data, ref cursor, end, out var fragmentGuid))
            return false;

        if (!TryReadCompactUInt64(data, cursor, out var dataElementSize, out var sizeBytes))
            return false;

        cursor += sizeBytes;

        if (!TryReadFileChunkReference(data, ref cursor, end, out var chunkStart, out var chunkLength))
            return false;

        result = new FsshttpbPayloadDecodeResult
        {
            Kind = "DataElementFragment",
            GuidText = fragmentGuid,
            ValueText = "DataElementSize=" + dataElementSize + ";FragmentBytesRemaining=" + Math.Max(0, end - cursor),
            ChunkStart = (long)chunkStart,
            ChunkLength = (long)chunkLength
        };

        return true;
    }

    private static bool TryReadFileChunkReference(
        byte[] data,
        ref int cursor,
        int end,
        out ulong start,
        out ulong length)
    {
        start = 0;
        length = 0;

        if (cursor >= end ||
            !TryReadCompactUInt64(data, cursor, out start, out var startBytes))
        {
            return false;
        }

        cursor += startBytes;

        if (cursor >= end ||
            !TryReadCompactUInt64(data, cursor, out length, out var lengthBytes))
        {
            return false;
        }

        cursor += lengthBytes;
        return true;
    }

    private static bool TryReadExtendedGuid(
        byte[] data,
        ref int cursor,
        int end,
        out string value)
    {
        value = "";

        if (cursor >= end)
            return false;

        var first = data[cursor];

        if (first == 0)
        {
            cursor++;
            value = "00000000-0000-0000-0000-000000000000:0";
            return true;
        }

        int integerValue;
        int guidOffset;
        int bytesConsumed;

        if ((first & 0x07) == 0x04)
        {
            integerValue = first >> 3;
            guidOffset = cursor + 1;
            bytesConsumed = 17;
        }
        else if ((first & 0x3F) == 0x20)
        {
            if (cursor + 1 >= end)
                return false;

            integerValue = ((data[cursor + 1] << 2) | (first >> 6));
            guidOffset = cursor + 2;
            bytesConsumed = 18;
        }
        else if ((first & 0x7F) == 0x40)
        {
            if (cursor + 2 >= end)
                return false;

            integerValue =
                (data[cursor + 1] << 1) |
                ((data[cursor + 2] & 0xFF) << 9) |
                (first >> 7);

            guidOffset = cursor + 3;
            bytesConsumed = 19;
        }
        else if (first == 0x80)
        {
            if (cursor + 4 >= end)
                return false;

            integerValue = BitConverter.ToInt32(data, cursor + 1);
            guidOffset = cursor + 5;
            bytesConsumed = 21;
        }
        else
        {
            return false;
        }

        if (cursor + bytesConsumed > end || guidOffset + 16 > end)
            return false;

        var guidBytes = new byte[16];
        Buffer.BlockCopy(data, guidOffset, guidBytes, 0, guidBytes.Length);
        cursor += bytesConsumed;
        value = new Guid(guidBytes).ToString("D") + ":" + integerValue;
        return true;
    }

    private static string GetSpecializedKnowledgeName(Guid guid)
    {
        if (guid == new Guid("327a35f6-0761-4414-9686-51e900667a4d"))
            return "CellKnowledge";

        if (guid == new Guid("3a76e90e-8032-4d0c-b9dd-f3c65029433e"))
            return "WaterlineKnowledge";

        if (guid == new Guid("0abe4f35-01df-4134-a24a-7c79f0859844"))
            return "FragmentKnowledge";

        if (guid == new Guid("10091f13-c882-40fb-9886-6533f934c21d"))
            return "ContentTagKnowledge";

        if (guid == new Guid("bf12e2c1-e64f-4959-8282-73b9a24a7c44"))
            return "VersionTokenKnowledge";

        return "UnknownSpecializedKnowledge";
    }

    private static void ParseSequentialHeadersForRow(
    ShreddedStreamRow row,
    int startOffset,
    StreamWriter writer)
    {
        var data = row.Content;
        var offset = startOffset;
        var stack = new Stack<string>();

        var maxObjects = 500;
        var count = 0;

        while (offset < data.Length && count < maxObjects)
        {
            count++;

            var decoded = TryDecodeSequentialHeader(data, offset);

            if (decoded == null)
            {
                writer.WriteLine(
                    string.Join(",",
                        row.HistVersion,
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        row.Content.Length,
                        startOffset,
                        offset,
                        stack.Count,
                        Csv("UnknownByte"),
                        Csv(data[offset].ToString("X2")),
                        "",
                        "",
                        "",
                        "",
                        "",
                        Csv(string.Join(">", stack.Reverse()))));

                offset++;
                continue;
            }

            var dataStart = offset + decoded.HeaderLength;
            var dataLength = Math.Max(0, Math.Min(decoded.Length, data.Length - dataStart));

            var preview = dataLength > 0
                ? ToHex(data.Skip(dataStart).Take(Math.Min(dataLength, 16)).ToArray())
                : "";

            if (decoded.IsEnd)
            {
                if (stack.Count > 0)
                    stack.Pop();

                writer.WriteLine(
                    string.Join(",",
                        row.HistVersion,
                        row.Type,
                        row.StreamId,
                        row.BSN,
                        row.Content.Length,
                        startOffset,
                        offset,
                        stack.Count,
                        Csv(decoded.Kind),
                        Csv(decoded.HeaderHex),
                        Csv(decoded.ObjectType),
                        Csv(decoded.ObjectName),
                        "",
                        "",
                        "",
                        Csv(string.Join(">", stack.Reverse()))));

                offset += decoded.HeaderLength;
                continue;
            }

            writer.WriteLine(
                string.Join(",",
                    row.HistVersion,
                    row.Type,
                    row.StreamId,
                    row.BSN,
                    row.Content.Length,
                    startOffset,
                    offset,
                    stack.Count,
                    Csv(decoded.Kind),
                    Csv(decoded.HeaderHex),
                    Csv(decoded.ObjectType),
                    Csv(decoded.ObjectName),
                    decoded.Compound ? "1" : "0",
                    decoded.Length,
                    Csv(preview),
                    Csv(string.Join(">", stack.Reverse()))));

            offset += decoded.HeaderLength;

            if (decoded.Length > 0)
                offset += decoded.Length;

            if (decoded.Compound)
                stack.Push(decoded.ObjectName);
        }
    }

    private static int FindLikelyCobaltStartOffset(byte[] data)
    {
        // 0x0104 appears in the CSV as HeaderHex 0104.
        // In little-endian bytes, this is 04 01.
        for (var i = 0; i + 1 < data.Length; i++)
        {
            if (data[i] == 0x04 && data[i + 1] == 0x01)
                return i;
        }

        // Data Element Package Start is commonly 0x02AC.
        // In little-endian bytes, this is AC 02.
        for (var i = 0; i + 1 < data.Length; i++)
        {
            if (data[i] == 0xAC && data[i + 1] == 0x02)
                return i;
        }

        return -1;
    }


    private static SequentialHeaderInfo? TryDecodeSequentialHeader(byte[] data, int offset)
    {
        if (offset >= data.Length)
            return null;

        var b0 = data[offset];

        // Known one-byte end markers from Cobalt/FSSHTTPD-style structures.
        if (b0 == 0x81)
        {
            return new SequentialHeaderInfo
            {
                IsEnd = true,
                HeaderLength = 1,
                Kind = "End8",
                HeaderHex = "81",
                ObjectType = "0x20",
                ObjectName = "RootNodeEnd"
            };
        }

        if (b0 == 0x7D)
        {
            return new SequentialHeaderInfo
            {
                IsEnd = true,
                HeaderLength = 1,
                Kind = "End8",
                HeaderHex = "7D",
                ObjectType = "0x1F",
                ObjectName = "IntermediateOrLeafNodeEnd"
            };
        }

        // Generic 8-bit stream object end.
        if ((b0 & 0x03) == 1)
        {
            var type = (b0 >> 2) & 0x3F;

            return new SequentialHeaderInfo
            {
                IsEnd = true,
                HeaderLength = 1,
                Kind = "End8",
                HeaderHex = b0.ToString("X2"),
                ObjectType = "0x" + type.ToString("X2"),
                ObjectName = GetFsshttpbObjectName(type)
            };
        }

        if (offset + 1 >= data.Length)
            return null;

        var value16 = BitConverter.ToUInt16(data, offset);

        // 16-bit stream object start. Header type low bits = 0.
        if ((value16 & 0x0003) == 0)
        {
            var compound = ((value16 >> 2) & 0x01) == 1;
            var type = (value16 >> 3) & 0x3F;
            var length = (value16 >> 9) & 0x7F;

            return new SequentialHeaderInfo
            {
                IsEnd = false,
                HeaderLength = 2,
                Kind = "Start16",
                HeaderHex = value16.ToString("X4"),
                ObjectType = "0x" + type.ToString("X2"),
                ObjectName = GetFsshttpbObjectName(type, value16),
                Compound = compound,
                Length = length
            };
        }

        if ((value16 & 0x0003) == 3)
        {
            var type = (value16 >> 2) & 0x3FFF;

            return new SequentialHeaderInfo
            {
                IsEnd = true,
                HeaderLength = 2,
                Kind = "End16",
                HeaderHex = value16.ToString("X4"),
                ObjectType = "0x" + type.ToString("X3"),
                ObjectName = GetFsshttpbObjectName(type)
            };
        }

        // 32-bit stream object start. Low two bits are 0x2; bit 2 is Compound;
        // bits 3..16 are Type; bits 17..31 are Length, with 32767 meaning
        // an additional compact unsigned 64-bit length follows.
        if ((value16 & 0x0003) == 2 && offset + 3 < data.Length)
        {
            var value32 = BitConverter.ToUInt32(data, offset);
            var compound = ((value32 >> 2) & 0x01) == 1;
            var type = (int)((value32 >> 3) & 0x3FFF);
            var length = (int)((value32 >> 17) & 0x7FFF);
            var headerLength = 4;
            long decodedLength = length;

            if (length == 0x7FFF)
            {
                if (!TryReadCompactUInt64(data, offset + 4, out var largeLength, out var bytesRead))
                    return null;

                headerLength += bytesRead;
                decodedLength = largeLength > int.MaxValue
                    ? int.MaxValue
                    : (long)largeLength;
            }

            return new SequentialHeaderInfo
            {
                IsEnd = false,
                HeaderLength = headerLength,
                Kind = "Start32",
                HeaderHex = value32.ToString("X8"),
                ObjectType = "0x" + type.ToString("X3"),
                ObjectName = GetFsshttpbObjectName(type),
                Compound = compound,
                Length = decodedLength > int.MaxValue ? int.MaxValue : (int)decodedLength
            };
        }

        return null;
    }


    private static string GetFsshttpbObjectName(int type, ushort? raw16 = null)
    {
        if (raw16.HasValue)
        {
            switch (raw16.Value)
            {
                case 0x0104:
                    return "RootNodeStart";

                case 0x00FC:
                    return "IntermediateOrLeafNodeStart";

                case 0x02AC:
                    return "DataElementPackageStart";
            }
        }

        switch (type)
        {
            case 0x01:
                return "DataElement";

            case 0x02:
                return "ObjectDataBlob";

            case 0x15:
                return "DataElementPackage";

            case 0x16:
                return "ObjectGroupData";

            case 0x19:
                return "RevisionManifestObjectGroupReference";

            case 0x1A:
                return "RevisionManifest";

            case 0x1E:
                return "ObjectGroup";

            case 0x20:
                return "RootNode";

            case 0x40:
                return "Request";

            case 0x41:
                return "SubResponse";

            case 0x42:
                return "SubRequest";

            case 0x43:
                return "ReadAccessResponse";

            case 0x44:
                return "SpecializedKnowledge";

            case 0x46:
                return "WriteAccessResponse";

            case 0x5A:
                return "PutChangesRequest";

            case 0x62:
                return "Response";

            case 0x6A:
                return "DataElementFragment";

            case 0x6B:
                return "FragmentKnowledge";

            case 0x6C:
                return "FragmentKnowledgeEntry";

            case 0x78:
                return "ObjectGroupMetadata";

            case 0x79:
                return "ObjectGroupMetadataDeclarations";

            case 0x8C:
                return "VersionTokenKnowledge";

            case 0x8E:
                return "FileHash";

            default:
                return "Type_" + type.ToString("X3");
        }
    }

    private static bool TryReadCompactUInt64(
        byte[] data,
        int offset,
        out ulong value,
        out int bytesRead)
    {
        value = 0;
        bytesRead = 0;

        if (offset < 0 || offset >= data.Length)
            return false;

        var shift = 0;

        for (var i = offset; i < data.Length && bytesRead < 10; i++)
        {
            var b = data[i];
            value |= ((ulong)(b & 0x7F)) << shift;
            bytesRead++;

            if ((b & 0x80) == 0)
                return true;

            shift += 7;

            if (shift >= 64)
                return false;
        }

        return false;
    }

    private sealed class SequentialHeaderInfo
    {
        public bool IsEnd { get; set; }

        public int HeaderLength { get; set; }

        public string Kind { get; set; } = "";

        public string HeaderHex { get; set; } = "";

        public string ObjectType { get; set; } = "";

        public string ObjectName { get; set; } = "";

        public bool Compound { get; set; }

        public int Length { get; set; }
    }

    private sealed class FsshttpbPayloadDecodeResult
    {
        public static readonly FsshttpbPayloadDecodeResult Empty = new();

        public string Kind { get; set; } = "";
        public string GuidText { get; set; } = "";
        public string ValueText { get; set; } = "";
        public long? ChunkStart { get; set; }
        public long? ChunkLength { get; set; }
    }

    private static int IndexOfBytes(byte[] data, byte[] pattern)
    {
        return IndexOfBytes(data, pattern, 0);
    }

    private static int IndexOfBytes(byte[] data, byte[] pattern, int startOffset)
    {
        if (pattern.Length == 0 || data.Length < pattern.Length)
            return -1;

        startOffset = Math.Max(0, startOffset);

        for (var i = startOffset; i <= data.Length - pattern.Length; i++)
        {
            var match = true;

            for (var j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
                return i;
        }

        return -1;
    }

    private static int LastIndexOfBytes(byte[] data, byte[] pattern)
    {
        if (pattern.Length == 0 || data.Length < pattern.Length)
            return -1;

        for (var i = data.Length - pattern.Length; i >= 0; i--)
        {
            var match = true;

            for (var j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j])
                {
                    match = false;
                    break;
                }
            }

            if (match)
                return i;
        }

        return -1;
    }

    private static string ToHex(byte[] data)
    {
        return BitConverter.ToString(data).Replace("-", "");
    }

    private static FsshttpbHeaderInfo? TryDecodeFsshttpbHeader(byte[] data, int offset)
    {
        if (offset >= data.Length)
            return null;

        var b0 = data[offset];
        var headerType = b0 & 0x03;

        // 8-bit stream object header end.
        // Low 2 bits = 1. Type is bits 2..7.
        if (headerType == 1)
        {
            var type = (b0 >> 2) & 0x3F;

            return new FsshttpbHeaderInfo
            {
                Offset = offset,
                BytesConsumed = 1,
                Kind = "End8",
                HeaderHex = b0.ToString("X2"),
                DecodedType = "0x" + type.ToString("X2"),
                Compound = "",
                LengthOrInfo = ""
            };
        }

        if (offset + 1 >= data.Length)
            return null;

        var value16 = BitConverter.ToUInt16(data, offset);

        // 16-bit stream object header start.
        // Low 2 bits = 0.
        if ((value16 & 0x0003) == 0)
        {
            var compound = ((value16 >> 2) & 0x01) == 1;
            var type = (value16 >> 3) & 0x3F;
            var length = (value16 >> 9) & 0x7F;

            return new FsshttpbHeaderInfo
            {
                Offset = offset,
                BytesConsumed = 2,
                Kind = "Start16",
                HeaderHex = value16.ToString("X4"),
                DecodedType = "0x" + type.ToString("X2"),
                Compound = compound ? "1" : "0",
                LengthOrInfo = length.ToString()
            };
        }

        // 16-bit stream object header end.
        // Low 2 bits = 3. Type is bits 2..15.
        if ((value16 & 0x0003) == 3)
        {
            var type = (value16 >> 2) & 0x3FFF;

            return new FsshttpbHeaderInfo
            {
                Offset = offset,
                BytesConsumed = 2,
                Kind = "End16",
                HeaderHex = value16.ToString("X4"),
                DecodedType = "0x" + type.ToString("X3"),
                Compound = "",
                LengthOrInfo = ""
            };
        }

        // 32-bit stream object header start.
        // Low 2 bits = 2. Bit 2 = compound. Type is bits 3..16.
        // Length is bits 17..31; 32767 means a compact UInt64 length follows.
        if ((value16 & 0x0003) == 2 && offset + 3 < data.Length)
        {
            var value32 = BitConverter.ToUInt32(data, offset);
            var compound = ((value32 >> 2) & 0x01) == 1;
            var type = (int)((value32 >> 3) & 0x3FFF);
            var length = (int)((value32 >> 17) & 0x7FFF);
            var bytesConsumed = 4;
            var lengthText = length.ToString();

            if (length == 0x7FFF &&
                TryReadCompactUInt64(data, offset + 4, out var largeLength, out var largeLengthBytes))
            {
                bytesConsumed += largeLengthBytes;
                lengthText = largeLength.ToString();
            }

            return new FsshttpbHeaderInfo
            {
                Offset = offset,
                BytesConsumed = bytesConsumed,
                Kind = "Start32",
                HeaderHex = value32.ToString("X8"),
                DecodedType = "0x" + type.ToString("X3"),
                Compound = compound ? "1" : "0",
                LengthOrInfo = lengthText
            };
        }

        return null;
    }

    private static string Csv(object? value)
    {
        var text = value?.ToString() ?? "";

        if (text.Contains("\""))
            text = text.Replace("\"", "\"\"");

        if (text.Contains(",") || text.Contains("\"") || text.Contains("\r") || text.Contains("\n"))
            return "\"" + text + "\"";

        return text;
    }

    private sealed class FsshttpbHeaderInfo
    {
        public int Offset { get; set; }

        public int BytesConsumed { get; set; }

        public string Kind { get; set; } = "";

        public string HeaderHex { get; set; } = "";

        public string DecodedType { get; set; } = "";

        public string Compound { get; set; } = "";

        public string LengthOrInfo { get; set; } = "";
    }

    private ShreddedExportResult TryExportViaHostBlobStoreEndpoint(
    CobaltCoreBinding binding,
    List<ShreddedStreamRow> rows,
    DocumentItemInfo item,
    string outputPath,
    string workFolder)
    {
        var strictItemGroupResult = AnalyzeFsshttpbRows(
            rows,
            item,
            outputPath,
            workFolder);

        if (strictItemGroupResult != null)
            return strictItemGroupResult;

        return new ShreddedExportResult
        {
            Success = false,
            OutputPath = outputPath,
            Message =
                "Microsoft.CobaltCore.dll was loaded, but its HostBlobStore import APIs throw Cobalt.MissingCodeException. " +
                "This means DocStreams rows cannot be imported through LocalHostBlobStore/PutBlob/GetFilename. " +
                "The next required step is direct FSSHTTPB/Cobalt stream-object parsing. " +
                "The cobalt work folder now contains row envelope, fragment map, and sequential parse CSVs for that parser work."
        };
    }

    private void PreseedPersistedLocalHostBlobStoreFolder(
    CobaltCoreBinding binding,
    List<ShreddedStreamRow> rows,
    BlobIdStrategy strategy,
    string strategyFolder)
    {
        if (binding.PersistedLocalHostBlobStoreType == null)
            throw new InvalidOperationException("PersistedLocalHostBlobStore type was not found.");

        Directory.CreateDirectory(strategyFolder);

        foreach (var oldFile in Directory.GetFiles(strategyFolder))
            File.Delete(oldFile);

        var getFilename = binding.PersistedLocalHostBlobStoreType.GetMethod(
            "GetFilename",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);

        if (getFilename == null)
            throw new InvalidOperationException("PersistedLocalHostBlobStore.GetFilename(...) was not found.");

        var blobChangeFrequency = CreatePreferredEnumValue(
            binding.BlobChangeFrequencyType,
            new[] { "Low", "Normal", "Unknown", "Default" });

        var blobTag = CreatePreferredEnumValue(
            binding.BlobTagType,
            new[] { "Nil", "None", "Default", "Unknown" });

        var blobCommitType = CreatePreferredEnumValue(
            binding.BlobCommitTypeType,
            new[] { "Committed", "Normal", "Default", "None" });

        if (blobChangeFrequency == null)
            throw new InvalidOperationException("Could not create BlobChangeFrequency value.");

        if (blobTag == null)
            throw new InvalidOperationException("Could not create BlobTag value.");

        if (blobCommitType == null)
            throw new InvalidOperationException("Could not create BlobCommitType value.");

        AnsiConsole.MarkupLine(
            $"[grey]Preseed BlobChangeFrequency:[/] [yellow]{Markup.Escape(blobChangeFrequency.ToString() ?? "")}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Preseed BlobTag:[/] [yellow]{Markup.Escape(blobTag.ToString() ?? "")}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]Preseed BlobCommitType:[/] [yellow]{Markup.Escape(blobCommitType.ToString() ?? "")}[/]");

        var type10Rows = rows
            .Where(x => x.Type == 10)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ToList();

        if (type10Rows.Count == 0)
            throw new InvalidOperationException("No Type 10 rows found.");

        var written = 0;
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in type10Rows)
        {
            var blobId = strategy.GetBlobId(row);
            var bsn = (ulong)row.BSN;

            var filenameObj = getFilename.Invoke(null, new object?[]
            {
            blobId,
            bsn,
            blobChangeFrequency,
            blobTag,
            blobCommitType
            });

            if (filenameObj == null)
                throw new InvalidOperationException("GetFilename returned null.");

            var filename = filenameObj.ToString();

            if (string.IsNullOrWhiteSpace(filename))
                throw new InvalidOperationException("GetFilename returned empty filename.");

            var fullPath = Path.Combine(strategyFolder, filename);

            if (!seenNames.Add(fullPath))
            {
                AnsiConsole.MarkupLine(
                    $"[yellow]Skipping duplicate persisted blob filename:[/] {Markup.Escape(filename)}");
                continue;
            }

            File.WriteAllBytes(fullPath, row.Content);

            AnsiConsole.MarkupLine(
                $"[grey]Preseed blob file:[/] blobId={blobId}, bsn={bsn}, streamId={row.StreamId}, bytes={row.Content.Length:N0}, file={Markup.Escape(filename)}");

            written++;
        }

        AnsiConsole.MarkupLine(
            $"[green]Preseeded persisted local blob files:[/] {written:N0}");
    }

    private object CreatePersistedHostBlobStore(
    CobaltCoreBinding binding,
    string folder,
    bool loadExisting)
    {
        Directory.CreateDirectory(folder);

        var filePath = CreateFilePath(binding, folder);

        Exception? persistedFailure = null;

        if (binding.PersistedLocalHostBlobStoreType != null)
        {
            try
            {
                foreach (var ctor in binding.PersistedLocalHostBlobStoreType.GetConstructors(
                             BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
                {
                    var p = ctor.GetParameters();

                    if (p.Length != 4)
                        continue;

                    AnsiConsole.MarkupLine("[grey]Trying PersistedLocalHostBlobStore constructor...[/]");

                    var args = new object?[4];

                    args[0] = filePath;
                    args[1] = CreateDefaultConfigObject(p[1].ParameterType);
                    args[2] = false; // synchronize
                    args[3] = loadExisting;

                    return ctor.Invoke(args);
                }
            }
            catch (Exception ex)
            {
                persistedFailure = ex;

                AnsiConsole.MarkupLine(
                    "[yellow]PersistedLocalHostBlobStore failed:[/] " +
                    Markup.Escape(GetUsefulExceptionMessage(ex)));
            }
        }

        if (binding.LocalHostBlobStoreType != null)
        {
            foreach (var ctor in binding.LocalHostBlobStoreType.GetConstructors(
                         BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                var p = ctor.GetParameters();

                // LocalHostBlobStore(Config config, FilePath dirPathForFileBackedBlobs, Boolean synchronize, HostBlobStoreIds hostBlobStoreIds)
                if (p.Length != 4)
                    continue;

                if (p[1].ParameterType != binding.FilePathType)
                    continue;

                AnsiConsole.MarkupLine("[grey]Trying LocalHostBlobStore constructor...[/]");

                var args = new object?[4];

                args[0] = CreateDefaultConfigObject(p[0].ParameterType);
                args[1] = filePath;
                args[2] = false; // synchronize
                args[3] = CreateHostBlobStoreIds(binding);

                return ctor.Invoke(args);
            }
        }

        if (persistedFailure != null)
        {
            throw new InvalidOperationException(
                "Could not create PersistedLocalHostBlobStore or LocalHostBlobStore. Persisted failure: " +
                GetUsefulExceptionMessage(persistedFailure),
                persistedFailure);
        }

        throw new InvalidOperationException("No usable LocalHostBlobStore/PersistedLocalHostBlobStore constructor was found.");
    }

    private void LoadRowsIntoHostBlobStore(
        CobaltCoreBinding binding,
        object hostBlobStore,
        List<ShreddedStreamRow> rows,
        BlobIdStrategy strategy,
        string strategyFolder)
    {
        var transactionContext = Stage(
            "CreateTransactionContextValue",
            () => CreatePreferredEnumValue(
                binding.TransactionContextType,
                new[] { "None", "Default", "Normal" }));

        if (transactionContext == null)
            throw new InvalidOperationException("Could not create TransactionContext value.");

        var blobTag = Stage(
            "CreateBlobTagValue",
            () => CreatePreferredEnumValue(
                binding.BlobTagType,
                new[] { "None", "Default", "Unknown" }));

        if (blobTag == null)
            throw new InvalidOperationException("Could not create BlobTag value.");

        AnsiConsole.MarkupLine(
            $"[grey]TransactionContext:[/] [yellow]{Markup.Escape(transactionContext.ToString() ?? "")}[/]");

        AnsiConsole.MarkupLine(
            $"[grey]BlobTag:[/] [yellow]{Markup.Escape(blobTag.ToString() ?? "")}[/]");

        var directPutBlob = hostBlobStore.GetType()
            .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(m =>
            {
                if (!m.Name.Equals("PutBlob", StringComparison.OrdinalIgnoreCase))
                    return false;

                var p = m.GetParameters();

                return p.Length == 4 &&
                       p[0].ParameterType == typeof(ulong) &&
                       binding.AtomType != null &&
                       p[1].ParameterType.IsAssignableFrom(binding.AtomType) &&
                       binding.BlobTagType != null &&
                       p[2].ParameterType == binding.BlobTagType &&
                       binding.TransactionContextType != null &&
                       p[3].ParameterType == binding.TransactionContextType;
            });

        if (directPutBlob == null)
        {
            directPutBlob = hostBlobStore.GetType()
                .GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                .FirstOrDefault(m =>
                {
                    if (!m.Name.Equals("PutBlob", StringComparison.OrdinalIgnoreCase))
                        return false;

                    var p = m.GetParameters();

                    return p.Length == 4 &&
                           p[0].ParameterType == typeof(ulong);
                });
        }

        if (directPutBlob == null)
            throw new InvalidOperationException("HostBlobStore.PutBlob(UInt64, Atom, BlobTag, TransactionContext) was not found.");

        var type10Rows = rows
            .Where(x => x.Type == 10)
            .OrderBy(x => x.StreamId)
            .ThenBy(x => x.BSN)
            .ToList();

        if (type10Rows.Count == 0)
            throw new InvalidOperationException("No Type 10 DocStreams rows found. Cannot load file data blobs.");

        var seenBlobIds = new HashSet<ulong>();

        foreach (var row in type10Rows)
        {
            var blobId = strategy.GetBlobId(row);

            if (!seenBlobIds.Add(blobId))
            {
                AnsiConsole.MarkupLine(
                    $"[yellow]Skipping duplicate blobId:[/] {blobId} [grey](streamId={row.StreamId}, bsn={row.BSN})[/]");
                continue;
            }

            var atom = CreateAtomFromBytes(binding, row.Content);

            try
            {
                AnsiConsole.MarkupLine(
                    $"[grey]HostBlobStore.PutBlob:[/] blobId={blobId}, type={row.Type}, streamId={row.StreamId}, bsn={row.BSN}, bytes={row.Content.Length:N0}");

                directPutBlob.Invoke(hostBlobStore, new[]
                {
                (object)blobId,
                atom,
                blobTag,
                transactionContext
            });
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"HostBlobStore.PutBlob failed. blobId={blobId}, type={row.Type}, streamId={row.StreamId}, bsn={row.BSN}, bytes={row.Content.Length:N0}",
                    ex);
            }
        }

        AnsiConsole.MarkupLine(
            $"[green]Loaded Type 10 rows into HostBlobStore:[/] {seenBlobIds.Count:N0}");
    }
    private object CreateFileAtomFromExisting(CobaltCoreBinding binding, string path)
    {
        if (binding.FileAtomType == null)
            throw new InvalidOperationException("FileAtom type was not found.");

        var fromExisting = binding.FileAtomType.GetMethod(
            "FromExisting",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            types: new[] { typeof(string) },
            modifiers: null);

        if (fromExisting == null)
            throw new InvalidOperationException("FileAtom.FromExisting(string) was not found.");

        var atom = fromExisting.Invoke(null, new object[] { path });

        if (atom == null)
            throw new InvalidOperationException("FileAtom.FromExisting returned null.");

        return atom;
    }

    private object CreateAtomFromBytes(CobaltCoreBinding binding, byte[] data)
    {
        if (binding.AtomType == null)
            throw new InvalidOperationException("Atom type was not found.");

        object? bytesObject = null;

        if (binding.BytesType != null)
            bytesObject = TryCreateBytesObject(binding.BytesType, data);

        var atomType = binding.AtomType;

        var candidateTypes = GetAllCobaltTypes()
            .Where(t =>
                t != null &&
                t != atomType &&
                !t.IsAbstract &&
                atomType.IsAssignableFrom(t))
            .Where(t =>
                // Very important: FileAtom is not usable here.
                // It causes Cobalt.MissingCodeException during PutBlob.
                t.FullName == null ||
                t.FullName.IndexOf("FileAtom", StringComparison.OrdinalIgnoreCase) < 0)
            .OrderBy(t =>
                (t.FullName ?? "").IndexOf("Byte", StringComparison.OrdinalIgnoreCase) >= 0 ? 0 :
                (t.FullName ?? "").IndexOf("Memory", StringComparison.OrdinalIgnoreCase) >= 0 ? 1 :
                2)
            .ThenBy(t => t.FullName)
            .ToList();

        var failures = new List<string>();

        // Static factory methods on concrete Atom-derived types
        foreach (var type in candidateTypes)
        {
            foreach (var method in type.GetMethods(
                         BindingFlags.Public |
                         BindingFlags.NonPublic |
                         BindingFlags.Static))
            {
                if (!atomType.IsAssignableFrom(method.ReturnType))
                    continue;

                var parameters = method.GetParameters();

                if (!TryBuildAtomFactoryArgumentsStrict(parameters, data, bytesObject, out var args))
                    continue;

                try
                {
                    var result = method.Invoke(null, args);

                    if (result != null)
                    {
                        AnsiConsole.MarkupLine(
                            $"[grey]Created byte-backed Atom using:[/] [yellow]{Markup.Escape(type.FullName ?? type.Name)}.{Markup.Escape(method.Name)}[/]");

                        return result;
                    }
                }
                catch (Exception ex)
                {
                    failures.Add(
                        $"{type.FullName}.{method.Name}: {GetUsefulExceptionMessage(ex)}");
                }
            }
        }

        // Constructors on concrete Atom-derived types
        foreach (var type in candidateTypes)
        {
            foreach (var ctor in type.GetConstructors(
                         BindingFlags.Public |
                         BindingFlags.NonPublic |
                         BindingFlags.Instance))
            {
                var parameters = ctor.GetParameters();

                if (!TryBuildAtomFactoryArgumentsStrict(parameters, data, bytesObject, out var args))
                    continue;

                try
                {
                    var result = ctor.Invoke(args);

                    if (result != null)
                    {
                        AnsiConsole.MarkupLine(
                            $"[grey]Created byte-backed Atom using constructor:[/] [yellow]{Markup.Escape(FormatConstructor(ctor))}[/]");

                        return result;
                    }
                }
                catch (Exception ex)
                {
                    failures.Add(
                        $"{type.FullName} constructor: {GetUsefulExceptionMessage(ex)}");
                }
            }
        }

        var message =
            "Could not create a byte-backed Atom from byte[]. " +
            "The previous FileAtom path is blocked because FileAtom.FromExisting causes Cobalt.MissingCodeException.";

        if (failures.Count > 0)
        {
            message += Environment.NewLine +
                       string.Join(Environment.NewLine, failures.Take(20));
        }

        throw new InvalidOperationException(message);
    }

    private static bool TryBuildAtomFactoryArgumentsStrict(
        ParameterInfo[] parameters,
        byte[] data,
        object? bytesObject,
        out object?[] args)
    {
        args = new object?[parameters.Length];

        if (parameters.Length == 0)
            return false;

        var usedBinarySource = false;

        for (var i = 0; i < parameters.Length; i++)
        {
            var type = parameters[i].ParameterType;

            if (type == typeof(byte[]))
            {
                args[i] = data;
                usedBinarySource = true;
                continue;
            }

            if (type == typeof(Stream))
            {
                args[i] = new MemoryStream(data);
                usedBinarySource = true;
                continue;
            }

            if (bytesObject != null && type.IsInstanceOfType(bytesObject))
            {
                args[i] = bytesObject;
                usedBinarySource = true;
                continue;
            }

            if (type == typeof(bool))
            {
                args[i] = false;
                continue;
            }

            if (type == typeof(int))
            {
                args[i] = data.Length;
                continue;
            }

            if (type == typeof(long))
            {
                args[i] = (long)data.Length;
                continue;
            }

            if (type == typeof(uint))
            {
                args[i] = (uint)data.Length;
                continue;
            }

            if (type == typeof(ulong))
            {
                args[i] = (ulong)data.Length;
                continue;
            }

            if (type.IsEnum)
            {
                var values = Enum.GetValues(type);

                args[i] = values.Length > 0
                    ? values.GetValue(0)
                    : Activator.CreateInstance(type);

                continue;
            }

            if (type.IsValueType)
            {
                args[i] = Activator.CreateInstance(type);
                continue;
            }

            // Important:
            // Do NOT allow string parameters here.
            // FileAtom.FromExisting(string) matched before because of this.
            // We only want factories/constructors that consume byte[], Stream, or Bytes.
            return false;
        }

        return usedBinarySource;
    }

    private static object? TryCreateBytesObject(Type bytesType, byte[] data)
    {
        foreach (var method in bytesType.GetMethods(
                     BindingFlags.Public |
                     BindingFlags.NonPublic |
                     BindingFlags.Static))
        {
            if (!bytesType.IsAssignableFrom(method.ReturnType))
                continue;

            var parameters = method.GetParameters();

            if (parameters.Length != 1)
                continue;

            if (parameters[0].ParameterType != typeof(byte[]))
                continue;

            try
            {
                var result = method.Invoke(null, new object[] { data });

                if (result != null)
                    return result;
            }
            catch
            {
            }
        }

        foreach (var ctor in bytesType.GetConstructors(
                     BindingFlags.Public |
                     BindingFlags.NonPublic |
                     BindingFlags.Instance))
        {
            var parameters = ctor.GetParameters();

            if (parameters.Length != 1)
                continue;

            if (parameters[0].ParameterType != typeof(byte[]))
                continue;

            try
            {
                return ctor.Invoke(new object[] { data });
            }
            catch
            {
            }
        }

        return null;
    }

    private object CreateDisposalEscrow(CobaltCoreBinding binding)
    {
        if (binding.DisposalEscrowType == null)
            throw new InvalidOperationException("DisposalEscrow type was not found.");

        return Activator.CreateInstance(binding.DisposalEscrowType)
               ?? throw new InvalidOperationException("Could not create DisposalEscrow.");
    }

    private Delegate CreateConstantDelegateForHostBlobStore(object hostBlobStore)
    {
        var delegateTypes = GetAllCobaltTypes()
            .Where(t =>
                typeof(MulticastDelegate).IsAssignableFrom(t.BaseType ?? typeof(object)) &&
                t.Name.IndexOf("CreateHostBlobStoreForPartition", StringComparison.OrdinalIgnoreCase) >= 0)
            .ToList();

        if (delegateTypes.Count == 0)
            throw new InvalidOperationException("CreateHostBlobStoreForPartition delegate type was not found.");

        return CreateConstantDelegate(delegateTypes[0], hostBlobStore);
    }

    private object CreateHostMetadataStore(
        CobaltCoreBinding binding,
        Delegate getStoreDelegate,
        object disposal)
    {
        if (binding.HostMetadataStoreOnHostBlobStoreType == null)
            throw new InvalidOperationException("HostMetadataStoreOnHostBlobStore type was not found.");

        var ctor = binding.HostMetadataStoreOnHostBlobStoreType.GetConstructors(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(c => c.GetParameters().Length == 2);

        if (ctor == null)
            throw new InvalidOperationException("HostMetadataStoreOnHostBlobStore constructor was not found.");

        return ctor.Invoke(new object[]
        {
        getStoreDelegate,
        disposal
        });
    }

    private object CreateHostLockingStore(
        CobaltCoreBinding binding,
        Delegate getStoreDelegate,
        object docArrayGuid,
        object metadataStore,
        object disposal)
    {
        if (binding.HostLockingStoreOnHostBlobStoreType == null)
            throw new InvalidOperationException("HostLockingStoreOnHostBlobStore type was not found.");

        var ctor = binding.HostLockingStoreOnHostBlobStoreType.GetConstructors(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(c => c.GetParameters().Length == 4);

        if (ctor == null)
            throw new InvalidOperationException("HostLockingStoreOnHostBlobStore constructor was not found.");

        return ctor.Invoke(new object[]
        {
        getStoreDelegate,
        docArrayGuid,
        metadataStore,
        disposal
        });
    }

    private object CreateEndpointImpl(
        CobaltCoreBinding binding,
        object disposal,
        Delegate getStoreDelegate,
        object hostBlobStore,
        object lockingStore,
        object metadataStore)
    {
        if (binding.EndpointImplType == null)
            throw new InvalidOperationException("EndpointUtils+EndpointImpl type was not found.");

        var constructors = binding.EndpointImplType.GetConstructors(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        foreach (var ctor in constructors)
        {
            var p = ctor.GetParameters();

            // We want the overload:
            // EndpointImpl(DisposalEscrow, CreateHostBlobStoreForPartition, IEnumerable<FilePartitionId>,
            //              BlobStore, CellStorageConfig, CellStorageMetrics, HostLockingStore, HostMetadataStore)
            if (p.Length != 8)
                continue;

            var filePartitionId = CreateKnownFilePartitionId(binding);
            var supportedPartitions = CreateTypedArray(filePartitionId.GetType(), filePartitionId);

            var args = new object?[8];

            args[0] = disposal;
            args[1] = getStoreDelegate;
            args[2] = supportedPartitions;
            args[3] = null; // BlobStore/MRT store
            args[4] = CreateCellStorageConfig(binding);
            args[5] = CreateDefaultValue(p[5].ParameterType); // CellStorageMetrics
            args[6] = lockingStore;
            args[7] = metadataStore;

            return ctor.Invoke(args);
        }

        throw new InvalidOperationException("Usable EndpointUtils+EndpointImpl constructor was not found.");
    }

    private object CreateGenericFda(
        CobaltCoreBinding binding,
        object endpoint)
    {
        if (binding.GenericFdaType == null)
            throw new InvalidOperationException("GenericFda type was not found.");

        var config = CreateGenericFdaConfig(binding);

        foreach (var ctor in binding.GenericFdaType.GetConstructors(
                     BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            var p = ctor.GetParameters();

            if (p.Length == 2 &&
                p[0].ParameterType.IsAssignableFrom(endpoint.GetType()))
            {
                return ctor.Invoke(new[]
                {
                endpoint,
                config
            });
            }
        }

        throw new InvalidOperationException("GenericFda(ICobaltGraphEndpoint, Config) compatible constructor was not found.");
    }

    private object CreateGenericFdaConfig(CobaltCoreBinding binding)
    {
        if (binding.GenericFdaConfigType == null)
            throw new InvalidOperationException("GenericFda.Config type was not found.");

        var config = Activator.CreateInstance(binding.GenericFdaConfigType)
                     ?? throw new InvalidOperationException("Could not create GenericFda.Config.");

        SetPropertyIfExists(config, "AllowGraphEndpoint", true);
        SetPropertyIfExists(config, "ComputeDataHashes", false);
        SetPropertyIfExists(config, "DelayLoadCells", false);
        SetPropertyIfExists(config, "DisableTopologyFetch", false);
        SetPropertyIfExists(config, "SingleUseOnFailure", true);

        return config;
    }

    private object CreateCellStorageConfig(CobaltCoreBinding binding)
    {
        if (binding.CellStorageConfigType == null)
            throw new InvalidOperationException("CellStorageConfig type was not found.");

        var config = Activator.CreateInstance(binding.CellStorageConfigType)
                     ?? throw new InvalidOperationException("Could not create CellStorageConfig.");

        SetPropertyIfExists(config, "ComputeDataHashes", false);
        SetPropertyIfExists(config, "DisableCorruptionChecks", false);
        SetPropertyIfExists(config, "O15SharePointHackMainStoreChangeNotPersisted", true);

        return config;
    }

    private void ExportGenericFdaContentStream(
        CobaltCoreBinding binding,
        object genericFda,
        string outputPath)
    {
        var getContentStream = binding.GenericFdaType!
            .GetMethod("GetContentStream", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        if (getContentStream == null)
            throw new InvalidOperationException("GenericFda.GetContentStream was not found.");

        var streamObj = getContentStream.Invoke(genericFda, Array.Empty<object>());

        if (streamObj == null)
            throw new InvalidOperationException("GenericFda.GetContentStream returned null.");

        var writeToFile = streamObj.GetType().GetMethod(
            "WriteToFile",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: new[] { typeof(string) },
            modifiers: null);

        if (writeToFile != null)
        {
            writeToFile.Invoke(streamObj, new object[] { outputPath });
            return;
        }

        var copyTo = streamObj.GetType().GetMethod(
            "CopyTo",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: new[] { typeof(Stream) },
            modifiers: null);

        if (copyTo == null)
            throw new InvalidOperationException("GenericFdaStream has neither WriteToFile(string) nor CopyTo(Stream).");

        using (var output = File.Create(outputPath))
        {
            copyTo.Invoke(streamObj, new object[] { output });
        }
    }

    private object CreateFilePath(CobaltCoreBinding binding, string path)
    {
        if (binding.FilePathType == null)
            throw new InvalidOperationException("FilePath type was not found.");

        Directory.CreateDirectory(path);

        var getFolder = binding.FilePathType.GetMethod(
            "GetFolder",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            types: new[] { typeof(string) },
            modifiers: null);

        if (getFolder != null)
        {
            var result = getFolder.Invoke(null, new object[] { path });

            if (result != null)
                return result;
        }

        throw new InvalidOperationException("FilePath.GetFolder(string) was not found or returned null.");
    }

    private object CreateArrayGuid(CobaltCoreBinding binding, Guid guid)
    {
        if (binding.ArrayGuidType == null)
            throw new InvalidOperationException("ArrayGuid type was not found.");

        var parse = binding.ArrayGuidType.GetMethod(
            "Parse",
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            types: new[] { typeof(string) },
            modifiers: null);

        if (parse == null)
            throw new InvalidOperationException("ArrayGuid.Parse(string) was not found.");

        var result = parse.Invoke(null, new object[] { guid.ToString("D") });

        if (result == null)
            throw new InvalidOperationException("ArrayGuid.Parse returned null.");

        return result;
    }

    private object CreateKnownFilePartitionId(CobaltCoreBinding binding)
    {
        if (binding.FilePartitionIdType == null)
            throw new InvalidOperationException("FilePartitionId type was not found.");

        var knownType = binding.FilePartitionIdType.GetNestedType(
            "Known",
            BindingFlags.Public | BindingFlags.NonPublic);

        if (knownType == null || !knownType.IsEnum)
            throw new InvalidOperationException("FilePartitionId.Known enum was not found.");

        var values = Enum.GetValues(knownType).Cast<object>().ToList();

        if (values.Count == 0)
            throw new InvalidOperationException("FilePartitionId.Known enum has no values.");

        // Prefer common names if present.
        var preferred =
            values.FirstOrDefault(v => v.ToString()!.IndexOf("Content", StringComparison.OrdinalIgnoreCase) >= 0) ??
            values.FirstOrDefault(v => v.ToString()!.IndexOf("Default", StringComparison.OrdinalIgnoreCase) >= 0) ??
            values[0];

        var ctor = binding.FilePartitionIdType.GetConstructor(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: new[] { knownType },
            modifiers: null);

        if (ctor == null)
            throw new InvalidOperationException("FilePartitionId(Known) constructor was not found.");

        return ctor.Invoke(new[] { preferred });
    }

    private static object CreateTypedArray(Type elementType, params object[] values)
    {
        var array = Array.CreateInstance(elementType, values.Length);

        for (var i = 0; i < values.Length; i++)
            array.SetValue(values[i], i);

        return array;
    }

    private static object? CreateDefaultValue(Type type)
    {
        if (type == typeof(string))
            return "";

        if (type.IsEnum)
        {
            var values = Enum.GetValues(type);

            return values.Length > 0
                ? values.GetValue(0)
                : Activator.CreateInstance(type);
        }

        if (type.IsValueType)
            return Activator.CreateInstance(type);

        var defaultCtor = type.GetConstructor(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: Type.EmptyTypes,
            modifiers: null);

        if (defaultCtor != null && !type.IsAbstract)
            return defaultCtor.Invoke(Array.Empty<object>());

        return null;
    }

    private static object CreateObjectWithoutDefaultConstructor(Type type)
    {
        if (type.IsValueType)
            return Activator.CreateInstance(type)!;

        var ctor = type.GetConstructor(
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            binder: null,
            types: Type.EmptyTypes,
            modifiers: null);

        if (ctor != null)
            return ctor.Invoke(Array.Empty<object>());

        return FormatterServices.GetUninitializedObject(type);
    }

    private static object? CreateDefaultConfigObject(Type type)
    {
        if (!type.IsAbstract)
        {
            var ctor = type.GetConstructor(
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                binder: null,
                types: Type.EmptyTypes,
                modifiers: null);

            if (ctor != null)
                return ctor.Invoke(Array.Empty<object>());
        }

        return null;
    }

    private static void SetPropertyIfExists(object target, string propertyName, object? value)
    {
        var prop = target.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        if (prop == null || !prop.CanWrite)
            return;

        prop.SetValue(target, value);
    }

    private static Delegate CreateConstantDelegate(Type delegateType, object returnValue)
    {
        var invoke = delegateType.GetMethod("Invoke");

        if (invoke == null)
            throw new InvalidOperationException("Delegate Invoke method was not found: " + delegateType.FullName);

        var parameters = invoke.GetParameters()
            .Select(p => Expression.Parameter(p.ParameterType, p.Name))
            .ToArray();

        var body = Expression.Convert(
            Expression.Constant(returnValue),
            invoke.ReturnType);

        return Expression.Lambda(delegateType, body, parameters).Compile();
    }

    private static string GetUsefulExceptionMessage(Exception ex)
    {
        var parts = new List<string>();

        var current = ex;

        while (current != null)
        {
            parts.Add(current.GetType().Name + ": " + current.Message);
            current = current.InnerException;
        }

        return string.Join(" -> ", parts);
    }

    private sealed class BlobIdStrategy
    {
        public BlobIdStrategy(string name, Func<ShreddedStreamRow, ulong> getBlobId)
        {
            Name = name;
            GetBlobId = getBlobId;
        }

        public string Name { get; }

        public Func<ShreddedStreamRow, ulong> GetBlobId { get; }
    }
    private void DumpDerivedTypes(CobaltCoreBinding binding)
    {
        var allTypes = GetAllCobaltTypes();

        DumpConcreteDerivedTypes(allTypes, binding.AtomType, "Concrete Atom implementations");
        DumpConcreteDerivedTypes(allTypes, binding.HostFileType, "Concrete HostFile implementations");
        DumpConcreteDerivedTypes(allTypes, binding.HostMetadataStoreType, "Concrete HostMetadataStore implementations");
        DumpConcreteDerivedTypes(allTypes, binding.HostLockingStoreType, "Concrete HostLockingStore implementations");
        DumpConcreteDerivedTypes(allTypes, binding.CobaltStorageType, "Concrete CobaltStorage implementations");
        DumpConcreteDerivedTypes(allTypes, binding.HostBlobStoreType, "Concrete HostBlobStore implementations");
        DumpConcreteDerivedTypes(allTypes, binding.BlobStoreType, "Concrete BlobStore implementations");
        DumpConcreteDerivedTypes(allTypes, binding.RequestProcessorType, "Concrete RequestProcessor implementations");
    }

    private static void DumpConcreteDerivedTypes(
        List<Type> allTypes,
        Type? baseType,
        string title)
    {
        if (baseType == null)
            return;

        var matches = allTypes
            .Where(t =>
                t != baseType &&
                !t.IsAbstract &&
                baseType.IsAssignableFrom(t))
            .OrderBy(t => t.FullName)
            .ToList();

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title(title);

        table.AddColumn("Type");
        table.AddColumn("Constructors");

        if (matches.Count == 0)
        {
            table.AddRow("[yellow]none found[/]", "");
            AnsiConsole.Write(table);
            return;
        }

        foreach (var type in matches)
        {
            var constructors = string.Join("\n",
                type.GetConstructors(
                        BindingFlags.Public |
                        BindingFlags.NonPublic |
                        BindingFlags.Instance)
                    .Select(SafeFormatConstructor));

            table.AddRow(
                Markup.Escape(type.FullName ?? type.Name),
                Markup.Escape(constructors));
        }

        AnsiConsole.Write(table);
    }
    private static void DumpDetailedTypeInfo(Type? type, string title)
    {
        if (type == null)
            return;

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title(title);

        table.AddColumn("Kind");
        table.AddColumn("Signature");
        table.AddRow("Type", Markup.Escape("FullName: " + (type.FullName ?? type.Name)));
        table.AddRow("Type", Markup.Escape("BaseType: " + (type.BaseType?.FullName ?? "")));
        table.AddRow("Type", Markup.Escape("IsAbstract: " + type.IsAbstract));

        foreach (var prop in type.GetProperties(
             BindingFlags.Public |
             BindingFlags.NonPublic |
             BindingFlags.Static |
             BindingFlags.Instance))
        {
            var getter = prop.GetGetMethod(true);
            var setter = prop.GetSetMethod(true);

            var access =
                getter != null && getter.IsAbstract ? "abstract " :
                getter != null && getter.IsVirtual ? "virtual " :
                "";

            var canSet = setter != null ? " set" : " readonly";

            table.AddRow(
                getter != null && getter.IsStatic ? "Static Property" : "Property",
                Markup.Escape(
                    access +
                    SafePrettyTypeName(prop.PropertyType) + " " +
                    prop.Name +
                    " [" + canSet + "]"));
        }

        foreach (var method in type.GetMethods(
                     BindingFlags.Public |
                     BindingFlags.NonPublic |
                     BindingFlags.Static |
                     BindingFlags.Instance)
                 .Where(m => !m.IsSpecialName)
                 .OrderBy(m => m.IsStatic ? 0 : 1)
                 .ThenBy(m => m.Name)
                 .Take(200))
        {
            table.AddRow(
                method.IsStatic ? "Static Method" : "Method",
                Markup.Escape(SafeFormatMethod(method)));
        }

        foreach (var prop in type.GetProperties(
                     BindingFlags.Public |
                     BindingFlags.NonPublic |
                     BindingFlags.Static |
                     BindingFlags.Instance))
        {
            table.AddRow(
                prop.GetMethod != null && prop.GetMethod.IsStatic ? "Static Property" : "Property",
                Markup.Escape(SafePrettyTypeName(prop.PropertyType) + " " + prop.Name));
        }

        AnsiConsole.Write(table);
    }

    private static string SafeFormatConstructor(ConstructorInfo ctor)
    {
        try
        {
            return FormatConstructor(ctor);
        }
        catch (Exception ex)
        {
            return ctor.Name + " [format failed: " + ex.GetType().Name + " - " + ex.Message + "]";
        }
    }

    private static string SafeFormatMethod(MethodInfo method)
    {
        try
        {
            return FormatMethod(method);
        }
        catch (Exception ex)
        {
            return method.Name + " [format failed: " + ex.GetType().Name + " - " + ex.Message + "]";
        }
    }

    private static string SafePrettyTypeName(Type type)
    {
        try
        {
            return PrettyTypeName(type);
        }
        catch
        {
            return type.FullName ?? type.Name;
        }
    }

    private List<Type> GetAllCobaltTypes()
    {
        var assemblies = new List<Assembly>();

        assemblies.Add(_assembly);

        foreach (var dll in Directory.GetFiles(_dllFolder, "*.dll"))
        {
            try
            {
                var name = AssemblyName.GetAssemblyName(dll);

                if (name.Name == null)
                    continue;

                if (name.Name.IndexOf("Cobalt", StringComparison.OrdinalIgnoreCase) < 0 &&
                    name.Name.IndexOf("Office", StringComparison.OrdinalIgnoreCase) < 0 &&
                    name.Name.IndexOf("Microsoft", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var alreadyLoaded = AppDomain.CurrentDomain
                    .GetAssemblies()
                    .FirstOrDefault(a =>
                        string.Equals(a.GetName().Name, name.Name, StringComparison.OrdinalIgnoreCase));

                assemblies.Add(alreadyLoaded ?? Assembly.LoadFrom(dll));
            }
            catch
            {
            }
        }

        return assemblies
            .Distinct()
            .SelectMany(a =>
            {
                try
                {
                    return a.GetTypes();
                }
                catch (ReflectionTypeLoadException ex)
                {
                    return ex.Types.Where(t => t != null)!;
                }
                catch
                {
                    return Array.Empty<Type>();
                }
            })
            .Where(t => t != null)
            .Cast<Type>()
            .ToList();
    }

    private readonly string _dllPath;
    private readonly Assembly _assembly;
    private readonly string _dllFolder;

    public MicrosoftCobaltCoreAdapter(string dllPath)
    {
        if (!File.Exists(dllPath))
            throw new FileNotFoundException("Microsoft.CobaltCore.dll was not found.", dllPath);

        _dllPath = dllPath;
        _dllFolder = Path.GetDirectoryName(dllPath)!;

        AppDomain.CurrentDomain.AssemblyResolve += ResolveCobaltDependency;

        _assembly = Assembly.LoadFrom(_dllPath);
    }

    public ShreddedExportResult TryExportFromSharePointDbRows(
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string outputPath,
        string workFolder)
    {
        try
        {
            Directory.CreateDirectory(workFolder);

            WriteRowsForDiagnostics(rows, workFolder);

            var binding = Bind();

            PrintBinding(binding);

            DumpDetailedTypeInfo(binding.DisposalEscrowType, "DisposalEscrow Details");
            DumpDetailedTypeInfo(binding.HostLockingStoreType, "HostLockingStore Details");
            DumpDetailedTypeInfo(binding.HostMetadataStoreType, "HostMetadataStore Details");
            DumpDetailedTypeInfo(binding.HostFileType, "HostFile Details");
            DumpDetailedTypeInfo(binding.CobaltFilePartitionConfigType, "CobaltFilePartitionConfig Details");
            DumpDetailedTypeInfo(binding.FilePartitionIdType, "FilePartitionId Details");
            DumpDetailedTypeInfo(binding.RequestProcessorType, "RequestProcessor Details");
            DumpDetailedTypeInfo(binding.GenericFdaType, "GenericFda Details");
            DumpDetailedTypeInfo(binding.CobaltStorageType, "CobaltStorage Details");
            DumpDetailedTypeInfo(binding.HostBlobStoreType, "HostBlobStore Details");
            DumpDetailedTypeInfo(binding.BlobStoreType, "BlobStore Details");
            DumpDetailedTypeInfo(binding.SchemaType, "Schema Details");
            DumpDetailedTypeInfo(binding.CellStorageConfigType, "CellStorageConfig Details");
            DumpDetailedTypeInfo(binding.HostFileUpdateType, "HostFileUpdate Details");
            DumpDetailedTypeInfo(binding.GenericFdaStreamType, "GenericFdaStream Details");
            DumpDetailedTypeInfo(binding.GenericFdaConfigType, "GenericFda.Config Details");
            DumpDetailedTypeInfo(binding.ArrayGuidType, "ArrayGuid Details");
            DumpDetailedTypeInfo(binding.FileTypeType, "FileType Details");
            DumpDetailedTypeInfo(binding.FileLockType, "FileLock Details");
            DumpDetailedTypeInfo(binding.FileCheckOutInfoType, "FileCheckOutInfo Details");
            DumpDetailedTypeInfo(binding.HostFileQueryType, "HostFileQuery Details");
            DumpDetailedTypeInfo(binding.QueryType, "Query Details");
            DumpDetailedTypeInfo(binding.UpdateType, "Update Details");
            DumpDetailedTypeInfo(binding.EndpointImplType, "EndpointImpl Details");
            DumpDetailedTypeInfo(binding.PersistedLocalHostBlobStoreType, "PersistedLocalHostBlobStore Details");
            DumpDetailedTypeInfo(binding.LocalHostBlobStoreType, "LocalHostBlobStore Details");
            DumpDetailedTypeInfo(binding.FilePathType, "FilePath Details");
            DumpDetailedTypeInfo(binding.HostBlobStoreIdsType, "HostBlobStoreIds Details");
            DumpDetailedTypeInfo(binding.CellStorageMetricsType, "CellStorageMetrics Details");
            DumpDetailedTypeInfo(binding.BlobTagType, "BlobTag Details");
            DumpDetailedTypeInfo(binding.TransactionContextType, "TransactionContext Details");
            DumpDetailedTypeInfo(binding.StoreUpdateOptionsType, "StoreUpdateOptions Details");
            DumpDetailedTypeInfo(binding.AtomType, "Atom Details");
            DumpDetailedTypeInfo(binding.BytesType, "Bytes Details");

            DumpEnumValues(binding.BlobCommitTypeType, "BlobCommitType Enum Values");
            DumpEnumValues(binding.BlobTagType, "BlobTag Enum Values");
            DumpEnumValues(binding.TransactionContextType, "TransactionContext Enum Values");
            DumpEnumValues(binding.BlobChangeFrequencyType, "BlobChangeFrequency Enum Values");

            DumpDerivedTypes(binding);
            /*
             * Important:
             * Microsoft.CobaltCore does not have a documented public method named
             * "OpenSharePointContentDatabaseDocStreams".
             *
             * Therefore the adapter has to bind one of these:
             *
             * 1. A CobaltFile factory that can open existing Cobalt/shredded storage.
             * 2. A Cobalt endpoint that can expose final content through GenericFda.
             * 3. A lower-level object storage API exposed by your DLL build.
             */

            var hostBlobResult = TryExportViaHostBlobStoreEndpoint(
    binding,
    rows,
    item,
    outputPath,
    workFolder);

            if (hostBlobResult.Success)
                return hostBlobResult;

            return hostBlobResult;
        }
        catch (Exception ex)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "CobaltCore export failed: " + ex
            };
        }
    }

    private Assembly? ResolveCobaltDependency(object? sender, ResolveEventArgs args)
    {
        var name = new AssemblyName(args.Name).Name + ".dll";
        var candidate = Path.Combine(_dllFolder, name);

        if (File.Exists(candidate))
            return Assembly.LoadFrom(candidate);

        return null;
    }

    private CobaltCoreBinding Bind()
    {
        var allTypes = GetAllCobaltTypes();

        var binding = new CobaltCoreBinding
        {
            BlobCommitTypeType = FindType(allTypes, "BlobCommitType"),
            BlobChangeFrequencyType = FindType(allTypes, "BlobChangeFrequency"),
            AtomType = FindType(allTypes, "Atom"),
            BytesType = FindType(allTypes, "Bytes"),
            AtomFromByteArrayType = FindType(allTypes, "AtomFromByteArray"),
            RequestBatchType = FindType(allTypes, "RequestBatch"),
            CobaltFileType = FindType(allTypes, "CobaltFile"),
            GenericFdaType = FindType(allTypes, "GenericFda"),
            FileAtomType = FindType(allTypes, "FileAtom"),
            DisposalEscrowType = FindType(allTypes, "DisposalEscrow"),
            FilePartitionIdType = FindType(allTypes, "FilePartitionId"),
            CobaltFilePartitionConfigType = FindType(allTypes, "CobaltFilePartitionConfig"),
            HostLockingStoreType = FindType(allTypes, "HostLockingStore"),
            RequestProcessorType = FindType(allTypes, "RequestProcessor"),
            HostMetadataStoreType = FindType(allTypes, "HostMetadataStore"),
            HostFileType = FindType(allTypes, "HostFile"),
            GetConfigForPartitionAndVersionType = FindType(allTypes, "GetConfigForPartitionAndVersion"),
            CobaltStorageType = FindType(allTypes, "CobaltStorage"),
            HostBlobStoreType = FindType(allTypes, "HostBlobStore"),
            BlobStoreType = FindType(allTypes, "BlobStore"),
            SchemaType = FindType(allTypes, "Schema"),
            CellStorageConfigType = FindType(allTypes, "CellStorageConfig"),
            HostFileUpdateType = FindType(allTypes, "HostFileUpdate"),
            GenericFdaStreamType = FindType(allTypes, "GenericFdaStream"),
            GenericFdaConfigType = FindNestedType(allTypes, "GenericFda", "Config"),
            ArrayGuidType = FindType(allTypes, "ArrayGuid"),
            FileTypeType = FindType(allTypes, "FileType"),
            FileLockType = FindType(allTypes, "FileLock"),
            FileCheckOutInfoType = FindType(allTypes, "FileCheckOutInfo"),
            HostFileQueryType = FindType(allTypes, "HostFileQuery"),
            QueryType = FindType(allTypes, "Query"),
            UpdateType = FindType(allTypes, "Update"),
            HostFileUpdateType2 = FindType(allTypes, "HostFileUpdate"),
            EndpointImplType = FindTypeContaining(allTypes, "EndpointUtils+EndpointImpl"),
            PersistedLocalHostBlobStoreType = FindType(allTypes, "PersistedLocalHostBlobStore"),
            LocalHostBlobStoreType = FindType(allTypes, "LocalHostBlobStore"),
            HostMetadataStoreOnHostBlobStoreType = FindType(allTypes, "HostMetadataStoreOnHostBlobStore"),
            HostLockingStoreOnHostBlobStoreType = FindType(allTypes, "HostLockingStoreOnHostBlobStore"),
            FilePathType = FindType(allTypes, "FilePath"),
            HostBlobStoreIdsType = FindType(allTypes, "HostBlobStoreIds"),
            CellStorageMetricsType = FindType(allTypes, "CellStorageMetrics"),
            BlobTagType = FindType(allTypes, "BlobTag"),
            TransactionContextType = FindType(allTypes, "TransactionContext"),
            StoreUpdateOptionsType = FindType(allTypes, "StoreUpdateOptions")
        };

        if (binding.GenericFdaType == null)
            throw new InvalidOperationException("GenericFda type was not found.");

        if (binding.CobaltFileType == null)
            throw new InvalidOperationException("CobaltFile type was not found.");

        return binding;
    }

    private static Type? FindType(IEnumerable<Type> allTypes, string simpleName)
    {
        return allTypes.FirstOrDefault(t =>
            string.Equals(t.Name, simpleName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(t.FullName, simpleName, StringComparison.OrdinalIgnoreCase) ||
            t.FullName?.EndsWith("." + simpleName, StringComparison.OrdinalIgnoreCase) == true);
    }

    private static Type? FindTypeContaining(IEnumerable<Type> allTypes, string text)
    {
        return allTypes.FirstOrDefault(t =>
            t.FullName?.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static Type? FindNestedType(IEnumerable<Type> allTypes, string parentSimpleName, string nestedSimpleName)
    {
        return allTypes.FirstOrDefault(t =>
            t.DeclaringType != null &&
            string.Equals(t.DeclaringType.Name, parentSimpleName, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(t.Name, nestedSimpleName, StringComparison.OrdinalIgnoreCase));
    }

    private static void PrintBinding(CobaltCoreBinding binding)
    {
        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("CobaltCore Adapter Binding");

        table.AddColumn("Required Type");
        table.AddColumn("Resolved");

        table.AddRow("BlobCommitType", binding.BlobCommitTypeType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("Atom", binding.AtomType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("Bytes", binding.BytesType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("AtomFromByteArray", binding.AtomFromByteArrayType?.FullName ?? "[red]missing[/]");
        table.AddRow("RequestBatch", binding.RequestBatchType?.FullName ?? "[red]missing[/]");
        table.AddRow("CobaltFile", binding.CobaltFileType?.FullName ?? "[red]missing[/]");
        table.AddRow("GenericFda", binding.GenericFdaType?.FullName ?? "[red]missing[/]");
        table.AddRow("FileAtom", binding.FileAtomType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("DisposalEscrow", binding.DisposalEscrowType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("FilePartitionId", binding.FilePartitionIdType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("CobaltFilePartitionConfig", binding.CobaltFilePartitionConfigType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostLockingStore", binding.HostLockingStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("RequestProcessor", binding.RequestProcessorType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostMetadataStore", binding.HostMetadataStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostFile", binding.HostFileType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("GetConfigForPartitionAndVersion", binding.GetConfigForPartitionAndVersionType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("CobaltStorage", binding.CobaltStorageType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostBlobStore", binding.HostBlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("BlobStore", binding.BlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("Schema", binding.SchemaType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("CellStorageConfig", binding.CellStorageConfigType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostFileUpdate", binding.HostFileUpdateType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("GenericFdaStream", binding.GenericFdaStreamType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("GenericFda.Config", binding.GenericFdaConfigType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("ArrayGuid", binding.ArrayGuidType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("FileType", binding.FileTypeType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("FileLock", binding.FileLockType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("FileCheckOutInfo", binding.FileCheckOutInfoType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostFileQuery", binding.HostFileQueryType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("Query", binding.QueryType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("Update", binding.UpdateType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("EndpointImpl", binding.EndpointImplType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("PersistedLocalHostBlobStore", binding.PersistedLocalHostBlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("LocalHostBlobStore", binding.LocalHostBlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostMetadataStoreOnHostBlobStore", binding.HostMetadataStoreOnHostBlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostLockingStoreOnHostBlobStore", binding.HostLockingStoreOnHostBlobStoreType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("FilePath", binding.FilePathType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("HostBlobStoreIds", binding.HostBlobStoreIdsType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("CellStorageMetrics", binding.CellStorageMetricsType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("BlobTag", binding.BlobTagType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("TransactionContext", binding.TransactionContextType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("StoreUpdateOptions", binding.StoreUpdateOptionsType?.FullName ?? "[yellow]not found[/]");
        table.AddRow("BlobChangeFrequency", binding.BlobChangeFrequencyType?.FullName ?? "[yellow]not found[/]");

        AnsiConsole.Write(table);
    }

    private object? TryCreateCobaltFileFromRows(
    CobaltCoreBinding binding,
    List<ShreddedStreamRow> rows,
    DocumentItemInfo item,
    string workFolder)
    {
        // No concrete HostFile exists in this CobaltCore build.
        // So CobaltFile constructor path is not usable.
        return null;
    }

    private object? TryCreateCobaltFileFromStaticFactory(
        CobaltCoreBinding binding,
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string workFolder)
    {
        var cobaltFileType = binding.CobaltFileType!;

        var methods = cobaltFileType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
            .Where(m =>
                m.ReturnType == cobaltFileType &&
                (
                    m.Name.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("Load", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("From", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    m.Name.IndexOf("Create", StringComparison.OrdinalIgnoreCase) >= 0
                ))
            .ToList();

        foreach (var method in methods)
        {
            var args = TryBuildArgumentsForMethod(
                method,
                rows,
                item,
                workFolder);

            if (args == null)
                continue;

            try
            {
                var result = method.Invoke(null, args);

                if (result != null)
                {
                    AnsiConsole.MarkupLine(
                        $"[green]Created CobaltFile using static factory:[/] {Markup.Escape(method.Name)}");

                    return result;
                }
            }
            catch (TargetInvocationException ex)
            {
                AnsiConsole.MarkupLine(
                    $"[yellow]Factory failed:[/] {Markup.Escape(method.Name)} - {Markup.Escape(ex.InnerException?.Message ?? ex.Message)}");
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine(
                    $"[yellow]Factory failed:[/] {Markup.Escape(method.Name)} - {Markup.Escape(ex.Message)}");
            }
        }

        return null;
    }

    private object? TryCreateCobaltFileFromConstructor(
    CobaltCoreBinding binding,
    List<ShreddedStreamRow> rows,
    DocumentItemInfo item,
    string workFolder)
    {
        return null;
    }

    private object[]? TryBuildArgumentsForMethod(
        MethodInfo method,
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string workFolder)
    {
        var parameters = method.GetParameters();
        var args = new object?[parameters.Length];

        for (var i = 0; i < parameters.Length; i++)
        {
            var value = TryBuildSimpleArgument(
                parameters[i].ParameterType,
                rows,
                item,
                workFolder);

            if (!value.Success)
                return null;

            args[i] = value.Value;
        }

        return args!;
    }

    private object[]? TryBuildArgumentsForConstructor(
        ConstructorInfo ctor,
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string workFolder,
        CobaltCoreBinding binding)
    {
        var parameters = ctor.GetParameters();
        var args = new object?[parameters.Length];

        for (var i = 0; i < parameters.Length; i++)
        {
            var type = parameters[i].ParameterType;

            var value = TryBuildSimpleArgument(
                type,
                rows,
                item,
                workFolder);

            if (value.Success)
            {
                args[i] = value.Value;
                continue;
            }

            /*
             * Special case:
             * Some CobaltFile constructors require Dictionary<FilePartitionId, CobaltFilePartitionConfig>.
             * We cannot safely create this without knowing your exact partition config constructor,
             * so this adapter returns null and prints constructor signatures.
             */
            return null;
        }

        return args!;
    }

    private ArgumentBuildResult TryBuildSimpleArgument(
        Type type,
        List<ShreddedStreamRow> rows,
        DocumentItemInfo item,
        string workFolder)
    {
        if (type == typeof(string))
            return ArgumentBuildResult.Ok(workFolder);

        if (type == typeof(Guid))
            return ArgumentBuildResult.Ok(item.Id);

        if (type == typeof(Stream))
        {
            var stream = BuildCombinedDiagnosticStream(rows);
            return ArgumentBuildResult.Ok(stream);
        }

        if (type == typeof(byte[]))
        {
            var bytes = BuildCombinedDiagnosticBytes(rows);
            return ArgumentBuildResult.Ok(bytes);
        }

        if (type == typeof(DirectoryInfo))
            return ArgumentBuildResult.Ok(new DirectoryInfo(workFolder));

        if (type == typeof(FileInfo))
            return ArgumentBuildResult.Fail();

        if (type.IsArray)
        {
            var elementType = type.GetElementType();

            if (elementType == null)
                return ArgumentBuildResult.Fail();

            return ArgumentBuildResult.Ok(Array.CreateInstance(elementType, 0));
        }

        if (type.IsGenericType &&
            type.GetGenericTypeDefinition() == typeof(Dictionary<,>))
        {
            var dictionary = Activator.CreateInstance(type);

            return dictionary == null
                ? ArgumentBuildResult.Fail()
                : ArgumentBuildResult.Ok(dictionary);
        }

        if (!type.IsValueType)
        {
            // Do not pass null for complex Cobalt host objects.
            // A CobaltFile created with null HostFile/HostMetadataStore/LockingStore is useless.
            if (type.FullName?.StartsWith("Cobalt.", StringComparison.OrdinalIgnoreCase) == true)
                return ArgumentBuildResult.Fail();

            return ArgumentBuildResult.Ok(null);
        }

        return ArgumentBuildResult.Fail();
    }

    private ShreddedExportResult ExportFromCobaltFile(
        CobaltCoreBinding binding,
        object cobaltFile,
        DocumentItemInfo item,
        string outputPath)
    {
        var cobaltEndpointProperty =
            cobaltFile.GetType().GetProperty("CobaltEndpoint", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);

        if (cobaltEndpointProperty == null)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "CobaltFile was created, but CobaltEndpoint property was not found."
            };
        }

        var cobaltEndpoint = cobaltEndpointProperty.GetValue(cobaltFile);

        if (cobaltEndpoint == null)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "CobaltFile.CobaltEndpoint returned null."
            };
        }

        var genericFdaType = binding.GenericFdaType!;

        var ctor = genericFdaType.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(c =>
            {
                var p = c.GetParameters();

                return p.Length == 2 &&
                       p[0].ParameterType.IsAssignableFrom(cobaltEndpoint.GetType());
            });

        if (ctor == null)
        {
            DumpConstructors(genericFdaType, "GenericFda constructors");

            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "Could not find GenericFda constructor compatible with CobaltEndpoint."
            };
        }

        var genericFda = ctor.Invoke(new[] { cobaltEndpoint, null });

        var getContentStream = genericFdaType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
            .FirstOrDefault(m =>
                m.Name.Equals("GetContentStream", StringComparison.OrdinalIgnoreCase) &&
                typeof(Stream).IsAssignableFrom(m.ReturnType));

        if (getContentStream == null)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "GenericFda.GetContentStream() was not found."
            };
        }

        var streamObj = getContentStream.Invoke(genericFda, Array.Empty<object>());

        if (streamObj is not Stream contentStream)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                Message = "GenericFda.GetContentStream() did not return a Stream."
            };
        }

        using (contentStream)
        using (var output = File.Create(outputPath))
        {
            contentStream.CopyTo(output);
        }

        var fileInfo = new FileInfo(outputPath);

        if (item.SizeBytes.HasValue && fileInfo.Length != item.SizeBytes.Value)
        {
            return new ShreddedExportResult
            {
                Success = false,
                OutputPath = outputPath,
                BytesWritten = fileInfo.Length,
                Message =
                    $"CobaltCore returned a stream, but size mismatch. Expected {item.SizeBytes.Value:N0}, got {fileInfo.Length:N0}."
            };
        }

        return new ShreddedExportResult
        {
            Success = true,
            OutputPath = outputPath,
            BytesWritten = fileInfo.Length,
            Message = "Exported through Microsoft.CobaltCore GenericFda.GetContentStream()."
        };
    }

    private static void DumpCobaltFileFactories(Type cobaltFileType)
    {
        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("CobaltFile Factories / Constructors");

        table.AddColumn("Member");

        foreach (var ctor in cobaltFileType.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            table.AddRow(Markup.Escape(FormatConstructor(ctor)));
        }

        foreach (var method in cobaltFileType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                     .Where(m =>
                         m.ReturnType == cobaltFileType ||
                         m.Name.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("Load", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("From", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("Create", StringComparison.OrdinalIgnoreCase) >= 0))
        {
            table.AddRow(Markup.Escape(FormatMethod(method)));
        }

        AnsiConsole.Write(table);
    }

    private static void DumpFileAtomFactories(Type? fileAtomType)
    {
        if (fileAtomType == null)
            return;

        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title("FileAtom Factories");

        table.AddColumn("Member");

        foreach (var method in fileAtomType.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                     .Where(m =>
                         m.Name.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("Load", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("From", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         m.Name.IndexOf("Create", StringComparison.OrdinalIgnoreCase) >= 0))
        {
            table.AddRow(Markup.Escape(FormatMethod(method)));
        }

        AnsiConsole.Write(table);
    }

    private static void DumpConstructors(Type type, string title)
    {
        var table = new Table()
            .Border(TableBorder.Rounded)
            .Title(title);

        table.AddColumn("Constructor");

        foreach (var ctor in type.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
        {
            table.AddRow(Markup.Escape(FormatConstructor(ctor)));
        }

        AnsiConsole.Write(table);
    }

    private static void WriteRowsForDiagnostics(List<ShreddedStreamRow> rows, string workFolder)
    {
        var dumpFolder = Path.Combine(workFolder, "db_rows");

        Directory.CreateDirectory(dumpFolder);

        foreach (var row in rows)
        {
            var fileName =
                $"hist_{row.HistVersion}_type_{row.Type}_stream_{row.StreamId}_bsn_{row.BSN}.bin";

            File.WriteAllBytes(Path.Combine(dumpFolder, fileName), row.Content);
        }
    }

    private static byte[] BuildCombinedDiagnosticBytes(List<ShreddedStreamRow> rows)
    {
        using var ms = new MemoryStream();

        foreach (var row in rows.OrderBy(x => x.StreamId))
            ms.Write(row.Content, 0, row.Content.Length);

        return ms.ToArray();
    }

    private static Stream BuildCombinedDiagnosticStream(List<ShreddedStreamRow> rows)
    {
        return new MemoryStream(BuildCombinedDiagnosticBytes(rows));
    }

    private static string FormatConstructor(ConstructorInfo ctor)
    {
        var args = string.Join(", ",
            ctor.GetParameters()
                .Select(FormatParameter));

        var visibility =
            ctor.IsPublic ? "public" :
            ctor.IsPrivate ? "private" :
            ctor.IsFamily ? "protected" :
            ctor.IsAssembly ? "internal" :
            "non-public";

        return visibility + " " + ctor.DeclaringType?.Name + "(" + args + ")";
    }

    private static string FormatMethod(MethodInfo method)
    {
        var args = string.Join(", ",
            method.GetParameters()
                .Select(FormatParameter));

        var visibility =
            method.IsPublic ? "public" :
            method.IsPrivate ? "private" :
            method.IsFamily ? "protected" :
            method.IsAssembly ? "internal" :
            "non-public";

        var staticText = method.IsStatic ? " static" : "";
        var abstractText = method.IsAbstract ? " abstract" : "";
        var virtualText = method.IsVirtual && !method.IsAbstract ? " virtual" : "";

        return visibility + staticText + abstractText + virtualText + " " +
               PrettyTypeName(method.ReturnType) + " " +
               method.Name + "(" + args + ")";
    }

    private static string FormatParameter(ParameterInfo parameter)
    {
        var prefix = "";

        if (parameter.IsOut)
            prefix = "out ";
        else if (parameter.ParameterType.IsByRef)
            prefix = "ref ";

        var parameterType = parameter.ParameterType.IsByRef
            ? parameter.ParameterType.GetElementType()
            : parameter.ParameterType;

        return prefix + PrettyTypeName(parameterType!) + " " + parameter.Name;
    }

    private static string PrettyTypeName(Type type)
    {
        if (type == typeof(void))
            return "void";

        if (!type.IsGenericType)
            return type.Name;

        var name = type.Name;
        var tick = name.IndexOf('`');

        if (tick >= 0)
            name = name.Substring(0, tick);

        return name + "<" + string.Join(", ", type.GetGenericArguments().Select(PrettyTypeName)) + ">";
    }

    private sealed class CobaltCoreBinding
    {
        public Type? BlobCommitTypeType { get; set; }
        public Type? AtomType { get; set; }
        public Type? BytesType { get; set; }
        public Type? BlobChangeFrequencyType { get; set; }
        public Type? EndpointImplType { get; set; }
        public Type? PersistedLocalHostBlobStoreType { get; set; }
        public Type? LocalHostBlobStoreType { get; set; }
        public Type? HostMetadataStoreOnHostBlobStoreType { get; set; }
        public Type? HostLockingStoreOnHostBlobStoreType { get; set; }
        public Type? FilePathType { get; set; }
        public Type? HostBlobStoreIdsType { get; set; }
        public Type? CellStorageMetricsType { get; set; }
        public Type? BlobTagType { get; set; }
        public Type? TransactionContextType { get; set; }
        public Type? StoreUpdateOptionsType { get; set; }
        public Type? AtomFromByteArrayType { get; set; }
        public Type? RequestBatchType { get; set; }
        public Type? CobaltFileType { get; set; }
        public Type? GenericFdaType { get; set; }
        public Type? FileAtomType { get; set; }
        public Type? DisposalEscrowType { get; set; }
        public Type? FilePartitionIdType { get; set; }

        public Type? CobaltFilePartitionConfigType { get; set; }

        public Type? HostLockingStoreType { get; set; }

        public Type? RequestProcessorType { get; set; }

        public Type? HostMetadataStoreType { get; set; }

        public Type? HostFileType { get; set; }

        public Type? GetConfigForPartitionAndVersionType { get; set; }
        public Type? CobaltStorageType { get; set; }
        public Type? HostBlobStoreType { get; set; }
        public Type? BlobStoreType { get; set; }
        public Type? SchemaType { get; set; }
        public Type? CellStorageConfigType { get; set; }
        public Type? HostFileUpdateType { get; set; }
        public Type? GenericFdaStreamType { get; set; }
        public Type? GenericFdaConfigType { get; set; }
        public Type? ArrayGuidType { get; set; }
        public Type? FileTypeType { get; set; }
        public Type? FileLockType { get; set; }
        public Type? FileCheckOutInfoType { get; set; }
        public Type? HostFileQueryType { get; set; }
        public Type? QueryType { get; set; }
        public Type? UpdateType { get; set; }
        public Type? HostFileUpdateType2 { get; set; }
    }

    private readonly struct ArgumentBuildResult
    {
        private ArgumentBuildResult(bool success, object? value)
        {
            Success = success;
            Value = value;
        }

        public bool Success { get; }

        public object? Value { get; }

        public static ArgumentBuildResult Ok(object? value)
        {
            return new ArgumentBuildResult(true, value);
        }

        public static ArgumentBuildResult Fail()
        {
            return new ArgumentBuildResult(false, null);
        }
    }
}

public static class Compat
{
    public static Task WriteAllBytesAsync(string path, byte[] bytes)
    {
        File.WriteAllBytes(path, bytes);
        return Task.CompletedTask;
    }
}
