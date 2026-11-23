using System.Reflection;
using CommandLine;
using CommandLine.Text;
using Windexter.Core;

namespace windexter_cli
{
    class Program
    {
        private static readonly string displayName = "Windexter CLI";
        private static readonly string githubBinaryRepo = "https://github.com/digitalsleuth/windexter";
        private static readonly Version? appVersion = new(Assembly.GetExecutingAssembly().GetName().Version!.ToString(3));

        public class Options
        {
            [Option('f', "file", Required = true, HelpText = "Path to the database.")]
            public string? File { get; set; }

            [Option('o', "output", Required = false, HelpText = "Path to the output directory.")]
            public string? Output { get; set; }

            [Option("all", HelpText = "Filter for all options.")]
            public bool AllFilterOptions { get; set; }

            [Option('a', "activity", HelpText = "Filter for user activity.")]
            public bool Activity { get; set; }

            [Option('c', "compinfo", HelpText = "Filter for Computer Info.")]
            public bool CompInfo { get; set; }

            [Option('g', "gps", HelpText = "Filter for GPS data.")]
            public bool Gps { get; set; }

            [Option('s', "summary", HelpText = "Filter for Search Summary.")]
            public bool Summary { get; set; }

            [Option('t', "timeline", HelpText = "Generate a Timeline for all data.")]
            public bool Timeline { get; set; }
            
            [Option('u', "url", HelpText = "Filter for URLs.")]
            public bool Url { get; set; }

            [Option("gather", HelpText = "Process Windows-gather.db if present (and -f is a Windows.db file).")]
            public bool Gather { get; set; }

            [Option("propmap", HelpText = "Process PropMap.db if present (and -f is a Windows.db file).")]
            public bool PropMap { get; set; }

            [Option("secstore", HelpText = "Process SecStore.db if present (and -f is a Windows.db file).")]
            public bool SecStore { get; set; }

            [Option("format", Default = "xlsx", HelpText = "Output format: xlsx, sqlite or both. The SQLite output always includes a view for every filter and the timeline.")]
            public string Format { get; set; } = "xlsx";

        }
        static async Task<int> Main(string[] args)
        {
            int exitCode = 0;
            var parser = new Parser(with => with.HelpWriter = null);
            var result = parser.ParseArguments<Options>(args);
            await result.WithParsedAsync(async opts => exitCode = await Execute(opts));
            result.WithNotParsed(errs =>
            {
                exitCode = errs.All(e => e is HelpRequestedError or VersionRequestedError) ? 0 : 1;
                var helpText = HelpText.AutoBuild(result, h =>
                {
                    h.AdditionalNewLineAfterOption = false;
                    h.Heading = $"{displayName} {appVersion}";
                    h.Copyright = $"Digital Sleuth - {githubBinaryRepo}";
                    return HelpText.DefaultParsingErrorsHandler(result, h);
                }, e => e);

                Console.Error.WriteLine(helpText);
            });
            return exitCode;
        }

        static async Task<int> Execute(Options opts)
        {
            try
            {
                return await RunOptions(opts) ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                return 1;
            }
        }

        static async Task<bool> RunOptions(Options opts)
        {
            string now = DateTime.Now.ToUniversalTime().ToString("yyyyMMdd-HHmmss");
            var filters = new SearchIndex.FilterOptions(
                opts.Url || opts.AllFilterOptions,
                opts.Gps || opts.AllFilterOptions,
                opts.Summary || opts.AllFilterOptions,
                opts.CompInfo || opts.AllFilterOptions,
                opts.Activity || opts.AllFilterOptions,
                opts.Timeline || opts.AllFilterOptions);
            string format = opts.Format.ToLowerInvariant();
            if (format is not ("xlsx" or "sqlite" or "both"))
            {
                Console.Error.WriteLine($"Unknown format '{opts.Format}'. Use xlsx, sqlite or both.");
                return false;
            }
            if (!File.Exists(opts.File))
            {
                Console.Error.WriteLine($"{opts.File} does not exist! Please check your path and try again.");
                return false;
            }
            SearchIndex.DbFile = new FileInfo(opts.File).FullName;
            string outputPath = string.IsNullOrEmpty(opts.Output) ? Path.GetDirectoryName(SearchIndex.DbFile)! : opts.Output;
            if (!Directory.Exists(outputPath))
            {
                Console.Error.WriteLine($"{outputPath} does not exist! Please check your path and try again.");
                return false;
            }
            Console.WriteLine($"{displayName} - {appVersion}");
            Console.WriteLine($"Digital Sleuth - {githubBinaryRepo}\n");
            Console.WriteLine("--- Processing Options ---");
            Console.WriteLine($"Start Time: {now}");
            Console.WriteLine($"File: {SearchIndex.DbFile}");
            Console.WriteLine($"Output: {outputPath}");
            Console.WriteLine($"Command Line: {Environment.CommandLine}");
            Console.WriteLine("Extracting data ...");
            SearchIndex.Reset();
            await Task.Run(SearchIndex.ExtractData);
            if (opts.Gather && SearchIndex.FindGatherFile() is string gatherFile)
            {
                Console.WriteLine("Windexter has found a Windows-gather.db in the same directory as the selected database and is processing it.");
                await Task.Run(() => SearchIndex.ExtractGather(gatherFile));
            }
            if (opts.PropMap && SearchIndex.FindPropMapFile() is string propMapFile)
            {
                Console.WriteLine("Windexter has found a PropMap.db in a sub-directory of the selected folder and is processing it.");
                await Task.Run(() => SearchIndex.ExtractPropMap(propMapFile));
            }
            if (opts.SecStore && SearchIndex.FindSecStoreFile() is string secStoreFile)
            {
                Console.WriteLine("Windexter has found a SecStore.db in a sub-directory of the selected folder and is processing it.");
                await Task.Run(() => SearchIndex.ExtractSecStore(secStoreFile));
            }
            Console.WriteLine("Parsing data ...");
            await SearchIndex.ParseDataAsync();
            List<string> outputFiles = [];
            if (format != "sqlite")
            {
                if (filters.Any)
                {
                    Console.WriteLine("Filtering data ...");
                    await SearchIndex.ParseSubData(filters, Console.WriteLine);
                }
                Console.WriteLine("Exporting to Excel ...");
                outputFiles.Add(Path.Combine(outputPath, SearchIndex.OutputFileName(now)));
                await SearchIndex.ExportToExcel(outputFiles[^1]);
            }
            if (format != "xlsx")
            {
                Console.WriteLine("Exporting to SQLite ...");
                outputFiles.Add(Path.Combine(outputPath, SearchIndex.OutputFileName(now, "sqlite")));
                await SearchIndex.ExportToSqlite(outputFiles[^1], $"{displayName} {appVersion}");
            }
            Console.WriteLine("Extraction and parsing completed!");
            outputFiles.ForEach(outputFile => Console.WriteLine($"Output File: {outputFile}"));
            Console.WriteLine($"Completed: {DateTime.Now.ToUniversalTime():yyyyMMdd-HHmmss}");
            return true;
        }
    }
}
