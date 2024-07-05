using CommandLine;
using CommandLine.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Excel2TextDiff
{
    class CommandLineOptions
    {
        [Option('t', SetName = "transform", Required = false, HelpText = "transform excel to text file")]
        public bool IsTransform { get; set; }

        [Option('d', SetName = "diff", Required = false, HelpText = "transform and diff file")]
        public bool IsDiff { get; set; }

        [Option('m', SetName = "merge", Required = false, HelpText = "transform and diff merge file")]
        public bool IsMerge { get; set; }

        [Option('p', Required = false, HelpText = "3rd diff program. default TortoiseMerge")]
        public string DiffProgram { get; set; }

        [Option('f', Required = false, HelpText = "3rd diff program argument format. default is TortoiseMerge format:'/base:{0} /mine:{1}'")]
        public string DiffProgramArgumentFormat { get; set; }

        [Option('g', SetName = "merge", Required = false, HelpText = "3rd diff program argument format. default is TortoiseMerge format:'/base:{0} /mine:{1}'")]
        public string MergeDiffArgumentFormat { get; set; }

        [Value(0)]
        public IList<string> Files { get; set; }

        [Usage()]
        public static IEnumerable<Example> Examples => new List<Example>
        {
            new Example("tranfrom to text", new CommandLineOptions { IsTransform = true, Files = new List<string>{"a.xlsx", "a.txt" } }),
            new Example("diff two excel file", new CommandLineOptions{ IsDiff = true, Files = new List<string>{"a.xlsx", "b.xlsx"}}),
            new Example("diff two excel file with TortoiseMerge", new CommandLineOptions{ IsDiff = true, DiffProgram = "TortoiseMerge",DiffProgramArgumentFormat = "/base:{0} /mine:{1}",  Files = new List<string>{"a.xlsx", "b.xlsx"}}),
        };
    }

    class Program
    {
        static void Main(string[] args)
        {
            // 打印args
            var options = ParseOptions(args);

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var writer = new Excel2TextWriter();

            if (options.IsTransform)
            {
                if (options.Files.Count != 2)
                {
                    Console.WriteLine("Usage: Excel2TextDiff -t <excel file> <text file>");
                    Environment.Exit(1);
                }

                writer.TransformToTextAndSave(options.Files[0], options.Files[1]);
            }
            else if (options.IsDiff)
            {
                if (options.Files.Count != 2)
                {
                    Console.WriteLine("Usage: Excel2TextDiff -d <excel file 1> <excel file 2> ");
                    Environment.Exit(1);
                }

                var diffProgame = options.DiffProgram ?? "TortoiseMerge.exe";

                if (!File.Exists(diffProgame))
                {
                    Console.WriteLine("Diff program not found");
                    Environment.Exit(1);
                }

                string ext = Path.GetExtension(options.Files[0]);
                string diffFile0 = options.Files[0];
                string diffFile1 = options.Files[1];
                if (ext.Equals(".xlsx"))
                {
                    diffFile0 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[0], diffFile0);

                    diffFile1 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[1], diffFile1);
                }

                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = diffProgame;
                string argsFormation = options.DiffProgramArgumentFormat ?? "/base:{0} /mine:{1}";
                startInfo.Arguments = string.Format(argsFormation, diffFile0, diffFile1);
                using (Process process = Process.Start(startInfo))
                {
                    // 等待进程退出
                    process.WaitForExit();

                    // 获取进程的退出代码
                    int exitCode = process.ExitCode;
                    Environment.Exit(exitCode);
                }
            }
            else if (options.IsMerge)
            {
                if (options.Files.Count != 4)
                {
                    Console.WriteLine("Usage: Excel2TextDiff -m <excel file 1> <excel file 2> <excel file 3> <excel file 4> ");
                    Environment.Exit(1);
                }

                var diffProgame = options.DiffProgram ?? "TortoiseMerge.exe";

                if (!File.Exists(diffProgame))
                {
                    Console.WriteLine("Diff program not found");
                    Environment.Exit(1);
                }

                string ext = Path.GetExtension(options.Files[0]);
                string diffFile0 = options.Files[0];
                string diffFile1 = options.Files[1];
                string diffFile2 = options.Files[2];
                string diffFile3 = options.Files[3];
                string argsFormation = options.DiffProgramArgumentFormat;
                bool forceError = false;
                if (ext.Equals(".xlsx"))
                {
                    diffFile0 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[0], diffFile0);

                    diffFile1 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[1], diffFile1);

                    diffFile2 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[2], diffFile2);

                    diffFile3 = Path.GetTempFileName();
                    writer.TransformToTextAndSave(options.Files[3], diffFile3);


                    argsFormation = options.MergeDiffArgumentFormat;
                    forceError = true;
                }
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = diffProgame;
                startInfo.Arguments = string.Format(argsFormation, diffFile0, diffFile1, diffFile2, diffFile3);
                using (Process process = Process.Start(startInfo))
                {
                    // 等待进程退出
                    process.WaitForExit();

                    // 获取进程的退出代码
                    int exitCode = process.ExitCode;
                    Environment.Exit(forceError ? 1 : exitCode);
                }
            }
            else
            {
                Console.WriteLine("Unknow Command");
                Environment.Exit(1);
            }
        }

        private static CommandLineOptions ParseOptions(String[] args)
        {
            var helpWriter = new StringWriter();
            var parser = new Parser(ps =>
            {
                ps.HelpWriter = helpWriter;
            });

            var result = parser.ParseArguments<CommandLineOptions>(args);
            if (result.Tag == ParserResultType.NotParsed)
            {
                Console.Error.WriteLine(helpWriter.ToString());
                Environment.Exit(1);
            }
            return ((Parsed<CommandLineOptions>)result).Value;
        }
    }
}
