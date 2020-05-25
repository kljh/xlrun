using System;
using Microsoft.Office.Interop.Excel;

namespace xlrun
{
    class Program
    {
        static string Usage() {
            var exe = System.AppDomain.CurrentDomain.FriendlyName;
            var usage = $@"
Usage examples: 
{exe}  -xlFileOpen MyWorkbook.xlsx  -xlRefreshLeftToRight  -xlRngGet Summary!TestStatus
{exe}  -xlFileOpen MyMacrobook.xlsm  -xlEvalMacro MyMacro  -xlRngGet Summary!B4  -xlFileSave  -timeout 10
{exe}  -xlFileNew  -xlRngSet A1 1.0  -xlRngGet A1  -xlRngSet A2 =today()  -xlRngGet A2 -xlFileSaveAs MyGeneratedBook.xlsx

";
            return usage;
        }

        static void Main(string[] args)
        {
            SetTimeoutWatchdog(args);
            MainTask(args);
        }

        static void MainTask(string[] args)
        {
            // display usage if no arguments
            if (args.Length==0) {
                Console.WriteLine(Usage());
                return;
            }

            var xlApp = new Microsoft.Office.Interop.Excel.Application();
            
            try
            {
                Microsoft.Office.Interop.Excel._Workbook wbk = null;
                Microsoft.Office.Interop.Excel.Range rng = null;

                xlApp.DisplayAlerts = false;
                xlApp.Visible = true;
                xlApp.Interactive = true;
                            
                // no-well established parsearg method
                for (int i=0; i<args.Length; i++) {
                    var cmd = args[i]; 
                    while (cmd[0]=='-' || cmd[0]=='/') 
                        cmd = cmd.Substring(1);
                    
                    switch (cmd) {
                        case "xlFileOpen":
                        case "xlFilePath":
                            var openPath = MakeFullPath(args[++i]);
                            Console.WriteLine($"> Open Workbook {openPath}");
                            if (wbk != null) 
                                wbk.Close(SaveChanges: false);
                            wbk = xlApp.Workbooks.Open(openPath);
                            break;

                        case "xlFileNew":
                            Console.WriteLine($"> New Workbook");
                            if (wbk != null) 
                                wbk.Close(SaveChanges: false);
                            wbk = xlApp.Workbooks.Add();
                            break;

                        case "xlFileSave":
                            Console.WriteLine($"> Save Workbook");
                            wbk.Save();
                            break;

                        case "xlFileSaveAs":
                            var savePath = MakeFullPath(args[++i]);
                            Console.WriteLine($"> Save Workbook as {savePath}");
                            wbk.SaveAs(savePath);
                            break;

                        case "xlEvalMacro":
                            var macro = args[++i];
                            Console.WriteLine($"> Evaluate macro {macro}");
                            xlApp.Evaluate(macro);
                            break;

                        case "xlRefreshLeftToRight":
                            Console.WriteLine($"> Refresh sheets left to right");
                            foreach (Microsoft.Office.Interop.Excel._Worksheet wsh in wbk.Worksheets)
                                if (wsh.Visible == XlSheetVisibility.xlSheetVisible)
                                    wsh.Calculate();
                            break;

                        case "xlRngGet":
                            var rngGetAddr = args[++i];
                            Console.WriteLine($"> Get Range {rngGetAddr}");
                            rng = xlApp.Range[rngGetAddr];
                            Console.WriteLine(rng.Value);
                            break;

                        case "xlRngSet":
                            var rngSetAddr = args[++i];
                            var rngSetValue = args[++i];
                            Console.WriteLine($"> Set Range {rngSetAddr} {rngSetValue}");
                            rng = xlApp.Range[rngSetAddr];
                            rng.Value = rngSetValue;
                            break;

                        case "timeout":
                            i++;
                            break;

                        default:
                            throw new Exception($"> Unexpected command {cmd}");
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Environment.Exit(1);
            }
            finally
            {
                foreach (Microsoft.Office.Interop.Excel._Workbook wbk in xlApp.Workbooks)
                    wbk.Close(SaveChanges: false);
                xlApp.Quit();
                Console.WriteLine("Done.");
            }
        }

        static string MakeFullPath(string path) {
            if (path.Contains("\\"))
                return path;
            else
                return System.IO.Directory.GetCurrentDirectory() + "\\" + path;
        }

        static void SetTimeoutWatchdog(string[] args)
        {
            int timeoutMs = -1;
            for (int i=0; i<args.Length; i++) {
                var cmd = args[i]; 
                if (cmd[0]!='-' && cmd[0]!='/') 
                    continue;
                while (cmd[0]=='-' || cmd[0]=='/') 
                    cmd = cmd.Substring(1);
                
                if (cmd=="timeout") {
                    timeoutMs = 1000 * Int32.Parse(args[++i]);
                    break;
                }
            }

            if (timeoutMs>0) {
                Console.WriteLine($"> Set Timeout {timeoutMs/1000.0}s");
                var timer = new System.Timers.Timer(timeoutMs);
                timer.AutoReset = false;
                timer.Elapsed += new System.Timers.ElapsedEventHandler((object source, System.Timers.ElapsedEventArgs eea) => {
                    Console.WriteLine($"> Timeout expired !!");
                    Environment.Exit(-1);
                });
                timer.Start();
            }
        }
    }
}
