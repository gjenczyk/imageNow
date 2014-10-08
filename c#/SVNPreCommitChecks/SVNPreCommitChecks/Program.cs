using System;
using System.IO;
using System.Text.RegularExpressions;

namespace TortoiseSVNCommitTests
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] allFiles = Directory.GetFiles(System.Environment.CurrentDirectory);
            string logName = "\\TortoiseSVNCommitTests_Results.txt";
            int errorCount = 0;
            bool errorFlag = false;
            string delimiterString = "*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*_*";

            StreamWriter logFile = new StreamWriter(System.Environment.CurrentDirectory + logName);
            Regex fileExtensionPattern = new Regex(@"^.*\.(js|jsh|ps1)$", RegexOptions.IgnoreCase);

            foreach (string file in allFiles)
            {
                if (fileExtensionPattern.IsMatch(file) && File.Exists(file))
                {

                    logFile.WriteLine(delimiterString);
                    logFile.WriteLine("Checking {0} for errors", file);
                    string[] fileToCheck = File.ReadAllLines(file);
                    int scriptLineCount = 0;

                    foreach (string line in fileToCheck)
                    {
                        if (!line.Contains(";"))
                        {
                            //Console.WriteLine("nope");
                        }

                        scriptLineCount++;
                        if (CheckForDebugLogsMismatch(line, logFile, scriptLineCount))
                        {
                            errorFlag = true;
                            errorCount++;
                        }
                    }
                }
            }

            if (errorFlag)
            {
                logFile.WriteLine("There were {0} errors detected by TortoiseSVNCommitTest", errorCount);
            }
            else
            {
                logFile.WriteLine("No errors found.  Files are ready to be committed.");
            }
            logFile.Close();
            //Console.ReadLine();
        }

        private static bool CheckForDebugLogsMismatch(string scriptLine, StreamWriter log, int lineNum)
        {

            //bit of manipulation to make the line easier to parse
            scriptLine = scriptLine.ToUpper();
            scriptLine = scriptLine.Replace("'","\"");

            if (scriptLine.Contains("DEBUG.LOG("))
            {
                Regex debugRegex = new Regex("DEBUG.LOG(.*);");
                var dR = debugRegex.Match(scriptLine);
                string dRr = dR.Groups[1].ToString();
                //Console.WriteLine(dRr);

                if (dRr == null)
                {
                    Console.WriteLine("Here's a problem");
                    Console.ReadLine();
                }

                // string manipulation to seperate the debug string and arguments
                int openingBrace = dRr.IndexOf('(');
                int closingBrace = dRr.LastIndexOf(')');
                if (closingBrace < 0)
                {
                    log.WriteLine("CB - NON-STANDARD DEBUG SYNTAX DETECTED @ line {0} \r\n {1}\r\n", lineNum, scriptLine);
                    return false;
                }
                int betweenBrace = closingBrace - openingBrace + 1;

                dRr = dRr.Substring(openingBrace, betweenBrace);
                //log.WriteLine(dRr + " * * ");

                int severityComma = dRr.IndexOf(',');
                if (severityComma < 0)
                {
                    log.WriteLine("SC - NON-STANDARD DEBUG SYNTAX DETECTED @ line {0} \r\n {1}\r\n", lineNum, scriptLine);
                    return false;
                }
                dRr = dRr.Substring(severityComma);

                int openingQuote = dRr.IndexOf('"');
                int closingQuote = dRr.LastIndexOf('"');
                if (closingQuote < 0)
                {
                    log.WriteLine("CQ - NON-STANDARD DEBUG SYNTAX DETECTED @ {0} \r\n {1}\r\n", lineNum, scriptLine);
                    return false;
                }
                int lengthQuote = closingQuote - openingQuote;

                string debugString = dRr.Substring(openingQuote + 1, lengthQuote - 1);

                string argsString = dRr.Substring(closingQuote + 1);

                //Console.WriteLine(debugString);
                //Console.WriteLine(argsString);

                int sCount = debugString.Split(new string[] { "%" }, StringSplitOptions.None).Length - 1;
                

                if (sCount > 0)
                {
                    int aCount = argsString.Split(',').Length - 1;
                    //Console.WriteLine(debugString + " " + sCount);
                    //Console.WriteLine(argsString + " " + aCount);
                    
                    if (sCount != aCount)
                    {
                        log.WriteLine("SYNTAX ERROR - ARGUMENTS NOT SUPPLIED FOR ALL SPECIFIERS @ line {0} - {1} {2}", lineNum, sCount, aCount);
                        log.WriteLine("Debug String: {0}\r\nArg String: {1}\r\n", debugString, argsString);
                        return true;
                    }

                }

                return false;
            }
            
            return false;
        }
    }
}

