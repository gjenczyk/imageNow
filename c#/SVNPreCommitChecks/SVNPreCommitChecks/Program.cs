using System;
using System.IO;
using System.Text.RegularExpressions;

namespace TortoiseSVNCommitTests
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] filesToCheck = File.ReadAllLines(args[0]);

            foreach (string line in filesToCheck)
            {

                if (CheckForDebugLogsMismatch(line))
                {
                    return false;
                }

            }

            Console.ReadLine();
        }


        private static bool CheckForDebugLogsMismatch(string scriptLine)
        {

            scriptLine = scriptLine.ToUpper();

            if (scriptLine.Contains("DEBUG.LOG"))
            {
                // string manipulation to seperate the debug string and arguments
                int openingBrace = scriptLine.IndexOf('(');
                int closingBrace = scriptLine.LastIndexOf(')');
                int betweenBrace = closingBrace - openingBrace;

                scriptLine = scriptLine.Substring(openingBrace, betweenBrace);
                //Console.WriteLine(scriptLine + " * * ");

                int severityComma = scriptLine.IndexOf(',');
                scriptLine = scriptLine.Substring(severityComma);

                int openingQuote = scriptLine.IndexOf('"');
                int closingQuote = scriptLine.LastIndexOf('"');
                int lengthQuote = closingQuote - openingQuote;

                string debugString = scriptLine.Substring(openingQuote + 1, lengthQuote - 1);

                string argsString = scriptLine.Substring(closingQuote + 1);

                //Console.WriteLine(debugString);
                //Console.WriteLine(argsString);

                int sCount = debugString.Split(new string[] { "%S" }, StringSplitOptions.None).Length - 1;
                int aCount = argsString.Split(',').Length - 1;

                if (sCount > 0)
                {
                    //Console.WriteLine(debugString + " " + sCount);
                    //Console.WriteLine(argsString + " " + aCount);
                    
                    if (sCount != aCount)
                    {
                        Console.WriteLine("SYNTAX ERROR - ARGUMENTS NOT SUPPLIED FOR ALL SPECIFIERS {0} {1}", sCount, aCount);
                        return true;
                    }

                }

                return false;
            }
            
            return false;
        }
    }
}

