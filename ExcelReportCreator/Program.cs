using System;
using System.Collections.Generic;
using System.Text;
using CLIUtility;
using ExcelReportCreator01.CmdLineCommands01;

namespace ExcelReportCreator01
{
    class Program
    {
        static int Main(string[] args)
        {
            int retval = 0;
            CommandList cl = new CommandList(true);
            cl.AddCommand(new DatabaseQueryCommand());
            retval = cl.DoCommand(args);
            return retval;
        }
    }
}
