using System;
using Conq.Noble.Spreadox.Samples;

internal static class Program
{
    private static int Main()
    {
#if NETFRAMEWORK
        string runtime = "net48";
#else
        string runtime = "net8.0";
#endif
        try
        {
            return FullDemoRunner.Run(runtime);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return 1;
        }
    }
}
