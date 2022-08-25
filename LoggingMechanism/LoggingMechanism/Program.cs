class LoggingMechanism
{
    public static void Main()
    {
        using (StreamWriter w = File.AppendText("C:\\Users\\addsi\\source\\repos\\LoggingMechanism\\LoggingMechanism\\bin\\Log\\log.txt"))
        {
            Log("Logging Process 1 Finished", w);
            Log("Logging Process 2 Finished", w);
            w.Close();
        }
        using (StreamReader r = File.OpenText("C:\\Users\\addsi\\source\\repos\\LoggingMechanism\\LoggingMechanism\\bin\\Log\\log.txt"))
        {
            DumpLog(r);
        }
    }

    public static void Log(string logMessage, TextWriter w)
    {
        w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
        w.WriteLine("  :{0}", logMessage);
        w.Flush();
    }

    public static void DumpLog(StreamReader r)
    {
        string line;
        while ((line = r.ReadLine()) != null)
        {
            Console.WriteLine(line);
        }
        r.Close();
    }
}
