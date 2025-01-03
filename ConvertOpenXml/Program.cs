//Program
class Program
{
    //Main
    static void Main(string[] args)
    {
        var argCmd = ArgsCmd.Parser(args);
        if (argCmd.cmds.Count <= 0)
        {
            argCmd.Add(string.Empty, "string[]");
            argCmd.Add("-filter", "string[]");
            argCmd.Add("-json");
            argCmd.Add("-cs");
            argCmd.Add("-cs-namespace", "string");
            argCmd.Add("-cs-baseclass", "string");

            foreach (var command in argCmd.cmds)
            {
                string value = argCmd.CmdIndex(command.Key);
                if (!string.IsNullOrEmpty(value))
                    Console.WriteLine(command.Key + " : " + value);
                else
                    Console.WriteLine(command.Key);
            }

            return;
        }

        //files
        List<string> files = new List<string>();
        List<string> cmd = argCmd.Cmd(string.Empty);
        if (cmd != null)
        {
            for (int i = 0; i < cmd.Count; ++i)
            {
                var path = cmd[i];

                if (File.Exists(path))
                    files.Add(path);
                else if (Directory.Exists(path))
                    files.AddRange(Directory.GetFiles(path, "*.xlsx", SearchOption.TopDirectoryOnly).ToList());
            }
        }

        //filters
        if (argCmd.CmdCount("-filter") <= 0)
            argCmd.Add("-filter", string.Empty);

        List<string> filters = argCmd.Cmd("-filter");

        //headerOnly
        bool headerOnly = (argCmd.CmdCount("-json") < 0);

        //run
        for (int i = 0; i < files.Count; ++i)
        {
            foreach (var filter in filters)
            {
                //data
                string file = files[i];
                var datas = ConvertOpenXml.ToData(file, filter, headerOnly);

                //json
                if (argCmd.CmdCount("-json") >= 0)
                {
                    var jsons = ConvertOpenXml.ToJson(datas);

                    string dir = Path.Combine(Path.GetDirectoryName(file), filter, "json");
                    Directory.CreateDirectory(dir);

                    foreach (var json in jsons)
                    {
                        string outputFile = Path.Combine(dir, json.Key + ".json");
                        Console.WriteLine("        " + outputFile);

                        // Write the string array to a new file named "WriteLines.txt".
                        using (StreamWriter stream = new StreamWriter(outputFile))
                        {
                            stream.Write(json.Value);
                        }
                    }
                }

                //cs
                if (argCmd.CmdCount("-cs") >= 0)
                {
                    var css = ConvertOpenXml.ToCSharp(datas, argCmd.CmdIndex("-cs-namespace"), argCmd.CmdIndex("-cs-baseclass"));

                    string dir = Path.Combine(Path.GetDirectoryName(file), filter, "cs");
                    Directory.CreateDirectory(dir);

                    foreach (var cs in css)
                    {
                        string outputFile = Path.Combine(dir, cs.Key + ".cs");
                        Console.WriteLine("        " + outputFile);

                        // Write the string array to a new file named "WriteLines.txt".
                        using (StreamWriter stream = new StreamWriter(outputFile))
                        {
                            stream.Write(cs.Value);
                        }
                    }
                }
            }
        }
    }
}
