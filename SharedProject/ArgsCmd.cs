//ArgsCmd
public class ArgsCmd
{
    public OrderedDictionary<string, List<string>> cmds = new OrderedDictionary<string, List<string>>();

    //IsKey
    static private bool IsKey(string arg, string s = "-")
    {
        if (!arg.StartsWith(s))
            return false;

        return arg.Length <= s.Length || !char.IsDigit(arg[s.Length]);
    }

    //Parser
    static public ArgsCmd Parser(string[] args, string s = "-")
    {
        ArgsCmd ret = new ArgsCmd();

        if (args == null)
            return ret;

        string key = string.Empty;
        for (int i = 0; i < args.Length; ++i)
        {
            var arg = args[i];

            if (IsKey(arg, s))
                ret.Add(key = arg);
            else
                ret.Add(key, arg);
        }

        return ret;
    }

    //Add
    public bool Add(string key)
    {
        if (!cmds.ContainsKey(key))
        {
            cmds[key] = new List<string>();
            return true;
        }

        return false;
    }

    public bool Add(string key, string value)
    {
        bool ret = Add(key);
        cmds[key].Add(value);
        return ret;
    }

    //Remove
    public bool Remove(string key)
    {
        return cmds.Remove(key);
    }

    public bool Remove(string key, string value)
    {
        if (cmds.ContainsKey(key))
            return cmds[key].Remove(value);

        return false;
    }

    //Cmd
    public List<string> Cmd(string key)
    {
        if (cmds.ContainsKey(key))
            return cmds[key];

        return null;
    }

    //CmdIndex
    public string CmdIndex(string key, int index = 0, string d = "")
    {
        var cmd = Cmd(key);
        if (cmd != null && cmd.Count > index)
            return cmd[index];

        return d;
    }

    //CmdCount
    public int CmdCount(string key)
    {
        var cmd = Cmd(key);
        if (cmd != null)
            return cmd.Count;

        return -1;
    }
}
