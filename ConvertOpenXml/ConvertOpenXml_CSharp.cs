//ConvertOpenXml
static partial class ConvertOpenXml
{
    //ToCSharp
    static public Dictionary<string, string> ToCSharp(Dictionary<string, Table> tables, string nameSpace = "", string baseClass = "")
    {
        var cs = new Dictionary<string, string>();

        if (tables != null)
        {
            foreach (var table in tables)
            {
                // Serialize to JSON
                if (table.Value.datas.Count > 1)
                {
                    cs[table.Key] = "";

                    string space = string.Empty;

                    cs[table.Key] += "using System.Collections.Generic;\n";
                    cs[table.Key] += "\n";

                    if (!string.IsNullOrEmpty(nameSpace))
                    {
                        cs[table.Key] += "namespace " + nameSpace + "\n";
                        cs[table.Key] += "{\n";
                        space = "    ";
                    }

                    if (!string.IsNullOrEmpty(baseClass))
                        cs[table.Key] += space + "public class " + table.Key + " : " + baseClass + "\n";
                    else
                        cs[table.Key] += space + "public class " + table.Key + "\n";

                    cs[table.Key] += space + "{\n";
                    cs[table.Key] += space + "    public class _\n";
                    cs[table.Key] += space + "    {\n";

                    string keyType = string.Empty;

                    foreach (var header in table.Value.headers)
                    {
                        if (keyType == string.Empty)
                            keyType = ConvertType(header.Value.format);

                        cs[table.Key] += space + "        public "
                            + ConvertType(header.Value.format)
                            + " "
                            + header.Value.name
                            + " = "
                            + ConvertInit(header.Value.format)
                            + ";\n";
                    }

                    cs[table.Key] += space + "    }\n";
                    cs[table.Key] += "\n";
                    cs[table.Key] += space + "    public Dictionary<"
                        + keyType
                        + ", _> datas = new Dictionary<"
                        + keyType
                        + ", _>();\n";

                    cs[table.Key] += space + "}\n";

                    if (!string.IsNullOrEmpty(nameSpace))
                        cs[table.Key] += "}\n";
                }
            }
        }

        return cs;
    }

    //ConvertDef
    static private string ConvertType(string fmt)
    {
        bool toList = fmt.Contains("[");
        if (toList)
        {
            int startIndex = fmt.IndexOf('[');
            fmt = fmt.Substring(0, startIndex).Trim();
        }

        string s = string.Empty;

        switch (fmt)
        {
            case "b":
                s = "bool";
                break;
            case "i32":
                s = "int";
                break;
            case "u32":
                s = "uint";
                break;
            case "i64":
                s = "long";
                break;
            case "u64":
                s = "ulong";
                break;
            case "f":
                s = "float";
                break;
            case "d":
                s = "double";
                break;
            case "s":
                s = "string";
                break;
        }

        if (toList)
            s = "List<" + s + ">";

        return s;
    }

    //ConvertInit
    static private string ConvertInit(string fmt)
    {
        bool toList = fmt.Contains('[');
        if (toList)
        {
            int startIndex = fmt.IndexOf('[');
            fmt = fmt.Substring(0, startIndex).Trim();
        }

        string s = string.Empty;

        switch (fmt)
        {
            case "b":
                s = toList ? "new List<bool>()" : "false";
                break;
            case "i32":
                s = toList ? "new List<int>()" : "0";
                break;
            case "u32":
                s = toList ? "new List<uint>()" : "0";
                break;
            case "i64":
                s = toList ? "new List<long>()" : "0";
                break;
            case "u64":
                s = toList ? "new List<ulong>()" : "0";
                break;
            case "f":
                s = toList ? "new List<float>()" : "0.0f";
                break;
            case "d":
                s = toList ? "new List<double>()" : "0.0";
                break;
            case "s":
                s = toList ? "new List<string>()" : "string.Empty";
                break;
        }

        return s;
    }
}
