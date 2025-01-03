using Newtonsoft.Json;

//ConvertOpenXml
static partial class ConvertOpenXml
{
    //ToJson
    static public Dictionary<string, string> ToJson(Dictionary<string, Table> tables)
    {
        var jsons = new Dictionary<string, string>();

        if (tables != null)
        {
            foreach (var table in tables)
            {
                // Serialize to JSON
                if (table.Value.datas.Count > 1)
                {
                    string key = table.Value.headers.First().Value.name;
                    OrderedDictionary<string, object> datas = new OrderedDictionary<string, object>();

                    for (int i = 0; i < table.Value.datas.Count; ++i)
                    {
                        var data = table.Value.datas[i];

                        try
                        {
                            datas.Add(data[key].ToString(), table.Value.datas[i]);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(table.Key + " " + data[key].ToString() + " - " + ex.ToString());
                        }
                    }

                    jsons[table.Key] = "{\"datas\":\n"
                        + JsonConvert.SerializeObject(datas, Formatting.Indented)
                        + "\n}\n";
                }
            }
        }

        return jsons;
    }
}
