using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace EInvoice.API.Services;

public class DocumentSerializationService
{
    public string Serialize(string content, JsonSerializerSettings jsonSettings)
    {
        var request = JsonConvert.DeserializeObject<JObject>(content, jsonSettings);

        var canonicalString = Serialize(request);

        return canonicalString;
    }

    public string Serialize(JObject request)
    {
        return SerializeToken(request);
    }

    private string SerializeToken(JToken request)
    {
        var serialized = "";

        if (request.Parent is null)
        {
            SerializeToken(request.First);
        }
        else
        {
            if (request.Type == JTokenType.Property)
            {
                string name = ((JProperty)request).Name.ToUpper();
                serialized += "\"" + name + "\"";
                foreach (var property in request)
                {
                    if (property.Type == JTokenType.Object)
                    {
                        serialized += SerializeToken(property);
                    }
                    if (property.Type == JTokenType.Boolean || property.Type == JTokenType.Integer || property.Type == JTokenType.Float || property.Type == JTokenType.Date)
                    {
                        serialized += "\"" + property.Value<string>() + "\"";
                    }
                    if (property.Type == JTokenType.String)
                    {
                        serialized += JsonConvert.ToString(property.Value<string>());
                    }
                    if (property.Type == JTokenType.Array)
                    {
                        foreach (var item in property.Children())
                        {
                            serialized += "\"" + ((JProperty)request).Name.ToUpper() + "\"";
                            serialized += SerializeToken(item);
                        }
                    }
                }
            }
            // Added to fix "References"
            if (request.Type == JTokenType.String)
            {
                serialized += JsonConvert.ToString(request.Value<string>());
            }
        }
        if (request.Type == JTokenType.Object)
        {
            foreach (var property in request.Children())
            {

                if (property.Type == JTokenType.Object || property.Type == JTokenType.Property)
                {
                    serialized += SerializeToken(property);
                }
            }
        }

        return serialized;
    }
}