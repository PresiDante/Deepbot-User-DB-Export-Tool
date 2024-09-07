using System;
using System.Threading;
using System.Text;
using Websocket.Client;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Linq;
using OfficeOpenXml;

public class DeepbotWebsocketDataExtract
{     
    static async Task Main(string[] args)
    {
        Task<DeepbotAPI> deepbot = DeepbotAPI.CreateAsync();
    }
}

public class DeepbotAPI
{
    string apiKey = "";
    int option = -1;
    string apiURL = "ws://localhost:3337/";
    List<User> allUsers = new List<User>();
    bool processing = false;
    int offset = 0;

    public static async Task<DeepbotAPI> CreateAsync()
    {
        DeepbotAPI x = new DeepbotAPI();
        await x.InitialiseAsync();
        return x;
    }

    private DeepbotAPI() { }

    private async Task InitialiseAsync()
    {
        SetupAPIKey();
        SetupAPIOption();
        await WebsocketClient();
    }

    private async Task WebsocketClient()
    {
        var exitEvent = new ManualResetEvent(false);
        Uri uri = new(apiURL);

        using (var client = new WebsocketClient(uri))
        {
            client.ReconnectTimeout = TimeSpan.FromSeconds(30);
            client.ReconnectionHappened.Subscribe(info => Console.WriteLine($"Reconnection happened, type: {info.Type}"));
            client.MessageReceived.Subscribe(msg =>
            {
                Console.WriteLine($"Message received: {msg}");
                JObject response = JObject.Parse(msg.ToString());

                switch(option)
                {
                    case 1:
                        if (response["function"].ToString() == "get_users" && processing)
                        {
                            if (response["msg"].Count() > 0)
                            {
                                List<User> users = JsonConvert.DeserializeObject<List<User>>(response["msg"].ToString());
                                foreach(User person in users)
                                {
                                    allUsers.Add(person);
                                }
                                offset += 100;
                                client.Send($"api|get_users|{offset}|100");
                            }
                            else
                            {
                                option = -1;
                                processing = false;
                                ProduceCSVFile();
                            }
                        }
                        break;
                    case 2:
                        if (response["function"].ToString() == "get_users" && processing)
                        {
                            if (response["msg"].Count() > 0)
                            {
                                List<User> users = JsonConvert.DeserializeObject<List<User>>(response["msg"].ToString());
                                foreach (User person in users)
                                {
                                    allUsers.Add(person);
                                }
                                offset += 100;
                                client.Send($"api|get_users|{offset}|100");
                            }
                            else
                            {
                                option = -1;
                                processing = false;
                                ProduceFirebotXLSXFile();
                                ReadFirebotUserDatabase();
                            }
                        }
                        break;
                    default:
                        SetupAPIOption();
                        break;
                };
            });

            if (option == 0)
            {
                SetupAPIOption();
            }     

            client.Start();

            Task.Run(() =>
            {
                client.Send($"api|register|{apiKey}");
                switch (option)
                {
                    case 1:
                    case 2:
                        allUsers = new List<User>();
                        processing = true;
                        offset = 0;
                        client.Send("api|get_users|0|100");
                        break;
                    default:
                        break;
                }
            });

            exitEvent.WaitOne();
        }
    }

    private void SetupAPIKey()
    {
        while (string.IsNullOrEmpty(apiKey))
        {
            Console.WriteLine("Please enter the API key that you can find in Deepbot's master settings: ");
            apiKey = Console.ReadLine() ?? "";
        }
    }

    private void ProduceCSVFile()
    {
        var sb = new StringBuilder();
        sb.AppendLine("Username, Points, Watch Time, Join Date, Last Seen");
        foreach(User user in allUsers)
        {
            sb.AppendLine($"{user.Username},{user.Points},{user.WatchTime},{user.JoinDate},{user.LastSeen}");
        }
        File.WriteAllText("test.csv", sb.ToString());
        Console.WriteLine("File created!");
    }

    private void ProduceFirebotXLSXFile()
    {
        var filename = "points.xlsx";
        var file = new FileInfo(filename);

        if (file.Exists)
        {
            file.Delete();
        }

        using (var package = new ExcelPackage(file))
        {
            var ws = package.Workbook.Worksheets.Add("Currency");
            ws.Cells["A1"].Value = "Name";
            ws.Cells["B1"].Value = "Rank";
            ws.Cells["C1"].Value = "Points";
            ws.Cells["D1"].Value = "Hours";
            int i = 2;
            foreach (User user in allUsers)
            {
                ws.Cells[$"A{i}"].Value = user.Username;
                ws.Cells[$"B{i}"].Value = "";
                ws.Cells[$"C{i}"].Value = user.Points;
                ws.Cells[$"D{i}"].Value = user.WatchTime / 60;
                i++;
            }

            package.Save();
        }
        Console.WriteLine("File created!");
    }

    private void ReadFirebotUserDatabase()
    {
        List<FirebotUserDB> users = new List<FirebotUserDB>();

        var filename = "users.db";
        try
        {
            using (StreamReader sr = new StreamReader(filename))
            {
                string line;

                while((line = sr.ReadLine()) != null)
                {
                    users.Add(JsonConvert.DeserializeObject<FirebotUserDB>(line));
                }
            }
            Console.WriteLine("test");

            foreach(FirebotUserDB user in users)
            {
                if(allUsers.Exists(a => a.Username == user.Username))
                {
                    var userToUpdate = allUsers.FirstOrDefault(a => a.Username == user.Username);
                    user.LastSeen = ((DateTimeOffset)userToUpdate.LastSeen).ToUnixTimeMilliseconds();
                    user.JoinDate = ((DateTimeOffset)userToUpdate.JoinDate).ToUnixTimeMilliseconds();
                    if (user.Currency.HasValues)
                    {
                        JToken token = user.Currency.First;
                        user.Currency[token.Path] = userToUpdate.Points;
                        //user.Currency[token.] = userToUpdate.Points;
                    }
                }
            }

            StringBuilder sb = new StringBuilder();

            foreach(FirebotUserDB user in users)
            {
                sb.AppendLine(JsonConvert.SerializeObject(user));
            }
            File.WriteAllText(filename, sb.ToString());
        }
        catch
        {

        }


    }
    private void SetupAPIOption()
    {
        bool valid = false;

        while (!valid)
        {
            Console.WriteLine("Please type the number of the operation you are trying to perform \n1: Retrieve all users and export to CSV\n2: Retrieve all users and export to XLSX (For Firebot)\n0: Exit Program");
            
            if(Int32.TryParse(Console.ReadLine(),out option))
            {
                valid = true;
            }
        }
    }
}

public class GetUsers
{
    [JsonProperty("msg")]
    public List<User>? Users { get; set; }
}

public class User
{
    [JsonProperty("user")]
    public string? Username { get; set; }
    [JsonProperty("points")]
    public decimal Points { get; set; }
    [JsonProperty("watch_time")]
    public decimal WatchTime { get; set; }
    [JsonProperty("join_date")]
    public DateTime JoinDate { get; set; }
    [JsonProperty("last_seen")]
    public DateTime LastSeen { get; set; }
}

public class FirebotUserDB
{
    [JsonProperty("_id")]
    public int Id { get; set; }
    [JsonProperty("chatMessages")]
    public int ChatMessages { get; set; }
    [JsonProperty("currency")]
    public JObject? Currency { get; set; }
    [JsonProperty("disableActiveUserList")]
    public bool DisableActiveUserList { get; set; }
    [JsonProperty("disableAutoStatAccrual")]
    public bool DisableAutoStatAccrual { get; set; }
    [JsonProperty("disableViewerList")]
    public bool DisableViewerList { get; set; }
    [JsonProperty("displayName")]
    public string? DisplayName { get; set; }
    [JsonProperty("joinDate")]
    public long JoinDate { get; set; }
    [JsonProperty("lastSeen")]
    public long LastSeen { get; set; }
    // Need to double check what can be contained in this meta data. Not needed for exporting points
    // Leaving as a list of strings until we find out more
    [JsonProperty("metadata")]
    public JObject? Metadata { get; set; }
    [JsonProperty("minutesInChannel")]
    public decimal MinutesInChannel { get; set; }
    [JsonProperty("online")]
    public bool Online { get; set; }
    [JsonProperty("onlineAt")]
    public long OnlineAt { get; set; }
    [JsonProperty("profilePicUrl")]
    public string ProfilePicUrl { get; set; }
    // Need to double check what can be contained in this ranks object. Not needed for exporting points
    // Leaving as a list of strings until we find out more
    [JsonProperty("ranks")]
    public JObject? Ranks { get; set; }
    [JsonProperty("twitch")]
    public bool Twitch { get; set; }
    // Need to double check what can be contained in this twitchRoles object. Not needed for exporting points
    // Leaving as a list of strings until we find out more
    [JsonProperty("twitchRoles")]
    public JObject? TwitchRoles { get; set; }
    [JsonProperty("username")]
    public string Username { get; set; }
}


