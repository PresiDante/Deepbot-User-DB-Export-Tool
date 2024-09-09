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
    static DeepbotAPI deepbot;
    static int option = -1;
    static int offset = 0;
    static bool processing = false;

    static async Task Main(string[] args)
    {
        deepbot = new DeepbotAPI();
        option = deepbot.SetupAPIOption();
        await WebsocketClient();
    }

    private static async Task WebsocketClient()
    {
        var exitEvent = new ManualResetEvent(false);
        Uri uri = new(deepbot.WebsocketURL());

        using (var client = new WebsocketClient(uri))
        {
            client.ReconnectTimeout = TimeSpan.FromSeconds(30);
            client.ReconnectionHappened.Subscribe(info => Console.WriteLine($"Reconnection happened, type: {info.Type}"));
            client.MessageReceived.Subscribe(async msg =>
            {
                JObject response = JObject.Parse(msg.ToString());

                switch (option)
                {
                    case 1: // Create general purpose .csv file with figures that can be acquired through Deepbot websocket API
                        if (response["function"].ToString() == "get_users" && processing)
                        {
                            if (response["msg"].Count() > 0)
                            {
                                List<User> users = JsonConvert.DeserializeObject<List<User>>(response["msg"].ToString());
                                await deepbot.UpdateAllUsersList(users);
                                offset += 100;
                                client.Send($"api|get_users|{offset}|100");
                            }
                            else
                            {
                                option = -1;
                                processing = false;
                                deepbot.ProduceCSVFile();
                                option = deepbot.SetupAPIOption();
                                SendWebsocketMessageOption(client);
                            }
                        }
                        break;
                    case 2: // Create .xlsx file for import to Firebot
                        if (response["function"].ToString() == "get_users" && processing)
                        {
                            if (response["msg"].Count() > 0)
                            {
                                List<User> users = JsonConvert.DeserializeObject<List<User>>(response["msg"].ToString());
                                await deepbot.UpdateAllUsersList(users);
                                offset += 100;
                                client.Send(deepbot.WebsocketGetUsersCall(offset));
                            }
                            else
                            {
                                option = -1;
                                processing = false;
                                deepbot.ProduceFirebotXLSXFile();
                                option = deepbot.SetupAPIOption();
                                SendWebsocketMessageOption(client);
                            }
                        }
                        break;
                    case 3: // Update Firebot user.db directly
                        if (response["function"].ToString() == "get_users" && processing)
                        {
                            if (response["msg"].Count() > 0)
                            {
                                List<User> users = JsonConvert.DeserializeObject<List<User>>(response["msg"].ToString());
                                await deepbot.UpdateAllUsersList(users);
                                offset += 100;
                                client.Send(deepbot.WebsocketGetUsersCall(offset));
                            }
                            else
                            {
                                option = -1;
                                processing = false;
                                deepbot.ReadFirebotUserDatabase();
                                option = deepbot.SetupAPIOption();
                                SendWebsocketMessageOption(client);
                            }
                        }
                        break;
                    default:
                        option = deepbot.SetupAPIOption();
                        SendWebsocketMessageOption(client);
                        break;
                };
            });

            if (option == 0)
            {
                //TODO: Close websocket connection properly to exit the program
                exitEvent.Close();
            }

            client.Start();

            Task.Run(async () =>
            {
                client.Send(deepbot.WebsocketAPIRegisterCall());
                SendWebsocketMessageOption(client);
            });

            exitEvent.WaitOne();
        }
    }

    public static async Task<Task> SendWebsocketMessageOption(WebsocketClient client)
    {
        switch (option)
        {
            case 1:
            case 2:
            case 3:
                await deepbot.ClearAllUsersList();
                processing = true;
                offset = 0;
                client.Send(deepbot.WebsocketGetUsersCall(offset));
                break;
            default:
                break;
        }

        return Task.CompletedTask;
    }

}

public class DeepbotAPI
{
    private string apiKey = "";
    private string apiIP = "localhost";
    private int apiPort = 3337;
    private List<User> allUsers = new List<User>();

    public DeepbotAPI() 
    {
        SetAPIKey();
        SetAPIIP();
    }

    public string WebsocketURL()
    {
        return $"ws://{apiIP}:{apiPort}/";
    }

    public string WebsocketAPIRegisterCall()
    {
        return $"api|register|{apiKey}";
    }

    public string WebsocketGetUsersCall(int offset)
    {
        return $"api|get_users|{offset}|100";
    }

    public async Task<Task> UpdateAllUsersList(List<User> users)
    {
        foreach (User person in users)
        {
            allUsers.Add(person);
        }

        Console.WriteLine($"Retrieved {allUsers.Count} users");
            
        return Task.CompletedTask;
    }

    public async Task<Task> ClearAllUsersList()
    {
        allUsers = new List<User>();
        return Task.CompletedTask;
    }

    private void SetAPIKey()
    {
        while (string.IsNullOrEmpty(apiKey))
        {
            Console.WriteLine("Please enter the API key that you can find in Deepbot's master settings: ");
            apiKey = Console.ReadLine() ?? "";
        }
    }

    private void SetAPIIP()
    {
        Console.WriteLine("Please enter the IP address where Deepbot is hosted. If this is running locally, press enter: ");
        var input = Console.ReadLine() ?? "";

        apiIP = string.IsNullOrEmpty(input) ? "localhost" : input;
    }

    public void ProduceCSVFile()
    { 
        var filename = "deepbotUsers.csv";
        var file = new FileInfo(filename);

        if (file.Exists)
        {
            file.Delete();
        }

        Console.WriteLine("Producing CSV file with retrieved data. Please wait...");

        var sb = new StringBuilder();
        sb.AppendLine("Username, Points, Watch Time, Join Date, Last Seen");
        
        foreach(User user in allUsers)
        {
            sb.AppendLine($"{user.Username},{user.Points},{user.WatchTime},{user.JoinDate},{user.LastSeen}");
        }

        File.WriteAllText(filename, sb.ToString());
        Console.WriteLine("File created!");
    }

    public void ProduceFirebotXLSXFile()
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

    /// <summary>
    /// This function currently works on the assumption that the user database is created after import.
    /// </summary>
    public void ReadFirebotUserDatabase()
    {
        List<FirebotUserDB> users = new List<FirebotUserDB>();
        List<string> usernamesNotImported = new List<string>();
        string filename = "users.db";
        string errorFilename = $"userDidntExistForImport{DateTime.Now.ToString("yyyyMMddTHHmmss")}.csv";
        int usersUpdated = 0;

        Console.WriteLine("Attempting to set values (Currency, Last Seen, Join date) in user.db.\nThis process can take a while. Please wait...");
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

            foreach (User DeepbotUser in allUsers)
            {
                if(!users.Exists(a => a.Username == DeepbotUser.Username))
                {
                    usernamesNotImported.Add(DeepbotUser.Username);
                }
            }

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
                    usersUpdated++;
                }
            }

            StringBuilder sb = new StringBuilder();

            foreach(FirebotUserDB user in users)
            {
                sb.AppendLine(JsonConvert.SerializeObject(user));
            }
            File.WriteAllText(filename, sb.ToString());
            Console.WriteLine($"Firebot user.db successfully modified. {usersUpdated} user(s) was/were updated.");

            if (usernamesNotImported.Any())
            {
                StringBuilder sb2 = new StringBuilder();

                foreach(string username in usernamesNotImported)
                {
                    sb2.AppendLine(username);
                }
                File.WriteAllText(errorFilename, sb2.ToString());
                Console.WriteLine($"{usernamesNotImported.Count} user(s) didn't exist in Firebot database. This could be because they no longer exist on Twitch. See {errorFilename} for further details of usernames not imported");
            }
        }
        catch (FileNotFoundException e)
        {
            Console.WriteLine("User.db was not found");
        }    
    }

    public int SetupAPIOption()
    {
        bool valid = false;
        int option = -1;

        while (!valid)
        {
            Console.WriteLine("Please type the number of the operation you are trying to perform \n1: Retrieve all users and export to CSV\n2: Retrieve all users and export to XLSX (For Firebot)\n3: Attempt to update Firebot user.db with currency information\n0: Exit Program");
            
            if(Int32.TryParse(Console.ReadLine(),out option))
            {
                if (option >= 0)
                {
                    valid = true;
                }
            }
        }

        return option;
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
    [JsonProperty("username")]
    public string Username { get; set; }
    [JsonProperty("_id")]
    public string Id { get; set; }
    [JsonProperty("displayName")]
    public string? DisplayName { get; set; }
    [JsonProperty("profilePicUrl")]
    public string ProfilePicUrl { get; set; }
    [JsonProperty("twitch")]
    public bool Twitch { get; set; }
    [JsonProperty("twitchRoles")]
    public JObject? TwitchRoles { get; set; }
    [JsonProperty("online")]
    public bool Online { get; set; }
    [JsonProperty("onlineAt")]
    public long OnlineAt { get; set; }
    [JsonProperty("lastSeen")]
    public long LastSeen { get; set; }
    [JsonProperty("joinDate")]
    public long JoinDate { get; set; }
    [JsonProperty("minutesInChannel")]
    public decimal MinutesInChannel { get; set; }
    [JsonProperty("chatMessages")]
    public int ChatMessages { get; set; }
    [JsonProperty("disableAutoStatAccrual")]
    public bool DisableAutoStatAccrual { get; set; }
    [JsonProperty("disableActiveUserList")]
    public bool DisableActiveUserList { get; set; }
    [JsonProperty("disableViewerList")]
    public bool DisableViewerList { get; set; }
    [JsonProperty("metadata")]
    public JObject? Metadata { get; set; }
    [JsonProperty("currency")]
    public JObject? Currency { get; set; }
    [JsonProperty("ranks")]
    public JObject? Ranks { get; set; }
}


