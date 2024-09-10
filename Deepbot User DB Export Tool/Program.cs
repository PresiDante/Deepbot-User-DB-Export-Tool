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
    static DeepbotAPI deepbot = new DeepbotAPI();
    static int option = -1;
    static int offset = 0;
    static bool processing = false;

    static async Task Main(string[] args)
    {
        deepbot.SetAPIKey();
        deepbot.SetAPIIP();
        option = SetupAPIOption();
        await WebsocketClient();
        Console.WriteLine("Program finished!");
    }

    private static async Task WebsocketClient()
    {
        var exitEvent = new ManualResetEvent(false);
        Uri uri = new(deepbot.WebsocketURL());

        using (var client = new WebsocketClient(uri))
        {
            client.ReconnectTimeout = TimeSpan.FromSeconds(30);
            client.ReconnectionHappened.Subscribe(info =>
            {
                Console.WriteLine($"Reconnection happened, type: {info.Type}");
                client.Send(deepbot.WebsocketAPIRegisterCall());
            });

            client.MessageReceived.Subscribe(async msg =>
            {
                JObject response;

                Console.WriteLine(msg);

                try
                {
                    response = JObject.Parse(msg.ToString());
                }
                catch
                {
                    Console.WriteLine($"Deepbot sent {msg} but we are choosing to ignore this.");
                    return;
                }

                if (response["function"] == null || response["msg"] == null)
                {
                    Console.WriteLine("Function or Msg params null. Continuing");
                    return;
                }

                var function = response["function"]?.ToString() ?? "";
                var message = response["msg"]?.ToString() ?? "";

                var updated = true;

                if (function == "get_users" && processing)
                {
                    if (message == "list empty")
                    {
                        updated = false;
                    }
                    else
                    {
                        updated = deepbot.UpdateAllUsersList(JsonConvert.DeserializeObject<List<User>>(message));

                        if (updated)
                        {
                            offset += 100;
                            client.Send(deepbot.WebsocketGetUsersCall(offset));
                            return;
                        }
                    }
                }

                switch (option)
                {
                    case 1: // Create general purpose .csv file with figures that can be acquired through Deepbot websocket API
                        if (!updated)
                        {
                            offset = 0;
                            option = -1;
                            processing = false;
                            deepbot.ProduceCSVFile();
                            option = SetupAPIOption();
                            await SendWebsocketMessageOption(client);
                        }
                        break;
                    case 2: // Create .xlsx file for import to Firebot
                        if (!updated)
                        {
                            option = -1;
                            processing = false;
                            deepbot.ProduceFirebotXLSXFile();
                            option = SetupAPIOption();
                            await SendWebsocketMessageOption(client);
                        }
                        break;
                    case 3: // Update Firebot user.db directly
                        if (!updated)
                        {
                            option = -1;
                            processing = false;
                            deepbot.ReadFirebotUserDatabase();
                            option = SetupAPIOption();
                            await SendWebsocketMessageOption(client);
                        }
                        break;
                    default:
                        break;
                };

                if (option == 0)
                {
                    await client.Stop(System.Net.WebSockets.WebSocketCloseStatus.NormalClosure, "Exit called!");
                    exitEvent.Set();
                }
            });

            await client.Start();

            await Task.Run(async () =>
            {
                client.Send(deepbot.WebsocketAPIRegisterCall());
                await SendWebsocketMessageOption(client);
            });

            exitEvent.WaitOne();
        }
    }

    public static async Task<Task> SendWebsocketMessageOption(WebsocketClient client)
    {
        switch (option)
        {
            case 0:
                //client.Send(deepbot.WebsocketGetUsersCall(offset));
                break;
            case 1:
            case 2:
            case 3:
                deepbot.ClearAllUsersList();
                processing = true;
                offset = 0;
                client.Send(deepbot.WebsocketGetUsersCall(offset));
                break;
            default:
                break;
        }

        return Task.CompletedTask;
    }

    private static int SetupAPIOption()
    {
        bool valid = false;
        int option = -1;

        while (!valid)
        {
            Console.WriteLine("Please type the number of the operation you are trying to perform \n1: Retrieve all users and export to CSV\n2: Retrieve all users and export to XLSX (For Firebot)\n3: Attempt to update Firebot user.db with currency information\n0: Exit Program");

            if (Int32.TryParse(Console.ReadLine(), out option))
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

public class DeepbotAPI
{
    private string apiKey = "";
    private string apiIP = "localhost";
    private int apiPort = 3337;
    private List<User> allUsers = new List<User>();

    public DeepbotAPI() { }

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

    /// <summary>
    /// Saves list of users to a variable for further use.
    /// </summary>
    /// <param name="users">List of users to process</param>
    /// <returns>False if there are more users to process, True otherwise</returns>
    public bool UpdateAllUsersList(List<User>? users)
    {
        if (users == null || users.Count == 0)
        {
            Console.WriteLine("No users to process.");
            return false;
        }

        foreach (User person in users)
        {
            if (string.IsNullOrEmpty(person.Username) || person.Username.Contains(" "))
            {
                Console.WriteLine($"User with potential invalid username skipped. Username: {person.Username}");
                continue;
            }
            allUsers.Add(person);
        }

        Console.WriteLine($"Retrieved {allUsers.Count} users");
        return true;
    }

    public void ClearAllUsersList()
    {
        allUsers = new List<User>();
    }

    public void SetAPIKey()
    {
        while (string.IsNullOrEmpty(apiKey))
        {
            Console.WriteLine("Please enter the API key that you can find in Deepbot's master settings: ");
            apiKey = Console.ReadLine() ?? "";
        }
    }

    public void SetAPIIP()
    {
        Console.WriteLine("Please enter the IP address where Deepbot is hosted. If this is running locally, press enter: ");
        var input = Console.ReadLine() ?? "";

        apiIP = string.IsNullOrEmpty(input) ? "localhost" : input;
    }

    //TODO: Potentially make use of async for the file operations
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
                string? line;

                while((line = sr.ReadLine()) != null)
                {
                    if (!string.IsNullOrEmpty(line))
                    {
                        var item = JsonConvert.DeserializeObject<FirebotUserDB>(line);
                        if (item != null)
                        {
                            users.Add(item);
                        }
                    }
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
                    
                    if (userToUpdate == default)
                    {
                        continue;
                    }

                    user.LastSeen = ((DateTimeOffset)userToUpdate.LastSeen).ToUnixTimeMilliseconds();
                    user.JoinDate = ((DateTimeOffset)userToUpdate.JoinDate).ToUnixTimeMilliseconds();
                    if (user.Currency.HasValues && user.Currency.First != null)
                    {
                        JToken token = user.Currency.First;
                        user.Currency[token.Path] = userToUpdate.Points;
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
        catch (FileNotFoundException)
        {
            Console.WriteLine("User.db was not found");
        }    
    }
}

public class GetUsers
{
    [JsonProperty("msg")]
    public List<User> Users { get; set; } = new List<User>();
}

public class User
{
    [JsonProperty("user")]
    public string Username { get; set; } = String.Empty;
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
    public string Username { get; set; } = String.Empty;
    [JsonProperty("_id")]
    public string Id { get; set; } = String.Empty;
    [JsonProperty("displayName")]
    public string DisplayName { get; set; } = String.Empty;
    [JsonProperty("profilePicUrl")]
    public string ProfilePicUrl { get; set; } = String.Empty;
    [JsonProperty("twitch")]
    public bool Twitch { get; set; }
    [JsonProperty("twitchRoles")]
    public JObject TwitchRoles { get; set; } = new JObject();
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
    public JObject Metadata { get; set; } = new JObject();
    [JsonProperty("currency")]
    public JObject Currency { get; set; } = new JObject();
    [JsonProperty("ranks")]
    public JObject Ranks { get; set; } = new JObject();
}


