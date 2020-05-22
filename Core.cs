using System;
using System.Collections.Generic;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace MS_Teams_API
{
    public class Core
    {
        private EnvironmentVariableTarget chromeOptions;
        private IWebDriver driver;
        private string currentURL;
        private WebDriverWait wait;
        private string teamsURL;
        private bool headless;
        private string username;
        private string password;
        private IList<IWebElement> teamNameElements;
        private string[] teamNames;
        private string command;
        private bool running;
        public Core(string[] args)
        {
            Console.WriteLine("Creating Teams API WebDriver Core");
            Console.WriteLine();
            // this.driver = new ChromeDriver("D:\\Source\\Libraries");
            this.currentURL = "";
            this.teamsURL = "http://teams.microsoft.com/";
            this.headless = true;
            this.username = "";
            this.password = "";
            this.command = "";
            this.running = true;
            this.HandleArgs(args);
            this.StartDriver();
            Console.WriteLine("Created WebDriver Core");
            Console.WriteLine();
        }
        public int StartDriver()
        {
            if (!this.headless)
            {
                this.driver = new ChromeDriver();
            }
            else
            {
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddArguments("headless");
                chromeOptions.AddArgument("--log-level=3");
                this.driver = new ChromeDriver(chromeOptions);
            }
            this.wait = new WebDriverWait(this.driver, TimeSpan.FromSeconds(20));
            return 0;
        }
        public int Help()
        {
            Console.WriteLine("Welcome to the (work in progress) alternate Microsoft Teams Client! At the moment, functionality is very limited.");
            Console.WriteLine("In the future there will be environment settings to be configured, but currently your email and password must be supplied with command line arguments");
            Console.WriteLine("Use the `/cred` switch and then provide your email and password. Ex. `MS_Teams_CLI /cred myEmail myPassword");
            Console.WriteLine("For debugging, you can view the Web Driver using the switch /head.");
            Console.WriteLine();
            Console.WriteLine("List of commands: ");
            Console.WriteLine("SetupSession - Starts Microsoft Teams and loads the home page.");
            Console.WriteLine("GetTeamNames - Gets the names of all teams you are in.");
            Console.WriteLine("DisplayTeamNames - Displays the names of all the teams you are in.");
            Console.WriteLine("Quit - Exits the application.");
            Console.WriteLine();
            return 0;
        }
        public int HandleArgs(string[] args)
        {
            int returnValue = 0;

            Console.WriteLine();
            switch (args.Length) {
                case 0:
                    returnValue = 0;
                    break;
                case 1:
                    if (args[0] == "/cred")
                    {
                        Console.WriteLine("No credentials provided with /cred switch.");
                        returnValue = 1;
                    }
                    else if (args[0] == "/head")
                    {
                        Console.WriteLine("Not running with headless configurations.");
                        this.headless = false;
                        returnValue = 0;
                    }
                    break;
                case 2:
                    if (args[0] == "/cred")
                    {
                        Console.WriteLine("Incorrect format for /cred switch. You most likely forgot to give either the username or password.");
                        returnValue = 1;
                    }
                    else if (args[0] == "/head")
                    {
                        Console.WriteLine("Switch /head does not take any arguments. Try again without `" + args[1] + "`.");
                        returnValue = 1;
                    }

                    break;
                case 3:
                    if (args[0] == "/cred")
                    {
                        Console.WriteLine("Setting username and password for this session to `" + args[1] + "` and `" + args[2] + "`.");
                        this.username = args[1];
                        this.password = args[2];
                        returnValue = 0;
                    }
                    else if (args[0] == "/head")
                    {
                        Console.WriteLine("Switch /head does not take any arguments. Try again without `" + args[1] + "and `" + args[2] + "`.");
                        returnValue = 1;
                    }
                    break;
                case 4:
                    if (args[0] == "/cred")
                    {
                        Console.WriteLine("Setting username and password for this session to `" + args[1] + "` and `" + args[2] + "`.");
                        this.username = args[1];
                        this.password = args[2];
                        if (args[3] == "/head")
                            {
                            this.headless = false;
                            returnValue = 0;
                        }
                        else
                        {
                            Console.WriteLine("Switch `/cred` takes two arguments, not three. Try again without `" + args[3] + "`.");
                            returnValue = 1;
                        }
                    }
                    if (args[0] == "/head")
                    {
                        Console.WriteLine("Not running with headless configurations.");
                        this.headless = false;
                        if (args[1] == "/cred")
                        {
                            Console.WriteLine("Setting username and password for this session to `" + args[1] + "` and `" + args[2] + "`.");
                            this.username = args[1];
                            this.password = args[2];
                            returnValue = 0;
                        }
                        else
                        {
                            Console.WriteLine("Switch `/head` takes no arguments. Try again without `" + args[1] + "`, `" + args[2] + "`, and " + args[3] + "`.");
                            returnValue = 1;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Bad switches(s)");
                        returnValue = 1;
                    }
                    break;

                default:
                    returnValue = 1;
                    break;
            }
            
            return returnValue;
        }
        public int Run()
        {
            while (this.running)
            {
                Console.WriteLine();
                Console.Write("~$ ");
                this.command = Console.ReadLine();
                switch (this.command)
                {
                    case ("Help"):
                        this.Help();
                        break;

                    case ("SetupSession"):
                        this.SetupSession();
                        break;

                    case ("GetTeamNames"):
                        this.GetTeamNames();
                        break;

                    case ("DisplayTeamNames"):
                        this.DisplayTeamNames();
                        break;

                    case ("Quit"):
                        this.running = false;
                        break;

                    default:
                        Console.WriteLine("`" + this.command + "` is not a recognised command.");
                        break;
                }
            }
       
            return 0;
        }
        public int SetupSession()
        {
            Console.WriteLine("Setting up session...");
            this.GoToTeams();
            this.EnterEmail();
            this.EnterPassword();
            this.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("team-name-text")));
            Console.WriteLine("Done!");
            return 0;
        }
        public int GoToTeams()
        {
            Console.WriteLine("Loading teams");
            driver.Navigate().GoToUrl(this.teamsURL);
            this.currentURL = this.teamsURL;
            return 0;
        }
        public int EnterEmail()
        {
            Console.WriteLine("Entering email");
            this.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Name("loginfmt")));
            driver.FindElement(By.Name("loginfmt")).SendKeys(this.username + Keys.Enter);
            return 0;
        }
        public int EnterPassword()
        {
            Console.WriteLine("Entering password");
            this.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.Name("Password")));;
            driver.FindElement(By.Name("Password")).SendKeys(this.password + Keys.Enter);
            return 0;
        }
        public int GetTeamNames()
        {
            this.driver.Navigate().GoToUrl("https://teams.microsoft.com/_#/school//?ctx=teamsGrid");
            this.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("team-name-text")));
            this.teamNameElements = driver.FindElements(By.ClassName("team-name-text"));

            this.teamNames = new String[teamNameElements.Count];

            int i = 0;
            foreach (IWebElement element in teamNameElements)
            {
                teamNames[i++] = element.Text;
            }

            return 0;
        }
        public int DisplayTeamNames()
        {
            int i = 0;
            foreach (string teamName in this.teamNames)
            {
                Console.WriteLine(i.ToString() + " - " + teamName);
                i++;
            }

            return 0;
        }
      
    }

}
