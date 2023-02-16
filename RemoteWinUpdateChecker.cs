using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please Provide an OU path");
            return;
        }

        if (args[0] == "--help")
        {
            Console.WriteLine("SimplyUpdate \nTo run use the following SimplyUpdate.exe \"<OUPATH>\"");
            return;
        }

        var entry = new DirectoryEntry("LDAP://" + args[0]);

        var searcher = new DirectorySearcher(entry)
        {
            Filter = "(&(objectClass=computer)(OperatingSystem=Windows*))",
            PageSize = int.MaxValue
        };

        var computers = searcher.FindAll().Cast<SearchResult>().Select(x => x.Properties["Name"][0].ToString()).ToList();

        var computersFullyUpdated = 0;
        var computersSupported = 0;
        var computersNoLongerSupported = 0;
        var computersUpdatesUnknown = 0;

        var computersOffline = 0;
        var computersOnline = 0;
        var totalComputers = computers.Count;

        Parallel.ForEach(computers, computerName =>
        {
            try
            {
                var ping = new Ping();
                var reply = ping.Send(computerName, 1000);
                var connectionTest = reply.Status == IPStatus.Success;

                if (connectionTest)
                {
                    var scope = new ManagementScope("\\\\" + computerName + "\\root\\cimv2");

                    var osQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");

                    var searcher2 = new ManagementObjectSearcher(scope, osQuery);

                    var results = searcher2.Get().Cast<ManagementObject>();

                    var rawVersionNumber = results.FirstOrDefault()?.GetPropertyValue("BuildNumber").ToString();

                    if (rawVersionNumber != null)
                    {
                        string versionNumber;
                        var outFormat = ConsoleColor.White;

                        switch (rawVersionNumber)
                        {
                            case "19041":
                                versionNumber = "2004";
                                outFormat = ConsoleColor.Red;
                                Interlocked.Increment(ref computersNoLongerSupported);
                                break;

                            case "19042":
                                versionNumber = "20H2";
                                outFormat = ConsoleColor.Red;
                                Interlocked.Increment(ref computersNoLongerSupported);
                                break;

                            case "19043":
                                versionNumber = "21H1";
                                outFormat = ConsoleColor.Red;
                                Interlocked.Increment(ref computersNoLongerSupported);
                                break;

                            case "19044":
                                versionNumber = "21H2";
                                outFormat = ConsoleColor.Yellow;
                                Interlocked.Increment(ref computersSupported);
                                break;

                            case "19045":
                                versionNumber = "22H2";
                                outFormat = ConsoleColor.Green;
                                Interlocked.Increment(ref computersFullyUpdated);
                                break;

                            default:
                                versionNumber = null;
                                Interlocked.Increment(ref computersUpdatesUnknown);
                                break;
                        }

                        if (versionNumber != null)
                        {
                            Console.ForegroundColor = outFormat;
                            Console.WriteLine($"{computerName} Version Number {versionNumber}");
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine($"{computerName} Version Number unknown");
                        }
                    }

                    Interlocked.Increment(ref computersOnline);
                }
                else
                {
                    Console.ForegroundColor = ConsoleColor.Gray;
                    Console.WriteLine($"{computerName} powered off");
                    Interlocked.Increment(ref computersOffline);
                    Interlocked.Increment(ref computersUpdatesUnknown);
                }
            }
            catch
            {
                Interlocked.Increment(ref computersOffline);
                Interlocked.Increment(ref computersUpdatesUnknown);
            }
        });

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine("");
        Console.WriteLine("");

        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine($"Update state unknown/Offline {computersUpdatesUnknown}");

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Computers on Unsupported versions {computersNoLongerSupported}");

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Computers on Supported versions {computersSupported}");

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Computers on Supported versions {computersFullyUpdated}");

        Console.WriteLine("");
        Console.ForegroundColor = ConsoleColor.Gray;

        Console.WriteLine($"Computers offline {computersOffline}");
        Console.WriteLine($"Computers online {computersOnline}");
        Console.WriteLine($"Total Computers {totalComputers}");

        Console.WriteLine("\nPress Enter To Continue");
        Console.ReadLine();

    }
}
