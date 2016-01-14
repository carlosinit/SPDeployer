using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security;

namespace SPDeployer
{
    class Program
    {
        static List<IDisposable> streams = new List<IDisposable>();

        static int Main(string[] args)
        {
            try
            {
                var stopwatch = new Stopwatch();
                stopwatch.Start();

                if (args.Length != 5)
                {
                    throw new InvalidOperationException(
                        @"You must pass 5 parameters in the following order [webUrl] [user] [password] [folder] [sourceFolderPath]. 
                        Example: 
                        SPDeployer.exe http://mysupersite.sharepoint.com/ user@domain.com P@55w0rd SiteAssets/SomePlace C:\temp\files");
                }

                var webSiteUrl = ParseWebUrl(args[0]);
                var user = ParseUser(args[1]);
                var securedPassword = ParsePassword(args[2]);
                var applicationFolderPath = ParseApplicationFolderPath(args[3]);
                var sourceFolderPath = ParseSourceFolderPath(args[4]);

                using (var clientContext = new ClientContext(webSiteUrl))
                {
                    clientContext.Credentials = new SharePointOnlineCredentials(user, securedPassword);

                    var web = clientContext.Web;
                    var applicationFolder = web.GetFolderByServerRelativeUrl(applicationFolderPath);

                    clientContext.Load(web);
                    clientContext.Load(applicationFolder);

                    try
                    {
                        clientContext.ExecuteQuery();
                    }
                    catch (ServerException exception)
                    {
                        if (exception.ServerErrorTypeName == "System.IO.FileNotFoundException")
                        {
                            applicationFolder = web.Folders.Add(applicationFolderPath);
                            Console.WriteLine("Could not find application folder, creating it...");
                        }
                        else
                        {
                            throw;
                        }
                    }

                    CreateFolderStructure(new DirectoryInfo(sourceFolderPath), applicationFolder);
                    Console.WriteLine("Sending the files to SharePoint...");
                    clientContext.ExecuteQuery();

                }

                stopwatch.Stop();
                Console.WriteLine($"Operation completed after {stopwatch.Elapsed.TotalSeconds}s");
                return 0;
            }
            catch (Exception exception)
            {
                Console.WriteLine($"{exception.GetType()}: {exception.Message}");
                return -1;
            }
            finally
            {
                foreach (var item in streams)
                {
                    item.Dispose();
                }
            }
        }

        static string ParseSourceFolderPath(string sourceFolderPathString)
        {
            if (string.IsNullOrWhiteSpace(sourceFolderPathString))
            {
                throw new ArgumentException("Invalid source folder path");
            }

            if (!Directory.Exists(sourceFolderPathString))
            {
                throw new ArgumentException("Source folder path does not exist");
            }

            return sourceFolderPathString;
        }

        static string ParseApplicationFolderPath(string applicationFolderPath)
        {
            if (string.IsNullOrWhiteSpace(applicationFolderPath))
            {
                throw new ArgumentException("Invalid application folder name");
            }

            if (applicationFolderPath.StartsWith("/", StringComparison.Ordinal))
            {
                applicationFolderPath = applicationFolderPath.Substring(1);
            }
            return applicationFolderPath;
        }

        static SecureString ParsePassword(string passwordString)
        {
            var securedPassword = new SecureString();
            foreach (var c in passwordString)
            {
                securedPassword.AppendChar(c);
            }
            return securedPassword;

        }

        static string ParseUser(string userString)
        {
            if (!userString.Contains("@"))
            {
                throw new ArgumentException("Invalid user name");
            }

            return userString;
        }

        static string ParseWebUrl(string urlString)
        {
            if (!urlString.ToLowerInvariant().StartsWith("https://", StringComparison.Ordinal))
            {
                throw new ArgumentException("Invalid web url");
            }
            return urlString;
        }

        static void CreateFolderStructure(DirectoryInfo input, Folder output)
        {
            foreach (var item in input.GetFileSystemInfos())
            {
                if (item is DirectoryInfo)
                {
                    CreateFolderStructure(item as DirectoryInfo, output.Folders.Add(item.Name));
                }
                else
                {
                    var itemStream = (item as FileInfo).OpenRead();
                    streams.Add(itemStream);

                    var creationInfo = new FileCreationInformation();
                    creationInfo.Overwrite = true;
                    creationInfo.Url = item.Name;
                    creationInfo.ContentStream = itemStream;
                    output.Files.Add(creationInfo);


                }
            }
        }
    }
}
