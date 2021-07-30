// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Threading.Tasks;

namespace GraphTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(".NET Core Graph Tutorial to fetch sensitivity labels of all files in a drive\n");

            var appId = "";
            var scopesString = "User.Read;Files.Read";
            var scopes = scopesString.Split(';');

            // Initialize Graph client
            GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            });

            var accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;
            // </InitializationSnippet>

            // <GetUserSnippet>
            // Get signed in user
            var user = GraphHelper.GetMeAsync().Result;
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

            int choice = -1;

            while (choice != 0) {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Fetch Files");

                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch(choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;
                    case 1:
                        GetLabelsInALibrary();
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }

        static void GetLabelsInALibrary()
        {
            var files = GraphHelper.GetAllFilesAsync();
        }
    }
}
