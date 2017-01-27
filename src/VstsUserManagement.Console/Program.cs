using System;
using System.Linq;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Identity;
using Microsoft.VisualStudio.Services.Identity.Client;
using Microsoft.VisualStudio.Services.Licensing;
using Microsoft.VisualStudio.Services.Licensing.Client;
using System.Diagnostics;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Channel;
using Microsoft.ApplicationInsights.Extensibility;
using Microsoft.ApplicationInsights.DataContracts;
using CommandLine;
using System.IO;
using Microsoft.TeamFoundation.Server;
using Microsoft.TeamFoundation.Core.WebApi;
using System.Collections.Generic;
using System.Text;

namespace VstsUserManagement.ConsoleApp
{
    public class Program
    {
        abstract class BaseOptions
        {
            [Option('n', "VssAccountName", Required = true, HelpText = "VssAccountName")]
            public string VssAccountName { get; set; }
            [Option('u', "VssAccountUrl", Required = true, HelpText = "VssAccountUrl")]
            public string VssAccountUrl { get; set; }
            [Option('p', "project", Required = false, HelpText = "Project to add users to")]
            public string Project { get; set; }
            [Option('g', "group", Required = false, HelpText = "Group in Project to add users to")]
            public string Group { get; set; }
        }

        [Verb("adduser", HelpText = "Add User Directly")]
        class AddUser : BaseOptions
        {
            [Option('m', "VssUserToAddMailAddress", Required = true, HelpText = "VssUserToAddMailAddress")]
            public string VssUserToAddMailAddress { get; set; }
            [Option('l', "VssLicense", Required = true, HelpText = "VssLicense")]
            public string VssLicense { get; set; }
            
        }
        [Verb("addusers", HelpText = "Add Users from file.")]
        class AddUsers : BaseOptions
        {
            [Option('c', "Csv", Required = true, HelpText = "CSV of users.. [VssUserToAddMailAddress],[VssLicense] .")]
            public string Csv { get; set; }
        }

        public static int Main(string[] args)
        {
           // Telemetry.Current.TrackEvent("ApplicationStart");
           // AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            DateTime startTime = DateTime.Now;
            Stopwatch mainTimer = new Stopwatch();
            mainTimer.Start();
            //////////////////////////////////////////////////
            Trace.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Trace.Listeners.Add(new TextWriterTraceListener(string.Format(@"{0}-{1}.log", DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"), "MigrationRun"), "myListener"));
            //////////////////////////////////////////////////
            Trace.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name, "VstsUserManagement");
            Trace.WriteLine(System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString(), "VstsUserManagement");
           // Trace.WriteLine(string.Format("Telemitery Enabled: {0}", Telemetry.Current.IsEnabled().ToString()), "VstsUserManagement");
           // Trace.WriteLine(string.Format("SessionID: {0}", Telemetry.Current.Context.Session.Id), "VstsUserManagement");
           // Trace.WriteLine(string.Format("User: {0}", Telemetry.Current.Context.User.Id), "VstsUserManagement");
            Trace.WriteLine(string.Format("Start Time: {0}", startTime.ToUniversalTime()), "VstsUserManagement");
            Trace.WriteLine("----------------------------------------------------------------", "VstsUserManagement");
            Trace.WriteLine("------------------------------START-----------------------------", "VstsUserManagement");
            Trace.WriteLine("----------------------------------------------------------------", "VstsUserManagement");
            //////////////////////////////////////////////////
            int result = (int)Parser.Default.ParseArguments<AddUser, AddUsers>(args).MapResult(
                (AddUser opts) => RunAddUserAndReturnExitCode(opts),
                (AddUsers opts) => RunAddUsersAndReturnExitCode(opts),
                errs => 1);
            //////////////////////////////////////////////////
            Trace.WriteLine("----------------------------------------------------------------", "VstsUserManagement");
            Trace.WriteLine("-------------------------------END------------------------------", "VstsUserManagement");
            Trace.WriteLine("----------------------------------------------------------------", "VstsUserManagement");
            mainTimer.Stop();
            //Telemetry.Current.TrackEvent("ApplicationEnd", null,
            //    new Dictionary<string, double> {
            //            { "ApplicationDuration", mainTimer.ElapsedMilliseconds }
            //    });
            //if (Telemetry.Current != null)
            //{
            //    Telemetry.Current.Flush();
            //    // Allow time for flushing:
            //    System.Threading.Thread.Sleep(1000);
            //}
            Trace.WriteLine(string.Format("Duration: {0}", mainTimer.Elapsed.ToString("c")), "VstsUserManagement");
            Trace.WriteLine(string.Format("End Time: {0}", startTime.ToUniversalTime()), "VstsUserManagement");
#if DEBUG
            Console.ReadKey();
#endif
            return result;
        }

        private static int RunAddUsersAndReturnExitCode(AddUsers opts)
        {
            CreateAuthConnection(opts.VssAccountUrl);
            var csvRows = File.ReadAllLines(opts.Csv);
            foreach (var row in csvRows)
            {
                string[] bits = row.Split(',');
                License VssLicense = ConvertStringToLicence(bits[1].ToLower());
                AddUserToAccount(opts.VssAccountName, bits[0], VssLicense);
                if (!string.IsNullOrEmpty(opts.Project))
                {
                    AddUserToSecurityGroup(opts.Project, opts.Group, bits[0]);
                }
                
            }
            return 0;
        }

        private static int RunAddUserAndReturnExitCode(AddUser opts)
        {
            CreateAuthConnection(opts.VssAccountUrl);
            License VssLicense = ConvertStringToLicence(opts.VssLicense);
            AddUserToAccount(opts.VssAccountName, opts.VssUserToAddMailAddress, VssLicense);
            if (!string.IsNullOrEmpty(opts.Project))
            {
                AddUserToSecurityGroup(opts.Project, opts.Group, opts.VssUserToAddMailAddress);
            }
            return 0;
        }

        static LicensingHttpClient licensingClient;
        static IdentityHttpClient identityClient;
        static VssConnection vssConnection;

        private static void CreateAuthConnection(string VssAccountUrl)
        {
            // Create a connection to the specified account.
            // If you change the false to true, your credentials will be saved.
            var creds = new VssClientCredentials(true);
            vssConnection = new VssConnection(new Uri(VssAccountUrl), creds);

            // We need the clients for two services: Licensing and Identity
            licensingClient = vssConnection.GetClient<LicensingHttpClient>();
            identityClient = vssConnection.GetClient<IdentityHttpClient>();
        }

        private static License ConvertStringToLicence(string license)
        {
            License VssLicense = AccountLicense.Express; // default to Basic license
            switch (license)
            {
                case "basic":
                    VssLicense = AccountLicense.Express;
                    break;
                case "professional":
                    VssLicense = AccountLicense.Professional;
                    break;
                case "advanced":
                    VssLicense = AccountLicense.Advanced;
                    break;
                case "msdn":
                    // When the user logs in, the system will determine the actual MSDN benefits for the user.
                    VssLicense = MsdnLicense.Eligible;
                    break;
                // Uncomment the code for Stakeholder if you are using VS 2013 Update 4 or newer.
                case "stakeholder":
                    VssLicense = AccountLicense.Stakeholder;
                    break;
                default:
                    Console.WriteLine("Error: License must be Basic, Professional, Advanced, or MSDN");
                     throw new Exception("Error: License must be Basic, Professional, Advanced, or MSDN");
            }
            return VssLicense;
        }

       static List<string> fubarAccounts = new List<string>();

            private static void AddUserToAccount(string VssAccountName, string VssUserToAddMailAddress, License VssLicense)
        {
            try
            {
                // The first call is to see if the user already exists in the account.
                // Since this is the first call to the service, this will trigger the sign-in window to pop up.
                Console.WriteLine("Sign in as the admin of account {0}. You will see a sign-in window on the desktop.",
                                  VssAccountName);
                var userIdentity = identityClient.ReadIdentitiesAsync(IdentitySearchFilter.AccountName,
                                                                      VssUserToAddMailAddress).Result.FirstOrDefault();
                if (userIdentity == null)
                {
                    var username = VssUserToAddMailAddress.Substring(0, VssUserToAddMailAddress.IndexOf("@"));
                    userIdentity = identityClient.ReadIdentitiesAsync(IdentitySearchFilter.MailAddress, VssUserToAddMailAddress).Result.FirstOrDefault();
                }
                // If the identity is null, this is a user that has not yet been added to the account.
                // We'll need to add the user as a "bind pending" - meaning that the email address of the identity is 
                // recorded so that the user can log into the account, but the rest of the details of the identity 
                // won't be filled in until first login.
                if (userIdentity == null)
                {
                    Console.WriteLine("Creating a new identity and adding it to the collection's licensed users group.");

                    // We are adding the user to a collection, and at the moment only one collection is supported per
                    // account in VSO.
                    var collectionScope = identityClient.GetScopeAsync(VssAccountName).Result;

                    // First get the descriptor for the licensed users group, which is a well known (built in) group.
                    var licensedUsersGroupDescriptor = new IdentityDescriptor(IdentityConstants.TeamFoundationType,
                                                                              GroupWellKnownSidConstants.LicensedUsersGroupSid);

                    // Now convert that into the licensed users group descriptor into a collection scope identifier.
                    var identifier = String.Concat(SidIdentityHelper.GetDomainSid(collectionScope.Id),
                                                   SidIdentityHelper.WellKnownSidType,
                                                   licensedUsersGroupDescriptor.Identifier.Substring(SidIdentityHelper.WellKnownSidPrefix.Length));

                    // Here we take the string representation and create the strongly-type descriptor
                    var collectionLicensedUsersGroupDescriptor = new IdentityDescriptor(IdentityConstants.TeamFoundationType,
                                                                                        identifier);


                    // Get the domain from the user that runs this code. This domain will then be used to construct
                    // the bind-pending identity. The domain is either going to be "Windows Live ID" or the Azure 
                    // Active Directory (AAD) unique identifier, depending on whether the account is connected to
                    // an AAD tenant. Then we'll format this as a UPN string.
                    var currUserIdentity = vssConnection.AuthorizedIdentity.Descriptor;
                    var directory = "Windows Live ID"; // default to an MSA (fka Live ID)
                    if (currUserIdentity.Identifier.Contains('\\'))
                    {
                        // The identifier is domain\userEmailAddress, which is used by AAD-backed accounts.
                        // We'll extract the domain from the admin user.
                        directory = currUserIdentity.Identifier.Split(new char[] { '\\' })[0];
                    }
                    var upnIdentity = string.Format("upn:{0}\\{1}", directory, VssUserToAddMailAddress);

                    // Next we'll create the identity descriptor for a new "bind pending" user identity.
                    var newUserDesciptor = new IdentityDescriptor(IdentityConstants.BindPendingIdentityType,
                                                                  upnIdentity);

                    // We are ready to actually create the "bind pending" identity entry. First we have to add the
                    // identity to the collection's licensed users group. Then we'll retrieve the Identity object
                    // for this newly-added user. Without being added to the licensed users group, the identity 
                    // can't exist in the account.
                    bool result = identityClient.AddMemberToGroupAsync(collectionLicensedUsersGroupDescriptor,
                                                                       newUserDesciptor).Result;
                    userIdentity = identityClient.ReadIdentitiesAsync(IdentitySearchFilter.AccountName,
                                                                      VssUserToAddMailAddress).Result.FirstOrDefault();
                }

                Console.WriteLine("Assigning license to user.");
                var entitlement = licensingClient.AssignEntitlementAsync(userIdentity.Id, VssLicense).Result;

                Console.WriteLine("Success!");
            }
            catch (Exception e)
            {
                using (StreamWriter sw = File.AppendText("fubarAccounts.txt"))
                {
                    sw.WriteLine("{0}, failed, {1}", VssUserToAddMailAddress, e.InnerException.Message);
                }
                Console.WriteLine("\r\nSomething went wrong...");
                Console.WriteLine(e.Message);
                if (e.InnerException != null)
                {
                    Console.WriteLine(e.InnerException.Message);
                }
            }
        }

        private static void AddUserToSecurityGroup(string VssTeamProjectName, string VssSecurityGroup, string VssUserToAddMailAddress)
        {
          var  projectClient = vssConnection.GetClient<ProjectHttpClient>();
            TeamProject teamProject = projectClient.GetProject(VssTeamProjectName).Result;
            var groups = identityClient.ListGroupsAsync(new Guid[] { teamProject.Id }).Result;
            var rGroup = groups.Where(x=> x.DisplayName.EndsWith(VssSecurityGroup)).SingleOrDefault();
            var userIdentity = identityClient.ReadIdentitiesAsync(IdentitySearchFilter.AccountName,
                                                                     VssUserToAddMailAddress).Result.FirstOrDefault();
            if (userIdentity!=null && rGroup != null)
            {
                bool result = identityClient.AddMemberToGroupAsync(rGroup.Descriptor,
                                                                      userIdentity.Descriptor).Result;
            }
            
          
        }
    }

}
