using System;
using System.Linq;
using Microsoft.TeamFoundation;
using Microsoft.TeamFoundation.Framework.Common;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Identity;
using Microsoft.VisualStudio.Services.Identity.Client;
using Microsoft.VisualStudio.Services.Licensing;
using Microsoft.VisualStudio.Services.Licensing.Client;

namespace AddUserToAccount
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                // For ease of running from the debugger, hard-code the account and the email address if not supplied
                // The account name here is just the name, not the URL.
                //args = new[] { "Awesome", "example@outlook.com", "basic" };
            }

            if (!Init(args))
            {
                Console.WriteLine("Add a licensed user to a Visual Studio Account");
                Console.WriteLine("Usage: [accountName] [userEmailAddress] <license>");
                Console.WriteLine("  accountName - just the name of the account, not the URL");
                Console.WriteLine("  userEmailAddress - email address of the user to be added");
                Console.WriteLine("  license - optional license (default is Basic): Basic, Professional, or Advanced");
                return;
            }

            AddUserToAccount();
        }

        private static void AddUserToAccount()
        {
            try
            {
                // Create a connection to the specified account.
                // If you change the false to true, your credentials will be saved.
                var creds = new VssClientCredentials(false);
                var vssConnection = new VssConnection(new Uri(VssAccountUrl), creds);

                // We need the clients for two services: Licensing and Identity
                var licensingClient = vssConnection.GetClient<LicensingHttpClient>();
                var identityClient = vssConnection.GetClient<IdentityHttpClient>();

                // The first call is to see if the user already exists in the account.
                // Since this is the first call to the service, this will trigger the sign-in window to pop up.
                Console.WriteLine("Sign in as the admin of account {0}. You will see a sign-in window on the desktop.",
                                  VssAccountName);
                var userIdentity = identityClient.ReadIdentitiesAsync(IdentitySearchFilter.AccountName, 
                                                                      VssUserToAddMailAddress).Result.FirstOrDefault();

                // If the identity is null, this is a user that has not yet been added to the account.
                // We'll need to add the user as a "bind pending" - meaning that the email address of the identity is 
                // recorded so that the user can log into the account, but the rest of the details of the identity 
                // won't be filled in until first login.
                if (userIdentity == null)
                {
                    Console.WriteLine("Creating a new identity and adding it to the collection's licensed users group.");

                    // We are adding the user to a collection, and at the moment only one collection is supported per
                    // account in VSO.
                    var collectionScope = identityClient.GetScopeAsync("DefaultCollection").Result;

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
                Console.WriteLine("\r\nSomething went wrong...");
                Console.WriteLine(e.Message);
                if (e.InnerException != null)
                {
                    Console.WriteLine(e.InnerException.Message);
                }
            }
        }

        private static bool Init(string[] args)
        {
            if (args == null || args.Length < 2)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(args[0]))
            {
                Console.WriteLine("Error: Invalid accountName");
                return false;
            }

            VssAccountName = args[0];

            // We need to talk to SPS in order to add a user and assign a license.
            VssAccountUrl = "https://" + VssAccountName + ".vssps.visualstudio.com/";

            if (string.IsNullOrWhiteSpace(args[1]))
            {
                Console.WriteLine("Error: Invalid userEmailAddress");
                return false;
            }

            VssUserToAddMailAddress = args[1];

            VssLicense = AccountLicense.Express; // default to Basic license
            if (args.Length == 3)
            {
                string license = args[2].ToLowerInvariant();
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
                    //case "Stakeholder":
                    //    VssLicense = AccountLicense.Stakeholder;
                    //    break;
                    default:
                        Console.WriteLine("Error: License must be Basic, Professional, Advanced, or MSDN");
                        //Console.WriteLine("Error: License must be Stakeholder, Basic, Professional, Advanced, or MSDN");
                        return false;
                }
            }

            return true;
        }

        public static string VssAccountUrl { get; set; }

        public static string VssAccountName { get; set; }

        public static string VssUserToAddMailAddress { get; set; }

        public static License VssLicense { get; set; }
    }


}
