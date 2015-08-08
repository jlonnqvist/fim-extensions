
using System;
using Microsoft.MetadirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace Mms_Metaverse
{
	/// <summary>
	/// Summary description for MVExtensionObject.
	/// </summary>
    public class MVExtensionObject : IMVSynchronization
    {
        public MVExtensionObject()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        void IMVSynchronization.Initialize ()
        {
            //
            // TODO: Add initialization logic here
            //
        }

        void IMVSynchronization.Terminate ()
        {
            //
            // TODO: Add termination logic here
            //
        }

        void IMVSynchronization.Provision(MVEntry mventry)
        {
            ConnectedMA ManagementAgent;
            int Connectors = 0;
            CSEntry csentry;
            ReferenceValue DN;

            log("Provisioning Start");
            // ACTIVE DIRECTORY
            ManagementAgent = mventry.ConnectedMAs["Active Directory"]; 
            Connectors = ManagementAgent.Connectors.Count;

            if (0 == Connectors)
            {
                log("ActiveDirectory > No Connector");
                log("ActiveDirectory > Check Object Type");
                if (mventry.ObjectType.ToLower() == "person")
                {
                    log("ActiveDirectory > Provision AD User");
                    DN = ManagementAgent.EscapeDNComponent("CN=" + mventry["displayName"].Value).Concat(mventry["ou"].Value);
                    csentry = ManagementAgent.Connectors.StartNewConnector("user");
                    csentry.DN = DN;

                    string username = "";
                    username = GenerateUsername(mventry["ADDomain"].Value, mventry["firstName"].Value, mventry["lastName"].Value, 2, 4);

                    csentry["samAccountName"].Value = username;
                    csentry["userPrincipalName"].Value = username + "@" + mventry["ADDomain"].Value;
                    csentry["unicodepwd"].Values.Add("somethingsomething"); //replace...

                    log("ActiveDirectory > Commit DN: " + DN);
                    csentry.CommitNewConnector();
                }
                if (mventry.ObjectType.ToLower() == "group")
                {
                    log("ActiveDirectory > Provision AD Group");
                    DN = ManagementAgent.EscapeDNComponent("CN=" + mventry["displayName"].Value).Concat(mventry["ou"].Value);
                    csentry = ManagementAgent.Connectors.StartNewConnector("group");
                    csentry.DN = DN;

                    csentry["samAccountName"].Value = mventry["displayName"].Value;

                    long GlobalSecurityGroup = -2147483646;
                    csentry["groupType"].IntegerValue = GlobalSecurityGroup;

                    log("ActiveDirectory > Commit DN: " + DN);
                    csentry.CommitNewConnector();
                }
                log("Active Directory > Completed");
            }


            // FEIDE (ADLDS)
            ManagementAgent = mventry.ConnectedMAs["FEIDE"];
            Connectors = ManagementAgent.Connectors.Count;

            if (0 == Connectors)
            {
                log("FEIDE > No Connector");
                log("FEIDE > Check Object Type");
                if (mventry.ObjectType.ToLower() == "person")
                {
                    if (mventry["accountName"].IsPresent) {
                        log("FEIDE > Provision ADLDS inetOrgPerson");
                        // We cannot project the user to ADLDS at the first import from the HR-system.
                        // Need to wait for the AD-projection first, so that a username is created and synced back to the metaverse.
                        DN = ManagementAgent.EscapeDNComponent("CN=" + mventry["displayName"].Value).Concat(mventry["ouADLDS"].Value);
                        csentry = ManagementAgent.Connectors.StartNewConnector("inetOrgPerson");
                        csentry.DN = DN;

                        csentry["eduPersonPrincipalName"].Value = mventry["accountName"].Value + mventry["ADLDSDomain"];

                        log("FEIDE > Commit DN: " + DN);
                        csentry.CommitNewConnector();
                    } else
                    {
                        log("FEIDE > The attribute AccountName is not yet present. Wait for AD-import.");
                    }
                }
                if (mventry.ObjectType.ToLower() == "organization")
                {
                    log("FEIDE > Provision ADLDS Organization");
                    DN = ManagementAgent.EscapeDNComponent("O=" + mventry["displayName"].Value).Concat(mventry["ouADLDS"].Value);
                    csentry = ManagementAgent.Connectors.StartNewConnector("organization", new string[4] { "organization", "top", "eduOrg", "norEduOrg" } );
                    csentry.DN = DN;

                    log("FEIDE > Commit DN: " + DN);
                    csentry.CommitNewConnector();
                }
                if (mventry.ObjectType == "organizationalUnit")
                {
                    log("FEIDE > Provision ADLDS OrganizationalUnit");
                    DN = ManagementAgent.EscapeDNComponent("OU=" + mventry["displayName"].Value).Concat(mventry["ouADLDS"].Value);
                    csentry = ManagementAgent.Connectors.StartNewConnector("organizationalUnit", new string[3] { "organizationalUnit", "top", "norEduOrgUnit" });
                    csentry.DN = DN;

                    log("FEIDE > Commit DN: " + DN);
                    csentry.CommitNewConnector();
                }
                log("FEIDE > Completed");
            }
            log("Provisioning End");
        }	

        bool IMVSynchronization.ShouldDeleteFromMV (CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: Add MV deletion logic here
            //
            throw new EntryPointNotImplementedException();
        }


        // CUSTOM FUNCTIONS
        bool CheckUsernameAvailability(string domain, string userName)
        {
            using (var domainContext = new PrincipalContext(ContextType.Domain, domain))
            {
                using (var foundUser = UserPrincipal.FindByIdentity(domainContext, IdentityType.SamAccountName, userName))
                {
                    return foundUser != null;
                }
            }
        }

        string GenerateUsername(string domain, string firstName, string lastName, int starti, int startj)
        {
            firstName = firstName.Replace(" ", "").Replace("ø","o").Replace("å","a").Replace("æ","e");
            lastName = lastName.Replace(" ", "").Replace("ø", "o").Replace("å", "a").Replace("æ", "e");

            if (firstName.Length < starti) { starti = firstName.Length; }
            if (lastName.Length < startj) { startj = lastName.Length; }

            log("GenerateUsername > " + firstName + ", " + lastName);
            string username = "";

            for (int i = starti; i < firstName.Length; i++) {
                for (int j = startj; j > 0; j--) {
                    username = (firstName.Substring(0, i) + lastName.Substring(0, j)).ToLower();
                    log("GenerateUsername > Check if available: " + username + "@" + domain); 
                    if (!CheckUsernameAvailability(domain, username)) {
                        log("GenerateUsername > Available!");
                        return username;
                    } else
                    {
                        log("GenerateUsername > Taken.");
                    }
                }
            }

            for (int n = 1; n < 9999; n++) {
                log("GenerateUsername > Available: " + username + "@" + domain);
                if (!CheckUsernameAvailability(domain, username+n)) { return username+n; }
            }

            return "gibberish";
        }

        void log(string message)
        {
            System.IO.StreamWriter file = new System.IO.StreamWriter("C:\\Temp\\MVExtension-ProvisionLog.txt", true);
            file.WriteLine(DateTime.Now + " :: " + message);
            file.Close();
        }

    }
}


