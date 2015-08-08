
using System;
using Microsoft.MetadirectoryServices;

namespace Mms_ManagementAgent_MGK_AD_Extension
{
    /// <summary>
    /// Summary description for MAExtensionObject.
	/// </summary>
	public class MAExtensionObject : IMASynchronization
	{
		public MAExtensionObject()
		{
            //
            // TODO: Add constructor logic here
            //
        }
		void IMASynchronization.Initialize ()
		{
            //
            // TODO: write initialization code
            //
        }

        void IMASynchronization.Terminate ()
        {
            //
            // TODO: write termination code
            //
        }

        bool IMASynchronization.ShouldProjectToMV (CSEntry csentry, out string MVObjectType)
        {
			//
			// TODO: Remove this throw statement if you implement this method
			//
			throw new EntryPointNotImplementedException();
		}

        DeprovisionAction IMASynchronization.Deprovision (CSEntry csentry)
        {
			//
			// TODO: Remove this throw statement if you implement this method
			//
			throw new EntryPointNotImplementedException();
        }	

        bool IMASynchronization.FilterForDisconnection (CSEntry csentry)
        {
            //
            // TODO: write connector filter code
            //
            throw new EntryPointNotImplementedException();
		}

		void IMASynchronization.MapAttributesForJoin (string FlowRuleName, CSEntry csentry, ref ValueCollection values)
        {
            //
            // TODO: write join mapping code
            //
            throw new EntryPointNotImplementedException();
        }

        bool IMASynchronization.ResolveJoinSearch (string joinCriteriaName, CSEntry csentry, MVEntry[] rgmventry, out int imventry, ref string MVObjectType)
        {
            //
            // TODO: write join resolution code
            //
            throw new EntryPointNotImplementedException();
		}

        void IMASynchronization.MapAttributesForImport( string FlowRuleName, CSEntry csentry, MVEntry mventry)
        {
            //
            // TODO: write your import attribute flow code
            //
            throw new EntryPointNotImplementedException();
        }

        void IMASynchronization.MapAttributesForExport (string FlowRuleName, MVEntry mventry, CSEntry csentry)
        {
            
            switch(FlowRuleName)
            { 
                case "userAccountControl":
                    // https://msdn.microsoft.com/en-us/library/windows/desktop/ms696026(v=vs.100).aspx
                    const long ADS_UF_PASSWD_NOTREQD = 0x0020; // No password is required
                    const long ADS_UF_NORMAL_ACCOUNT = 0x0200; // Typical enabled user account
                    const long ADS_UF_ACCOUNTDISABLE = 0x0002; // Disabled account
                    long CurrentUAC = 0;

                    if (csentry["userAccountControl"].IsPresent)
                    {
                        CurrentUAC = csentry["userAccountControl"].IntegerValue;
                    } 

                    switch (mventry["employeeStatus"].Value.ToLower())
                    {
                        case "active": //check current value first. DO nothing if already active in both
                            csentry["userAccountControl"].IntegerValue = (CurrentUAC | ADS_UF_NORMAL_ACCOUNT) 
                                                                         & ~ADS_UF_ACCOUNTDISABLE;
                            break;
                        case "inactive":
                            csentry["userAccountControl"].IntegerValue = CurrentUAC 
                                                                        | ADS_UF_ACCOUNTDISABLE
                                                                        | ADS_UF_PASSWD_NOTREQD;
                            break;
                    }
                    break;
            }
      
        }
	}
}
