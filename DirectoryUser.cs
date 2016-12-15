/*
 * MIT License
 * 
 * Copyright (c) 2016 Drew Chapin <drew@drewchapin.com>
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.ComponentModel;

namespace ActiveDirectoryEx
{
	public class DirectoryUser: DirectoryEntryEx, IDisposable
	{
		/*
		 * Constructors
		 */
		protected DirectoryUser()
		{
			// Nothing to do.
		}
		public DirectoryUser( DirectoryEntry entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryUser( DirectoryEntryEx entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryUser( DirectoryUser entry, bool leaveOpen = false )
			: base(entry.entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryUser( object adsObject )
			: base(adsObject)
		{
			// Nothing to do.
		}
		public DirectoryUser( string path )
			: base(path)
		{
			// Nothing to do.
		}
		public DirectoryUser( string path, string username, string password )
			: base(path,username,password)
		{
			// Nothing to do.
		}
		public DirectoryUser( string path, string username, string password, AuthenticationTypes authenticationType )
			: base(path,username,password,authenticationType)
		{
			// Nothing to do.
		}

		/*
		 * Operators
		 */
		public static explicit operator DirectoryUser( DirectoryEntry entry )
		{
			return new DirectoryUser(entry);
		}
		public static explicit operator DirectoryEntry( DirectoryUser entry )
		{
			return entry.entry;
		}

		/*
		 * Extended Properties
		 */
		public string AssistantPhone
		{
			get {
				return entry.Properties.Contains("telephoneAssistant") ? (string)entry.Properties["telephoneAssistant"].Value : null;
			}
			set {
				entry.Properties["telephoneAssistant"].Value = value;
			}
		}
		public string City
		{
			get {
				return entry.Properties.Contains("L") ? (string)entry.Properties["L"].Value : null;
			}
			set {
				entry.Properties["L"].Value = value;
			}
		}
		public string Company
		{
			get {
				return entry.Properties.Contains("company") ? (string)entry.Properties["company"].Value : null;
			}
			set {
				entry.Properties["company"].Value = value;
			}
		}
		public string Country
		{
			get {
				return entry.Properties.Contains("c") ? (string)entry.Properties["c"].Value : null;
			}
			set {
				entry.Properties["c"].Value = value;
			}
		}
		public Int32 CountryCode
		{
			get {
				return entry.Properties.Contains("countryCode") ? (Int32)entry.Properties["countryCode"].Value : 0;
			}
			set {
				entry.Properties["countryCode"].Value = value;
			}
		}
		public string CountryName
		{
			get {
				return entry.Properties.Contains("co") ? (string)entry.Properties["co"].Value : null;
			}
			set {
				entry.Properties["co"].Value = value;
			}
		}
		public string Department
		{
			get {
				return entry.Properties.Contains("department") ? (string)entry.Properties["department"].Value : null;
			}
			set {
				entry.Properties["department"].Value = value;
			}
		}
		public string EmailAddress
		{
			get {
				return entry.Properties.Contains("mail") ? (string)entry.Properties["mail"].Value : null;
			}
			set {
				entry.Properties["mail"].Value = value;
			}
		}
		public string EmployeeId
		{
			get {
				return entry.Properties.Contains("employeeID") ? (string)entry.Properties["employeeID"].Value : null;
			}
			set {
				entry.Properties["employeeID"].Value = value;
			}
		}
		public bool Enabled
		{
			get {
				return 0 == (UserAccountControl.Value & (int)UserAccountControlFlags.AccountDisabled);
			}
			set {
				if( value )
					entry.Properties["userAccountControl"].Value = UserAccountControl.Value & ~(int)UserAccountControlFlags.AccountDisabled;
				else
					entry.Properties["userAccountControl"].Value = UserAccountControl.Value | (int)UserAccountControlFlags.AccountDisabled;
			}
		}
		public string ExchangeAlias
		{
			get {
				return entry.Properties.Contains("mailNickname") ? (string)entry.Properties["mailNickname"].Value : null;
			}
			set {
				entry.Properties["mailNickname"].Value = value;
			}
		}
		public string ExchangeAssistant
		{
			get {
				return entry.Properties.Contains("msExchAssistantName") ? (string)entry.Properties["msExchAssistantName"].Value : null;
			}
			set {
				entry.Properties["msExchAssistantName"].Value = value;
			}
		}
		public DateTime? ExpirationDate
		{
			get {
				if( entry.Properties.Contains("accountExpires") )
				{
					ActiveDs.LargeInteger value = (ActiveDs.LargeInteger)entry.Properties["accountExpires"].Value;
					Int64 time = ((Int64)value.HighPart << 32) | ((Int64)value.LowPart & 0xFFFFFFFF);
					return DateTime.FromFileTime(time);
				}
				return null;
			}
			set {
				if( value != null )
				{
					ActiveDs.LargeInteger expire = new ActiveDs.LargeInteger();
					Int64 fileTime = value.Value.ToFileTime();
					expire.HighPart = (int)(fileTime >> 32);
					expire.LowPart = (int)(fileTime & 0xFFFFFFFF);
					entry.Properties["accountExpires"].Value = expire;
				}
				else
					entry.Properties["accountExpires"].Value = null;
			}
		}
		public string FaxNumber
		{
			get {
				return entry.Properties.Contains("facsimileTelephoneNumber") ? (string)entry.Properties["facsimileTelephoneNumber"].Value : null;
			}
			set {
				entry.Properties["facsimileTelephoneNumber"].Value = value;
			}
		}
		public string Firstname
		{
			get {
				return entry.Properties.Contains("givenName") ? (string)entry.Properties["givenName"].Value : null;
			}
			set {
				entry.Properties["givenName"].Value = value;
			}
		}
		public string HomeDirectory
		{
			get {
				return entry.Properties.Contains("homeDirectory") ? (string)entry.Properties["homeDirectory"].Value : null;
			}
			set {
				entry.Properties["homeDirectory"].Value = value;
			}
		}
		public string HomeDrive
		{
			get {
				return entry.Properties.Contains("homeDrive") ? (string)entry.Properties["homeDrive"].Value : null;
			}
			set {
				entry.Properties["homeDrive"].Value = value;
			}
		}
		public string HomePhone
		{
			get {
				return entry.Properties.Contains("homePhone") ? (string)entry.Properties["homePhone"].Value : null;
			}
			set {
				entry.Properties["homePhone"].Value = value;
			}
		}
		public string Initials
		{
			get {
				return entry.Properties.Contains("initials") ? (string)entry.Properties["initials"].Value : null;
			}
			set {
				entry.Properties["initials"].Value = value;
			}
		}		
		public string IpPhone
		{
			get {
				return entry.Properties.Contains("ipPhone") ? (string)entry.Properties["ipPhone"].Value : null;
			}
			set {
				entry.Properties["ipPhone"].Value = value;
			}
		}
		public DateTime? LastLogon
		{
			get {
				if( entry.Properties.Contains("lastLogon") )
				{
					ActiveDs.IADsLargeInteger value = (ActiveDs.IADsLargeInteger)entry.Properties["lastLogon"].Value;
					Int64 time = ((Int64)value.HighPart << 32) | ((Int64)value.LowPart & 0xFFFFFFFF);
					return DateTime.FromFileTime(time);
				}
				return null;
			}
		}
		public DateTime? LastLogonTimestamp
		{
			get {
				if( entry.Properties.Contains("lastLogonTimestamp") )
				{
					ActiveDs.IADsLargeInteger value = (ActiveDs.IADsLargeInteger)entry.Properties["lastLogonTimestamp"].Value;
					Int64 time = ((Int64)value.HighPart << 32) | ((Int64)value.LowPart & 0xFFFFFFFF);
					return DateTime.FromFileTime(time);
				}
				return null;
			}
		}
		public string Lastname
		{
			get {
				return entry.Properties.Contains("sn") ? (string)entry.Properties["sn"].Value : null;
			}
			set {
				entry.Properties["sn"].Value = value;
			}
		}
		public string LoginScript
		{
			get {
				return entry.Properties.Contains("scriptPath") ? (string)entry.Properties["scriptPath"].Value : null;
			}
			set {
				entry.Properties["scriptPath"].Value = value;
			}
		}
		public string MobilePhone
		{
			get {
				return entry.Properties.Contains("mobile") ? (string)entry.Properties["mobile"].Value : null;
			}
			set {
				entry.Properties["mobile"].Value = value;
			}
		}
		public string Notes
		{
			get {
				return entry.Properties.Contains("info") ? (string)entry.Properties["info"].Value : null;
			}
			set {
				entry.Properties["info"].Value = value;
			}
		}
		public string Office
		{
			get {
				return entry.Properties.Contains("physicalDeliveryOfficeName") ? (string)entry.Properties["physicalDeliveryOfficeName"].Value : null;
			}
			set {
				entry.Properties["physicalDeliveryOfficeName"].Value = value;
			}
		}
		public IEnumerable<string> OtherFax
		{
			get {
				return entry.Properties.Contains("otherFacsimileTelephoneNumber") ? entry.Properties["otherFacsimileTelephoneNumber"].Cast<string>() : null;
			}
			set {
				entry.Properties["otherFacsimileTelephoneNumber"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}		
		public IEnumerable<string> OtherHomePhone
		{
			get {
				return entry.Properties.Contains("otherHomePhone") ? entry.Properties["otherHomePhone"].Cast<string>().ToList() : null;
			}
			set {
				entry.Properties["otherHomePhone"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public IEnumerable<string> OtherIpPhone
		{
			get {
				return entry.Properties.Contains("otherIpPhone") ? entry.Properties["otherIpPhone"].Cast<string>() : null;
			}
			set {
				entry.Properties["otherIpPhone"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public IEnumerable<string> OtherMobile
		{
			get {
				return entry.Properties.Contains("otherMobile") ? entry.Properties["otherMobile"].Cast<string>() : null;
			}
			set {
				entry.Properties["otherMobile"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public IEnumerable<string> OtherPager
		{
			get {
				return entry.Properties.Contains("otherPager") ? entry.Properties["otherPager"].Cast<string>() : null;
			}
			set {
				entry.Properties["otherPager"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public IEnumerable<string> OtherTelephoneNumbers 
		{
			get {
				return entry.Properties.Contains("otherTelephone") ? entry.Properties["otherTelephone"].Cast<string>() : null;
			}
			set {
				entry.Properties["otherTelephone"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public string PagerNumber
		{
			get {
				return entry.Properties.Contains("pager") ? (string)entry.Properties["pager"].Value : null;
			}
			set {
				entry.Properties["pager"].Value = value;
			}
		}
		public string PostalCode
		{
			get {
				return entry.Properties.Contains("postalCode") ? (string)entry.Properties["postalCode"].Value : null;
			}
			set {
				entry.Properties["postalCode"].Value = value;
			}
		}
		public string PostOfficeBox
		{
			get {
				return entry.Properties.Contains("postOfficeBox") ? (string)entry.Properties["postOfficeBox"].Value : null;
			}
			set {
				entry.Properties["postOfficeBox"].Value = value;
			}
		}
		public string PrimarySipAddress
		{
			get {
				IEnumerable<string> addr = ProxyAddresses;
				if( addr != null )
				{
					foreach( string x in addr )
					{
						if( x.StartsWith("SIP:",StringComparison.Ordinal) )
							return x.Substring(4);
					}
					foreach( string x in addr )
						if( x.StartsWith("sip:",StringComparison.OrdinalIgnoreCase) )
							return x.Substring(4);
				}
				return null;
			}
			set {
				List<string> addr = ProxyAddresses.ToList();
				bool updated = false;
				if( addr != null )
				{
					for( int i = 0; i < addr.Count(); i++ )
					{
						if( addr[i].Equals("sip:"+value,StringComparison.OrdinalIgnoreCase) )
						{
							addr[i] = "SIP:" + value;
							updated = true;
						}
						else if( addr[i].StartsWith("SIP:",StringComparison.Ordinal) )
							addr[i] = "sip:" + addr[i].Substring(4);
					}
				}
				else
				{
					addr = new List<string>();
				}
				if( !updated ) 
					addr.Add("SIP:"+value);
				ProxyAddresses = addr;

			}
		}
		public string PrimarySmtpAddress
		{
			get {
				IEnumerable<string> addr = ProxyAddresses;
				if( addr != null )
				{
					foreach( string x in addr )
					{
						if( x.StartsWith("SMTP:",StringComparison.Ordinal) )
							return x.Substring(5);
					}
				}
				return null;
			}
			set {
				List<string> addr = ProxyAddresses.ToList();
				bool updated = false;
				if( addr != null )
				{
					for( int i = 0; i < addr.Count(); i++ )
					{
						if( addr[i].Equals("smtp:"+value,StringComparison.OrdinalIgnoreCase) )
						{
							addr[i] = "SMTP:" + value;
							updated = true;
						}
						else if( addr[i].StartsWith("SMTP:",StringComparison.Ordinal) )
							addr[i] = "smtp:" + addr[i].Substring(5);
					}
				}
				else
				{
					addr = new List<string>();
				}
				if( !updated ) 
					addr.Add("SMTP:"+value);
				ProxyAddresses = addr;
			}
		}
		public string ProfilePath 
		{
			get {
				return entry.Properties.Contains("profilePath") ? (string)entry.Properties["proflePath"].Value : null; 
			}
			set {
				entry.Properties["profilePath"].Value = value;
			}
		}
		public IEnumerable<string> ProxyAddresses
		{
			get {
				return entry.Properties.Contains("proxyAddresses") ? entry.Properties["proxyAddresses"].Cast<string>() : null;
			}
			set {
				entry.Properties["proxyAddresses"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public int? SamAccountType
		{
			get {
				return entry.Properties.Contains("sAMAccountType") ? (int?)entry.Properties["sAMAccountType"].Value : null;
			}
		}
		public string SamAccountName
		{
			get {
				return entry.Properties.Contains("sAMAccountName") ? (string)entry.Properties["sAMAccountName"].Value : null;
			}
			set {
				entry.Properties["sAMAccountName"].Value = value;
			}
		}
		public string State
		{
			get {
				return entry.Properties.Contains("st") ? (string)entry.Properties["st"].Value : null;
			}
			set {
				entry.Properties["st"].Value = value;
			}
		}
		public string Street
		{
			get {
				return entry.Properties.Contains("streetAddress") ? (string)entry.Properties["streetAddress"].Value : null;
			}
			set {
				entry.Properties["streetAddress"].Value = value;
			}
		}
		public string Title
		{
			get {
				return entry.Properties.Contains("title") ? (string)entry.Properties["title"].Value : null;
			}
			set {
				entry.Properties["title"].Value = value;
			}
		}
		public Int32? UserAccountControl
		{
			get {
				return entry.Properties.Contains("userAccountControl") ? (Int32?)entry.Properties["userAccountControl"].Value : null;
			}
			set {
				entry.Properties["userAccountControl"].Value = value;
			}
		}
		public string UserPrincipalName
		{
			get {
				return entry.Properties.Contains("userPrincipalName") ? (string)entry.Properties["userPrincipalName"].Value : null;
			}
			set {
				entry.Properties["userPrincipalName"].Value = value;
			}
		}

		/*
		 * Extended Methods
		 */
		public new DirectoryUser CopyTo( DirectoryEntry newParent )
		{
			return new DirectoryUser(entry.CopyTo(newParent));
		}
		public new DirectoryUser CopyTo( DirectoryEntry newParent, string newName )
		{
			return new DirectoryUser(entry.CopyTo(newParent,newName));
		}
		public new DirectoryUser CopyTo( DirectoryEntryEx newParent )
		{
			return new DirectoryUser(entry.CopyTo(newParent.GetDirectoryEntry()));
		}
		public new DirectoryUser CopyTo( DirectoryEntryEx newParent, string newName )
		{
			return new DirectoryUser(entry.CopyTo(newParent.GetDirectoryEntry(),newName));
		}
		public new DirectoryUser CopyTo( DirectoryOU newParent )
		{
			return new DirectoryUser(entry.CopyTo(newParent.GetDirectoryEntry()));
		}
		public new DirectoryUser CopyTo( DirectoryOU newParent, string newName )
		{
			return new DirectoryUser(entry.CopyTo(newParent.GetDirectoryEntry(),newName));
		}
		public string GetExtensionAttribute( int index )
		{
			if( index < 1 || index > 15 )
				throw new IndexOutOfRangeException("index must be between 1 and 15");
			string name = String.Format("extensionAttribute{0}",index);
			return entry.Properties.Contains(name) ? (string)entry.Properties[name].Value : null;
		}
		public void SetExtensionAttribute( int index, string value )
		{
			if( index < 1 || index > 15 )
				throw new IndexOutOfRangeException("index must be between 1 and 15");
			string name = String.Format("extensionAttribute{0}",index);
			entry.Properties[name].Value = value;
		}
	}
}
