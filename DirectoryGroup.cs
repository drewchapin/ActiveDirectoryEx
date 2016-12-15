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
	public class DirectoryGroup: DirectoryEntryEx, IDisposable
	{
		/*
		 * Constructors
		 */
		protected DirectoryGroup()
		{
			// Nothing to do.
		}
		public DirectoryGroup( DirectoryEntry entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryGroup( DirectoryEntryEx entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryGroup( DirectoryGroup entry, bool leaveOpen = false )
			: base(entry.entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryGroup( object adsObject )
			: base(adsObject)
		{
			// Nothing to do.
		}
		public DirectoryGroup( string path )
			: base(path)
		{
			// Nothing to do.
		}
		public DirectoryGroup( string path, string username, string password )
			: base(path,username,password)
		{
			// Nothing to do.
		}
		public DirectoryGroup( string path, string username, string password, AuthenticationTypes authenticationType )
			: base(path,username,password,authenticationType)
		{
			// Nothing to do.
		}

		/*
		 * Operators
		 */
		public static explicit operator DirectoryGroup( DirectoryEntry entry )
		{
			return new DirectoryGroup(entry);
		}
		public static explicit operator DirectoryEntry( DirectoryGroup entry )
		{
			return entry.entry;
		}

		/*
		 * Extended Properties
		 */
		public string EmailAddress
		{
			get {
				return entry.Properties.Contains("mail") ? (string)entry.Properties["mail"].Value : null;
			}
			set {
				entry.Properties["mail"].Value = value;
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
		public string Notes
		{
			get {
				return entry.Properties.Contains("info") ? (string)entry.Properties["info"].Value : null;
			}
			set {
				entry.Properties["info"].Value = value;
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

		/*
		 * Extended Methods
		 */
		public new DirectoryGroup CopyTo( DirectoryEntry newParent )
		{
			return new DirectoryGroup(entry.CopyTo(newParent));
		}
		public new DirectoryGroup CopyTo( DirectoryEntry newParent, string newName )
		{
			return new DirectoryGroup(entry.CopyTo(newParent,newName));
		}
		public new DirectoryGroup CopyTo( DirectoryEntryEx newParent )
		{
			return new DirectoryGroup(entry.CopyTo(newParent.GetDirectoryEntry()));
		}
		public new DirectoryGroup CopyTo( DirectoryEntryEx newParent, string newName )
		{
			return new DirectoryGroup(entry.CopyTo(newParent.GetDirectoryEntry(),newName));
		}
		public DirectoryGroup CopyTo( DirectoryOU newParent )
		{
			return new DirectoryGroup(entry.CopyTo(newParent.GetDirectoryEntry()));
		}
		public DirectoryGroup CopyTo( DirectoryOU newParent, string newName )
		{
			return new DirectoryGroup(entry.CopyTo(newParent.GetDirectoryEntry(),newName));
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
