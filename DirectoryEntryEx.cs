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

namespace ActiveDirectoryEx
{
	public class DirectoryEntryEx
	{
		/*
		 * Memebers
		 */
		private bool disposed = false;
		protected DirectoryEntry entry = null;
		private bool leaveOpen = false;

		/*
		 * Constructors
		 */
		protected DirectoryEntryEx()
		{
			// Nothing to do.
		}
		public DirectoryEntryEx( DirectoryEntry entry, bool leaveOpen = false )
		{
			this.entry = entry;
			this.leaveOpen = leaveOpen;
		}
		public DirectoryEntryEx( DirectoryEntryEx entry, bool leaveOpen = false )
		{
			this.entry = entry.entry;
			this.leaveOpen = leaveOpen;
		}
		public DirectoryEntryEx( object adsObject )
		{
			entry = new DirectoryEntry(adsObject);
		}
		public DirectoryEntryEx( string path )
		{
			entry = new DirectoryEntry(path);
		}
		public DirectoryEntryEx( string path, string username, string password )
		{
			entry = new DirectoryEntry(path,username,password);
		}
		public DirectoryEntryEx( string path, string username, string password, AuthenticationTypes authenticationType )
		{
			entry = new DirectoryEntry(path,username,password,authenticationType);
		}

		/*
		 * Operators
		 */
		public static explicit operator DirectoryEntryEx( DirectoryEntry entry )
		{
			return new DirectoryEntryEx(entry);
		}
		public static explicit operator DirectoryEntry( DirectoryEntryEx entry )
		{
			return entry.entry;
		}

		/*
		 * Base Properties
		 */
		public AuthenticationTypes AuthenticationType 
		{ 
			get { 
				return entry.AuthenticationType;
			}
			set {
				entry.AuthenticationType = value;
			}
		}
		public DirectoryEntries Children
		{ 
			get { 
				return entry.Children;
			}
		}
		public Guid Guid
		{ 
			get { 
				return entry.Guid;
			}
		}
		public string Name
		{ 
			get { 
				return entry.Name;
			}
		}
		public string NativeGuid
		{ 
			get { 
				return entry.NativeGuid;
			}
		}
		public object NativeObject
		{ 
			get { 
				return entry.NativeObject;
			}
		}
		public ActiveDirectorySecurity ObjectSecurity
		{ 
			get { 
				return entry.ObjectSecurity;
			}
			set {
				entry.ObjectSecurity = value;
			}
		}
		public DirectoryEntryConfiguration Options
		{ 
			get {
				return entry.Options;
			}
		}
		public DirectoryEntryEx Parent
		{ 
			get {
				return entry.Parent != null ? new DirectoryEntryEx(entry.Parent) : null;
			}
		}
		public string Password 
		{ 
			set {
				entry.Password = value;
			}
		}
		public string Path
		{ 
			get { 
				return entry.Path;
			}
			set {
				entry.Path = value;
			}
		}
		public PropertyCollection Properties
		{ 
			get {
				return entry.Properties;
			}
		}
		public string SchemaClassName 
		{ 
			get {
				return entry.SchemaClassName;
			}
		}
		public DirectoryEntry SchemaEntry 
		{ 
			get {
				return entry.SchemaEntry;
			}
		}
		public bool UsePropertyCache
		{ 
			get { 
				return entry.UsePropertyCache;
			}
			set {
				entry.UsePropertyCache = value;
			}
		}
		public string Username
		{ 
			get { 
				return entry.Username;
			}
			set {
				entry.Username = value;
			}
		}

		/*
		 * Extended Properties
		 */
		public string CanonicalName
		{
			get {
				if( !entry.Properties.Contains("canonicalName") )
					entry.RefreshCache(new string[]{"canonicalName"});
				return entry.Properties.Contains("canonicalName") ? (string)entry.Properties["canonicalName"].Value : null;
			}
		}
		public DateTime? CreationDate
		{
			get {
				return entry.Properties.Contains("whenCreated") ? (DateTime?)entry.Properties["whenCreated"].Value : null;
			}
		}
		public string Description
		{
			get {
				return entry.Properties.Contains("description") ? entry.Properties["description"].Value as string : null;
			}
			set {
				entry.Properties["description"].Value = value;
			}
		}
		public string DisplayName
		{
			get {
				return entry.Properties.Contains("displayName") ? (string)entry.Properties["displayName"].Value : null;
			}
			set {
				entry.Properties["displayName"].Value = value;
			}
		}
		public string Manager
		{
			get {
				return entry.Properties.Contains("manager") ? (string)entry.Properties["manager"].Value : null;
			}
			set {
				entry.Properties["manager"].Value = value;
			}
		}
		public DateTime? ModifiedDate
		{
			get {
				return entry.Properties.Contains("modifyimeStamp") ? (DateTime?)entry.Properties["modifyimeStamp"].Value : null;
			}
		}
		public IEnumerable<string> OtherWebPages
		{
			get {
				return entry.Properties.Contains("url") ? entry.Properties["url"].Cast<string>() : null;
			}
			set {
				entry.Properties["url"].Value = (value != null ? value.Reverse().ToArray() : new string[]{});
			}
		}
		public string TelephoneNumber
		{
			get {
				return entry.Properties.Contains("telephoneNumber") ? (string)entry.Properties["telephoneNumber"].Value : null;
			}
			set {
				entry.Properties["telephoneNumber"].Value = value;
			}
		}
		public string WebPage
		{
			get {
				return entry.Properties.Contains("wWWHomePage") ? (string)entry.Properties["wWWHomePage"].Value : null;
			}
			set {
				entry.Properties["wWWHomePage"].Value = value;
			}
		}

		/*
		 * Base Methods
		 */
		public void Close()
		{
			entry.Close();
		}
		public void CommitChanges()
		{
			entry.CommitChanges();
		}
		public void DeleteTree()
		{
			entry.DeleteTree();
		}
		public static bool Exists( string path )
		{
			return DirectoryEntry.Exists(path);
		}
		public object Invoke( string methodName, params object[] args )
		{
			return entry.Invoke(methodName,args);
		}
		public object InvokeGet( string propertyName )
		{
			return entry.InvokeGet(propertyName);
		}
		public void InvokeSet( string propertyName, params object[] args )
		{
			entry.InvokeSet(propertyName,args);
		}
		public void MoveTo( DirectoryEntry newParent )
		{
			entry.MoveTo(newParent);
		}
		public void MoveTo( DirectoryEntry newParent, string newName )
		{
			entry.MoveTo(newParent,newName);
		}
		public void RefreshCache()
		{
			entry.RefreshCache();
		}
		public void RefreshCache( string[] propertyNames )
		{
			entry.RefreshCache(propertyNames);
		}
		public void Rename( string newName )
		{
			entry.Rename(newName);
		}

		/*
		 * Extended Methods
		 */
		public virtual DirectoryEntryEx CopyTo( DirectoryEntry newParent )
		{
			DirectoryEntry copy = entry.CopyTo(newParent);
			return new DirectoryEntryEx(copy);
		}
		public virtual DirectoryEntryEx CopyTo( DirectoryEntryEx newParent )
		{
			DirectoryEntry copy = entry.CopyTo(newParent.entry);
			return new DirectoryEntryEx(copy);
		}
		public virtual DirectoryEntryEx CopyTo( DirectoryEntry newParent, string newName )
		{
			DirectoryEntry copy = entry.CopyTo(newParent,newName);
			return new DirectoryEntryEx(copy);
		}
		public virtual DirectoryEntryEx CopyTo( DirectoryEntryEx newParent, string newName )
		{
			DirectoryEntry copy = entry.CopyTo(newParent.entry,newName);
			return new DirectoryEntryEx(copy);
		}
		public virtual DirectoryEntryEx CopyTo( DirectoryOU newParent, string newName )
		{
			DirectoryEntry copy = entry.CopyTo(newParent.entry,newName);
			return new DirectoryEntryEx(copy);
		}
		public virtual DirectoryEntryEx CopyTo( DirectoryOU newParent, string newName )
		{
			DirectoryEntry copy = entry.CopyTo(newParent.entry,newName);
			return new DirectoryEntryEx(copy);
		}
		public void Dispose()
		{
			Dispose(true);
		}
		protected virtual void Dispose( bool disposing )
		{
			if( disposed )
				return;
			if( disposing && !this.leaveOpen )
				entry.Dispose();
			disposed = true;
		}
		public DirectoryEntry GetDirectoryEntry()
		{
			return entry;
		}
		public DirectoryOU GetTopLevelOU()
		{
			DirectoryEntry parent = entry.Parent;
			while( parent.Parent != null && String.Equals("organizationalUnit",parent.Parent.SchemaClassName,StringComparison.OrdinalIgnoreCase) )
				parent = parent.Parent;
			return parent != null ? new DirectoryOU(parent) : null;
		}
		public void MoveTo( DirectoryEntryEx newParent )
		{
			entry.MoveTo(newParent.entry);
		}
		public void MoveTo( DirectoryEntryEx newParent, string newName )
		{
			entry.MoveTo(newParent.entry,newName);
		}
		public void MoveTo( DirectoryOU newParent )
		{
			entry.MoveTo(newParent.entry);
		}
		public void MoveTo( DirectoryOU newParent, string newName )
		{
			entry.MoveTo(newParent.entry,newName);
		}
		public void RefreshCache( string propertyNames )
		{
			entry.RefreshCache(new string[]{propertyNames});
		}
		public override string ToString()
		{
			return entry.ToString();
		}
	}
}
