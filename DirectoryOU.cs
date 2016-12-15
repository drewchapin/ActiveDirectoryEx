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
	public class DirectoryOU: DirectoryEntryEx, IDisposable
	{
		/*
		 * Constructors
		 */
		protected DirectoryOU()
		{
			// Nothing to do.
		}
		public DirectoryOU( DirectoryEntry entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryOU( DirectoryEntryEx entry, bool leaveOpen = false )
			: base(entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryOU( DirectoryOU entry, bool leaveOpen = false )
			: base(entry.entry,leaveOpen)
		{
			// Nothing to do.
		}
		public DirectoryOU( object adsObject )
			: base(adsObject)
		{
			// Nothing to do.
		}
		public DirectoryOU( string path )
			: base(path)
		{
			// Nothing to do.
		}
		public DirectoryOU( string path, string username, string password )
			: base(path,username,password)
		{
			// Nothing to do.
		}
		public DirectoryOU( string path, string username, string password, AuthenticationTypes authenticationType )
			: base(path,username,password,authenticationType)
		{
			// Nothing to do.
		}

		/*
		 * Operators
		 */
		public static explicit operator DirectoryOU( DirectoryEntry entry )
		{
			return new DirectoryOU(entry);
		}

		/*
		 * Extended Properties
		 */
		public string City
		{
			get {
				return entry.Properties.Contains("L") ? (string)entry.Properties["L"].Value : null;
			}
			set {
				entry.Properties["L"].Value = value;
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
		public string FaxNumber
		{
			get {
				return entry.Properties.Contains("facsimileTelephoneNumber") ? (string)entry.Properties["facsimileTelephoneNumber"].Value : null;
			}
			set {
				entry.Properties["facsimileTelephoneNumber"].Value = value;
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
				return entry.Properties.Contains("street") ? (string)entry.Properties["street"].Value : null;
			}
			set {
				entry.Properties["street"].Value = value;
			}
		}

		/*
		 * Extended Methods
		 */
		public new DirectoryOU CopyTo( DirectoryEntry newParent )
		{
			return new DirectoryOU(entry.CopyTo(newParent));
		}
		public new DirectoryOU CopyTo( DirectoryEntry newParent, string newName )
		{
			return new DirectoryOU(entry.CopyTo(newParent,newName));
		}
		public new DirectoryOU CopyTo( DirectoryEntryEx newParent )
		{
			return new DirectoryOU(entry.CopyTo(newParent.GetDirectoryEntry()));
		}
		public new DirectoryOU CopyTo( DirectoryEntryEx newParent, string newName )
		{
			return new DirectoryOU(entry.CopyTo(newParent.GetDirectoryEntry(),newName));
		}
		public DirectoryOU CopyTo( DirectoryOU newParent )
		{
			return new DirectoryOU(entry.CopyTo(newParent.entry));
		}
		public DirectoryOU CopyTo( DirectoryOU newParent, string newName )
		{
			return new DirectoryOU(entry.CopyTo(newParent.entry,newName));;
		}

	}
}
