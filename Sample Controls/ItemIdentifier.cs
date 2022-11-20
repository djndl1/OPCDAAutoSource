//============================================================================
// TITLE: ItemIdentifier.cs
//
// CONTENTS:
// 
// A class that contains a unique identifier for an item.
//
// (c) Copyright 2002-2003 The OPC Foundation
// ALL RIGHTS RESERVED.
//
// DISCLAIMER:
//  This code is provided by the OPC Foundation solely to assist in 
//  understanding and use of the appropriate OPC Specification(s) and may be 
//  used as set forth in the License Grant section of the OPC Specification.
//  This code is provided as-is and without warranty or support of any sort
//  and is subject to the Warranty and Liability Disclaimers which appear
//  in the printed OPC Specification.
//
// MODIFICATION LOG:
//
// Date       By    Notes
// ---------- ---   -----
// 2003/11/21 RSA   First release.

using System;
using System.Text;

namespace Opc.Da.SampleClient
{
	/// <summary>
	/// A uniquely identifiable item.
	/// </summary>
	[Serializable]
	public class ItemIdentifier : ICloneable
	{
		/// <summary>
		/// The primary identifier for an item within the server namespace.
		/// </summary>
		public string ItemName = null;

		/// <summary>
		/// An secondary identifier for an item within the server namespace.
		/// </summary>
		public string ItemPath = null;

		/// <summary>
		/// A unique item identifier assigned by the client.
		/// </summary>
		public object ClientHandle = null;

		/// <summary>
		/// A unique item identifier assigned by the server.
		/// </summary>
		public object ServerHandle = null;

		/// <summary>
		/// Creates a shallow copy of the object.
		/// </summary>
		public virtual object Clone() { return MemberwiseClone(); }

		/// <summary>
		/// Create a string that can be used as index in a hash table for the item.
		/// </summary>
		public string Key
		{ 
			get 
			{
				return new StringBuilder(64)
					.Append((ItemName == null)?"null":ItemName)
					.Append("\r\n")
					.Append((ItemPath == null)?"null":ItemPath)
					.ToString();
			}
		}

		/// <summary>
		/// Initializes the object with default values.
		/// </summary>
		public ItemIdentifier() {}

		/// <summary>
		/// Initializes the object with the specified item name.
		/// </summary>
		public ItemIdentifier(string itemName)
		{
			ItemPath = null;
			ItemName = itemName;
		}

		/// <summary>
		/// Initializes the object with the specified item path and item name.
		/// </summary>
		public ItemIdentifier(string itemPath, string itemName)
		{
			ItemPath = itemPath;
			ItemName = itemName;
		}
		
		/// <summary>
		/// Initializes the object with the specified item identifier.
		/// </summary>
		public ItemIdentifier(ItemIdentifier itemID)
		{
			if (itemID != null)
			{
				ItemPath     = itemID.ItemPath;
				ItemName     = itemID.ItemName;
				ClientHandle = itemID.ClientHandle;
				ServerHandle = itemID.ServerHandle;
			}
		}
	}
}
