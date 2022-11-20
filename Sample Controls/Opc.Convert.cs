//============================================================================
// TITLE: Convert.cs
//
// CONTENTS:
// 
// Contains classes that facilitate type conversion.
//
// (c) Copyright 2003 The OPC Foundation
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
// 2003/03/26 RSA   Initial implementation.
//

using System;
using System.Xml;
using System.Collections;
using System.Text;

namespace Opc
{
	/// <summary>
	/// Declares constants for common XML Schema and OPC namespaces.
	/// </summary>
	public class Convert
	{
		/// <summary>
		/// Performs a deep copy of an object if possible.
		/// </summary>
		public static object Clone(object source)
		{
			if (source == null)               return null;
			if (source.GetType().IsValueType) return source;

			if (source.GetType().IsArray || source.GetType() == typeof(Array))
			{
				Array array = (Array)((Array)source).Clone();

				for (int ii = 0; ii < array.Length; ii++)
				{
					array.SetValue(Convert.Clone(array.GetValue(ii)), ii);
				}

				return array;
			}

			try   { return ((ICloneable)source).Clone(); }
			catch { throw new NotSupportedException("Object cannot be cloned."); }
		}

		/// <summary>
		/// Does a deep comparison between two objects.
		/// </summary>
		public static bool Compare(object a, object b)
		{
			if (a == null || b == null) return (a == null && b == null);

			System.Type type1 = a.GetType();
			System.Type type2 = b.GetType();

			if (type1 != type2) return false;

			if (type1.IsArray && type2.IsArray)
			{
				Array array1 = (Array)a;
				Array array2 = (Array)b;

				if (array1.Length != array2.Length) return false;

				for (int ii = 0; ii < array1.Length; ii++)
				{
					if (!Compare(array1.GetValue(ii), array2.GetValue(ii))) return false;
				}

				return true;
			}

			return a.Equals(b);
		}

		/// <summary>
		/// Converts an object to the specified type and returns a deep copy.
		/// </summary>
		public static object ChangeType(object source, System.Type newType)
		{
			// check for null source object.
			if (source == null)
			{
				if (newType != null && newType.IsValueType)
				{
					return Activator.CreateInstance(newType);
				}

				return null;
			}

			// check for null type or 'object' type.
			if (newType == null || newType == typeof(object)) 
			{
				return Clone(source);
			}

			System.Type type = source.GetType();

			// convert between array types.
			if (type.IsArray && newType.IsArray)
			{
				ArrayList array = new ArrayList(((Array)source).Length);

				foreach (object element in (Array)source)
				{
					array.Add(Opc.Convert.ChangeType(element, newType.GetElementType()));
				}

				return array.ToArray(newType.GetElementType());
			}

			// convert scalar value to an array type.
			if (!type.IsArray && newType.IsArray)
			{
				ArrayList array = new ArrayList(1);
				array.Add(Convert.ChangeType(source, newType.GetElementType()));
				return array.ToArray(newType.GetElementType());
			}

			// convert single element array type to scalar type.
			if (type.IsArray && !newType.IsArray && ((Array)source).Length == 1)
			{
				return System.Convert.ChangeType(((Array)source).GetValue(0), newType);
			}

			// use default conversion.
			return System.Convert.ChangeType(source, newType);
		}

		/// <summary>
		/// Formats an item or property value as a string.
		/// </summary>
		public static string ToString(object source)
		{
			// check for null
			if (source == null) return "";

			System.Type type = source.GetType();

			// check for invalid values in date times.
			if (type == typeof(DateTime))
			{
				if (((DateTime)source) == DateTime.MinValue)
				{
					return String.Empty;
				}

				return ((DateTime)source).ToString("yyyy/MM/dd hh:mm:ss.fff");
			}

			// use only the local name for qualified names.
			if (type == typeof(XmlQualifiedName))
			{
				return ((XmlQualifiedName)source).Name;
			}		
		
			// use only the name for system types.
			if (type.FullName == "System.RuntimeType")
			{
				return ((System.Type)source).Name;
			}

			// treat byte arrays as a special case.
			if (type == typeof(byte[]))
			{
				byte[] bytes = (byte[])source;

				StringBuilder buffer = new StringBuilder(bytes.Length*3);

				for (int ii = 0; ii < bytes.Length; ii++)
				{
					buffer.Append(bytes[ii].ToString("X2"));
					buffer.Append(" ");
				}

				return buffer.ToString();
			}
		
			// show the element type and length for arrays.
			if (type.IsArray)
			{
				return String.Format("{0}[{1}]", type.GetElementType().Name, ((Array)source).Length);
			}

			// instances of array are always treated as arrays of objects.
			if (type == typeof(Array))
			{
				return String.Format("Object[{0}]", ((Array)source).Length);
			}

			// default behavoir.
			return source.ToString();
		}
	}
}
