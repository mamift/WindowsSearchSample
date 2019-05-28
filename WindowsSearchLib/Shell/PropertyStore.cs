using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

// Shell Property System:https://msdn.microsoft.com/en-us/library/windows/desktop/ff728898(v=vs.85).aspx
// Microsoft hasn't provided a good shell wrapper nor does the type library work: http://stackoverflow.com/questions/4450121/c-sharp-4-0-dynamic-object-and-winapi-interfaces-like-ishellitem-without-defini
// Help in managing VARIANT from managed code: https://limbioliong.wordpress.com/2011/09/04/using-variants-in-managed-code-part-1/

namespace WindowsSearch.Shell
{
    // Wrapper Class for IPropertyStore
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class PropertyStore : IDisposable
    {
        public static PropertyStore Open(string filename, bool writeAccess = false)
        {
            NativeMethods.IPropertyStore store;
            var iPropertyStoreGuid = typeof(NativeMethods.IPropertyStore).GUID;
            NativeMethods.SHGetPropertyStoreFromParsingName(filename, (IntPtr) 0,
                writeAccess ? GETPROPERTYSTOREFLAGS.GPS_READWRITE : GETPROPERTYSTOREFLAGS.GPS_BESTEFFORT,
                ref iPropertyStoreGuid, out store);
            return new PropertyStore(store);
        }

        private NativeMethods.IPropertyStore m_IPropertyStore;

        public PropertyStore(NativeMethods.IPropertyStore propertyStore)
        {
            m_IPropertyStore = propertyStore;
        }

        public int Count
        {
            get {
                uint value;
                m_IPropertyStore.GetCount(out value);
                return (int) value;
            }
        }

        public PROPERTYKEY GetAt(int index)
        {
            PROPERTYKEY key;
            m_IPropertyStore.GetAt((uint) index, out key);
            return key;
        }

        public object GetValue(PROPERTYKEY key)
        {
            var pv = IntPtr.Zero;
            object value = null;
            try {
                pv = Marshal.AllocCoTaskMem(16);
                m_IPropertyStore.GetValue(key, pv);
                try {
                    value = PropVariant.ToObject(pv);
                }
                catch (Exception err) {
                    throw new ApplicationException("Unsupported property data type", err);
                }
            }
            finally {
                if (pv != (IntPtr) 0) {
                    try {
                        NativeMethods.PropVariantClear(pv);
                    }
                    catch {
                        Debug.Fail("VariantClear failure");
                    }

                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }

            return value;
        }

        public void SetValue(PROPERTYKEY key, object value)
        {
            var pv = IntPtr.Zero;
            try {
                pv = NativeMethods.PropVariantFromObject(value);
                m_IPropertyStore.SetValue(key, pv);
            }
            finally {
                if (pv != IntPtr.Zero) {
                    NativeMethods.PropVariantClear(pv);
                    Marshal.FreeCoTaskMem(pv);
                    pv = IntPtr.Zero;
                }
            }
        }

        public void Commit()
        {
            m_IPropertyStore.Commit();
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertyStore()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertyStore != null) {
                if (!disposing) {
                    Debug.Fail("Failed to dispose PropertyStore");
                }

                Marshal.FinalReleaseComObject(m_IPropertyStore);
                m_IPropertyStore = null;
            }

            if (disposing) {
                GC.SuppressFinalize(this);
            }
        }
    } // PropertyStore

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY
    {
        public Guid fmtid;
        public UInt32 pid;
    }

    public enum GETPROPERTYSTOREFLAGS : uint
    {
        // If no flags are specified (GPS_DEFAULT), a read-only property store is returned that includes properties for the file or item.
        // In the case that the shell item is a file, the property store contains:
        //     1. properties about the file from the file system
        //     2. properties from the file itself provided by the file's property handler, unless that file is offline,
        //     see GPS_OPENSLOWITEM
        //     3. if requested by the file's property handler and supported by the file system, properties stored in the
        //     alternate property store.
        //
        // Non-file shell items should return a similar read-only store
        //
        // Specifying other GPS_ flags modifies the store that is returned
        GPS_DEFAULT = 0x00000000,
        GPS_HANDLERPROPERTIESONLY = 0x00000001, // only include properties directly from the file's property handler
        GPS_READWRITE = 0x00000002, // Writable stores will only include handler properties
        GPS_TEMPORARY = 0x00000004, // A read/write store that only holds properties for the lifetime of the IShellItem object

        GPS_FASTPROPERTIESONLY =
            0x00000008, // do not include any properties from the file's property handler (because the file's property handler will hit the disk)

        GPS_OPENSLOWITEM =
            0x00000010, // include properties from a file's property handler, even if it means retrieving the file from offline storage.

        GPS_DELAYCREATION =
            0x00000020, // delay the creation of the file's property handler until those properties are read, written, or enumerated

        GPS_BESTEFFORT =
            0x00000040, // For readonly stores, succeed and return all available properties, even if one or more sources of properties fails. Not valid with GPS_READWRITE.
        GPS_NO_OPLOCK = 0x00000080, // some data sources protect the read property store with an oplock, this disables that
        GPS_MASK_VALID = 0x000000FF,
    }

    public enum PROPDESC_ENUMFILTER : uint
    {
        PDEF_ALL = 0,
        PDEF_SYSTEM = 1,
        PDEF_NONSYSTEM = 2,
        PDEF_VIEWABLE = 3,
        PDEF_QUERYABLE = 4,
        PDEF_INFULLTEXTQUERY = 5,
        PDEF_COLUMN = 6
    }

    [Flags]
    public enum PROPDESC_FORMAT_FLAGS : uint
    {
        PDFF_DEFAULT = 0,
        PDFF_PREFIXNAME = 0x1,
        PDFF_FILENAME = 0x2,
        PDFF_ALWAYSKB = 0x4,
        PDFF_RESERVED_RIGHTTOLEFT = 0x8,
        PDFF_SHORTTIME = 0x10,
        PDFF_LONGTIME = 0x20,
        PDFF_HIDETIME = 0x40,
        PDFF_SHORTDATE = 0x80,
        PDFF_LONGDATE = 0x100,
        PDFF_HIDEDATE = 0x200,
        PDFF_RELATIVEDATE = 0x400,
        PDFF_USEEDITINVITATION = 0x800,
        PDFF_READONLY = 0x1000,
        PDFF_NOAUTOREADINGORDER = 0x2000
    }
}