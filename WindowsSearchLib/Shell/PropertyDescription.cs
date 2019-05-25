using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace WindowsSearch.Shell
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class PropertyDescription : IDisposable
    {
        private NativeMethods.IPropertyDescription m_IPropertyDescription;

        public PropertyDescription(NativeMethods.IPropertyDescription iPropertyDescription)
        {
            m_IPropertyDescription = iPropertyDescription;
        }

        public PROPERTYKEY PropertyKey
        {
            get {
                PROPERTYKEY value;
                m_IPropertyDescription.GetPropertyKey(out value);
                return value;
            }
        }

        public string CanonicalName
        {
            get {
                var pszName = (IntPtr) 0;
                try {
                    m_IPropertyDescription.GetCanonicalName(out pszName);
                    return Marshal.PtrToStringUni(pszName);
                }
                finally {
                    if (pszName != (IntPtr) 0) {
                        Marshal.FreeCoTaskMem(pszName);
                        pszName = (IntPtr) 0;
                    }
                }
            }
        }

        public string DisplayName
        {
            get {
                var pszName = (IntPtr) 0;
                try {
                    m_IPropertyDescription.GetDisplayName(out pszName);
                    return Marshal.PtrToStringUni(pszName);
                }
                finally {
                    if (pszName != (IntPtr) 0) {
                        Marshal.FreeCoTaskMem(pszName);
                        pszName = (IntPtr) 0;
                    }
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertyDescription()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertyDescription != null) {
                if (!disposing) {
                    Debug.Fail("Failed to dispose PropertyDescription");
                }

                Marshal.FinalReleaseComObject(m_IPropertyDescription);
                m_IPropertyDescription = null;
            }

            if (disposing) {
                GC.SuppressFinalize(this);
            }
        }
    }
}