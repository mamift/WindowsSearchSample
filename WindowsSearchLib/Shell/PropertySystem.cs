using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace WindowsSearch.Shell
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    public class PropertySystem : IDisposable
    {
        private NativeMethods.IPropertySystem m_IPropertySystem;

        public PropertySystem()
        {
            var IID_IPropertySystem = typeof(NativeMethods.IPropertySystem).GUID;
            NativeMethods.PSGetPropertySystem(ref IID_IPropertySystem, out m_IPropertySystem);
        }

        public PropertyDescription GetPropertyDescription(PROPERTYKEY propKey)
        {
            var IID_IPropertyDescription = typeof(NativeMethods.IPropertyDescription).GUID;
            NativeMethods.IPropertyDescription iPropertyDescription;
            m_IPropertySystem.GetPropertyDescription(propKey, ref IID_IPropertyDescription, out iPropertyDescription);
            return new PropertyDescription(iPropertyDescription);
        }

        public PropertyDescription GetPropertyDescriptionByName(string canonicalName)
        {
            var IID_IPropertyDescription = typeof(NativeMethods.IPropertyDescription).GUID;
            NativeMethods.IPropertyDescription iPropertyDescription;
            m_IPropertySystem.GetPropertyDescriptionByName(canonicalName, ref IID_IPropertyDescription,
                out iPropertyDescription);
            return new PropertyDescription(iPropertyDescription);
        }

        public PROPERTYKEY GetPropertyKeyByName(string canonicalName)
        {
            using (var pd = GetPropertyDescriptionByName(canonicalName)) {
                return pd.PropertyKey;
            }
        }

        /*
        public PropertyDescriptionList GetPropertyDescriptionListFromString(string propList)
        {
            throw new NotImplementedException();
        }

        public PropertyDescriptionList void EnumeratePropertyDescriptions(PROPDESC_ENUMFILTER)
        {
            throw new NotImplementedException();
        }

        public string FormatForDisplay(PROPERTYKEY propKey, PropVariant propvar, PROPDESC_FORMAT_FLAGS pdff)
        {
            throw new NotImplementedException();
        }

        public void RegisterPropertySchema(string path)
        {
            throw new NotImplementedException();
        }

        public void UnregisterPropertySchema(string path)
        {
            throw new NotImplementedException();
        }

        public void RefreshPropertySchema()
        {
            throw new NotImplementedException();
        }
        */

        public void Dispose()
        {
            Dispose(true);
        }

        ~PropertySystem()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (m_IPropertySystem != null) {
                if (!disposing) {
                    Debug.Fail("Failed to dispose PropertySystem");
                }

                Marshal.FinalReleaseComObject(m_IPropertySystem);
                m_IPropertySystem = null;
            }

            if (disposing) {
                GC.SuppressFinalize(this);
            }
        }
    }
}