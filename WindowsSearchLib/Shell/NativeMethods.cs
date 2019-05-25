using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WindowsSearch.Shell
{
    public static class NativeMethods
    {
        /*
        // The C++ Version
        interface IPropertyStore : IUnknown
        {
            HRESULT GetCount([out] DWORD *cProps);
            HRESULT GetAt([in] DWORD iProp, [out] PROPERTYKEY *pkey);
            HRESULT GetValue([in] REFPROPERTYKEY key, [out] PROPVARIANT *pv);
            HRESULT SetValue([in] REFPROPERTYKEY key, [in] REFPROPVARIANT propvar);
            HRESULT Commit();
        }
        */
        [ComImport, Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyStore
        {
            void GetCount([Out] out uint cProps);

            void GetAt([In] uint iProp, out PROPERTYKEY pkey);

            void GetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void SetValue([In] ref PROPERTYKEY key, [In] IntPtr pv);

            void Commit();
        }

        /*
        MIDL_INTERFACE("ca724e8a-c3e6-442b-88a4-6fb0db8035a3")
        IPropertySystem : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescription( 
                __RPC__in REFPROPERTYKEY propkey,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionByName( 
                __RPC__in_string LPCWSTR pszCanonicalName,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyDescriptionListFromString( 
                __RPC__in_string LPCWSTR pszPropList,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE EnumeratePropertyDescriptions( 
                PROPDESC_ENUMFILTER filterOn,
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__out_ecount_full_string(cchText) LPWSTR pszText,
                __RPC__in_range(0,0x8000) DWORD cchText) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplayAlloc( 
                __RPC__in REFPROPERTYKEY key,
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdff,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RegisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE UnregisterPropertySchema( 
                __RPC__in_string LPCWSTR pszPath) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE RefreshPropertySchema( void) = 0;
        
        };
        */
        [ComImport, Guid("ca724e8a-c3e6-442b-88a4-6fb0db8035a3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertySystem
        {
            void GetPropertyDescription([In] ref PROPERTYKEY propkey, [In] ref Guid riid,
                [Out] out IPropertyDescription rPropertyDescription);

            void GetPropertyDescriptionByName([In] [MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName,
                [In] ref Guid riid, [Out] out IPropertyDescription rPropertyDescription);

            void GetPropertyDescriptionListFromString([In] [MarshalAs(UnmanagedType.LPWStr)] string pszPropList,
                [In] ref Guid riid, [Out] out IPropertyDescriptionList rPropertyDescriptionList);

            void EnumeratePropertyDescriptions([In] PROPDESC_ENUMFILTER filterOn, [In] ref Guid riid,
                [Out] out IPropertyDescriptionList rPropertyDescriptionList);

            void FormatForDisplay([In] ref PROPERTYKEY key, [In] IntPtr propvar, [In] PROPDESC_FORMAT_FLAGS pdff,
                [In] IntPtr pszText, ushort cchText);

            void FormatForDisplayAlloc([In] ref PROPERTYKEY key, [In] IntPtr propvar, [In] PROPDESC_FORMAT_FLAGS pdff,
                [Out] out IntPtr ppszText);

            void RegisterPropertySchema([In] [MarshalAs(UnmanagedType.LPWStr)] string pszPath);

            void UnregisterPropertySchema([In] [MarshalAs(UnmanagedType.LPWStr)] string pszPath);

            void RefreshPropertySchema();
        }

        /*
        MIDL_INTERFACE("6f79d558-3e96-4549-a1d1-7d75d2288814")
        IPropertyDescription : public IUnknown
        {
        public:
            virtual HRESULT STDMETHODCALLTYPE GetPropertyKey( 
                __RPC__out PROPERTYKEY *pkey) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetCanonicalName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetPropertyType( 
                __RPC__out VARTYPE *pvartype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayName( 
                __RPC__deref_out_opt_string LPWSTR *ppszName) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEditInvitation( 
                __RPC__deref_out_opt_string LPWSTR *ppszInvite) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetTypeFlags( 
                PROPDESC_TYPE_FLAGS mask,
                __RPC__out PROPDESC_TYPE_FLAGS *ppdtFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetViewFlags( 
                __RPC__out PROPDESC_VIEW_FLAGS *ppdvFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDefaultColumnWidth( 
                __RPC__out UINT *pcxChars) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetDisplayType( 
                __RPC__out PROPDESC_DISPLAYTYPE *pdisplaytype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetColumnState( 
                __RPC__out SHCOLSTATEF *pcsFlags) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetGroupingRange( 
                __RPC__out PROPDESC_GROUPING_RANGE *pgr) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescriptionType( 
                __RPC__out PROPDESC_RELATIVEDESCRIPTION_TYPE *prdt) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetRelativeDescription( 
                __RPC__in REFPROPVARIANT propvar1,
                __RPC__in REFPROPVARIANT propvar2,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc1,
                __RPC__deref_out_opt_string LPWSTR *ppszDesc2) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescription( 
                __RPC__out PROPDESC_SORTDESCRIPTION *psd) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetSortDescriptionLabel( 
                BOOL fDescending,
                __RPC__deref_out_opt_string LPWSTR *ppszDescription) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetAggregationType( 
                __RPC__out PROPDESC_AGGREGATION_TYPE *paggtype) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetConditionType( 
                __RPC__out PROPDESC_CONDITION_TYPE *pcontype,
                __RPC__out CONDITION_OPERATION *popDefault) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE GetEnumTypeList( 
                __RPC__in REFIID riid,
                __RPC__deref_out_opt void **ppv) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE CoerceToCanonicalValue( 
                _Inout_  PROPVARIANT *ppropvar) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE FormatForDisplay( 
                __RPC__in REFPROPVARIANT propvar,
                PROPDESC_FORMAT_FLAGS pdfFlags,
                __RPC__deref_out_opt_string LPWSTR *ppszDisplay) = 0;
        
            virtual HRESULT STDMETHODCALLTYPE IsValueCanonical( 
                __RPC__in REFPROPVARIANT propvar) = 0;
        
        };
        */
        [ComImport, Guid("6f79d558-3e96-4549-a1d1-7d75d2288814"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IPropertyDescription
        {
            void GetPropertyKey([Out] out PROPERTYKEY pkey);

            void GetCanonicalName([Out] out IntPtr ppszName);

            void GetPropertyType([Out] out ushort vartype);

            void GetDisplayName([Out] out IntPtr ppszName);

            // === All Other Methods Deferred Until Later! ===
        }

        public interface IPropertyDescriptionList { }

        [DllImport("shell32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void SHGetPropertyStoreFromParsingName(
            [In] [MarshalAs(UnmanagedType.LPWStr)] string pszPath,
            [In] IntPtr zeroWorks,
            [In] GETPROPERTYSTOREFLAGS flags,
            [In] ref Guid iIdPropStore,
            [Out] out IPropertyStore propertyStore);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PropVariantInit([In] IntPtr pvarg);

        [DllImport(@"ole32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void PropVariantClear([In] IntPtr pvarg);

        [DllImport("propsys.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall, PreserveSig = false)]
        public static extern void
            PSGetPropertySystem([In] ref Guid iIdPropertySystem, [Out] out IPropertySystem propertySystem);

        // Converts a string to a PropVariant with type LPWSTR instead of BSTR
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromString(string value)
        {
            var pstr = IntPtr.Zero;
            var pv = IntPtr.Zero;
            try {
                // In managed code, new automatically zeros the contents.
                var propvariant = new PropVariant.PROPVARIANT();

                // Allocate the string
                pstr = Marshal.StringToCoTaskMemUni(value);

                // Allocate the PropVariant
                pv = Marshal.AllocCoTaskMem(16);

                // Transfer ownership of the string
                propvariant.vt = 31; // VT_LPWSTR - not documented but this is to be allocated using CoTaskMemAlloc.
                propvariant.dataIntPtr = pstr;
                Marshal.StructureToPtr(propvariant, pv, false);
                pstr = IntPtr.Zero;

                // Transfer ownership to the result
                var result = pv;
                pv = IntPtr.Zero;
                return result;
            }
            finally {
                if (pstr != IntPtr.Zero) {
                    Marshal.FreeCoTaskMem(pstr);
                    pstr = IntPtr.Zero;
                }

                if (pv != IntPtr.Zero) {
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
        }

        // Converts an object to a PropVariant including special handling for strings
        // The resulting variant must be cleared using PropVariantClear and freed using Marshal.FreeCoTaskMem
        public static IntPtr PropVariantFromObject(object value)
        {
            var strValue = value as string;
            if (strValue != null) {
                return PropVariantFromString(strValue);
            }
            else {
                var pv = IntPtr.Zero;
                try {
                    pv = Marshal.AllocCoTaskMem(16);
                    Marshal.GetNativeVariantForObject(value, pv);
                    var result = pv;
                    pv = IntPtr.Zero;
                    return result;
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
            }
        }
    }
}