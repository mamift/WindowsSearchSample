using System;
using System.Runtime.InteropServices;

namespace WindowsSearch.Shell
{
    public static class PropVariant
    {
        public static object ToObject(IntPtr pv)
        {
            // Copy to structure
            var v = (PROPVARIANT) Marshal.PtrToStructure(pv, typeof(PROPVARIANT));

            object value = null;
            switch (v.vt) {
                case 0: // VT_EMPTY
                case 1: // VT_NULL
                case 2: // VT_I2
                case 3: // VT_I4
                case 4: // VT_R4
                case 5: // VT_R8
                case 6: // VT_CY
                case 7: // VT_DATE
                case 8: // VT_BSTR
                case 10: // VT_ERROR
                case 11: // VT_BOOL
                case 14: // VT_DECIMAL
                case 16: // VT_I1
                case 17: // VT_UI1
                case 18: // VT_UI2
                case 19: // VT_UI4
                case 20: // VT_I8
                case 21: // VT_UI8
                case 22: // VT_INT
                case 23: // VT_UINT
                case 24: // VT_VOID
                case 25: // VT_HRESULT
                    value = Marshal.GetObjectForNativeVariant(pv);
                    break;

                case 30: // VT_LPSTR
                    value = Marshal.PtrToStringAnsi(v.dataIntPtr);
                    break;

                case 31: // VT_LPWSTR
                    value = Marshal.PtrToStringUni(v.dataIntPtr);
                    break;

                case 0x101f: // VT_VECTOR|VT_LPWSTR
                {
                    var strings = new string[v.cElems];
                    for (var i = 0; i < v.cElems; ++i) {
                        var strPtr = Marshal.ReadIntPtr(v.pElems + i * Marshal.SizeOf(typeof(IntPtr)));
                        strings[i] = Marshal.PtrToStringUni(strPtr);
                    }

                    value = strings;
                }
                    break;

                case 0x1005: // VT_Vector|VT_R8
                {
                    var doubles = new double[v.cElems];
                    Marshal.Copy(v.pElems, doubles, 0, (int) v.cElems);
                    value = doubles;
                }
                    break;

                case 64: // VT_FILETIME
                    value = DateTime.FromFileTime(v.dataInt64);
                    break;

                default:
                    try {
                        value = Marshal.GetObjectForNativeVariant(pv);
                        if (value == null) value = "(null)";
                        value = $"(Supported type 0x{v.vt:x4}): {value.ToString()}";
                    }
                    catch {
                        // Get the variant type
                        value = $"(Unsupported type 0x{v.vt:x4})";
                    }

                    break;
            }

            return value;
        }

        /*
        // C++ version
        typedef struct PROPVARIANT {
            VARTYPE vt;
            WORD    wReserved1;
            WORD    wReserved2;
            WORD    wReserved3;
            union {
                // Various types of up to 8 bytes
            }
        } PROPVARIANT;
        */
        [StructLayout(LayoutKind.Explicit)]
        public struct PROPVARIANT
        {
            [FieldOffset(0)] public ushort vt;
            [FieldOffset(2)] public ushort wReserved1;
            [FieldOffset(4)] public ushort wReserved2;
            [FieldOffset(6)] public ushort wReserved3;
            [FieldOffset(8)] public Int32 data01;
            [FieldOffset(12)] public Int32 data02;

            // IntPtr (for strings and the like)
            [FieldOffset(8)] public IntPtr dataIntPtr;

            // For FileTime and Int64
            [FieldOffset(8)] public long dataInt64;

            // Vector-style arrays (for VT_VECTOR|VT_LPWSTR and such)
            [FieldOffset(8)] public uint cElems;
            [FieldOffset(12)] public IntPtr pElems;
        }
    }
}