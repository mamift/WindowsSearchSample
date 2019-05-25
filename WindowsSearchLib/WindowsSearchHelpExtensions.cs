using System.Data.OleDb;
using System.IO;

namespace WindowsSearch
{
    public static class WindowsSearchHelpExtensions
    {
        public static void WriteColumnNamesToCsv(this OleDbDataReader reader, TextWriter writer)
        {
            var fieldCount = reader.FieldCount;
            for (var i = 0; i < fieldCount; ++i) {
                if (i > 0) writer.Write(',');
                writer.Write(reader.GetName(i));
            }

            writer.WriteLine();
        }

        static readonly char[] sCsvSpecialChars = new char[] {',', '"', '\r', '\n'};

        public static int WriteRowsToCsv(this OleDbDataReader reader, TextWriter writer)
        {
            var rowCount = 0;
            while (reader.Read()) {
                ++rowCount;

                var values = new object[reader.FieldCount];
                reader.GetValues(values);

                for (var i = 0; i < values.Length; ++i) {
                    var value = values[i].ToString();
                    if (value == null) {
                        // Do nothing
                    }
                    else if (value.IndexOfAny(sCsvSpecialChars) >= 0) {
                        writer.Write('"');
                        if (value.IndexOf('"') >= 0)
                            writer.Write(value.Replace("\"", "\"\""));
                        else
                            writer.Write(value);
                        writer.Write('"');
                    }
                    else {
                        writer.Write(value);
                    }

                    if (i < values.Length - 1)
                        writer.Write(',');
                }

                writer.WriteLine();
            }

            reader.Close();
            return rowCount;
        }
    }
}