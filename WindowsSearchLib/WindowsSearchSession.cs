using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;

namespace WindowsSearch
{
    public class WindowsSearchSession : IDisposable
    {
        private OleDbConnection _mDbConnection = null;
        private readonly string _mPathInUrlForm;
        private readonly string _mHostPrefix;

        public WindowsSearchSession(string path)
        {
            path = Path.GetFullPath(path);
            _mPathInUrlForm = path.Replace('\\', '/');

            // Get host prefix (empty string if localhost)
            if (_mPathInUrlForm.StartsWith("//", StringComparison.Ordinal))
            {
                var slash = _mPathInUrlForm.IndexOf('/', 2);
                if (slash > 1)
                {
                    _mHostPrefix = string.Concat(_mPathInUrlForm.Substring(2, slash - 2), ".");
                }
                else
                {
                    throw new ArgumentException($"WindowsSearchSession - Invalid Path: '{path}'", "path");
                }
            }
            else
            {
                _mHostPrefix = string.Empty;
            }

            _mDbConnection = new OleDbConnection("Provider=Search.CollatorDSO;Extended Properties='Application=Windows';");
            _mDbConnection.Open();
        }

        public string[] GetAllKeywords()
        {
            var query = $"SELECT System.Keywords FROM {_mHostPrefix}SystemIndex WHERE SCOPE='file:{_mPathInUrlForm}'";
            Debug.WriteLine(query);

            var keywords = new HashSet<string>();

            var nReads = 0;
            var nValues = 0;
            var maxValuesPerRead = 0;

            using (var cmd = new OleDbCommand(query, _mDbConnection))
            {
                using (var rdr = cmd.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        ++nReads;

                        var values = rdr[0] as string[];
                        if (values != null)
                        {
                            foreach (var value in values)
                            {
                                ++nValues;
                                keywords.Add(value);
                            }
                            if (maxValuesPerRead < values.Length) maxValuesPerRead = values.Length;
                        }
                    }
                    rdr.Close();
                }
            }

            Debug.WriteLine("{0} reads, {1} values, {2} maxValuesPerRead, {3} distinct values", nReads, nValues, maxValuesPerRead, keywords.Count);

            var kwList = new List<string>(keywords);
            kwList.Sort();

            return kwList.ToArray();
        }

        private static readonly Regex SRxSystemIndex = new Regex(@"\sFROM\s+""?SystemIndex""?\s+WHERE\s+", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
        private static readonly Regex SRxSystemIndex2 = new Regex(@"\sFROM\s+""?SystemIndex""?\s*$", RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        public OleDbDataReader Query(string sql)
        {
            // Update the scope in the SQL statement
            if (SRxSystemIndex.Match(sql).Success)
            {
                sql = SRxSystemIndex.Replace(sql, $@" FROM {_mHostPrefix}SystemIndex WHERE SCOPE='file:{_mPathInUrlForm}' AND ");
            }
            else if (SRxSystemIndex2.Match(sql).Success)
            {
                sql = SRxSystemIndex2.Replace(sql, $@" FROM {_mHostPrefix}SystemIndex WHERE SCOPE='file:{_mPathInUrlForm}'");
            }
            else
            {
                 throw new ApplicationException("SQL Statement didn't match expected syntax.");
            }
            Debug.WriteLine(sql);
            using (var cmd = new OleDbCommand(sql, _mDbConnection))
            {
                cmd.CommandTimeout = 600; // Seconds
                return cmd.ExecuteReader();
            }
        }

        public void Dispose()
        {
            Dispose(true);
        }

        ~WindowsSearchSession()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (_mDbConnection != null)
            {
                _mDbConnection.Dispose();
                _mDbConnection = null;
                GC.SuppressFinalize(this);
#if DEBUG
                if (!disposing)
                {
                    Debug.Fail("Failed to dispose WindowsSearchSession.");
                }
#endif
            }
        }

    }
}
