using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Search.Interop;
using System.Data.OleDb;

namespace WindowsSearch
{
    public static class Search
    {
        /// <summary>
        /// Performs a full-text search against the Windows Search Index, by a given file system <paramref name="path"/>
        /// and <paramref name="query"/>. This generates an SQL query that is passed to the Windows Search query processor.
        /// <para>If you know in advance the SQL query you want to pass, use <see cref="PerformQuery"/>.</para>
        /// </summary>
        /// <param name="path">A path in the local file system.</param>
        /// <param name="query">Any full-text search string or search operators accepted by the Windows Search query engine.</param>
        /// <param name="progress">Optionally provide an <see cref="IProgress{T}"/> object to view progress output.</param>
        public static void PerformSearch(string path, string query, IProgress<string> progress = null)
        {
            string sqlQuery;
            CSearchManager srchMgr = null;
            CSearchCatalogManager srchCatMgr = null;
            CSearchQueryHelper queryHelper = null;
            try {
                srchMgr = new CSearchManager();
                srchCatMgr = srchMgr.GetCatalog("SystemIndex");
                queryHelper = srchCatMgr.GetQueryHelper();
                sqlQuery = queryHelper.GenerateSQLFromUserQuery(query);
            }
            finally {
                if (queryHelper != null) {
                    Marshal.FinalReleaseComObject(queryHelper);
                    queryHelper = null;
                }

                if (srchCatMgr != null) {
                    Marshal.FinalReleaseComObject(srchCatMgr);
                    srchCatMgr = null;
                }

                if (srchMgr != null) {
                    Marshal.FinalReleaseComObject(srchMgr);
                    srchMgr = null;
                }
            }

            progress?.Report($"Full query: {sqlQuery}");

            PerformQuery(path, sqlQuery);
        }

        /// <summary>
        /// Performs a full-text search against the Windows Search Index, by a given file system <paramref name="path"/>
        /// and <paramref name="sqlQuery"/> string. 
        /// </summary>
        /// <param name="path">A path in the local file system.</param>
        /// <param name="sqlQuery">Valid SQL query.</param>
        /// <param name="silentOutput">Suppress output. If an <see cref="IProgress{T}"/> object is provided, no output is given.</param>
        /// <param name="progress">Optionally provide an <see cref="IProgress{T}"/> object to view progress output.</param>
        public static void PerformQuery(string path, string sqlQuery, bool silentOutput = false, 
            IProgress<string> progress = null)
        {
            using (var session = new WindowsSearchSession(path)) {
                var startTicks = Environment.TickCount;
                var ticksToFirstRead = 0;
                using (var reader = session.Query(sqlQuery)) {
                    reader.WriteColumnNamesToCsv(Console.Out);
                    // Need unchecked because tickcount can wrap around - nevertheless it still generates a valid result
                    unchecked {
                        ticksToFirstRead = Environment.TickCount - startTicks;
                    }

                    var rowCount = silentOutput ? SilentlyReadAllRows(reader) : reader.WriteRowsToCsv(Console.Out);

                    var output = $"{rowCount} rows.";
                    progress?.Report(output);
                    Debug.WriteLine(output);
                }

                int elapsedTicks;
                unchecked {
                    elapsedTicks = Environment.TickCount - startTicks;
                }

                progress?.Report($"{ticksToFirstRead / 1000:d}.{ticksToFirstRead % 1000:d3} until first read.");
                Debug.WriteLine($"{ticksToFirstRead / 1000:d}.{ticksToFirstRead % 1000:d3} until first read.");

                progress?.Report($"{elapsedTicks / 1000:d}.{elapsedTicks % 1000:d3} seconds elapsed.");
                Debug.WriteLine($"{elapsedTicks / 1000:d}.{elapsedTicks % 1000:d3} seconds elapsed.");
            }
        }

        /// <summary>
        /// Reads rows from an <see cref="OleDbDataReader"/> and return the row count.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private static int SilentlyReadAllRows(OleDbDataReader reader)
        {
            var rowCount = 0;
            while (reader.Read()) {
                ++rowCount;
                var values = new object[reader.FieldCount];
                reader.GetValues(values);
            }

            return rowCount;
        }
    }
}