using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Search.Interop;
using System.Data.OleDb;
using System.Linq;

namespace WindowsSearch
{
    /// <summary>
    /// Static class that serves as main entry point for this library.
    /// <para>Call <see cref="PerformSearch"/> to do a full-text search using keywords.</para>
    /// <para>Call <see cref="PerformQuery"/> to do a search using Windows Search SQL.</para>
    /// </summary>
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
        public static List<SearchResult> PerformSearch(string path, string query, IProgress<string> progress = null)
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

            return PerformQuery(path, sqlQuery, progress);
        }

        /// <summary>
        /// Performs a full-text search against the Windows Search Index, by a given file system <paramref name="path"/>
        /// and <paramref name="sqlQuery"/> string. 
        /// </summary>
        /// <param name="path">A path in the local file system.</param>
        /// <param name="sqlQuery">Valid SQL query.</param>
        /// <param name="progress">Optionally provide an <see cref="IProgress{T}"/> object to view progress output.</param>
        public static List<SearchResult> PerformQuery(string path, string sqlQuery, IProgress<string> progress = null)
        {
            using (var session = new WindowsSearchSession(path)) {
                var startTicks = Environment.TickCount;
                var ticksToFirstRead = 0;
                List<SearchResult> rows;
                using (var reader = session.Query(sqlQuery)) {
                    //reader.WriteColumnNamesToCsv(Console.Out);
                    // Need unchecked because tickcount can wrap around - nevertheless it still generates a valid result
                    unchecked {
                        ticksToFirstRead = Environment.TickCount - startTicks;
                    }

                    rows = reader.OutputRowsToSearchResults();

                    var output = $"{rows.Count} rows.";
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

                return rows;
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