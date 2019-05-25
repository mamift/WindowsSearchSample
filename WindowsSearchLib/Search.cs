using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Search.Interop;
using System.Data.OleDb;

namespace WindowsSearch
{
    public static class Search
    {
        public static void PerformSearch(string libPath, string query)
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

            Console.Error.WriteLine(sqlQuery);
            Console.Error.WriteLine();

            PerformQuery(libPath, sqlQuery);
        }

        public static void PerformQuery(string libPath, string sqlQuery, bool sSilent = false)
        {
            using (var session = new WindowsSearchSession(libPath)) {
                var startTicks = Environment.TickCount;
                var ticksToFirstRead = 0;
                using (var reader = session.Query(sqlQuery)) {
                    reader.WriteColumnNamesToCsv(Console.Out);
                    // Need unchecked because tickcount can wrap around - nevertheless it still generates a valid result
                    unchecked {
                        ticksToFirstRead = Environment.TickCount - startTicks;
                    }

                    int rowCount;
                    if (!sSilent) {
                        rowCount = reader.WriteRowsToCsv(Console.Out);
                    }
                    else {
                        rowCount = SilentlyReadAllRows(reader);
                    }

                    Console.Error.WriteLine();
                    Console.Error.WriteLine("{0} rows.", rowCount);
                    Debug.WriteLine("{0} rows.", rowCount);
                }

                int elapsedTicks;
                unchecked {
                    elapsedTicks = Environment.TickCount - startTicks;
                }

                Console.Error.WriteLine($"{ticksToFirstRead / 1000:d}.{ticksToFirstRead % 1000:d3} until first read.");
                Debug.WriteLine($"{ticksToFirstRead / 1000:d}.{ticksToFirstRead % 1000:d3} until first read.");

                Console.Error.WriteLine($"{elapsedTicks / 1000:d}.{elapsedTicks % 1000:d3} seconds elapsed.");
                Debug.WriteLine($"{elapsedTicks / 1000:d}.{elapsedTicks % 1000:d3} seconds elapsed.");
            }
        }

        public static int SilentlyReadAllRows(OleDbDataReader reader)
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