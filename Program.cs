using System;
using System.IO;
using WindowsSearch;

namespace WindowsSearchSample
{
    public static class Program
    {
        const string c_syntax =
@"Syntax:
    WindowsSearchSample -lib <libraryRoot> [options]
Options:
    -h                  Write this help text.
    -lib <LibraryRoot>  Path to the root of the folder tree to be searched
    -s <windows search> Perform a search using Windows Search syntax
    -q <SQL Query>      Perform a search using SQL syntax
    -x                  Silent output - just reads all rows to time how long the query takes.
";
        // 78 Columns                                                                |

        /* Sample Command Lines:
        -lib \\Ganymede\Archive\Photos -q "SELECT System.ItemPathDisplay, System.Photo.CameraModel, System.Photo.CameraManufacturer, System.Photo.DateTaken FROM SystemIndex WHERE CONTAINS(System.Photo.CameraModel, '\"EZ Controller\"',1033) AND System.Photo.CameraManufacturer = 'NORITSU KOKI' AND System.Photo.DateTaken = '2013/11/20 18:15:06'"
        -lib \\Ganymede\Archive\Photos -s "cameramodel:\"EZ Controller\" cameramaker:\"NORITSU KOKI\" datetaken:11/20/2013 11:15 AM"
        */

        static bool s_silent = false;

        static void Main(string[] args)
        {
            bool writeSyntax = false;
            string libraryPath = null;
            string winSearch = null;
            string sqlQuery = null;

            try
            {
                for (int nArg = 0; nArg < args.Length; ++nArg)
                {
                    switch (args[nArg].ToLower())
                    {
                        case "-h":
                            writeSyntax = true;
                            break;

                        case "-lib":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-lib'");
                            libraryPath = Path.GetFullPath(args[nArg]);
                            if (!Directory.Exists(libraryPath))
                            {
                                throw new ArgumentException(String.Format("Folder '{0}' not found.", libraryPath));
                            }
                            break;

                        case "-s":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-s'");
                            winSearch = args[nArg];
                            break;

                        case "-q":
                            ++nArg;
                            if (nArg >= args.Length) throw new ArgumentException("Command-Line Syntax Error: No value specified for '-q'");
                            sqlQuery = args[nArg];
                            break;

                        case "-x":
                            s_silent = true;
                            break;


                        default:
                            throw new ArgumentException(string.Format("Unexpected command-line parameter '{0}'", args[nArg]));
                    }
                }

                if (writeSyntax)
                {
                    // Do nothing here
                }
                else if (libraryPath == null)
                {
                    throw new ArgumentException("Missing -lib argument.");
                }
                else if (winSearch != null)
                {
                    var results = Search.PerformSearch(libraryPath, winSearch);

                    foreach (var result in results)
                        Console.WriteLine($"{result.FilePath}");
                }
                else if (sqlQuery != null)
                {
                    var results = Search.PerformQuery(libraryPath, sqlQuery);
                    foreach (var result in results)
                        Console.WriteLine($"{result.FilePath}");
                }
                else
                {
                    throw new ArgumentException("No operation option specified.");
                }
            }
            catch (Exception err)
            {
                if (err is ArgumentException) writeSyntax = true;
#if DEBUG
                Console.Error.WriteLine(err.ToString());
#else
        Console.Error.WriteLine(err.Message);
#endif
                Console.Error.WriteLine();
            }

            if (writeSyntax) Console.Error.Write(c_syntax);

            if (Win32Interop.ConsoleHelper.IsSoleConsoleOwner)
            {
                Console.Error.WriteLine();
                Console.Error.Write("Press any key to exit.");
                Console.ReadKey(true);
            }
        }
    }
}

