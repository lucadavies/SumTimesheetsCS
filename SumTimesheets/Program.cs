using DynamicObj;
using ExcelDataReader;
using Plotly.NET;
using System.Data;
using System.Text;

internal class Program
{
    private static bool debugCellRead = true;
    private static bool debugHourCount = true;
    private static bool showByDay = false;
    private static string timesheetsLocation = GetTimesheetDirPath();
    private static Dictionary<int, string> indToDay = new Dictionary<int, string>()
    {
        { 0, "Sun" },
        { 1, "Mon" },
        { 2, "Tue" },
        { 3, "Wed" },
        { 4, "Thu" },
        { 5, "Fri" },
        { 6, "Sat" }
    };
    private static void Main(string[] args)
    {
        // Required by ExcelDataReader for reading .xls files
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        Dictionary<int, int> hours = GenHourDict();
        Dictionary<int, Dictionary<int, int>> hoursByDay = GenHoursByDayDicts();
        int fileCount = 0;

        DirectoryInfo timesheetDir = new DirectoryInfo(timesheetsLocation);
        if (timesheetDir.Exists && timesheetDir is not null)
        {
            DataSet readData;
            DataRowCollection timeCells;
            double readSheetTotal = 0;

            // Iterate over all files in timesheet directory
            foreach (FileInfo f in timesheetDir.GetFiles())
            {
                fileCount++;
                using (var stream = File.Open(f.FullName, FileMode.Open, FileAccess.Read))
                {
                    if (debugCellRead || debugHourCount)
                    {
                        Console.WriteLine("\n[" + fileCount + "] | " + f.Name);
                    }
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        readData = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            UseColumnDataType = false,

                            /* Filtered such that we have 8 rows of 9 columns each.
                             * First 7 rows are each day: each one has pairs of start/end times, and one get-out length at the end
                             * Last row only contains one wanted value, at index 5: total hours on sheet
                             */
                            ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = false,
                                FilterColumn = (IExcelDataReader columnReader, int num) =>
                                {
                                    if (num >= 1 && num <= 8 || num == 13)
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                },
                                FilterRow = (IExcelDataReader rowReader) =>
                                {
                                    if (rowReader.Depth >= 5 && rowReader.Depth <= 11 || rowReader.Depth == 13)
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }

                                },
                                TransformValue = (IExcelDataReader tableReader, int n, object value) =>
                                {
                                    DateTime dt = new();
                                    if (DateTime.TryParse(value?.ToString(), out dt))
                                    {
                                        if (n == 8) // if get-out column, read minutes and hours, and seconds as minutes
                                        {
                                            DateTime getOutDt = new();
                                            getOutDt = getOutDt.AddHours(dt.Minute);
                                            getOutDt = getOutDt.AddMinutes(dt.Second);
                                            return getOutDt.TimeOfDay;
                                        }

                                        return dt.TimeOfDay;
                                    }
                                    else if (value is null)
                                    {
                                        return 0;
                                        //return new DateTime(0).TimeOfDay;
                                    }
                                    else
                                    {
                                        return value;
                                    }
                                }
                            }
                        });
                        timeCells = readData.Tables[0].Rows;
                        readSheetTotal = Math.Round((double)timeCells[7][5], 1); // Read total number of hours on timesheet (for debug/checking purposes)
                    }
                }

                if (debugCellRead)
                {
                    PrintCells(timeCells);
                }
                //CountWorkedHours(hours, hoursByDay, timeCells, readSheetTotal);
            }
        }
    }

    private static void CountWorkedHours(Dictionary<int, int> hours, Dictionary<int, Dictionary<int, int>> hoursByDay, DataRowCollection cells, int readSheetTotal)
    {
        int timeSheetHours = 0;
        int startTime;
        int endTime;

        for (int day = 0; day < 7; day++)
        {
            if (debugHourCount)
            {
                Console.Write(indToDay[day] + ": ");
            }

            // For each shift start/end time pair...
            for (int shift = 0; shift < 8; shift += 2)
            {
                // Check shift has both start AND end time
                //if (cells[day][shift] != 0) && (cells[day][shift + 1] != 0)
                //{

                //}
                //startTime = cells[day][shift] != 0) && 
            }
        }
    }

    private static void PrintCells(DataSet cells)
    {
        DateTime getOutAdjusted;
        for (int row = 0; row < indToDay.Count; row++)
        {
            Console.Write(indToDay[row] + " ");
            for (int col = 0; col < 13; col++)
            {
                getOutAdjusted = new();
                var val = cells.Tables[0].Rows[row][col];
                if (val is DateTime dt)
                {
                    Console.Write(dt.TimeOfDay + " ");  
                }
                else if (col == 12)
                {
                    // TODO This is absolutley rank. Fix it.
                    getOutAdjusted = getOutAdjusted.AddHours(((TimeSpan)val).Minutes);
                    getOutAdjusted = getOutAdjusted.AddMinutes(((TimeSpan)val).Seconds);
                    Console.Write(getOutAdjusted.TimeOfDay + " ");
                }
                else
                {
                    Console.Write(val?.ToString() + " ");
                }                
            }
            Console.WriteLine();
        }
    }

    private static void PrintCells(DataRowCollection cells)
    {
        for (int r = 0; r < 7; r++)
        {
            Console.Write(indToDay[r] + ":");
            for (int c = 0; c < 9; c++)
            {
                Console.Write(" " + cells[r][c].ToString().PadLeft(8));
            }
            Console.WriteLine();
        }
    }

    /// <summary>
    /// Returns the full path of directory containing timesheets to be processed.
    /// Default is in a root level folder "Excel"
    /// </summary>
    /// <returns>The path to the timesheets folder.</returns>
    private static string GetTimesheetDirPath()
    {
        string workingDir = Environment.CurrentDirectory;
        string projDir = Directory.GetParent(workingDir).Parent.Parent.FullName;
        return projDir + "\\Excel";
    }

    /// <summary>
    /// Generates a fresh dictionary for each hour of the day.
    /// </summary>
    /// <returns>A fresh dictionary with 24 entries.</returns>
    private static Dictionary<int, int> GenHourDict()
    {
        Dictionary<int, int> h = new Dictionary<int, int>();

        for (int hour = 0; hour < 24; hour++)
        {
            h[hour] = 0;
        }

        return h;
    }

    /// <summary>
    /// Generates a fresh dictionary of dictionaries for each hour of each day.
    /// </summary>
    /// <returns>A fresh dictionary with 7 entries, each containing a fresh dictionary of 24 entries.</returns>
    private static Dictionary<int, Dictionary<int, int>> GenHoursByDayDicts()
    {
        Dictionary<int, Dictionary<int, int>> d = new Dictionary<int, Dictionary<int, int>>();
        
        for (int day = 0; day < 7; day++)
        {
            d[day] = GenHourDict();
        }

        return d;
    }
}