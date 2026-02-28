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

        Dictionary<int, double> hours = GenHourDict();
        Dictionary<int, Dictionary<int, double>> hoursByDay = GenHoursByDayDicts();
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
                                    if (value is not double && DateTime.TryParse(value?.ToString(), out dt))
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
                CountWorkedHours(hours, hoursByDay, timeCells, readSheetTotal);
            }
        }
    }

    private static void CountWorkedHours(Dictionary<int, double> hours, Dictionary<int, Dictionary<int, double>> hoursByDay, DataRowCollection cells, double readSheetTotal)
    {
        double timesheetHours = 0;
        double startTime;
        double endTime;

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
                if ((cells[day][shift] is not DBNull) && (cells[day][shift + 1] is not DBNull))
                {
                    startTime = ((TimeSpan)cells[day][shift]).Hours + Math.Round((double)(((TimeSpan)cells[day][shift]).Minutes / 60), 2);
                    endTime = ((TimeSpan)cells[day][shift + 1]).Hours + Math.Round((double)(((TimeSpan)cells[day][shift + 1]).Minutes / 60), 2);

                    // Account for a shift finishing at midnight (00:00:00)
                    if (Math.Truncate(endTime) == 0)
                    {
                        endTime += 24;
                    }

                    // For each hour spanned by the shift, add one to relevant hour
                    for (int hr = (int)startTime; hr < (int)endTime; hr++)
                    {
                        hours[hr] += 1;
                        hoursByDay[day][hr] += 1;
                        timesheetHours += 1;
                    }

                    // If start time was not on-the-hour, subtract part-hour from total time
                    if (startTime % 1 > 0)
                    {
                        hours[(int)startTime] -= startTime % 1;
                        hoursByDay[day][(int)startTime] -= startTime % 1;
                        timesheetHours -= startTime % 1;
                    }

                    // If end time was not on-the-hour, add part-hour to total time
                    if (endTime % 1 > 0)
                    {
                        hours[(int)endTime] += endTime % 1;
                        hoursByDay[day][(int)endTime] -= endTime % 1;
                        timesheetHours -= endTime % 1;
                    }

                    if (debugHourCount)
                    {
                        Console.Write(endTime - startTime + " ");
                    }
                }
            }

            // Count get-outs. Takes start time from end of evening/night shift (whichever is later, if one exists), or else assumes 10pm
            // If there's a get-out at all
            if (cells[day][8] is not DBNull && ((TimeSpan)cells[day][8]).Hours != 0)
            {
                startTime = 22;


                if (cells[day][7] is not DBNull) // If night shift exists...
                {
                    startTime = ((TimeSpan)cells[day][7]).Hours + Math.Round((double)((TimeSpan)cells[day][7]).Minutes / 60, 2);
                }
                else if (cells[day][5] is not DBNull) // then if evening shift exists...
                {
                    startTime = ((TimeSpan)cells[day][5]).Hours + Math.Round((double)((TimeSpan)cells[day][5]).Minutes / 60, 2);
                }
                else // no evening/night shift, assume get-out starts at 10pm
                {
                    startTime = 22;
                }

                endTime = startTime + ((TimeSpan)cells[day][8]).Hours + Math.Round((double)((TimeSpan)cells[day][8]).Minutes / 60, 2);

                // For each hour spanned by the shift, add one to relevant hour
                for (int hr = (int)startTime; hr < (int)endTime; hr++)
                {
                    hours[hr % 24] += 1;
                    hoursByDay[(day + (hr / 24)) % 7][hr % 24] += 1;
                    timesheetHours += 1;
                }

                // If start time was not on-the-hour, subtract part-hour from total time (roll over hour to next day as needed)
                if (startTime % 1 > 0)
                {
                    hours[(int)startTime % 24] -= startTime % 1;
                    hoursByDay[(day + ((int)startTime / 24)) % 7][(int)startTime % 24] -= startTime % 1;
                    timesheetHours -= startTime % 1;
                }

                // If end time was not on-the-hour, add part-hour to total time (roll ove hour to next day as needed)
                if (endTime % 1 > 0)
                {
                    hours[(int)endTime % 24] += endTime % 1;
                    hoursByDay[(day + ((int)endTime / 24)) % 7][(int)endTime % 24] += endTime % 1;
                    timesheetHours += endTime % 1;
                }

                if (debugHourCount)
                {
                    Console.Write("GO: " + (endTime - startTime));
                }
            }

            if (debugHourCount)
            {
                Console.WriteLine();
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
                Console.Write(" " + cells[r][c].ToString()?.PadLeft(8));
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
    private static Dictionary<int, double> GenHourDict()
    {
        Dictionary<int, double> h = new Dictionary<int, double>();

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
    private static Dictionary<int, Dictionary<int, double>> GenHoursByDayDicts()
    {
        Dictionary<int, Dictionary<int, double>> d = new Dictionary<int, Dictionary<int, double>>();
        
        for (int day = 0; day < 7; day++)
        {
            d[day] = GenHourDict();
        }

        return d;
    }
}