using static System.Console;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.SqlServer;
using Excel = Microsoft.Office.Interop.Excel;

#region DatabaseGet
public class MyDbContext : DbContext
{
    public DbSet<Table> Table { get; set; }
    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
    {
        optionsBuilder.UseSqlServer(@"Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=New Database;Integrated Security=True;Connect Timeout=30;Encrypt=False;Trust Server Certificate=False;Application Intent=ReadWrite;Multi Subnet Failover=False");
    }
}

public class Table
{
    public int Id { get; set; }
    public string Name { get; set; }
    public string Email { get; set; }
    public int? Number { get; set; }
}
#endregion

public class menuOption
{
    public string DisplayMessage;
    public System.Action Selected;

    public menuOption(string displayMessage, System.Action selected)
    {
        DisplayMessage = displayMessage;
        Selected = selected;
    }
}


public class Program
{
    public MyDbContext dbContext = new MyDbContext();
    menuOption[] m_MenuOptions;
    int menuIndex = 0, lastMenuIndex = 0;

    static void Main(string[] args)
    {
        var program = new Program();
        program.Awaken();
    }

    #region Navigation & Rendering
    void AwaitingMenuInput()
    {
        //CursorVisible = false;
        bool choosingOption = true;

        while (choosingOption)
        {
            MenuSelection();
            var input = ReadKey(true).Key;
            if (input == ConsoleKey.DownArrow && menuIndex < m_MenuOptions.Length - 1)
            {
                lastMenuIndex = menuIndex;
                menuIndex++;
            }
            else if (input == ConsoleKey.UpArrow && menuIndex > 0)
            {
                lastMenuIndex = menuIndex;
                menuIndex--;
            }
            else if (input == ConsoleKey.Enter)
            {
                lastMenuIndex = menuIndex;
                m_MenuOptions[menuIndex].Selected.Invoke();
                menuIndex = 0;
                lastMenuIndex = 0;
            }
        }
    }

    void MenuSelection()
    {
        SetCursorPosition(2, lastMenuIndex + 3);
        ResetColor();
        Write("");
        Write(m_MenuOptions[lastMenuIndex].DisplayMessage);

        SetCursorPosition(2, menuIndex + 3);
        Write("");
        BackgroundColor = ConsoleColor.White;
        ForegroundColor = ConsoleColor.Black;
        Write(m_MenuOptions[menuIndex].DisplayMessage);

        ResetColor();
        SetCursorPosition(2, 5 + m_MenuOptions.Length);
    }

    void NewRenderMenu()
    {
        Clear();
        menuIndex = 0;
        lastMenuIndex = 0;
        WriteLine("Use up/down arrow keys to choose. Enter to select. Esc to exit.");
        WriteLine("\nChoose an option:");
        for (int i = 0; i < m_MenuOptions.Length; i++)
        {
            Write("");
            WriteLine($"> {m_MenuOptions[i].DisplayMessage}");
        }

        ///Ensure that the whole console stays black.
        Write("");
        BackgroundColor = ConsoleColor.Black;
        ForegroundColor = ConsoleColor.White;
        MenuSelection();
    }
    #endregion

    #region Actions
    void Awaken()
    {
        if(dbContext != null)
        {
            WriteLine("Successfully connected to Database");
        }
        else
        {
            WriteLine("Failed");
        }
        ReadKey(true);
        Start();
    }

    void Start()
    {
        m_MenuOptions = new menuOption[]
        {
                new menuOption("Read", () => Read_DB()),
                new menuOption("Export", () => Exoort_DB()),
                new menuOption("Exit", () => DoExit()),
        };

        NewRenderMenu();
        AwaitingMenuInput();
    }

    void DoThing(string msg)
    {
        Write(msg);
        m_MenuOptions = new menuOption[]
        {
                new menuOption("stin", () => DoThing("1")),
                new menuOption("asjio", () => DoThing("2")),
                new menuOption("Back", () => Start())
        };
        NewRenderMenu();
    }

    void Read_DB()
    {
        Clear();

        WriteLine("Searching...");
        var tables = dbContext.Table.ToList();

        if (tables != null)
        {
            WriteLine("Found Tables:");
            foreach (var table in tables)
            {
                WriteLine($"ID: {table.Id}, Name: {table.Name}, Email: {table.Email}, Number: {table.Number}");
            }

        }
        else
        {
            Write(
            $"Failed"
            );
        }
        WriteLine("[Enter] Continue =>");

        ReadKey(true);
        m_MenuOptions = new menuOption[]
        {
                new menuOption("Back", () => Start())
        };
        NewRenderMenu();
    }

    void Exoort_DB()
    {
        Clear();
        WriteLine("Starting...");

        //ExcelExport();

        //string currentDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
        //string logFolder = @"C:\Users\Elias\Downloadss";
        string filePath = @"C:\Users\Elias\Downloads\dbExportToExcel.XLSX";
        var tables = dbContext.Table.ToList();


        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWB = excelApp.Workbooks.Add();
            Excel.Worksheet excelWS = (Excel.Worksheet)excelWB.ActiveSheet;

            excelWS.Cells[1, 1] = "ID";
            excelWS.Cells[1, 2] = "Name";
            excelWS.Cells[1, 3] = "Email";
            excelWS.Cells[1, 4] = "Number";

            //for (int i = 0; i < tables.Count; i++)
            //{
            //    excelWS.Cells[(i + 1), 1] = tables[i].Id.ToString();
            //    excelWS.Cells[(i + 1), 2] = tables[i].Name.ToString();
            //    excelWS.Cells[(i + 1), 3] = tables[i].Email.ToString();
            //    excelWS.Cells[(i + 1), 4] = tables[i].Number.ToString();
            //}

            //excelWB.SaveCopyAs(filePath);
            excelWB.Close();
            excelApp.Quit();

            WriteLine("Creation Complete");
        }
        catch (Exception ex)
        {
            WriteLine("Creation Failed");
            WriteLine(ex.ToString());
        }

        WriteLine("...Done");

        WriteLine("[Enter] Continue =>");
        ReadKey(true);

        m_MenuOptions = new menuOption[]
        {
                new menuOption("Back", () => Start())
        };
        NewRenderMenu();
    }

    void DoExit()
    {
        Environment.Exit(0);
    }
    #endregion

    //void ExcelExport()
    //{
    //    string currentDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
    //    string logFolder = @"D:\Files\Logs";
    //    string filePath = @"D:\Files\dbExportToExcel.XLSX";

    //    try
    //    {
    //        if (File.Exists(filePath))
    //            File.Delete(filePath);

    //        var tables = dbContext.Table.ToList();

    //        Excel.Application excelApp = new Excel.Application();
    //        Excel.Workbook excelWB = excelApp.Workbooks.Add();
    //        Excel.Worksheet excelWS = (Excel.Worksheet)excelWB.ActiveSheet;

    //        excelWS.Cells[1, 1] = "ID";
    //        excelWS.Cells[1, 2] = "Name";
    //        excelWS.Cells[1, 3] = "Email";
    //        excelWS.Cells[1, 4] = "Number";

    //        for (int i = 0; i < tables.Count; i++)
    //        {
    //            excelWS.Cells[(i + 1), 1] = tables[i].Id.ToString();
    //            excelWS.Cells[(i + 1), 2] = tables[i].Name.ToString();
    //            excelWS.Cells[(i + 1), 3] = tables[i].Email.ToString();
    //            excelWS.Cells[(i + 1), 4] = tables[i].Number.ToString();
    //        }

    //        excelWB.SaveCopyAs(filePath);
    //        excelWB.Close();
    //        excelApp.Quit();

    //        WriteLine("Creation Complete");
    //    }
    //    catch (Exception ex)
    //    {
    //        WriteLine("Creation Failed");
    //        WriteLine(ex.ToString());
    //    }
    //}
}
