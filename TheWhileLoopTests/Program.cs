using static System.Console;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.SqlServer;
using OfficeOpenXml;
using System.Drawing;
// if you have a commercial license
//ExcelPackage.LicenseContext = LicenseContext.Commercial;
// if you are using epplus for noncommercial purposes, see https://polyformproject.org/licenses/noncommercial/1.0.0/
//ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


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
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
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

        WriteLine("Searching...\n");
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
        WriteLine("\n[Enter] Continue =>");

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
        WriteLine("Starting... \n");
        var tables = dbContext.Table.ToList();

        try
        {
            ExcelPackage excel = new ExcelPackage();

            var excelWS = excel.Workbook.Worksheets.Add("MySheet");

            // Set the cell value using row and column.
            excelWS.Cells[1, 1].Value = "ID";
            excelWS.Cells[1, 2].Value = "Name";
            excelWS.Cells[1, 3].Value = "Email";
            excelWS.Cells[1, 4].Value = "Number";

            for (int i = 0; i < tables.Count; i++)
            {
                excelWS.Cells[(i + 2), 1].Value = tables[i].Id.ToString();
                excelWS.Cells[(i + 2), 2].Value = tables[i].Name.ToString();
                excelWS.Cells[(i + 2), 3].Value = tables[i].Email.ToString();
                excelWS.Cells[(i + 2), 4].Value = tables[i].Number.ToString();
            }

            // The style object is used to access most cells formatting and styles.
            excelWS.Cells[1, 1].Style.Font.Bold = true;
            excelWS.Cells[1, 2].Style.Font.Bold = true;
            excelWS.Cells[1, 3].Style.Font.Bold = true;
            excelWS.Cells[1, 4].Style.Font.Bold = true;

            excelWS.Cells[1, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelWS.Cells[1, 5].Style.Fill.BackgroundColor.SetColor(Color.Red);
            excelWS.Cells[1, 6].Style.Border.DiagonalUp = true;
            excelWS.Cells[1, 6].Style.Border.DiagonalDown = true;

            // file name with .xlsx extension  
            string p_strPath = @"C:\Users\Elias\Documents\myworkbook.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package 
            excel.Dispose();

            WriteLine("* Export Successful *");
        }
        catch (Exception ex)
        {
            WriteLine("* Export Failed *");
            WriteLine(ex.ToString());
        }

        WriteLine("\n...Done\n");
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
