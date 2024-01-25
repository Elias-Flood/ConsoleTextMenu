using static System.Console;

menuOption[] m_MenuOptions;
int menuIndex = 0, lastMenuIndex = 0;

Start();
void Start()
{
	m_MenuOptions = new menuOption[]
	{
		new menuOption("Inventory", () => DoThing("1")),
		new menuOption("Travel", () => DoThing("2")),
		new menuOption("Explore", () => DoThing("3")),
		new menuOption("Exit", () => DoExit()),
	};

	NewRenderMenu();
	AwaitingMenuInput();
}

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
			lastMenuIndex=menuIndex;
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
void DoThing(string msg)
{
	//Write(msg);
	m_MenuOptions = new menuOption[]
	{
		new menuOption("stin", () => DoThing("1")),
		new menuOption("asjio", () => DoThing("2")),
		new menuOption("Back", () => Start())
	};
	NewRenderMenu();
}

void DoExit()
{
	Environment.Exit(0);
}

public class menuOption
{
	public string DisplayMessage;
	public Action Selected;

	public menuOption(string displayMessage, Action selected)
	{
		DisplayMessage = displayMessage;
		Selected = selected;
	}
}