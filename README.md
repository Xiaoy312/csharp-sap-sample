# SAP Inspector.linq
> This is a LINQPad script that allow you to inspect the current SAP window.

Note: 
[`SAP.GUI.Scripting.net`](https://www.nuget.org/packages/SAP.GUI.Scripting.net/)  is created to be referenced by LINQPad, as it doesn't support type and COM library. And, it targets specifically `730 Final Release version` of SAP. If you do need a different version of them you can convert them using the Type Library Importer (Tlbimp.exe) shipped with visual studio. For more information, check out these links:
- [Importing a Type Library as an Assembly](https://docs.microsoft.com/en-us/dotnet/framework/interop/importing-a-type-library-as-an-assembly)
- [Tlbimp.exe (Type Library Importer)](https://docs.microsoft.com/en-us/dotnet/framework/tools/tlbimp-exe-type-library-importer)

## Features:
- Browsing the "visual tree"
- Inspect the controls' name, content, and type
- Generate code to get this control: `.FindByName<__TYPE__>("__NAME__")`
- Highlight (draw red rectangle around) a control in SAP
- Custom:
  - `GuiTableControl`:
    - list table columns
    - export table columns as text:
      ```
      Index Type          Title            Tooltip        
      0     GuiTextField  Début            Date de début  
      1     GuiTextField  Fin              Date de fin    
      2     GuiCTextField Mode communicat. Mode communicat.
      3     GuiTextField  Désignation      Désignation    
      4     GuiTextField  ID du système    ID/N° long      
      5     GuiTextField  CB               Code de blocage
      ```
    - view table content
  - Visualizer for "Label Map"
    > Some reports are made entirely with GuiLabels placed a virtual grid.
    > They will be converted into a single string and displayed in monospace for inspection or copying (without losing padding).

# Xy.Sap
> This namespace contains a set of helper and extension method to interact the SAP system.

Retrieving employee identity:
```csharp
    // obtain a reference to an opened SAP connection
    var session = SapHelper.GetActiveSession();

    // launch "PA20: Display HR Naster Data" on the active window
    const string TransactionTitle = "Afficher données de base personnel";
    var window = session.BeginTransaction("PA20", "Display HR Master Data");

    // fill search form with data
    window.FindByName<GuiCTextField>("RP50G-PERNR").Text = employeeId.ToString();
    window.FindByName<GuiCTextField>("RP50G-CHOIC").Text = "0002"; // identity
    window.FindByName<GuiCTextField>("RP50G-SUBTY").Text = null;

    // search: F7 (7) for list view, Ctrl+F8 (20) for detail view
    window.SendVKey(7);

    // if the title remains the same, that means we have failed to access the data
    if (window.Text == TransactionTitle)
    {
        const string Message = "The employee's personal data can not be accessed. The employee's ID may be incorrect, or there is no record for the requested infotype. Or. you do not have the correct permission to access this infotype.";
        var context = new Dictionary<string, object>
        {
            { "EmployeeID", employeeID },
            { "InfoType", it },
            { "SubType", sty },
            { "AsListView", asListView },
            { "StatusMessage", window.FindByName<GuiStatusbar>("sbar").Text } // main reason is explained here
        };

        Logger.Instance.Error("Unable to access infotype: " + context.Prettify());
        throw new InvalidOperationException(Message + Environment.NewLine + context.Prettify())
            .BindContext(context);
    }

    // extract identity data from the record
    var result = new Person();
    var prefix = detail.FindByName<GuiComboBox>("Q0002-ANREX").Key;
    result.Gender = new Dictionary<string, Gender?>
        {
            { "M.", Gender.Male },
            { "Mme", Gender.Female },
        }
        .FirstOrDefault(x => x.Key == prefix).Value;
    result.Name = Regex.Replace(detail.FindByName<GuiTextField>("P0001-ENAME").Text, @"^(M\.|Mme) ", "");
```

# HQ.CSRH.SimpleDocument.Services.Sap
> This namespace contains a few examples on how to encapsulate Xy.SAP.

Same code as above to retrieve employee identity:
```
    public Person GetIdentity(int employeeID)
    {
        Logger.Instance.LogCurrentMethod(LogLevel.Info);

        var detail = GetHrDetail(employeeID, "Identité (0002)");

        var result = new Person();
        using (result.Changed.Log(Logger.Instance.Info))
        {
            var prefix = detail.FindByName<GuiComboBox>("Q0002-ANREX").Key;
            result.Gender = new Dictionary<string, Gender?>
                {
                    { "M.", Gender.Male },
                    { "Mme", Gender.Female },
                }
                .FirstOrDefault(x => x.Key == prefix).Value;
            result.Name = Regex.Replace(detail.FindByName<GuiTextField>("P0001-ENAME").Text, @"^(M\.|Mme) ", "");
        }

        return result;
    }
```

Another example but with `GuiTableControl`:
```csharp
    public static IEnumerable<HrEntry> GetHrEntries(int employeeID, string it, string sty = null)
    {
        Logger.Instance.LogCurrentMethod(LogLevel.Debug);

        var window = QueryPA20(employeeID, it, sty, true);
        var usr = window.FindByName<GuiUserArea>("usr");
        var tableName = usr.Children.OfType<GuiTableControl>().Single().Name;

        // Scrolling, or navigating to another page and navigate back
        // all these resets the reference
        Func<GuiTableControl> getTable = () => window.FindByName<GuiTableControl>(tableName);

        // todo: optimize this to scroll only when item go out of page
        var rows = getTable().VerticalScrollbar.Maximum;
        for (int i = 0; i <= rows; i++)
        {
            /* zero-indexed position (screen value = this+1) */
            getTable().VerticalScrollbar.Position = i;

            yield return new HrEntry(window, (GuiTableRow)getTable().Rows.Item(0));
        }
    }

    public Dictionary<EmailType, string> GetEmailAddresses(int employeeID)
    {
        Logger.Instance.LogCurrentMethod(LogLevel.Info);

        var entries = GetHrEntries(employeeID, "Communication  (0105)");

        /*  Index Type          Title            Tooltip        
            0     GuiTextField  Début            Date de début  
            1     GuiTextField  Fin              Date de fin    
            2     GuiCTextField Mode communicat. Mode communicat.
            3     GuiTextField  Désignation      Désignation    
            4     GuiTextField  ID du système    ID/N° long      
            5     GuiTextField  CB               Code de blocage */
        const int DescriptionIndex = 3, ValueIndex = 4;
        var emailTypes = new Dictionary<string, EmailType>
        {
            { "Courrier électronique HQ", EmailType.Work },
            { "Adresse courriel personnelle", EmailType.Personal },
        };

        var result = new ReactiveList<KeyValuePair<EmailType, string>>();
        using (result.ItemsAdded.Subscribe(x => Logger.Instance.Info(x.Key + " => " + x.Value)))
        {
            foreach (var entry in entries)
            {
                var description = entry[DescriptionIndex];
                var type = default(EmailType);

                if (emailTypes.TryGetValue(description, out type) && !result.Any(x => x.Key == type))
                {
                    result.Add(KVPair.Create(type, entry[ValueIndex]));

                    // early exit if all found
                    if (result.Count == emailTypes.Count)
                        break;
                }
            }
        }

        return result.ToDictionary(x => x.Key, x => x.Value);
    }
```