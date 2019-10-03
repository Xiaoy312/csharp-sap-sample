#region namespace HQ.CSRH.SimpleDocument.Services.Sap
namespace HQ.CSRH.SimpleDocument.Services.Sap
{
    using ReactiveUI;
    using sapfewse;
    using saprotwr.net;
    using Splat;
    using HQ.CSRH.SimpleDocument.Models;
    using Xy.Logging;
    using Xy.Sap;

    public static class SapServices
    {
        public static IPA20Service PA20 { get { return Locator.Current.GetService<IPA20Service>(Contract); } }

        public static void Initialize()
        {
            Locator.CurrentMutable.RegisterLazySingleton(() => new PA20Service(), typeof(IPA20Service), RealContract);
            Locator.CurrentMutable.RegisterLazySingleton(() => new FakePA20Service(), typeof(IPA20Service), FakeContract);
        }

        public static TService GetService<TService, TRealService, TFakeService>()
            where TRealService : TService, new()
            where TFakeService : TService, new()
        {
            return !App.Config.TestMode ? (TService)new TRealService() : new TFakeService();
        }

        private const string RealContract = "Sap", FakeContract = "Faker";
        private static string Contract { get { return !App.Config.TestMode ? RealContract : FakeContract; } }
    }

    public interface IPA20Service
    {
        Person GetIdentity(int employeeID);
        Gender GetGender(int employeeID);
        Address GetAddress(int employeeID);
        Dictionary<EmailType, string> GetEmailAddresses(int employeeID);
    }
    public partial class PA20Service : IPA20Service
    {
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
        public Gender GetGender(int employeeID)
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
            }

            return result.Gender.Value;
        }
        public Address GetAddress(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Info);

            var detail = GetHrDetail(employeeID, "Adresses  (0006)");

            var result = new Address();
            using (result.Changed.Log(Logger.Instance.Info))
            {
                result.StreetNumberName = string.Join(" ", new string[]
                {
                    detail.FindByName<GuiTextField>("P0006-STRAS").Text.Trim(), //Address
                    detail.FindByName<GuiTextField>("P0006-LOCAT").Text.Trim(), //Additional Address
                }
                .Where(x => !string.IsNullOrWhiteSpace(x)));
                result.City = detail.FindByName<GuiTextField>("P0006-ORT01").Text.Trim();
                result.Province = detail.FindByName<GuiTextField>("T005U-BEZEI").Text.Trim();
                result.PostalCode = detail.FindByName<GuiTextField>("P0006-PSTLZ").Text.Trim().ToUpper();
            }

            return result;
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
    }
    public partial class PA20Service
    {
        public static IEnumerable<HrEntry> GetHrEntries(int employeeID, int it, string sty = null)
        {
            return GetHrEntries(employeeID, it.ToString("D4"), sty);
        }
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

        public static HrDetail GetHrDetail(int employeeID, int it, string sty = null)
        {
            return GetHrDetail(employeeID, it.ToString("D4"), sty);
        }
        public static HrDetail GetHrDetail(int employeeID, string it, string sty = null)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Debug);

            var window = QueryPA20(employeeID, it, sty, false);

            return new HrDetail(window);
        }

        /// <summary>Performs a PA20 query</summary>
        /// <param name="it">Infotype ID</param>
        /// <param name="sty">Subtype ID, optional</param>
        /// <param name="asListView">True for list view. False for detail view. The default is to use list view</param>
        public static GuiFrameWindow QueryPA20(int employeeID, string it, string sty = null, bool asListView = false)
        {
            const string TransactionTitle = "Afficher données de base personnel";

            var session = SapHelper.GetActiveSession();
            var window = session.BeginTransaction("PA20", TransactionTitle);

            Logger.Instance.Debug("Accessing personal data: emp={0}, it={1}, sty={2}, mode={3}", employeeID, it, sty ?? "null", asListView ? "list" : "detail");
            window.FindByName<GuiCTextField>("RP50G-PERNR").Text = employeeID.ToString();
            window.FindByName<GuiCTextField>("RP50G-CHOIC").Text = it;
            window.FindByName<GuiCTextField>("RP50G-SUBTY").Text = sty;
            window.SendVKey(asListView ? 20 : 7); // Ctrl+F8 vs F7

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
                    { "StatusMessage", window.FindByName<GuiStatusbar>("sbar").Text }
                };

                Logger.Instance.Error("Unable to access infotype: " + context.Prettify());
                throw new InvalidOperationException(Message + Environment.NewLine + context.Prettify())
                    .BindContext(context);
            }

            return window;
        }

        #region nested types

        public class HrEntry
        {
            protected GuiFrameWindow window;
            protected GuiTableRow row;

            public HrEntry(GuiFrameWindow window, GuiTableRow row)
            {
                this.window = window;
                this.row = row;
            }

            public string this[int column, bool isCText = false]
            {
                get
                {
                    return isCText ? row.GetCText(column) : row.GetText(column);
                }
            }

            public HrDetail InspectDetail()
            {
                row.Selected = true;
                window.SendVKey(2); // F2

                return new HrDetail(window);
            }
        }
        public class HrDetail
        {
            protected GuiFrameWindow window;

            public HrDetail(GuiFrameWindow window)
            {
                this.window = window;
            }

            public TSapControl FindByName<TSapControl>(string name)
            {
                return window.FindByName<TSapControl>(name);
            }

            public void Dispose()
            {
                window.SendVKey(3); // F3
            }
        }

        #endregion

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }
    public class FakePA20Service : IPA20Service
    {
        public Person GetIdentity(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = new Person();
            using (result.Changed.Log(Logger.Instance.Warn))
            {
                result.Gender = Faker.EnumFaker.SelectFrom<Gender>();
                result.Name = result.Gender == Gender.Male ? Faker.NameFaker.MaleName() : Faker.NameFaker.FemaleName();
            }

            return result;
        }
        public Gender GetGender(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = new Person();
            using (result.Changed.Log(Logger.Instance.Warn))
            {
                result.Gender = Faker.EnumFaker.SelectFrom<Gender>();
            }

            return result.Gender.Value;
        }
        public Address GetAddress(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = new Address();
            using (result.Changed.Log(Logger.Instance.Warn))
            {
                result.StreetNumberName = Faker.LocationFaker.Street();
                result.City = Faker.LocationFaker.City();
                result.Province = Faker.LocationFaker.Country();
                result.PostalCode = Faker.StringFaker.Randomize("?#? #?#").ToUpper();
            }

            return result;
        }
        public Dictionary<EmailType, string> GetEmailAddresses(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = new ReactiveList<KeyValuePair<EmailType, string>>();
            using (result.ItemsAdded.Subscribe(x => Logger.Instance.Warn(x.Key + " => " + x.Value)))
            {
                result.Add(KVPair.Create(EmailType.Work, Faker.StringFaker.Randomize("????????.###@fake.email.com")));
                result.Add(KVPair.Create(EmailType.Personal, Faker.StringFaker.Randomize("????????.###@fake.email.biz")));
            }

            return result.ToDictionary(x => x.Key, x => x.Value);
        }

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }

    public interface ICustomSapService
    {
        KeyValuePair<EmploymentType, DateTime>? GetModificationEvent(int employeeID);
        bool? HasHealthInsurance(int employeeID);
        bool? HasDentalInsurance(int employeeID);
        DisabilityCoverage? GetDisabilityCoverage(int employeeID);
    }
    public class CustomSapService : ICustomSapService
    {
        public KeyValuePair<EmploymentType, DateTime>? GetModificationEvent(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Info);

            var entries = PA20Service.GetHrEntries(employeeID, "Mesures  (0000)");

            /*  Index Type          Title             Tooltip            
                0     GuiTextField  Début             Date de début      
                1     GuiTextField  Fin               Date de fin        
                2     GuiCTextField Mes.              Catégorie de mesure
                3     GuiTextField  Dés. cat. mesure  Dés. cat. mesure    
                4     GuiCTextField MotMe             Motif mesure        
                5     GuiTextField  Dés. motif mesure Dés. motif mesure
                6     GuiCTextField Client            Statut propre client
                7     GuiCTextField Activité          Statut d'activité  
                8     GuiCTextField Paiement          Statut paiement part */
            var actionToEmploymentType = new Dictionary<string, EmploymentType>
            {
                { "A5", EmploymentType.Temporary }, // Cess. réemb. sans bris /CCE
                { "A6", EmploymentType.Permanent }, // Mouvement de personnel /CCE
            };
            Logger.Instance.Info("Searching for a record with action of [A5 or A6] and reason of [90]");
            var match = entries
                .Select(x => new
                {
                    StartDate = x[0],
                    EndDate = x[1],
                    Action = x[2, true],
                    ActionText = x[3],
                    Reason = x[4, true],
                    ReasonText = x[5],
                })
                .FirstOrDefault(x => actionToEmploymentType.Keys.Contains(x.Action));

            if (match == null)
            {
                Logger.Instance.Warn("-> There is no record matching the criteria.");
                Logger.Instance.Info("=> null");
                return null;
            }

            Logger.Instance.Info("-> " + match);
            var result = KVPair.Create(
                actionToEmploymentType[match.Action],
                DateTime.ParseExact(match.StartDate, "yyyy/MM/dd", CultureInfo.InvariantCulture));

            Logger.Instance.Info("=> " + result);
            return result;
        }
        public bool? HasHealthInsurance(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Info);

            try
            {
                var detail = PA20Service.GetHrDetail(employeeID, "Régimes de santé  (0167)", "MEDI");

                // retirement health plan is not considered valid
                if (detail.FindByName<GuiTextField>("Q0167-BPLAN").Text == "MALH")
                {
                    Logger.Instance.Warn("-> The employee doesn't have a valid health plan: [{0}]{1}",
                        detail.FindByName<GuiTextField>("Q0167-BPLAN").Text,
                        detail.FindByName<GuiTextField>("T5UCA-LTEXT").Text);

                    Logger.Instance.Info("=> null");
                    return null;
                }

                // an expired plan will have an end date other than 9999-12-31
                if (!detail.FindByName<GuiCTextField>("P0167-ENDDA").Text.StartsWith("9999"))
                {
                    Logger.Instance.Info("-> Employee's last health plan has expired on " +
                        detail.FindByName<GuiCTextField>("P0167-ENDDA").Text);

                    Logger.Instance.Warn("=> false");
                    return false;
                }

                var options = new Dictionary<string, bool>
                {
                    { "EXEM", false }, // Exempté de participation
                    { "BASE", true },  // Option de base
                    { "MOD1", true },  // Module de base
                    { "MOD2", true },  // Module bonifié
                    { "MOD3", true },  // Module enrichi
                    { "NONP", false }, // Non participant
                };
                var option = detail.FindByName<GuiCTextField>("P0167-BOPTI").Text;
                var optionText = detail.FindByName<GuiTextField>("T5UCE-LTEXT").Text;

                bool result;
                if (!options.TryGetValue(option, out result))
                {
                    const string Message = "The application doesn't know to how to process this information.";
                    var context = new Dictionary<string, string>
                    {
                        { "Option", option },
                        { "Description", optionText },
                    };

                    Logger.Instance.Error("Invalid health plan option:" + context.Prettify());
                    throw new ArgumentOutOfRangeException(Message + Environment.NewLine + context.Prettify())
                        .BindContext(context);
                }

                Logger.Instance.Info("-> [{0}]{1}", option, optionText);
                Logger.Instance.Info("=> " + result);
                return result;
            }
            catch (InvalidOperationException e)
            {
                const string NoRecordStatusMessage = "Aucune donnée existe pour Régimes de santé  (0167) (dans période sélectionnée)";
                if (e.Data["StatusMessage"] as string == NoRecordStatusMessage)
                {
                    Logger.Instance.Warn("-> " + NoRecordStatusMessage);
                    Logger.Instance.Info("=> false");
                    return false;
                }

                throw;
            }
        }
        public bool? HasDentalInsurance(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Info);

            try
            {
                var detail = PA20Service.GetHrDetail(employeeID, "Régimes de santé  (0167)", "DENT");

                // only '[DEN0]Dentaire Bureau/Métier/Techn.' is taken into consideration
                if (detail.FindByName<GuiCTextField>("P0167-BPLAN").Text != "DEN0")
                {
                    // TODO: ask for the reason why this is invalid
                    const string Message = "The dental plan type is invalid.";
                    var context = new Dictionary<string, string>
                    {
                        { "Type", detail.FindByName<GuiTextField>("Q0167-BPLAN").Text },
                        { "Description", detail.FindByName<GuiTextField>("T5UCA-LTEXT").Text },
                    };

                    Logger.Instance.Error("The dental plan type is invalid: " + context.Prettify());
                    throw new ArgumentOutOfRangeException(Message + Environment.NewLine + context.Prettify())
                        .BindContext(context);
                }

                var options = new Dictionary<string, bool>
                {
                    { "EXEM", false }, // Exempté de participation
                    { "BASE", true },  // Option de base
                    { "MOD1", true },  // Module de base
                    { "MOD2", true },  // Module bonifié
                    { "MOD3", true },  // Module enrichi
                    { "NONP", false }, // Non participant
                };
                var option = detail.FindByName<GuiCTextField>("P0167-BOPTI").Text;
                var optionText = detail.FindByName<GuiTextField>("T5UCE-LTEXT").Text;

                bool result;
                if (!options.TryGetValue(option, out result))
                {
                    const string Message = "The dental plan option could not be processed.";
                    var context = new Dictionary<string, string>
                    {
                        { "Option", option },
                        { "Description", optionText },
                    };

                    Logger.Instance.Error("Invalid dental plan option:" + context.Prettify());
                    throw new ArgumentOutOfRangeException(Message + Environment.NewLine + context.Prettify())
                        .BindContext(context);
                }

                Logger.Instance.Info("-> [{0}]{1}", option, optionText);
                Logger.Instance.Info("=> " + result);
                return result;
            }
            catch (Exception e)
            {
                const string NoRecordStatusMessage = "Aucune donnée existe pour Régimes de santé  (0167) (dans période sélectionnée)";
                if (e.Data["StatusMessage"] as string == NoRecordStatusMessage)
                {
                    Logger.Instance.Warn("-> " + NoRecordStatusMessage);
                    Logger.Instance.Info("=> false");
                    return false;
                }

                throw;
            }
        }
        public DisabilityCoverage? GetDisabilityCoverage(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Info);

            try
            {
                var detail = PA20Service.GetHrDetail(employeeID, "Contingents d'absences  (2006)", "15");

                var valueText = detail.FindByName<GuiTextField>("P2006-ANZHL").Text.Trim();

                Logger.Instance.Info("-> {0} < 99999", valueText);
                var result = double.Parse(valueText) < 99999 ? DisabilityCoverage.Weeks52 : DisabilityCoverage.Weeks26;

                Logger.Instance.Info("=> " + result);
                return result;
            }
            catch (Exception e)
            {
                const string NoRecordStatusMessage = "Aucune donnée existe pour Contingents d'absences  (2006) (dans période sélectionnée)";
                if (e.Data["StatusMessage"] as string == NoRecordStatusMessage)
                {
                    Logger.Instance.Warn("-> " + NoRecordStatusMessage);
                    Logger.Instance.Info("=> null");
                    return null;
                }

                throw;
            }

        }

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }
    public class FakeCustomSapService : ICustomSapService
    {
        public KeyValuePair<EmploymentType, DateTime>? GetModificationEvent(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = KVPair.Create(
                Faker.EnumFaker.SelectFrom<EmploymentType>(),
                Faker.DateTimeFaker.DateTime().Date);

            Logger.Instance.Warn("=> " + result);
            return result;
        }
        public bool? HasHealthInsurance(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = Faker.BooleanFaker.Boolean();

            Logger.Instance.Warn("=> " + result);
            return result;
        }
        public bool? HasDentalInsurance(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = Faker.BooleanFaker.Boolean();

            Logger.Instance.Warn("=> " + result);
            return result;
        }
        public DisabilityCoverage? GetDisabilityCoverage(int employeeID)
        {
            Logger.Instance.LogCurrentMethod(LogLevel.Warn);

            var result = Faker.EnumFaker.SelectFrom<DisabilityCoverage>();

            Logger.Instance.Warn("=> " + result);
            return result;
        }

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }


    public static class KVPair
    {
        public static KeyValuePair<TKey, TValue> Create<TKey, TValue>(TKey key, TValue value)
        {
            return new KeyValuePair<TKey, TValue>(key, value);
        }
    }
}
#endregion

#region namespace Xy.Sap
namespace Xy.Sap
{
    using System.Runtime.InteropServices;
    using System.Reflection;
    using sapfewse;
    using saprotwr.net;
    using Xy.Logging;
    using COMException = System.Runtime.InteropServices.COMException;

    public static class SapHelper
    {
        /// <summary>Return an already opened sap session</summary>
        public static GuiSession GetActiveSession()
        {
            var rot = new CSapROTWrapper().GetROTEntry("SAPGUI");
            if (rot == null)
            {
                const string Message = "The application can no connect to SAP. Make sure SAP is up and running, and then try again.";
                Logger.Instance.Fatal("The application can no connect to SAP");

                throw new InvalidComObjectException(Message);
            }

            var app = (GuiApplication)rot.GetType().InvokeMember("GetScriptingEngine", BindingFlags.InvokeMethod, null, rot, null);
            var connectedSession = app.Connections.Cast<GuiConnection>()
                .SelectMany(x => x.Children.Cast<GuiSession>())
                .Where(x => !string.IsNullOrEmpty(x.Info.User))
                .FirstOrDefault();

            if (connectedSession == null)
            {
                const string Message = "Could not find an opened SAP session. Make sure you are logged in, and then try again.";
                Logger.Instance.Fatal("Could not find an opened SAP session");

                throw new InvalidComObjectException(Message);
            }

            return connectedSession;
        }

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }
    public static class SapExtensions
    {
        #region GuiFrameWindow

        public static TSapControl FindByName<TSapControl>(this GuiFrameWindow window, string name)
        {
            try
            {
                return (TSapControl)window.FindByName(name, typeof(TSapControl).Name);
            }
            catch (COMException e)
            {
                const string Message = "The control could not be found by name and type.";
                var context = new Dictionary<string, object>
                {
                    { "Name", name },
                    { "Type", typeof(TSapControl).Name },
                }.Prettify();

                Logger.Instance.ErrorException("The control could not be found by name and type: " + context, e);
                throw new InvalidOperationException(Message + Environment.NewLine + context, e);
            }
        }

        #endregion
        #region GuiTableControl

        /// <summary>Note: Do not iterate through this ienumerable more than once</summary>
        public static IEnumerable<GuiTableRow> AsEnumerable(this GuiTableControl table)
        {
            var container = table.Parent as dynamic;
            string name = table.Name, type = table.Type;
            int rowCount = table.VerticalScrollbar.Maximum;

            Func<GuiTableControl> getTable = () => container.FindByName(name, type) as GuiTableControl;

            for (int i = 0; i <= rowCount; i++)
            {
                getTable().VerticalScrollbar.Position = i;
                yield return getTable().Rows.Item(0) as GuiTableRow;
            }
        }

        public static TSapControl GetCell<TSapControl>(this GuiTableRow row, int column)
        {
            return (TSapControl)row.Item(column);
        }
        public static string GetCText(this GuiTableRow row, int column)
        {
            return row.GetCell<GuiCTextField>(column).Text;
        }
        public static string GetText(this GuiTableRow row, int column)
        {
            return row.GetCell<GuiTextField>(column).Text;
        }
        #endregion

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }

    public static class CommonSapService
    {
        /// <summary>Start a new transaction</summary>
        /// <param name="transactionID">Transaction ID</param>
        /// <param name="expectedTitle">Title of the transaction. This is used for validation.</param>
        public static GuiFrameWindow BeginTransaction(this GuiSession session, string transactionID, string expectedTitle)
        {
            // force current transaction to end, preventing any blocking(eg: model dialog)
            session.EndTransaction();

            Logger.Instance.Debug("Starting transaction : " + transactionID);
            session.StartTransaction(transactionID);

            var window = session.ActiveWindow;
            if (window.Text != expectedTitle)
            {
                const string Message = "You do not have the correct permission to access this transaction. Or, such transaction does not exist.";
                var context = new Dictionary<string, object>
                {
                    { "TransactionID", transactionID },
                    { "Expected Title", expectedTitle },
                    { "WindowTitle", window.Text },
                    { "StatusMessage", window.FindByName<GuiStatusbar>("sbar").Text },
                };

                Logger.Instance.Error("Unable to access personal data");
                throw new InvalidOperationException(Message + Environment.NewLine + context.Prettify())
                    .BindContext(context);
            }

            return window;
        }

        /// <summary>Performs a PA20 query</summary>
        /// <param name="it">Infotype ID</param>
        /// <param name="sty">Subtype ID, optional</param>
        /// <param name="asListView">True for list view. False for detail view. The default is to use list view</param>
        [Obsolete]
        public static GuiFrameWindow QueryPA20(this GuiSession session, int employeeID, string it, string sty = null, bool asListView = false)
        {
            const string TransactionTitle = "Afficher données de base personnel";

            var window = session.BeginTransaction("PA20", TransactionTitle);

            Logger.Instance.Debug("Accessing personal data: emp={0}, it={1}, sty={2}, mode={3}", employeeID, it, sty ?? "null", asListView ? "list" : "detail");
            window.FindByName<GuiCTextField>("RP50G-PERNR").Text = employeeID.ToString();
            window.FindByName<GuiCTextField>("RP50G-CHOIC").Text = it;
            window.FindByName<GuiCTextField>("RP50G-SUBTY").Text = sty;
            window.SendVKey(asListView ? 20 : 7); // Ctrl+F8 vs F7

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
                    { "StatusMessage", window.FindByName<GuiStatusbar>("sbar").Text }
                };

                Logger.Instance.Error("Unable to access infotype." + context.Prettify());
                throw new InvalidOperationException(Message + Environment.NewLine + context.Prettify())
                    .BindContext(context);
            }

            return window;
        }

        private static readonly TypeNamedLogger Logger = new TypeNamedLogger();
    }
}
#endregion