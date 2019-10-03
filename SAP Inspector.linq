<Query Kind="Program">
  <Reference>&lt;RuntimeDirectory&gt;\Microsoft.VisualBasic.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\mscorlib.dll</Reference>
  <Reference>&lt;RuntimeDirectory&gt;\System.dll</Reference>
  <GACReference>CustomMarshalers, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a</GACReference>
  <GACReference>System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089</GACReference>
  <NuGetReference>SAP.GUI.Scripting.net</NuGetReference>
  <Namespace>Microsoft.VisualBasic</Namespace>
  <Namespace>sapfewse</Namespace>
  <Namespace>saprotwr.net</Namespace>
  <Namespace>System.Runtime.InteropServices</Namespace>
  <Namespace>System.Windows.Forms</Namespace>
</Query>

GuiSession session = SAP.GetSession();
Stack<string> history = new Stack<string>();

void Main()
{
	Util.AutoScrollResults = false;

	var id = Util.ReadLine("Enter an control id to inspect", session.ActiveWindow.Id + "/usr");
	InspectComponent(string.IsNullOrWhiteSpace(id)? session.ActiveWindow.Id : id);
}

void InspectComponent(string id)
{
	var component = session.FindById(id);
	
	// push id in history if we are not refreshing
	if (history.FirstOrDefault() != id)
		history.Push(id);
	
	Util.ClearResults();
	
	// Window Title
	for (GuiComponent c = component; c != null; c = c.Parent as GuiComponent)
	{
		if (c is GuiFrameWindow)
		{
			Util.WithStyle((c as GuiFrameWindow).Text, "font-weight: bold;").Dump();
			break;
		}
	}
	
	// Navigation bar
	Util.HorizontalRun(true,
		BuildInspectLinq(history.Skip(1).FirstOrDefault(), "Back"),
		BuildInspectLinq(id, "Refresh"),
		BuildInspectLinq((component.Parent as GuiComponent).Id, "Parent"),
		Util.OnDemand("History", () => history.Take(5).Select(x => BuildInspectLinq(x)))
		).Dump("Navigation");
	
	if (component is GuiTableControl)
		InspectTable((GuiTableControl)component);
	else if (component.ContainerType)
	{
		var children = ((component as dynamic).Children as GuiComponentCollection).Cast<GuiComponent>();
		if (children.All(x => x is GuiLabel))
			InspectLabelMap(children.Cast<GuiLabel>());
		
		children.Select(x => new
		{
			x.Type, x.Name, 
			Selector = new Hyperlinq(() => Clipboard.SetText(string.Format(".FindByName<{0}>(\"{1}\")", x.Type, x.Name)), "Copy"),
			Highlight = BuildHighlightLinq(x),
			Content = GetContent(x), 
			Id = BuildInspectLinq(x.Id),
		}).Dump("Children");
	}
	
	Util.OnDemand("Click to expand", () => component.SmartWrap()).Dump("Properties : " + id);
	ExposeMethods(component).OnDemand("Click to expand").Dump("ExposeMethods");
	BuildHighlightLinq(component).Dump("Highlight");
}

void InspectTable(GuiTableControl table)
{
	Util.VerticalRun(
		table.Columns.Cast<GuiTableColumn>()
			.Select((x,i) => new
			{
				Index = i,
				x.Item(0).Type,
				x.Title,
				x.Tooltip
			}),
		table.Columns.Cast<GuiTableColumn>()
			.Select((x,i) => new
			{
				Index = i,
				x.Item(0).Type,
				x.Title,
				x.Tooltip
			})
			.ToPaddedTable()
			.AsCopyLinq()
		)
		.Dump("Table Headers");
	
	Util.OnDemand("Click to expand", () => {
		var datatable = new DataTable();
		foreach(var column in table.Columns.Cast<GuiTableColumn>())
			datatable.Columns.Add(column.Title);
		
		foreach(var row in table.AsEnumerable())
		{
			var datarow = datatable.NewRow();
			for (int i = 0; i < datatable.Columns.Count; i++)
				datarow[i] = row.GetCell<dynamic>(i).Text;
			
			datatable.Rows.Add(datarow);
		}
		
		return datatable;
	}).Dump("Table Content");
}
void InspectLabelMap(IEnumerable<GuiLabel> labels)
{
	var contents = labels
		.Select(x => new 
		{
			Row = x.CharTop, 
			Column = x.CharLeft, 
			x.Text
		})
		.OrderBy(x => x.Row)
		.ThenBy(x => x.Column)
		.ToList();
	
	var buffer = Enumerable.Repeat(new string(' ', contents.Max(x => x.Column) + 1), contents.Max(x => x.Row) + 1).ToList();
	foreach (var content in contents)
		buffer[content.Row] = buffer[content.Row].Insert(content.Column, content.Text);
	
	for (int i = 0; i < buffer.Count; i++)
		buffer[i] = buffer[i].TrimEnd();
	
	Util.WithStyle(buffer, "font-family:consolas").Dump("Label Map");
}


Hyperlinq BuildHighlightLinq(GuiComponent target)
{
	bool toggle = false;
	if (target is GuiVComponent)
		return new Hyperlinq(() => (target as GuiVComponent).Visualize(toggle = !toggle), "Highlight");
	
	return null;
}
Hyperlinq BuildInspectLinq(string id, string description = null)
{
	if (id == null) return null;
	
	return new Hyperlinq(() => InspectComponent(id), description == null ? id : description);
}

object GetContent(GuiComponent component)
{
	if (component is GuiLabel) return (component as GuiLabel).Text;
	if (component is GuiTextField) return (component as GuiTextField).Text;
	if (component is GuiCTextField) return (component as GuiCTextField).Text;
	if (component is GuiTab) return (component as GuiTab).Text;
	if (component is GuiRadioButton) return (component as GuiRadioButton).Text;
	if (component is GuiButton) return (component as GuiButton).Text + "\n" + (component as GuiButton).Tooltip;
	if (component is GuiComboBox) return string.Format("KeyValuePair[{0}, {1}]", (component as GuiComboBox).Key, (component as GuiComboBox).Value);
	if (component is GuiBox) return (component as GuiBox).Text;
	if (component is GuiStatusbar) return (component as GuiStatusbar).Text;
	
	if (component is GuiSimpleContainer)
	{
		var firstChild = (component as GuiSimpleContainer).Children.Item(0);
		return (firstChild is GuiBox)? (firstChild as GuiBox).Text : null;
	}
	
	return null;
}

public static class Extensions
{
	public static string ToPaddedTable<T>(this IEnumerable<T> source)
	{
		var properties = typeof(T).GetProperties()
			.Select((x, i) => new
			{
				x.Name,
				Getter = x.GetGetMethod(),
				Index = i
			});
		var rows = source.Select(x => properties.Select(p => 
				p.Getter.Invoke(x, null).ToString()).ToArray()).ToList();
		var columnWidths = properties.ToDictionary(x => x.Name, 
			p => Math.Max(
				p.Name.Length, 
				rows.Max(x => x[p.Index].Length)));
		
		var format = string.Join(" ", columnWidths.Select((x, i) => string.Format("{{{0},-{1}}}", i, x.Value)));
		
		
		return string.Format(format, properties.Select(x => x.Name).ToArray()) + "\n" +
			string.Join("\n", rows.Select(x => string.Format(format, x)));
	}
	
	public static Hyperlinq AsCopyLinq(this string value, string description = "Copy to clipboard")
	{
		return new Hyperlinq(() => Clipboard.SetText(value), description);
	}
}

public class SAP
{
	public static GuiSession GetSession()
	{
		var wrapper = new CSapROTWrapper();
		var rot = wrapper.GetROTEntry("SAPGUI");
		var engine = rot.GetType().InvokeMember("GetSCriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, rot, null);
		var app = engine as GuiApplication;
		var connection = app.Children.ElementAt(0) as GuiConnection;
		return connection.Children.ElementAt(0) as GuiSession;
	}
}

public static IEnumerable ExposeMethods(object obj)
{
	var typeFormatter = new Dictionary<Type, string>
	{
		{ typeof(string), "string" },
		{ typeof(int), "int" },
		{ typeof(void), "void" },
	};
	Func<Type, string> getFamiliarTypeName = t => typeFormatter.ContainsKey(t) ? typeFormatter[t] : t.Name;
	
	/* will block lazy-execution
	return DispatchUtility.GetType(obj, true).GetMethods()
		.Select(m => string.Format("{0} {1}({2})", 
			getFamiliarTypeName(m.ReturnType),
			m.Name,
			string.Join(", ", m.GetParameters().Select(p => getFamiliarTypeName(p.ParameterType) + " " + p.Name))
			).Dump());*/
	
	DispatchUtility.GetType(obj, true).Name.Dump();
	foreach (var method in DispatchUtility.GetType(obj, true).GetMethods())
	{
		yield return new
		{
			Signature = string.Format("{0} {1}({2})",
				getFamiliarTypeName(method.ReturnType),
				method.Name,
				string.Join(", ", //args
					method.GetParameters().Select(p => getFamiliarTypeName(p.ParameterType) + " " + p.Name))),
			Invoke = method.GetParameters().Any() ? null : new Lazy<object>(() => method.Invoke(obj, null))
		};
	}
}

// credit : Bill Menees @http://stackoverflow.com/a/14208030/561113
public static class DispatchUtility
{
    private const int S_OK = 0; //From WinError.h
    private const int LOCALE_SYSTEM_DEFAULT = 2 << 10; //From WinNT.h == 2048 == 0x800

    public static bool ImplementsIDispatch(object obj)
    {
        bool result = obj is IDispatchInfo;
        return result;
    }

    public static Type GetType(object obj, bool throwIfNotFound)
    {
        RequireReference(obj, "obj");
        Type result = GetType((IDispatchInfo)obj, throwIfNotFound);
        return result;
    }

    public static bool TryGetDispId(object obj, string name, out int dispId)
    {
        RequireReference(obj, "obj");
        bool result = TryGetDispId((IDispatchInfo)obj, name, out dispId);
        return result;
    }

    public static object Invoke(object obj, int dispId, object[] args)
    {
        string memberName = "[DispId=" + dispId + "]";
        object result = Invoke(obj, memberName, args);
        return result;
    }

    public static object Invoke(object obj, string memberName, object[] args)
    {
        RequireReference(obj, "obj");
        Type type = obj.GetType();
        object result = type.InvokeMember(memberName,
            BindingFlags.InvokeMethod | BindingFlags.GetProperty,
            null, obj, args, null);
        return result;
    }

    private static void RequireReference<T>(T value, string name) where T : class
    {
        if (value == null)
        {
            throw new ArgumentNullException(name);
        }
    }

    private static Type GetType(IDispatchInfo dispatch, bool throwIfNotFound)
    {
        RequireReference(dispatch, "dispatch");

        Type result = null;
        int typeInfoCount;
        int hr = dispatch.GetTypeInfoCount(out typeInfoCount);
        if (hr == S_OK && typeInfoCount > 0)
        {
            dispatch.GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, out result);
        }

        if (result == null && throwIfNotFound)
        {
            // If the GetTypeInfoCount called failed, throw an exception for that.
            Marshal.ThrowExceptionForHR(hr);

            // Otherwise, throw the same exception that Type.GetType would throw.
            throw new TypeLoadException();
        }

        return result;
    }

    private static bool TryGetDispId(IDispatchInfo dispatch, string name, out int dispId)
    {
        RequireReference(dispatch, "dispatch");
        RequireReference(name, "name");

        bool result = false;

        Guid iidNull = Guid.Empty;
        int hr = dispatch.GetDispId(ref iidNull, ref name, 1, LOCALE_SYSTEM_DEFAULT, out dispId);

        const int DISP_E_UNKNOWNNAME = unchecked((int)0x80020006); //From WinError.h
        const int DISPID_UNKNOWN = -1; //From OAIdl.idl
        if (hr == S_OK)
        {
            result = true;
        }
        else if (hr == DISP_E_UNKNOWNNAME && dispId == DISPID_UNKNOWN)
        {
            result = false;
        }
        else
        {
            Marshal.ThrowExceptionForHR(hr);
        }

        return result;
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00020400-0000-0000-C000-000000000046")]
    private interface IDispatchInfo
    {
        [PreserveSig]
        int GetTypeInfoCount(out int typeInfoCount);

        void GetTypeInfo(int typeInfoIndex, int lcid, [MarshalAs(UnmanagedType.CustomMarshaler,
            MarshalTypeRef = typeof(System.Runtime.InteropServices.CustomMarshalers.TypeToTypeInfoMarshaler))] out Type typeInfo);

        [PreserveSig]
        int GetDispId(ref Guid riid, ref string name, int nameCount, int lcid, out int dispId);

        // NOTE: The real IDispatch also has an Invoke method next, but we don't need it.
    }
}

#region SapExtension
public static class SapExtension
{
	#region GuiSession
	/// Start a new transaction and return the active window
	public static GuiFrameWindow BeginTransaction(this GuiSession session, string transactionID, string expectedTitle)
	{
		return session.BeginTransaction(transactionID, 
			x => x.Text == expectedTitle, 
			x => string.Format("Unable to open transaction : {0}\n\tExpected title : {1}\n\tActive window : {2}",
				transactionID, expectedTitle, x.Text));
	}
	public static GuiFrameWindow BeginTransaction(this GuiSession session, string transactionID, Predicate<GuiFrameWindow> assumption, Func<GuiFrameWindow, string> formatErrorMessage)
	{
		// force end current transaction to prevent any blocking
		session.EndTransaction();
		
		session.StartTransaction(transactionID);
		
		var window = session.ActiveWindow;
		if (!assumption(window))
			throw new Exception(formatErrorMessage(window));
		
		return window;
	}
	#endregion
	#region GuiFrameWindow
	public static TSapControl FindById<TSapControl>(this GuiFrameWindow window, string id)
	{
		return (TSapControl)window.FindById(id);
	}
    public static TSapControl FindByName<TSapControl>(this GuiFrameWindow window, string name)
	{
		return (TSapControl)window.FindByName(name, typeof(TSapControl).Name);
	}
	#endregion
    #region GuiTab
    public static TSapControl FindByName<TSapControl>(this GuiTab tab, string name)
	{
		return (TSapControl)tab.FindByName(name, typeof(TSapControl).Name);
	}
    #endregion
	#region GuiTableControl
	public static IEnumerable<GuiTableRow> AsEnumerable(this GuiTableControl table)
	{
		var container = table.Parent as dynamic;
		string name = table.Name, type = table.Type;
		int rowCount = table.VerticalScrollbar.Maximum;
		
		Func<GuiTableControl> getTable = () => container.FindByName(name, type) as GuiTableControl;
		
		for(int i = 0; i <= rowCount; i++)
		{
			getTable().VerticalScrollbar.Position = i;
			yield return getTable().Rows.Item(0) as GuiTableRow;
		}
	}
	#region GuiTableRow
	public static TSapControl GetCell<TSapControl>(this GuiTableRow row, int column)
	{
		return (TSapControl)row.Item(column);
	}
	#endregion
	#endregion
	
	private const string LibraryName = "sapfewse";
	public static T Wrap<T>(this object comObject)
	{
		return (T)Marshal.CreateWrapperOfType(comObject, typeof(T).Dump());
	}
    public static object SmartWrap(this object comObject)
	{
        if(comObject == null) return comObject;
        
        var typename = (comObject as GuiComponent).Type;
		return Marshal.CreateWrapperOfType(comObject, 
            typeof(GuiApplication).Assembly.GetType(LibraryName + "." + typename + "Class"));
	}
    public static void DumpSap(this object comObject, int level = 3)
    {
        comObject.SmartWrap().Dump();
    }
}
#endregion