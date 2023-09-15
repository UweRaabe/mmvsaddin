namespace MMVSAddIn {
	using System;
	using EnvDTE;
#if VS2005
    // using EnvDTE80;
    using Microsoft.VisualStudio.CommandBars;
#else
	using Microsoft.Office.Core;
#endif
	using System.Windows.Forms;
	using System.Diagnostics;
	using System.Drawing;
	using System.Resources;
	using System.Collections;
	using System.Collections.Specialized;
	using System.Runtime.InteropServices;
	using System.IO;
	using System.Reflection;
	using MMTools.Utils;

	delegate void MMCommandEvent(int Cmd);
	

	internal class MMCommands {
		internal const int RunMMPascal = 0;
		internal const int RunMMCSharp = 1;

		internal const int JumpToModelMaker = 10;
		internal const int InvokeModelSearch = 11;

		internal const int AddToModel = 12;
		internal const int AddFilesToModel = 13;
		internal const int RefreshInModel = 14;
		internal const int SynchronizeModel = 15;

		internal const int ConvertToModel = 16;
		internal const int ConvertProjectToModel = 17;
		internal const int CreateSequenceDiagram = 18;
		internal const int Properties = 20;
		}

	internal class MMAction  {

		internal MMAction(string AName, string ACaption, string AHint, string AImageName, string ABinding, bool ASeparator, int ACmd) : 
			this(AName, ACaption, AHint, AImageName, ABinding, ASeparator, ACmd, Color.Fuchsia) {}

		internal MMAction(string AName, string ACaption, string AHint, string AImageName, string ABinding, bool ASeparator, int ACmd, Color AMaskColor) {
			Name = AName;
			Caption = ACaption;
			Hint = AHint;
			if (Hint ==  "") Hint = Caption;
			Cmd = ACmd;
			Separator = ASeparator; // Start a new group with this item = add separator before this item.
			Binding = ABinding;
			MaskColor = AMaskColor;
			ReadImage(AImageName);
		}

		internal string Binding;
		internal string Caption;
		internal int Cmd;
		internal string Hint;
		internal Bitmap Image = null;
		internal bool IsNew;
		internal Color MaskColor;
		internal string Name;
		internal bool Separator;
		internal string Tag;

		internal void DisposeImage() {
			if (Image != null) {
				Image.Dispose();
				Image = null;
			}
		}

		private void ReadImage(string imageName) {
			if (imageName ==  "") return;
			Stream imgStream = null;
			Assembly a = Assembly.GetExecutingAssembly();
			
			//			string [] resNames = a.GetManifestResourceNames();
			//			foreach (string s in resNames)
			//			  Debug.WriteLine(s);
			
			// attach to stream to the resource in the manifest
#if VS2012
			string nameSpace = "MMVS2012AddIn.";
#else
#if VS2010
			string nameSpace = "MMVS2010AddIn.";
#else
#if VS2008 
			string nameSpace = "MMVS2008AddIn.";
#else
#if VS2005
			string nameSpace = "MMVS2005AddIn.";
#else
			string nameSpace = "MMVS2003AddIn.";
#endif
#endif
#endif
#endif
			imgStream = a.GetManifestResourceStream(nameSpace + imageName + ".bmp");
			// if the resource is not found, imgStream == null, no exceptions raised
			if( imgStream != null ) {                    
				Image = Bitmap.FromStream(imgStream) as Bitmap;
				imgStream.Close();
				imgStream = null;
				// Image.MakeTransparent(MaskColor) results in "OutOfMemory"
			}            
		}
	}
	
	internal class MMCommandHandler {
		internal MMCommandHandler(_DTE DTEApplication, AddIn instance, object Client, bool doUseToolsMenu) {
			useToolsMenu = doUseToolsMenu;
			actions = new ArrayList();
			unsupported = new ArrayList();		  
			wiredButtons = new ArrayList();
			log = new StringCollection();
			application = DTEApplication;
			addIn = instance;
			InitCmdPrefix();
			DefineActions();
			UpdateNamedCommands();
			InsertCommandBars();
			HookEvents();
		}

		private bool useToolsMenu = false;
		private const string loadCmdName = "Load_MM";
		private const string MainMenuCaption = "ModelMaker";
		private ArrayList actions;
		private AddIn addIn;
		private _DTE application;
		private StringCollection log;
		private SolutionEvents solutionEvents;
		private Timer timer;
		private ArrayList unsupported;
		private ArrayList wiredButtons;
		private string namedCmdPrefix; // auto updated from ProgID
		private CommandBarPopup mmMenu;

		internal void ClearBindings(string Value) {
			Commands commands = application.Commands;
			foreach (Command cmd in commands) {
				string Name = cmd.Name;
				if (ConvertNamedCmd(ref Name)) {
					cmd.Bindings = new Object[0]; // VS2005 does not like "" , VS2003 accepts the empty object array.
				}
			}
		}

		internal void Disconnect() {
			FreeTimer();
			UnhookEvents();
			UnwireButtons();
			
			if (mmMenu != null) mmMenu.Delete(false);
			
			application = null;
			addIn = null;
			if (actions != null) {
				DisposeImages(); // before clearing images!
				actions.Clear();
				actions = null;
			}
			if (log != null) {
				log.Clear();
				log = null;
			}
			if (unsupported != null) {
				unsupported.Clear();
				unsupported = null;
			}
		}

		internal bool Execute(string Cmd) {
			if (ConvertNamedCmd(ref Cmd)) {
				int Index = 0;
				return Cmd.Equals(loadCmdName) || ( (FindAction(Cmd, ref Index) ) && ExecuteCmd((actions[Index] as MMAction).Cmd) );
			}
			return false;
		}

		internal string GetActionsAsStr() {
			string Result = "";
			ResetActionsTags();
			Commands commands = application.Commands;
			
			int Index = 0;
			Array Bindings;
			foreach (Command cmd in commands) {
				Bindings = cmd.Bindings as Array;
				if ((Bindings == null) || (Bindings.Length == 0)) continue;
				string Name = cmd.Name;
				if (ConvertNamedCmd(ref Name)) {
					if (FindAction(Name, ref Index)) {
						string Tag = "";
						foreach (string S in Bindings) {
							Tag += (Tag == "") ? S : "," + S;
						}  
						(actions[Index] as MMAction).Tag = Tag;
					}
				}
			}
			foreach (MMAction Action in actions) {
				Result = Result + Action.Caption + "=" + Action.Binding + "=" + Action.Tag + "\n";
			}
		  
		  
			return Result;
		}

		internal void HookEvents() {
			EnvDTE.Events events = application.Events;
			solutionEvents = events.SolutionEvents;
			solutionEvents.AfterClosing += new _dispSolutionEvents_AfterClosingEventHandler(SolutionEvents_AfterClosing);
			solutionEvents.Opened += new _dispSolutionEvents_OpenedEventHandler(SolutionEvents_Opened);
		}

		internal vsCommandStatus QueryStatus(string Cmd) {
			vsCommandStatus status = vsCommandStatus.vsCommandStatusUnsupported;
			if (ConvertNamedCmd(ref Cmd)) {
				// TODO: VS2010: invisible does not work 
				status = vsCommandStatus.vsCommandStatusInvisible;
				if (!Cmd.Equals(loadCmdName)) {
					status = (vsCommandStatus)vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
				}
			}
			return status;
		}

		internal void SetCommandData(string Value) {
			// == Remove Named Commands
			unsupported.Clear();
			Commands commands = application.Commands;
			foreach (Command cmd in commands) {
				string Name = cmd.Name;
				if (ConvertNamedCmd(ref Name)) {
					unsupported.Add(cmd);
				}
			}
			DeleteUnsupportedCmds(this, null);
		}

		internal void UnhookEvents() {
			if (solutionEvents != null) {
				solutionEvents.AfterClosing -= new _dispSolutionEvents_AfterClosingEventHandler(SolutionEvents_AfterClosing);
				solutionEvents.Opened -= new _dispSolutionEvents_OpenedEventHandler(SolutionEvents_Opened);
				solutionEvents = null;
			}  
		}

		private Command AddNamedCommand(string Name, string BtnText, string Hint) {
			try {
				object []contextGUIDS = new object[] { };
				Commands commands = application.Commands;
				Command command = commands.AddNamedCommand(addIn, Name, BtnText, Hint, true, 0, ref contextGUIDS, (int)vsCommandStatus.vsCommandStatusSupported+(int)vsCommandStatus.vsCommandStatusEnabled);
				log.Add("New ModelMaker Command available for keyboard shortcut: " + BtnText);
				return command;
			}  
			catch(System.Exception e) {
				MessageBox.Show("Error adding Named command:\n" + e.Message, VersionResource.ThisVersionResource.FileDescription, MessageBoxButtons.OK, MessageBoxIcon.Error);
				return null;
			}
		}

		private void ButtonClick(CommandBarButton Ctrl, ref bool Cancel) {
			int Index = 0;
			Cancel = ! ( (FindAction(Ctrl.Tag, ref Index)) && ExecuteCmd((actions[Index] as MMAction).Cmd) );
		}

		private bool ContainsAction(string AName) {
			int Index = 0;
			return FindAction(AName, ref Index);
		}

		private bool ConvertNamedCmd(ref string Cmd) {
			if ((Cmd != null) && (Cmd != "") && Cmd.StartsWith(namedCmdPrefix)) {
				Cmd = Cmd.Remove(0, namedCmdPrefix.Length);
				return true;
			}
			return false;
		}

		private void DefineActions() {
			//actions.Add(new MMAction("RunModelMakerPascal", "Run ModelMaker Pascal Edition", "Run ModelMaker Pascal Edition", "run_mm", "", false, MMCommands.RunMMPascal, Color.White));
			actions.Add(new MMAction("RunModelMakerCS", "Run ModelMaker C# Edition", "Run ModelMaker C# Edition", "run_mm", "", false, MMCommands.RunMMCSharp));
			
			actions.Add(new MMAction("JumpToModelMaker", "Jump To ModelMaker", "", "JumpMM", "Text Editor::Ctrl+Shift+M", true, MMCommands.JumpToModelMaker));
			actions.Add(new MMAction("InvokeModelSearch", "Model Search", "", "searchbar", "", false, MMCommands.InvokeModelSearch));

			actions.Add(new MMAction("AddCurrentToModel", "Add To Model", "Add To Model", "add_doc", "", true, MMCommands.AddToModel));
			actions.Add(new MMAction("AddFilesToModel", "Add files to Model...", "Add files to Model...", "add_doc_stack", "", false, MMCommands.AddFilesToModel));
			
			actions.Add(new MMAction("ConvertCurrentToModel", "Convert to Model", "Convert to Model", "convert_doc", "", false, MMCommands.ConvertToModel));
			actions.Add(new MMAction("ConvertProjectToModel", "Convert Project to Model", "Convert Project to Model", "convert_doc_stack", "", false, MMCommands.ConvertProjectToModel));

			actions.Add(new MMAction("RefreshInModel", "Refresh in Model", "Refresh in Model", "refresh", "", true, MMCommands.RefreshInModel));
			actions.Add(new MMAction("MMSynchronizeModel", "Synchronize Model", "Synchronize modified files in Model", "Refresh_all", "", false, MMCommands.SynchronizeModel));

			actions.Add(new MMAction("CreateSequenceDiagram", "Create Sequence diagram", "Create Sequence diagram from call stack", "seq_diagram", "seq_diagram", true, MMCommands.CreateSequenceDiagram));

			actions.Add(new MMAction("ModelMakerProperties", "ModelMaker Integration Properties", "ModelMaker Integration Properties", "options", "", true, MMCommands.Properties));
		}

		private void DeleteUnsupportedCmds(object sender, EventArgs e) {
			FreeTimer();
			for (int I = 0; I < unsupported.Count; I++)
				(unsupported[I] as Command).Delete();
			unsupported.Clear();
		}

		private void DisposeImages() {
			MMAction Action;
			for (int I = 0; I < actions.Count; I++) {
				Action = actions[I] as MMAction;
				Action.DisposeImage();
			}
		}

		private bool ExecuteCmd(int Cmd) {
			if (OnCommand != null) {
				OnCommand(Cmd);
				return true;
			}
			else
				return false;
		}

		private bool FindAction(string AName, ref int Index) {
			for (int I = 0; I < actions.Count; I++)
				if (String.Compare((actions[I] as MMAction).Name, AName, true) == 0) {
					Index = I;
					return true;
				}
			return false;
		}

		private void FreeTimer() {
			if (timer != null) {
				timer.Enabled = false;
				timer.Dispose();
				timer = null;
			}
		}

		private void InitCmdPrefix() {
#if VS2012
			namedCmdPrefix = "MMVS2012AddIn.MMVSClient."; // VS2012 uses the add-in class name
#else
#if VS2010
			namedCmdPrefix = "MMVS2010AddIn.MMVSClient."; // VS2010 uses the add-in class name
#else
#if VS2008
			namedCmdPrefix = "MMVS2008AddIn.MMVSClient."; // VS2008 uses the add-in class name
#else
#if VS2005
			namedCmdPrefix = "MMVS2005AddIn.MMVSClient."; // VS2005 uses the add-in class name
#else
			namedCmdPrefix = "MMVS2003AddIn.MMVSClient."; // VS2003 uses the add-in ProgId
#endif
#endif
#endif
#endif

		}

		private void InsertCommandBars() {
			//			foreach (CommandBar _bar in FApplication.CommandBars) {
			//				Debug.WriteLine("Commandbar: " + _bar.Name);
			//			}

			// determine insertion position in mainMenu
			CommandBars cmdbars = application.CommandBars as CommandBars;
			CommandBar mainMenu = cmdbars["MenuBar"];
			CommandBar toolsMenu = cmdbars["Tools"];

			if (useToolsMenu)
			{
				// default insertionPos = 1
				mmMenu = OfficeHelper.CreateCommandBarPopup(toolsMenu, MainMenuCaption, null, 1, false);
			}
			else
			{
				// The tools menu Index returns an exception
				// MessageBox.Show(toolsMenu.Index.ToString());
				int insertionPos = mainMenu.Controls.Count;
#if VS2010
				// In VS2010 the Id returns 0 for all controls -> use InstanceId instead.
				// InstanceId cannot be used on VS2003
				// For VS 2005 nothing works.
				int index = 0;
				int toolsId = toolsMenu.InstanceId;
				foreach (CommandBarControl control in mainMenu.Controls)
				{
					index++;
					CommandBarPopup menuBar = control as CommandBarPopup;
					if (menuBar != null)
					{
						int menuId = menuBar.InstanceId;
						if (toolsId == menuId)
						{
							insertionPos = index + 1; // VS2010 uses a 0-based index, whereas older versions seem to use 1-based (see below)
							break;
						}
					}
				}
#else
				// This does not work for VS 2005. Tried match on GetHashCode(), Id, InstanceId (Name is not supported for CommandBarControl)
				int toolsId = toolsMenu.Id;
				for (int I = 1 ; I <= mainMenu.Controls.Count ; I ++ ) {
					if (mainMenu.Controls[I].Id == toolsId) {
						insertionPos = I + 1; // + 1 is required for VS2003
						break;
					}
				}
#endif
				mmMenu = OfficeHelper.CreateCommandBarPopup(mainMenu, MainMenuCaption, null, insertionPos, false);
			}
			

			CommandBarButton button = null;
			MMAction Action;

			for (int I = 0; I < actions.Count; I++)
			{
				Action = actions[I] as MMAction;
				button = OfficeHelper.CreateCommandButtonIconAndCaption(mmMenu, Action.Caption, Action.Image, mmMenu.Controls.Count + 1, Action.Separator, Action.MaskColor);
				if (button != null)
				{
					button.Click += new _CommandBarButtonEvents_ClickEventHandler(ButtonClick);
					wiredButtons.Add(button);
					button.TooltipText = Action.Hint;
					button.Tag = Action.Name;
				}
			}  
		}

		private int LimitToRange(int Value, int LowBound, int HighBound) {
			if (Value < LowBound)
				return LowBound;
			else
				if (Value > HighBound) 
				return HighBound;
			else
				return Value;
		}

		private void MarkActionsAsNew() {
			foreach (object Action in actions)
				(Action as MMAction).IsNew = true;
		}

		private void ResetActionsTags() {
			foreach (object Action in actions)
				(Action as MMAction).Tag = "";
		}

		private void RethinkCommandBars() {
#if VS2010
			// There's no need for this anymore in VS2010
#else
			UnwireButtons();
			WireButtons();
#endif
		}

		private void SolutionEvents_AfterClosing() {
			RethinkCommandBars();
		}

		private void SolutionEvents_Opened() {
			RethinkCommandBars();
		}

		private void UnwireButtons() {
			if (wiredButtons == null) return;
			foreach (CommandBarButton button in wiredButtons) {
				button.Click -= new _CommandBarButtonEvents_ClickEventHandler(ButtonClick);
			}
		}

		private void UpdateNamedCommands() {
			Commands commands = application.Commands;
			int Index = 0;
			MarkActionsAsNew();
			bool addLoadMM = true;
			// remove unsupported commands
			unsupported.Clear();
			foreach (Command cmd in commands) {
				string Name = cmd.Name;
				if (ConvertNamedCmd(ref Name)) {
					if (FindAction(Name, ref Index)) 
						(actions[Index] as MMAction).IsNew = false;
					else
						if (Name.Equals(loadCmdName))
							addLoadMM = false;
						else // to remove all commands put a ; after this else
							unsupported.Add(cmd);
				}
			}
			// Immediately deleting the unsupported commands causes an AV when the command that 
			// invoked the Add-in to load is deleted- because that object is still on the stack.
			// => Decoupled deleting with a timer.
			if (unsupported.Count > 0) {
				timer = new Timer();
				timer.Interval = 250;
				timer.Tick += new EventHandler(DeleteUnsupportedCmds);
				timer.Enabled = true;
			}
			  
			MMAction Action = null;  
			Command command = null;
			bool canAddBinding = true;
			bool hasNewCmd = false;
			for (int I = 0; I < actions.Count; I++) {
				Action = actions[I] as MMAction;
				if (Action.IsNew) {
					command = AddNamedCommand(Action.Name, Action.Caption, Action.Hint);
					hasNewCmd = true;
					try {
						if (canAddBinding && (Action.Binding != "") ) 
							command.Bindings = Action.Binding;
					}
					catch /*(System.Exception e)*/ {
						canAddBinding = false;
					}
				}
			}  
			if (!canAddBinding) {
				log.Add("\nError assigning default keybinding(s).\nCould you be using the default keyboard scheme?\nTo assign short cuts manually, go to Tools|Options|Environment|Keyboard.\n");
			}
			if (hasNewCmd) {
				log.Add("To (re)define default ModelMaker keyboard short cuts manually, go to Tools|Options|Environment|Keyboard.\n");
			}
			
			// add LoadModelMaker named command to toolbar
			if (addLoadMM) {
				command = AddNamedCommand(loadCmdName, "Load ModelMaker Integration", "Load ModelMaker Integration");
				CommandBars cmdbars = application.CommandBars as CommandBars;
				CommandBar commandBar = (CommandBar)cmdbars["Tools"];
				// There's a problem in VS2010: Although the command is available, it is not displayed in the Tools menu when the addin is not loaded.
				// But adding another control to the command causes a duplicate - persistent ! - menu item. => add control only if command not found.
				CommandBarControl commandBarControl = (CommandBarControl) command.AddControl(commandBar, 1);
			}	
		}

		private void WireButtons() {
			if (wiredButtons == null) return;
			foreach (CommandBarButton button in wiredButtons) {
				button.Click += new _CommandBarButtonEvents_ClickEventHandler(ButtonClick);
			}
		}

		internal StringCollection Log { 
			get { 
				return log; } 
		}

		internal event MMCommandEvent OnCommand;
	}
}