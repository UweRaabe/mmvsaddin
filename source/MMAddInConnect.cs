
#if VS2012
namespace MMVS2012AddIn
#else
#if VS2010
namespace MMVS2010AddIn
#else
#if VS2008
namespace MMVS2008AddIn
#else
#if VS2005
namespace MMVS2005AddIn
#else
namespace MMVS2003AddIn
#endif
#endif
#endif
#endif
{
	using System;
	using Microsoft.Office.Core;
	using Extensibility;
	using System.Runtime.InteropServices;
	using EnvDTE;
	using System.Windows.Forms;
	using System.Diagnostics;
	using System.Text;
	using MMVSAddIn;
	using MMVSExpert;

	internal class ProjectFileFinder: Object
	{
		internal ProjectFileFinder(Project aProject)
		{
			project = aProject;
		}

		private string fileName;
		private readonly Project project;
		private ProjectItem result;

		internal bool ContainsFile(string FileName)
		{
			return (LocateFile(FileName) != null);
		}
		internal ProjectItem LocateFile(string FileName)
		{
			fileName = FileName;
			result = null;
			IterateProjectItems(project.ProjectItems);
			return result;
		}

		private void IterateProjectItems(ProjectItems projectItems)
		{
			ProjectItem projectItem;
			for (int i = 1; i <= projectItems.Count; i++)
			{
				if (result != null) return;
				projectItem = projectItems.Item(i);
				for (short j = 1; j <= projectItem.FileCount; j++)
				{
					if (String.Equals(fileName, projectItem.get_FileNames(j)))
					{
						result = projectItem;
						return;
					}
				}
				ProjectItems subItems = projectItem.ProjectItems;
				if (subItems != null)
					IterateProjectItems(subItems);
			}
		}
	}


#if VS2012
	[GuidAttribute("02BA5347-5A33-44D2-A615-BE8A8C58C437"), ProgId("MMVS2012AddIn.MMVSClient")]
#else
#if VS2010
	[GuidAttribute("5D05C547-36FC-497F-A747-55B84BA4A069"), ProgId("MMVS2010AddIn.MMVSClient")]
#else
#if VS2008
	[GuidAttribute("028A5453-93A9-4417-B64E-F2D5E9886D79"), ProgId("MMVS2008AddIn.MMVSClient")]
else
#if VS2005
	[GuidAttribute("832947A2-8E6F-4590-94E6-24C56F8AB96F"), ProgId("MMVS2005AddIn.MMVSClient")]
#else
	[GuidAttribute("822F44DF-590A-4AE4-9CB9-BF81AA066E59"), ProgId("MMVS2003AddIn.MMVSClient")]
#endif
#endif
#endif
#endif
	public class MMVSClient : Object, Extensibility.IDTExtensibility2, IDTCommandTarget, IMMVSClient, IMMVSClientEx
	{
		private IMMVSExpert mmExpert;
		private _DTE application;
		private AddIn addInInstance;
		private MMVSEditorInterface editorInterface;
		private MMCommandHandler commandHandler;
		private DocumentEvents documentEvents;

		private int documentSaveLock;

		public void OnConnection(object DTEApplication, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
		{
			Debug.WriteLine("============= MM8 MMVSClient OnConnection #1");
			application = (_DTE)DTEApplication;
			addInInstance = (AddIn)addInInst;
			Debug.WriteLine("============= MM8 MMVSClient OnConnection #2");
			try
			{

				if (IsVSInSpecialMode())
				{
					//					MessageBox.Show("Unable to initialize ModelMaker integration add-in.\nWas Visual Studio started in a non-visual mode?", "ModelMaker IDE Integration", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}


				editorInterface = new MMVSEditorInterface(application);
				Debug.WriteLine("============= MM8 MMVSClient OnConnection #3");
				mmExpert = new MMVSExpertClass();
				Debug.WriteLine("============= MM8 MMVSClient OnConnection #4");
				mmExpert.Connect(this);
				Debug.WriteLine("============= MM8 MMVSClient OnConnection #5");
				bool useToolsMenu = mmExpert.GetUseToolsMenu() != 0;
				commandHandler = new MMCommandHandler(application, addInInstance, this, useToolsMenu);
				commandHandler.OnCommand += new MMCommandEvent(CommandHandler_OnCommand);

				documentEvents = application.Events.get_DocumentEvents(null);
				documentEvents.DocumentSaved += new _dispDocumentEvents_DocumentSavedEventHandler(DocumentSaved);

			}
			catch (System.Exception e)
			{
				MessageBox.Show("A fatal error occured loading ModelMaker Integration\n\n" +
					"Please retry loading ModelMaker Integration and contact\n" +
					"ModelMaker Tools: mailto:support@modelmakertools.com \n" +
					"reporting this message:\n" +
					"Msg: \"" + e.Message + "\"",
					"MM Fatal error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			Debug.WriteLine("============= MM8 MMVSClient OnDisconnection #1");
			if (mmExpert != null)
			{
				mmExpert.Disconnect();
				mmExpert = null;
			}
			Debug.WriteLine("============= MM8 MMVSClient OnDisconnection #2");
			editorInterface = null;
			Debug.WriteLine("============= MM8 MMVSClient OnDisconnection #3");

			if (commandHandler != null)
			{
				commandHandler.Disconnect();
				commandHandler.OnCommand -= new MMCommandEvent(CommandHandler_OnCommand);
				commandHandler = null;
			}

			if (documentEvents != null)
			{
				documentEvents.DocumentSaved -= new _dispDocumentEvents_DocumentSavedEventHandler(DocumentSaved);
				documentEvents = null;
			}

			addInInstance = null;
			application = null;
			GC.Collect();
		}

		public void OnAddInsUpdate(ref System.Array custom)
		{
		}

		public void OnStartupComplete(ref System.Array custom)
		{
		}

		public void OnBeginShutdown(ref System.Array custom)
		{
		}

		public string GetModuleCode(string FileName)
		{
			string Code = "";
			editorInterface.GetModuleCode(FileName, out Code);
			return Code;
		}

		public void LocateLineColumn(string FileName, int TopLine, int FocusLine, int Column)
		{
			if (editorInterface.OpenModule(FileName))
			{
				editorInterface.SetScrollPos(TopLine, FocusLine, Column);
			}
		}

		public void OpenModule(string FileName)
		{
			editorInterface.OpenModule(FileName);
		}

		public void PerformCompileAction(int action)
		{
			switch (action)
			{
				case 0:// "SyntaxCheck" is treated as "Compile"
				case 2:
					application.ExecuteCommand("Build.BuildSolution", "");
					break;
				case 1:
					application.ExecuteCommand("Build.RebuildSolution", "");
					break;
			}
		}

		public void ReloadModule(string FileName)
		{
			editorInterface.ReloadModule(FileName);
		}

		public int GetMainWindowHandle()
		{
			return application == null ? 0 : application.MainWindow.HWnd;
		}

		public int GetIsDebugging()
		{
			return (application.Mode == vsIDEMode.vsIDEModeDesign) ? 0 : 1;
		}

		public void ShowClassHelp(string topic)
		{
			application.ExecuteCommand("Help.Index", "");
		}



		public void Exec(string commandName, vsCommandExecOption executeOption, ref object varIn, ref object varOut, ref bool handled)
		{
			handled = false;
			if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
			{
				if (commandHandler != null)
					handled = commandHandler.Execute(commandName);
			}
		}

		public string GetClientData(int Index)
		{
			return "";
		}

		public string GetModifiedBuffers()
		{
			StringBuilder builder = new StringBuilder();
			Documents documents = application.Documents;
			foreach (Document document in documents)
			{
				if (!document.Saved)
				{
					builder.Append(document.FullName);
					builder.Append("\n");
				}
			}
			return builder.ToString();
		}

		public void SetClientData(int Index, string Value)
		{
		}

		public void QueryStatus(string commandName, EnvDTE.vsCommandStatusTextWanted neededText, ref EnvDTE.vsCommandStatus status, ref object commandText)
		{
			if (neededText == EnvDTE.vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
			{
				if (commandHandler != null)
					status = commandHandler.QueryStatus(commandName);
				else
					status = vsCommandStatus.vsCommandStatusInvisible | vsCommandStatus.vsCommandStatusInvisible;
			}
		}

		public void UpdateProjectFiles(string ProjectName, string FileList)
		{
			if ((FileList == null) || (FileList == "")) return;
			Projects projects = (Projects)application.GetObject("CSharpProjects");
			char[] lineSeparators = new char[] { '\n', '\r' };
			foreach (Project project in projects)
			{
				if (project.FullName.ToLower().Equals(ProjectName.ToLower()) )
				{
					// found the project: now add fileLines: split FileList by line number
					string[] fileLines = FileList.Split(lineSeparators);
					ProjectFileFinder fileFinder = new ProjectFileFinder(project);
					ProjectItem projectItem;
					foreach (string fileLine in fileLines)
					{
						if ((fileLine == null) || (fileLine == "")) continue;
						// fileLine is a non-empty entry	, first position contains + or - (cmd)
						char Cmd = fileLine[0];
						string file = fileLine.Remove(0, 1);
						try
						{
							switch (Cmd)
							{
								case '+':
									if (!fileFinder.ContainsFile(file))
										project.ProjectItems.AddFromFile(file);
									break;
								case '-':
									projectItem = fileFinder.LocateFile(file);
									if (projectItem != null) projectItem.Remove();
									break;
							}

						}
						catch (Exception e)
						{
							MessageBox.Show(e.Message, "Error updating project files", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					break; // current project only.
				}
			}
		}

		private void DocumentSaved(EnvDTE.Document document)
		{
			if ((documentSaveLock == 0) && (mmExpert != null))
			{
				mmExpert.ModuleSaved(document.FullName);
			}
		}

		private void CommandHandler_OnCommand(int Cmd)
		{
			if (mmExpert != null)
			{
				Document document = null;
				switch (Cmd)
				{
					case MMCommands.RunMMPascal:
						mmExpert.RunModelMaker(0);
						break;
					case MMCommands.RunMMCSharp:
						mmExpert.RunModelMaker(1);
						break;
					case MMCommands.JumpToModelMaker:
						document = SaveActiveDocument(false);
						if (document != null)
						{
							int lineNumber = 0;
							int column = 0;
							editorInterface.GetCursorPos(out lineNumber, out column);
							mmExpert.LocateLineColumn(document.FullName, lineNumber + 1, column);
						}
						break;
					case MMCommands.InvokeModelSearch:
						document = SaveActiveDocument(false);
						if (document != null)
						{
							int lineNumber = 0;
							int column = 0;
							editorInterface.GetCursorPos(out lineNumber, out column);
							mmExpert.InvokeModelSearch(document.FullName, lineNumber, column);
						}
						break;
					case MMCommands.AddToModel:
						document = SaveActiveDocument(true);
						if (document != null)
							mmExpert.AddToModel(document.FullName);
						break;
					case MMCommands.AddFilesToModel:
						application.Documents.SaveAll();
						mmExpert.AddFilesToModel(BuildSolutionFileList());
						break;
					case MMCommands.ConvertToModel:
						document = SaveActiveDocument(true);
						if (document != null)
							mmExpert.ConvertToModel(document.FullName);
						break;
					case MMCommands.ConvertProjectToModel:
						application.Documents.SaveAll();
						mmExpert.ConvertProjectToModel(BuildActiveProjectFileList());
						break;
					case MMCommands.RefreshInModel:
						document = SaveActiveDocument(true);
						if (document != null)
							mmExpert.RefreshInModel(document.FullName);
						break;
					case MMCommands.SynchronizeModel:
						application.Documents.SaveAll();
						mmExpert.SynchronizeModel();
						break;
					case MMCommands.Properties:
						mmExpert.EditProperties();
						break;
					case MMCommands.CreateSequenceDiagram:
						mmExpert.CreateSequenceDiagram(StackTraceBuilder.BuildStackTrace(application));
						break;
					default:
						Debug.Assert(false);
						break;
				}
			}
		}

		private string BuildSolutionFileList()
		{
			StringBuilder builder = new StringBuilder();
			Projects projects = (Projects)application.GetObject("CSharpProjects");
			foreach (Project project in projects)
			{
				BuildProjectFileList(project, builder);
			}
			builder.Append("project=Editor Files\n");
			foreach (Document document in application.Documents)
			{
				builder.AppendFormat("file={0}\n", document.FullName);
			}
			return builder.ToString();
		}

		private string BuildActiveProjectFileList()
		{
			StringBuilder builder = new StringBuilder();
			System.Array projects;
			projects = application.ActiveSolutionProjects as System.Array;
			foreach (Project project in projects)
			{
				BuildProjectFileList(project, builder);
			}
			return builder.ToString();
		}

		private void BuildProjectFileList(Project project, StringBuilder output)
		{
			output.AppendFormat("project={0}\n", project.FullName);
			AppendProjectItems(project.ProjectItems, output);
		}

		private void AppendProjectItems(ProjectItems projectItems, StringBuilder output)
		{
			ProjectItem projectItem;
			for (int i = 1; i <= projectItems.Count; i++)
			{
				projectItem = projectItems.Item(i);
				for (short j = 1; j <= projectItem.FileCount; j++)
				{
					output.AppendFormat("file={0}\n", projectItem.get_FileNames(j));
					ProjectItems subItems = projectItem.ProjectItems;
					if (subItems != null)
					{
						AppendProjectItems(subItems, output);
					}
				}
			}
		}

		private bool IsVSInSpecialMode()
		{
			try
			{
				if (application.Windows == null) return true;
				Window window = application.Windows.Item(Constants.vsWindowKindOutput);
				return (window == null);
			}
			catch
			{
				return true;
			}
		}

		private Document SaveActiveDocument(bool lockModuleSaveEvent)
		{
			if (lockModuleSaveEvent) documentSaveLock++;
			Document document = editorInterface.SaveActiveDocument();
			if (lockModuleSaveEvent) documentSaveLock--;
			return document;
		}
	}
}