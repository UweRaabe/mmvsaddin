using System;
using System.Collections;
using Extensibility;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using EnvDTE;

namespace MMVSAddIn
{
	internal class StackTraceBuilder
	{
		private static string assemblyFileName;

		private static void DefaultSplitFunctionName(ref string functionName, ref string className)
		{
			if ((functionName == null) || (functionName == "")) return;
			char[] separators = new char[] { '.' };
			string[] sections = functionName.Split(separators);
			ArrayList list = new ArrayList();
			foreach (string section in sections)
			{
				if ((section == null) || (section == "")) continue;
				list.Add(section);
			}
			if (list.Count <= 1) return;
			if ( (list.Count >= 3) &&  (list[list.Count - 1] as string).Equals("get") || ((list[list.Count - 1] as string).Equals("set"))) 
			{
				className = list[list.Count - 3] as string;
				functionName = list[list.Count - 1] as string + "." + list[list.Count - 2] as string;
			}
			else
			{
				className = list[list.Count - 2] as string;
				functionName = list[list.Count - 1] as string;
			}
      }

		private static void SplitFunctionName(string moduleName, ref string functionName, ref string className)
		{
			assemblyFileName = moduleName;
			Assembly assembly = null;
			try
			{
				// 32-bit add-in cannot load 64-bit target assemblies.
				assembly = Assembly.LoadFile(moduleName);
			}
			catch
			{
				DefaultSplitFunctionName(ref functionName, ref className);
				return;
			}
			Type[] types;
			try {
				types = assembly.GetTypes(); // GetExportedTypes for just publics
				// use reversed types to pick up nested types before containing types.
				// The containing type will match also for the nested type.
				ArrayList reversedTypes = new ArrayList();
				foreach (Type type in types)
				{
					reversedTypes.Add(type);
				}
				reversedTypes.Reverse();
				foreach (Type type in reversedTypes)
				{
					if (type == null) continue;
					string qualifiedName = type.FullName.Replace('+', '.'); // nested types are concatenated with + rather than .
					if (functionName.StartsWith(qualifiedName + '.')) {
						className = type.Name;
						functionName = functionName.Remove(0, qualifiedName.Length + 1);
						// revert PropName.set|get
						if (functionName.EndsWith(".get") || functionName.EndsWith(".set"))
						{
							string prefix = functionName.Substring(functionName.Length - 3);
							// Remove(int32) not supported in v1.1
							functionName = prefix + '.' + functionName.Remove(functionName.Length - 4, 4);
						}
						return;
					}
				}

			}
			catch (ReflectionTypeLoadException)
			{
				// ignore and revert to default behaviour
                DefaultSplitFunctionName(ref functionName, ref className);
			}
		}


		internal static string BuildStackTrace(_DTE dte)
		{
			EnvDTE.Thread thread = dte.Debugger.CurrentThread;
			if (thread == null) throw new Exception("No program is being debugged.\nStack trace is only available when debugger hits breakpoint.");
			if (thread.StackFrames.Count == 0) throw new Exception("Stack trace is only available when debugger hits breakpoint.");
			// use reversed order to create sequence diagram
			ArrayList frames = new ArrayList(thread.StackFrames.Count);
			foreach (StackFrame frame in thread.StackFrames) frames.Add(frame);
			frames.Reverse();
			StringBuilder output = new StringBuilder();
			// build stack using reversed order
			foreach (StackFrame frame in frames)
			{
				// skip native <-> managed transitions
				if (frame.FunctionName.StartsWith("[")) continue;
				output.Append("\n");
				output.Append(frame.Module);
				output.Append("|");
				string functionName = frame.FunctionName;
				string className = "";
				SplitFunctionName(frame.Module, ref functionName, ref className);
				if (!((className == null) || (className == "")))
				{
					output.Append(className);
					output.Append(".");
				}
				output.Append(functionName);
				output.Append("(");
				EnvDTE.Expressions expressions = frame.Arguments;
				bool firstParam = true;
				foreach (EnvDTE.Expression exp in expressions)
				{
					if (firstParam)
						firstParam = false;
					else
						output.Append(", ");
					if (exp.IsValidValue && ((exp.DataMembers == null) || (exp.DataMembers.Count == 0)))
						output.Append(exp.Value);
					else
						output.Append(exp.Name);
				}
				output.Append(")");
			}
			return output.ToString();
		}

		private Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args) {
			// Debug.WriteLine("Resolve Assembly: " + args.Name);

			string assemblyName = args.Name;
			int index = assemblyName.IndexOf(',');
			if (index >= 0)
				assemblyName = assemblyName.Remove(index, assemblyName.Length - index);
			assemblyName = assemblyName.Trim() + ".dll";
			assemblyName = Path.GetDirectoryName(assemblyFileName) + '\\' + assemblyName;
			if (File.Exists(assemblyName)) {
				try {
					Assembly	assembly	= Assembly.LoadFile(assemblyName);
					return assembly;
				}
				catch {
					return null;
				}
			}
			return null;
		}
	}
}
