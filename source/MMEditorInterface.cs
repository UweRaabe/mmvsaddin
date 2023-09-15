using System;
using EnvDTE;
using System.IO;
using System.Windows.Forms;
using Extensibility;
using System.Diagnostics;

namespace MMVSAddIn
{
	internal class MMVSEditorInterface: System.Object {

		public MMVSEditorInterface(_DTE application) {
			this.application = application;
		}

        internal Document ActiveDocument
        {
            get
            {
                if (application == null) return null;

                try
                {
                    return application.ActiveDocument;
                }
                catch
                {
                    return null;
                }
            }
        }
        public bool GetCursorPos(out int LineNr, out int Column)
        {
			TextSelection selection = ActiveSelection();
			if (selection == null) {
				LineNr = 0;
				Column = 0;
				return false;
			}
			else {
				LineNr = selection.ActivePoint.Line - 1;
				Column = selection.ActivePoint.LineCharOffset - 1;
				return true;
			}
		}

		public bool GetModuleCode(string FileName, out string Code) {
			TextDocument textDoc = (FileName != "") && File.Exists(FileName)? FindTextDocument(FileName) : null;
			return GetDocumentCode(textDoc, out Code);
		}

		public string GetTopModule() {
            Document document = ActiveDocument;
			return (document == null) ? "" : document.FullName;
		}

		public bool IsBufferedModule(string FileName) {
			return  (application != null) &&
							(FileName != "") &&
							File.Exists(FileName) &&
							application.ItemOperations.IsFileOpen(FileName, Constants.vsViewKindCode);
		}

		public bool OpenModule(string FileName) {
			if (application == null) return false;
			try {
				return File.Exists(FileName) && (application.ItemOperations.OpenFile(FileName, Constants.vsViewKindCode) != null);
			}
			catch {
				return false;
			}
		}

		public void ReloadModule(string fileName) {
			TextDocument textDocument = FindTextDocument(fileName);
			if (textDocument == null ) return;
			int topLine = textDocument.Selection.TextPane.StartPoint.Line;
			int cursorLine = textDocument.Selection.ActivePoint.Line;
			int cursorColumn = textDocument.Selection.ActivePoint.VirtualDisplayColumn;
			// close the document
			textDocument.Parent.Close(vsSaveChanges.vsSaveChangesNo);
			// and reopen it.
			application.ItemOperations.OpenFile(fileName, Constants.vsViewKindPrimary );
			textDocument = FindTextDocument(fileName);
			if (textDocument == null) return;
			// relocate topline and cursor
			TextSelection selection = textDocument.Selection;
			documentSize = selection.Parent.EndPoint.Line;
			try {
				selection.MoveToLineAndOffset(TransformLineNr(topLine - 1), 1, false);
				selection.Collapse();
				selection.TextPane.TryToShow(selection.AnchorPoint, vsPaneShowHow.vsPaneShowTop, null);
				selection.MoveToLineAndOffset(TransformLineNr(cursorLine - 1), 1, false);
				// MoveToDisplayColumn can fail of display column is beyond EndOfLine
				selection.EndOfLine(false);
				if (cursorColumn > selection.ActivePoint.VirtualDisplayColumn)
					cursorColumn = selection.ActivePoint.VirtualDisplayColumn;
				if (cursorColumn < 1) cursorColumn = 1;
				selection.MoveToDisplayColumn(TransformLineNr(cursorLine - 1), cursorColumn, false);
			}
			catch (System.Exception e) {
				MessageBox.Show(e.Message);
				// ignore exceptions
			}
		}

		public Document SaveActiveDocument() {
			Document document = ActiveDocument;
			if ( (document != null) && (!document.Saved) && (!document.ReadOnly) ) {
				try {
					document.Save("");
				}
				catch (System.Exception e) { // all
					MessageBox.Show(e.Message);
				}
			}
			return document;
		}

		public void SetCursorPos(int LineNr, int Column) {
			TextSelection selection = ActiveSelection();
			if (selection == null) return;
			UpdateCurDocSize();
			selection.Collapse();
			selection.MoveToLineAndOffset(TransformLineNr(LineNr), TransformColumn(Column), false);
		}

		public void SetScrollPos(int TopLine, int FocusLine, int Column) {
			TextSelection selection = ActiveSelection();
			if (selection == null) return;
			UpdateCurDocSize();
			selection.MoveToLineAndOffset(TransformLineNr(TopLine), 1, false);
			selection.Collapse();
			selection.TextPane.TryToShow(selection.AnchorPoint, vsPaneShowHow.vsPaneShowTop, null);
			selection.MoveToLineAndOffset(TransformLineNr(FocusLine), TransformColumn(Column), false);
			// vsPaneShowAsIs: The lines displayed remain the same unless it is necessary to move the display to show the text.
			// selection.TextPane.TryToShow(selection.AnchorPoint, vsPaneShowHow.vsPaneShowAsIs, null);
		}

		internal string GetEditorOptions(string Language) {
			string Result = "";
			if (application == null) return Result;
			Properties Props = null;
			try {
				Props = application.get_Properties("TextEditor", Language);
			}
			catch {
				Props = null;
			}
			if (Props == null) Props = application.get_Properties("TextEditor", "PlainText");
			if (Props != null) {
				foreach (Property P in Props) {
					Result += P.Name + '=' + P.Value + "\n";
				}
			}
			return Result;
		}

		private _DTE application;
		private int documentSize;

		private TextSelection ActiveSelection() {
			Document document = ActiveDocument;
			TextSelection selection;
			selection = document == null ? null : document.Selection as TextSelection;
			return selection;
		}

		private TextDocument ActiveTextDocument() {
			Document document = ActiveDocument;
			return document == null ? null : document.Object("") as TextDocument;
		}

		private Document FindDocument(string fileName) {
			for (int I = 1 ; I <= application.Documents.Count ; I++) {
				if (SameText(application.Documents.Item(I).FullName, fileName))
					return application.Documents.Item(I);
			}
			return null;
		}

		private TextDocument FindTextDocument(string fileName) {
			Document document = FindDocument(fileName);
			return document == null ? null : document.Object("") as TextDocument;
		}

		private bool GetDocumentCode(TextDocument textDocument, out string code) {
			if (textDocument != null) {
				EditPoint editPoint = textDocument.CreateEditPoint(textDocument.StartPoint);
				code = editPoint.GetText(textDocument.EndPoint);
				return true;
			}
			else {
				code = "";
				return false;
			}
		}

		private bool SameText(string Str1, string Str2) {
			return String.Compare(Str1, Str2, true) == 0;
		}

		private void SetVirtualScrollPos(TextSelection selection, int TopLine, int FocusLine, int Column) {
			documentSize = selection.Parent.EndPoint.Line;
			selection.MoveToLineAndOffset(TransformLineNr(TopLine), 1, false);
			selection.Collapse();
			selection.TextPane.TryToShow(selection.AnchorPoint, vsPaneShowHow.vsPaneShowTop, null);
			selection.MoveToDisplayColumn(TransformLineNr(FocusLine), TransformColumn(Column), false);
			// vsPaneShowAsIs: The lines displayed remain the same unless it is necessary to move the display to show the text.
			// selection.TextPane.TryToShow(selection.AnchorPoint, vsPaneShowHow.vsPaneShowAsIs, null);
		}

		private int TransformColumn(int Column) {
			return Column < 0 ? 1 : Column + 1;
		}

		private int TransformLineNr(int LineNr) {
			LineNr = LineNr + 1; // 0 based to 1-based
			if (LineNr < 1)
				return 1 ;
			else
				if (LineNr > documentSize)
				return documentSize;
			else
				return LineNr;
		}

		private void UpdateCurDocSize() {
			TextDocument textDocument = ActiveTextDocument();
			if (textDocument != null)
				documentSize = textDocument.EndPoint.Line;
			else
				documentSize = 1;
		}
	}
}
