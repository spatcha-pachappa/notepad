using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

using System.Reflection;
using System.Runtime.CompilerServices;
using System.Resources;

[assembly: AssemblyTitle("メモ帳 (ANSI)")]
[assembly: AssemblyVersion("1.1.0.0")]

/*

1.1.0.0
挙動を変更した。
	Open(), CheckChanged(), Corner()
カーソル位置表示をSendMessage()を用いる方法に変更した。
テキストボックスのフォントについて、変更時とフォーム終了時に明示的に破棄するようにした。
「メモ帳で開く」メニューを実行した際の「/A」オプションを外した(文字コードはメモ帳で自動識別される)。
ZoomFont()を実装した(未導入)。
	導入時には、BaseFontSizeを記憶しておき、それをもとに倍率を増減させる仕組みにするといいかもしれない。
	フォントダイアログによる単純なフォント変更とは区別される。

*/

namespace CSharp
{

public class Program
{
	[STAThread]
	public static void Main(string[] args)
	{
		Application.EnableVisualStyles();
		Application.SetCompatibleTextRenderingDefault(false);
		var arg = (0 < args.Length) ? args[0] : "";
		Application.Run(new FormMain(arg));
	}
}

public class FormMain : Form
{
	private const int CSW = 1024, CSH = 512, LBLH = 16, LBLPOSW = 48;
	private const string AppName = "メモ帳";
	private const string Filter = "テキストファイル|*.txt;*.cmd;*.cs|全てのファイル (*.*)|*.*";
	private TextBoxEx TBxMain;
	private Label LblInfo, LblRow, LblColumn;
	private MenuStrip MenuMain;
	private ToolStripMenuItem MenuFile, MenuOption;
	// ファイル名
	private string _FileName = "";
	private string FileName
	{
		get
		{
			return this._FileName;
		}
		set
		{
			if (value == "")
			{
				this.Text = AppName;
			}
			else
			{
				this.Text = string.Format("{0} - {1}", AppName, Path.GetFileName(value));
			}
			this._FileName = value;
		}
	}
	// テキストのハッシュ値
	private byte[] Hash = {};
	// バイナリかも
	private bool IsBinary = false;

	public FormMain(string fileName)
	{
		this.FormBorderStyle = FormBorderStyle.Sizable;
		this.StartPosition = FormStartPosition.CenterScreen;
		this.ClientSize = new Size(CSW, CSH);
		this.AllowDrop = true;
		this.Text = AppName;

		this.MenuMain = new MenuStrip();

		this.MenuFile = new ToolStripMenuItem
		{
			Text = "ファイル (&F)",
		};
		this.MenuMain.Items.Add(this.MenuFile);
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuNew",
				Text = "新規作成 (&N)",
				ShortcutKeys = Keys.Control | Keys.N,
				ShowShortcutKeys = true,
			}
		);
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuNewWindow",
				Text = "新しいウィンドウ (&W)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.N,
				ShowShortcutKeys = true,
			}
		);
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuOpen",
				Text = "開く (&O)",
				ShortcutKeys = Keys.Control | Keys.O,
				ShowShortcutKeys = true,
			}
		);
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuOpenAgain",
				Text = "開き直す (&G)",
			}
		);
		this.MenuFile.DropDownItems.Add(new ToolStripSeparator());
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuSave",
				Text = "保存 (&S)",
				ShortcutKeys = Keys.Control | Keys.S,
				ShowShortcutKeys = true,
			}
		);
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuSaveAs",
				Text = "名前を付けて保存 (&A)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.S,
				ShowShortcutKeys = true,
			}
		);
		this.MenuFile.DropDownItems.Add(new ToolStripSeparator());
		this.MenuFile.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuExit",
				Text = "終了 (&X)",
			}
		);

		this.MenuOption = new ToolStripMenuItem
		{
			Text = "機能 (&O)",
		};
		this.MenuMain.Items.Add(this.MenuOption);
		this.MenuOption.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuNotepad",
				Text = "メモ帳で開く (&M)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.M,
				ShowShortcutKeys = true,
			}
		);
		this.MenuOption.DropDownItems.Add(new ToolStripSeparator());
		this.MenuOption.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuSpaceReplace",
				Text = "半角スペース x4 をタブに置換する (&T)",
				ShortcutKeys = Keys.Control | Keys.T,
				ShowShortcutKeys = true,
			}
		);
		this.MenuOption.DropDownItems.Add(new ToolStripSeparator());
		this.MenuOption.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuFont",
				Text = "フォント (&F)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.F,
				ShowShortcutKeys = true,
			}
		);
		this.MenuOption.DropDownItems.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuBackColor",
				Text = "背景色 (&B)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.B,
				ShowShortcutKeys = true,
			}
		);

		this.MenuMain.Items.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuCorner",
				Text = "右下に寄せる (&C)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.C,
				Alignment = ToolStripItemAlignment.Right,
			}
		);

		this.MenuMain.Items.Add
		(
			new ToolStripMenuItem
			{
				Name = "MenuTop",
				Text = "最前面切り替え (&T)",
				ShortcutKeys = Keys.Control | Keys.Shift | Keys.T,
				Alignment = ToolStripItemAlignment.Right,
			}
		);

		this.Controls.Add(this.MenuMain);
		this.MainMenuStrip = this.MenuMain;

		int dy = this.MenuMain.Height;
		var font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

		this.TBxMain = new TextBoxEx
		{
			Width = this.ClientSize.Width,
			Height = this.ClientSize.Height - dy - LBLH,
			Left = 0,
			Top = dy,
			Font = font,
			Multiline = true,
			MaxLength = 0,
			AcceptsReturn = true,
			ReadOnly = false,
			AutoSize = false,
			TabStop = false,
			ScrollBars = ScrollBars.Vertical,
			Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right),
			AllowDrop = true,
			Text = "",
			ForeColor = Color.Black,
			BackColor = Color.White,
		};
		this.Controls.Add(this.TBxMain);

		this.LblColumn = new Label
		{
			Width = LBLPOSW,
			Height = LBLH,
			Left = this.ClientSize.Width - LBLPOSW,
			Top = this.TBxMain.Bottom,
			Font = font,
			Anchor = (AnchorStyles.Bottom | AnchorStyles.Right),
			TextAlign = ContentAlignment.MiddleRight,
			AllowDrop = true,
		};
		this.Controls.Add(this.LblColumn);

		this.LblRow = new Label
		{
			Width = LBLPOSW,
			Height = LBLH,
			Left = this.ClientSize.Width - LBLPOSW * 2,
			Top = this.TBxMain.Bottom,
			Font = font,
			Anchor = (AnchorStyles.Bottom | AnchorStyles.Right),
			TextAlign = ContentAlignment.MiddleRight,
			AllowDrop = true,
		};
		this.Controls.Add(this.LblRow);

		this.LblInfo = new Label
		{
			Width = this.ClientSize.Width - LBLPOSW * 2,
			Height = LBLH,
			Left = 0,
			Top = this.TBxMain.Bottom,
			Font = font,
			Anchor = (AnchorStyles.Bottom | AnchorStyles.Left),
			TextAlign = ContentAlignment.MiddleLeft,
			AllowDrop = true,
		};
		this.Controls.Add(this.LblInfo);

		// タブの幅を変更する
		int tabSize = 4;
		int tabWidth = (int)(tabSize * 3 * 96 / 72);
		Win32.SendMessage(this.TBxMain.Handle, Win32.EM_SETTABSTOPS, 1, new int[]{tabWidth, });

		this.Load += this.FormMain_Load;
		this.FormClosing += this.FormMain_FormClosing;
		this.DragEnter += this.Controls_DragEnter;
		this.DragDrop += this.Controls_DragDrop;

		foreach (var m in this.MenuMain.Items)
		{
			if (m.GetType().Name == "ToolStripMenuItem")
			{
				((ToolStripMenuItem)m).Click += this.MenuMain_Click;
			}
		}
		foreach (var m in this.MenuFile.DropDownItems)
		{
			if (m.GetType().Name == "ToolStripMenuItem")
			{
				((ToolStripMenuItem)m).Click += this.MenuMain_Click;
			}
		}
		foreach (var m in this.MenuOption.DropDownItems)
		{
			if (m.GetType().Name == "ToolStripMenuItem")
			{
				((ToolStripMenuItem)m).Click += this.MenuMain_Click;
			}
		}

		this.TBxMain.TextChanged += this.TBxMain_TextChanged;
		this.TBxMain.KeyPress += this.TBxMain_KeyPress;
		this.TBxMain.KeyDown += this.TBxMain_KeyDown;
		this.TBxMain.MouseDown += this.TBxMain_MouseDown;

		this.TBxMain.DragEnter += this.Controls_DragEnter;
		this.TBxMain.DragDrop += this.Controls_DragDrop;
		this.LblInfo.DragEnter += this.Controls_DragEnter;
		this.LblInfo.DragDrop += this.Controls_DragDrop;
		this.LblRow.DragEnter += this.Controls_DragEnter;
		this.LblRow.DragDrop += this.Controls_DragDrop;
		this.LblColumn.DragEnter += this.Controls_DragEnter;
		this.LblColumn.DragDrop += this.Controls_DragDrop;

		// 初期状態のハッシュ
		this.Hash = this.GetHash(this.TBxMain.Text);

		// 有効なファイルが渡されていたら開く
		if (File.Exists(fileName))
		{
			this.FileName = fileName;
			this.Open(false);
		}
	}

	private void FormMain_Load(object sender, EventArgs e)
	{
		this.ActiveControl = this.TBxMain;
		this.TBxMain.SelectionStart = this.TBxMain.SelectionLength = 0;
		this.TBxMain_TextChanged(sender, e);
		this.Corner();
		this.SetLabel();
	}

	private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
	{
		if (!this.CheckChanged())
		{
			e.Cancel = true;
			return;
		}

		this.TBxMain.Font.Dispose();
	}

	private void Controls_DragEnter(object sender, DragEventArgs e)
	{
		if (e.Data.GetDataPresent(DataFormats.FileDrop))
		{
			e.Effect = DragDropEffects.Copy;
		}
		else
		{
			e.Effect = DragDropEffects.None;
		}
	}

	private void Controls_DragDrop(object sender, DragEventArgs e)
	{
		string[] fns = (string[])e.Data.GetData(DataFormats.FileDrop, false);
		if (File.Exists(fns[0]))
		{
			if (!this.CheckChanged())
			{
				return;
			}
			this.FileName = fns[0];
			this.Open(false);
		}
	}

	private void MenuMain_Click(object sender, EventArgs e)
	{
		var m = (ToolStripMenuItem)sender;
		switch (m.Name)
		{
			case "MenuNew":
			case "MenuOpen":
				if (!this.CheckChanged())
				{
					return;
				}
				if (m.Name == "MenuNew")
				{
					this.New();
				}
				else if (m.Name == "MenuOpen")
				{
					this.Open(true);
				}
				break;
			case "MenuNewWindow":
				Process.Start(Assembly.GetExecutingAssembly().Location);
				break;
			case "MenuOpenAgain":
				if (!this.CheckChanged())
				{
					return;
				}
				this.Open(false);
				break;
			case "MenuSave":
				this.Save(false);
				break;
			case "MenuSaveAs":
				this.Save(true);
				break;
			case "MenuExit":
				this.Close();
				break;
			case "MenuNotepad":
				if (this.FileName != "")
				{
					var fn = Environment.GetFolderPath(Environment.SpecialFolder.System).TrimEnd('\\') + @"\notepad.exe";
//					var args = string.Format("/A \"{0}\"", this.FileName);
					var args = string.Format("\"{0}\"", this.FileName);
					Process.Start(fn, args);
				}
				break;
			case "MenuSpaceReplace":
				this.TBxMain.Text = this.TBxMain.Text.Replace(new string(' ', 4), "\t");
				break;
			case "MenuFont":
				this.ChangeFont();
				break;
			case "MenuBackColor":
				this.ChangeBackColor();
				break;
			case "MenuCorner":
				this.Corner();
				break;
			case "MenuTop":
				this.TopMost = !this.TopMost;
				break;
		}
	}

	private void TBxMain_TextChanged(object sender, EventArgs e)
	{
		this.SetLabel();
		this.ShowCursorPos();
	}

	private void TBxMain_KeyPress(object sender, KeyPressEventArgs e)
	{
		// 選択状態でタブキーが押されたらキャンセルする
		if ((e.KeyChar == '\t') && (0 < this.TBxMain.SelectionLength))
		{
			e.Handled = true;
		}
	}

	private void TBxMain_KeyDown(object sender, KeyEventArgs e)
	{
		// 選択状態で(シフト+)タブキーが押されたら全ての行のタブを制御する
		if ((e.KeyCode == Keys.Tab) && (0 < this.TBxMain.SelectionLength))
		{
			var text = this.TBxMain.Text + Environment.NewLine;
			int selectstart = this.TBxMain.SelectionStart - 1;
			if (selectstart < 0)
			{
				selectstart = 0;
			}
			int st = text.LastIndexOf(Environment.NewLine, selectstart);
			st = (st < 0) ? 0 : st + Environment.NewLine.Length;
			int selectionEnd = this.TBxMain.SelectionStart + this.TBxMain.SelectionLength;
			int en;

			if ((0 < selectionEnd) && text.Substring(0, selectionEnd).EndsWith(Environment.NewLine))
			{
				// 選択範囲の終了位置が次の行の先頭なら、その行は含めない
				en = selectionEnd - Environment.NewLine.Length;
			}
			else
			{
				en = text.IndexOf(Environment.NewLine, selectionEnd);

				if (en < 0)
				{
					en = text.Length;
				}
			}
			this.TBxMain.Select(st, en - st);
			text = this.TBxMain.SelectedText;

			if (e.Shift)
			{
				if (text.StartsWith("\t"))
				{
					text = text.Remove(0, 1);
				}
				text = text.Replace(Environment.NewLine + "\t", Environment.NewLine);
			}
			else
			{
				var lines = text.Split(new string[]{"\r\n", "\r", "\n"}, StringSplitOptions.None);
				var result = new StringBuilder();

				for (int i = 0; i < lines.Length; i ++)
				{
					if (!string.IsNullOrWhiteSpace(lines[i]))
					{
						result.Append('\t');
					}

					result.Append(lines[i]);

					if (i < lines.Length - 1)
					{
						result.Append(Environment.NewLine);
					}
				}
				text = result.ToString();
			}
			this.TBxMain.SelectedText = text;
			this.TBxMain.Select(st, text.Length);
			this.TBxMain.Modified = true;
			this.TBxMain_TextChanged(sender, e);
		}

		// カーソル位置を取得する
		switch (e.KeyCode)
		{
			case Keys.Alt:
			case Keys.Back:
			case Keys.ControlKey:
			case Keys.Delete:
			case Keys.Down:
			case Keys.End:
			case Keys.Enter:
			case Keys.Escape:
			case Keys.Home:
			case Keys.Insert:
			case Keys.Left:
			case Keys.PageDown:
			case Keys.PageUp:
			case Keys.Pause:
			case Keys.PrintScreen:
			case Keys.Right:
			case Keys.Shift:
			case Keys.Up:
				this.ShowCursorPos();
				break;
		}
	}

	private void TBxMain_MouseDown(object sender, MouseEventArgs e)
	{
		this.ShowCursorPos();
	}

	private void SetLabel()
	{
		var s1 = (this.FileName == "") ? "無題" : this.FileName;
		var s2 = (this.IsBinary) ? " [B]" : "";
		var s3 = (this.TBxMain.Modified) ? " *" : "";
		this.LblInfo.Text = string.Format("{0}{1}{2}", s1, s2, s3);
	}

/*
	private void ShowCursorPos()
	{
		// カーソルの位置を取得する
		// 変更がないときはtrueを返す
		if (this.Running)
		{
			return;
		}
		this.Running = true;
		var sw = new Stopwatch();
		sw.Start();
		int r = 1, c = 1;
		try
		{
			int startPos = 0, selectPos = this.TBxMain.SelectionStart;
			int i = 1;
			for (int endPos = 0; ((endPos = this.TBxMain.Text.IndexOf('\n', startPos)) < selectPos) && (-1 < endPos); i ++)
			{
				startPos = endPos + 1;
				if (200 < sw.ElapsedMilliseconds)
				{
					return;
				}
			}

			r = i;
			c = selectPos - startPos + 1;
			this.LblRow.Text = r.ToString();
			this.LblColumn.Text = c.ToString();
		}
		catch
		{
		}
		finally
		{
			sw.Stop();
			this.Running = false;
		}
	}
*/

	private void ShowCursorPos()
	{
		// カーソルの位置を取得する
		int pos = this.TBxMain.SelectionStart;

		int row = (int)Win32.SendMessage(this.TBxMain.Handle, Win32.EM_LINEFROMCHAR, pos, 0).ToInt64();
		int index = (int)Win32.SendMessage(this.TBxMain.Handle, Win32.EM_LINEINDEX, row, 0).ToInt64();

		this.LblRow.Text = (row + 1).ToString();
		this.LblColumn.Text = (pos - index + 1).ToString();
	}

	private bool ShowOpenDialog(string DefaultDirectory, out string ReturnValue)
	{
		// ---- 開くダイアログを表示する ----
		var ofd = new OpenFileDialog
		{
			DefaultExt = "txt",
			FileName = "",
			Filter = Filter,
			InitialDirectory = DefaultDirectory,
			Title = "開く",
		};
		var b = (ofd.ShowDialog(this) == DialogResult.OK);
		if (b)
		{
			ReturnValue = ofd.FileName;
		}
		else
		{
			ReturnValue = "";
		}
		return b;
	}

	private bool ShowSaveDialog(string DefaultFileName, string DefaultDirectory, out string ReturnValue)
	{
		// ---- 保存ダイアログを表示する ----
		var sfd = new SaveFileDialog
		{
			DefaultExt = "txt",
			FileName = DefaultFileName,
			Filter = Filter,
			InitialDirectory = DefaultDirectory,
			Title = "保存",
		};
		var b = (sfd.ShowDialog(this) == DialogResult.OK);
		if (b)
		{
			ReturnValue = sfd.FileName;
		}
		else
		{
			ReturnValue = "";
		}
		return b;
	}

	private bool CheckChanged()
	{
		// TextBox.Modifiedを確認した後にTextのハッシュを比較して変更の有無を判断する
		if (this.TBxMain.Modified)
		{
			byte[] b = this.GetHash(this.TBxMain.Text);
			if (this.Hash.SequenceEqual(b))
			{
				this.TBxMain.Modified = false;
				this.SetLabel();
				return true;
			}
			var dr = MessageBox.Show("変更を保存しますか？", AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
			if (dr == DialogResult.Yes)
			{
				return this.Save(false);
			}
			else if (dr == DialogResult.Cancel)
			{
				return false;
			}
		}
		return true;
	}

	private void New()
	{
		this.FileName = "";
		this.Hash = this.GetHash(this.TBxMain.Text = "");
		this.IsBinary = false;
		this.TBxMain.Modified = false;
		this.SetLabel();
		this.ShowCursorPos();
	}

	private bool Open(bool showDialog)
	{
		// showDialogがfalseのときはダイアログを表示せずにthis.FileNameを開く
		var dd = (this.FileName == "") ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) : Path.GetDirectoryName(this.FileName);
		var fn = "";
		if (showDialog)
		{
			if (!this.ShowOpenDialog(dd, out fn))
			{
				return false;
			}
			this.FileName = fn;
		}
		else
		{
			if (!File.Exists(this.FileName))
			{
				return false;
			}
		}

		byte[] bs;
		try
		{
			using (var fs = new FileStream(this.FileName, FileMode.Open, FileAccess.Read))
			{
				bs = new byte[fs.Length];
				int offset = 0;

				while (offset < bs.Length)
				{
					int read = fs.Read(bs, offset, bs.Length - offset);

					if (read == 0)
					{
						break;
					}

					offset += read;
				}

				if (offset != bs.Length)
				{
					throw new IOException("ファイルを完全に読み込めませんでした。");
				}
			}

			this.IsBinary = false;
			// 0x00のバイトは0x20で表示する
			for (int i = 0; i < bs.Length; i ++)
			{
				if (!this.IsBinary)
				{
					// バイナリかものフラグ
					this.IsBinary |= this.IsCtrlChar(bs[i]);
				}
				bs[i] = (bs[i] == 0x00) ? (byte)0x20 : bs[i];
			}
			this.TBxMain.Text = Encoding.GetEncoding(932).GetString(bs);
			this.TBxMain.SelectionStart = this.TBxMain.SelectionStart;
		}
		catch
		{
			this.FileName = "";
			this.TBxMain.Modified = false;
			this.SetLabel();
			this.ShowCursorPos();
			return false;
		}
		this.Hash = this.GetHash(this.TBxMain.Text);
		this.TBxMain.Modified = false;
		this.SetLabel();
		this.ShowCursorPos();
		return true;
	}

	private bool Save(bool showDialog)
	{
		// showDialogがfalseのときはダイアログを表示せずにthis.FileNameで保存する
		if (this.IsBinary)
		{
			var mt = "バイナリファイルを上書きしようとしている可能性があります。続行しますか？";
			var dr = MessageBox.Show(mt, AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
			switch (dr)
			{
				case DialogResult.Yes:
					break;
				case DialogResult.No:
					showDialog = true;
					break;
				case DialogResult.Cancel:
					return false;
			}
		}
		if ((this.FileName == "") || (showDialog))
		{
			var df = (this.FileName == "") ? DateTime.Now.ToString("yyyyMMddHHmmss") : Path.GetFileName(this.FileName);
			var dd = (this.FileName == "") ? Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) : Path.GetDirectoryName(this.FileName);
			var fn = "";
			if (!this.ShowSaveDialog(df, dd, out fn))
			{
				return false;
			}
			this.FileName = fn;
		}
		File.WriteAllText(this.FileName, this.TBxMain.Text, Encoding.GetEncoding(932));
		this.Hash = this.GetHash(this.TBxMain.Text);
		this.TBxMain.Modified = false;
		this.IsBinary = false;
		this.SetLabel();
		this.ShowCursorPos();
		return true;
	}

	private void ChangeBackColor()
	{
		var cd = new ColorDialog
		{
			Color = this.TBxMain.BackColor,
		};
		if (cd.ShowDialog(this) == DialogResult.OK)
		{
			this.TBxMain.BackColor = cd.Color;
		}
	}

	private void ChangeFont()
	{
		var fd = new FontDialog
		{
			Font = this.TBxMain.Font,
			Color = this.TBxMain.ForeColor,
			FontMustExist = true,
			ShowColor = true,
			ShowEffects = true,
		};
		if (fd.ShowDialog(this) == DialogResult.OK)
		{
			var oldFont = this.TBxMain.Font;
			this.TBxMain.Font = new Font(fd.Font, fd.Font.Style);
			this.TBxMain.ForeColor = fd.Color;
			oldFont.Dispose();
		}
	}

	private void ZoomFont(float scale)
	{
		var oldFont = this.TBxMain.Font;
		this.TBxMain.Font = new Font(oldFont.FontFamily, oldFont.Size * scale, oldFont.Style);
		oldFont.Dispose();
	}

/*
	private void Corner()
	{
		int x = Screen.GetWorkingArea(this).Width - this.Width;
		int y = Screen.GetWorkingArea(this).Height - this.Height;
		x = (x < 0) ? 0 : x;
		y = (y < 0) ? 0 : y;
		// なぜか7ピクセルだけ隙間が空く
		x += 7;
		y += 7;
		this.SetDesktopLocation(x, y);
	}
*/
	private void Corner()
	{
		var area = Screen.GetWorkingArea(this);
		int x = area.Right - this.Width;
		int y = area.Bottom - this.Height;
		// なぜか7ピクセルだけ隙間が空く
		x += 7;
		y += 7;
		this.SetDesktopLocation(x, y);
	}

	private byte[] GetHash(string src)
	{
		byte[] dst = {};
		using (var sha1 = new SHA1CryptoServiceProvider())
		{
			dst = sha1.ComputeHash(Encoding.GetEncoding(932).GetBytes(src));
		}
		return dst;
	}

	private bool IsCtrlChar(byte b)
	{
		// ---- ASCII制御文字かどうか(ただし、水平タブとCRLFを除く) ----
		return ((b < 0x09) || (b == 0x11) || (b == 0x12) || ((0x13 < b) && (b < 0x20)) || (b == 0x7F));
	}
}

public class TextBoxEx : TextBox
{
	public TextBoxEx()
	{
		this.DoubleBuffered = true;
	}
}

public static class Win32
{
	public const int EM_LINEINDEX = 0x00BB;
	public const int EM_LINEFROMCHAR = 0x00C9;
	public const int EM_SETTABSTOPS = 0x00CB;

	[DllImport("kernel32.dll", CharSet = CharSet.Auto)]
	public static extern bool IsDBCSLeadByte(byte TestChar);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	public static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

	[DllImport("user32.dll", CharSet = CharSet.Auto)]
	public static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, int[] lParam);
}

}

