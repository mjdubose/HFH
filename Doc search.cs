using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Automation;
using RichTextBoxLinks;


namespace HFH
{
    public partial class DocSearch : Form
    {
        private readonly BackgroundWorker _docfilespicker;
        private readonly BackgroundWorker _worker;
        private List<TreeNode> _checkedNodes;
        private int _filecounter;
        private string _folderName;
        private readonly List<string> _wordfilepathholder; 
        public DocSearch()
        {
            InitializeComponent();
            _worker = new BackgroundWorker {WorkerReportsProgress = true, WorkerSupportsCancellation = true};
            _worker.DoWork += worker_DoWork;
            _worker.ProgressChanged += worker_ProgressChanged;
            _worker.RunWorkerCompleted += worker_RunWorkerCompleted;
            _docfilespicker = new BackgroundWorker();
            _docfilespicker.DoWork += docfilespicker_DoWork;
            _docfilespicker.RunWorkerCompleted += docfilespicker_RunWorkerCompleted;
            _wordfilepathholder = new List<string>();
            _checkedNodes = new List<TreeNode>();
        }

        private void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progress.Value = e.ProgressPercentage;
        }

        private static void docfilespicker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
        }

        private void docfilespicker_DoWork(object sender, DoWorkEventArgs e)
        {
            FileDialogueWork();
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ProgressVisible();
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var help = (Helper) e.Argument;
            using (var automation = new WordAutomation())
            {
                automation.CreateWordApplication();
                ViewDirectories(automation,_wordfilepathholder, help.Argument);
                automation.CloseWordApp();
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            if (label1.Text == "")
                return;
            _filecounter = 0;
            string textToFind = textSearch.Text;

            string startpath = _folderName;

            if (!Directory.Exists(startpath)) return;
            var dir = new DirectoryInfo(startpath);


            if (_worker.IsBusy) return;
            var help = new Helper {Argument = textToFind, Dir = dir};
            progress.Value = 0;

            progress.Maximum = Convert.ToInt32(label1.Text);

            progress.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            label4.Text = 0.ToString(CultureInfo.InvariantCulture);
            _worker.RunWorkerAsync(help);
        }


        private void ViewDirectories(WordAutomation automation,IEnumerable<string> wordfilepath, string textToFind)
        {
            string[] wordstofind = textToFind.Split(',');

            try
            {
                // automation.CreateWordApplication();
                foreach (var file in wordfilepath)
                {
                    try
                    {
                        if (!_worker.CancellationPending)
                            SearchDocument(automation, file, wordstofind);
                    }
                        // ReSharper disable once EmptyGeneralCatchClause
                    catch (Exception)
                    {
                    }
                }
            }

            catch (UnauthorizedAccessException)
            {
            }
            catch (AccessViolationException)
            {
            }
// ReSharper disable once EmptyGeneralCatchClause
            catch (Exception)
            {
            }
         
        }

        private void ProgressVisible()
        {
            if (progress.InvokeRequired)
            {
                var d = new ProgressVisibleCallBack(ProgressVisible);
                Invoke(d);
            }
            else
            {
                progress.Visible = false;
            }
        }

        private void SetProgress(int amount)
        {
            if (progress.InvokeRequired)
            {
                var d = new ProgressCallBack(SetProgress);
                Invoke(d, new object[] {amount});
            }
            else
            {
                if (amount > progress.Maximum)
                {
                    amount = progress.Maximum;
                }
                progress.Value = amount;
            }
        }

        private void UpdateTextBox(RichTextBoxEx text, Textboxexhelper tbeh)
        {
            if (text.InvokeRequired)
            {
                var d = new SetTextCallback(UpdateTextBox);
                Invoke(d, new object[] {text, tbeh});
            }
            else
            {
                text.InsertLink(tbeh.Filename, tbeh.uri.AbsoluteUri);
                text.SelectedText = " " + tbeh.SearchWord + " " + tbeh.Count + " times" + Environment.NewLine;
                //  text.Text = text.Text +@" " +tbeh.SearchWord +@" "+tbeh.Count+@" times." + Environment.NewLine;
            }
        }

        private void SearchDocument(WordAutomation automation, string filename, IEnumerable<string> textToFind)
        {
            int filewithwordscount = 0;
            _filecounter = _filecounter + 1;

            //  automation.CreateWordApplication();

            automation.CreateWordDoc(filename, false);
            foreach (string word in textToFind)
            {
                var whelper = new Wordcounthelper {Count = 0, Word = ""};
                string temp = word;
                if (!temp.Contains("&"))
                {
                    whelper.Count = automation.GetWordCount(word);
                    whelper.Word = word;
                }
                else
                {
                    bool hasallwords = true;
                    string[] tempholder = temp.Split('&');
                    foreach (int count in tempholder.Select(automation.GetWordCount).Where(count => count <= 0))
                    {
                        hasallwords = false;
                    }
                    if (hasallwords)
                    {
                        whelper.Word = word;
                        whelper.Count = 1;
                    }
                }
                if (whelper.Count <= 0) continue;

                var uri = new Uri(filename, UriKind.RelativeOrAbsolute);
                var tbeh = new Textboxexhelper
                {
                    Count = whelper.Count,
                    SearchWord = whelper.Word,
                    uri = uri,
                    Filename = Path.GetFileName(filename)
                };
                filewithwordscount = filewithwordscount + 1;
                if (filewithwordscount > 1)
                    filewithwordscount = 1;

                UpdateTextBox(textBox1, tbeh);
            }

            SetLabelforWordFileCount(filewithwordscount);
            SetProgress(_filecounter);
            automation.CloseWordDoc(false);
        }

        private void OpenCheckedNodes(IEnumerable nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Checked)
                {
                    node.BackColor = Color.Yellow;
                    if (!_checkedNodes.Contains(node))
                        _checkedNodes.Add(node);
                }
                else
                {
                    node.BackColor = Color.White;
                    OpenCheckedNodes(node.Nodes);
                }
            }
        }

        private void OpenFiles()
        {
            foreach (string treeNodeName in from checkedNode in _checkedNodes
                select checkedNode.ToString().Replace("TreeNode: ", String.Empty)
                into treeNodeName
                let extension = Path.GetExtension(treeNodeName)
                where extension == ".doc" || extension == ".docx"
                select treeNodeName)
            {
                Process.Start(treeNodeName);
            }
        }

        private void ListDirectory(string defaultdir,ICollection<string> wordpath, out int count)
        {
            var di = new DirectoryInfo(defaultdir);
            bool addbecauseofdoc;
            TreeNode temp = CreateDirectoryNode(di,wordpath, out addbecauseofdoc, out count);
            AddNodesToTreeView(temp);
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo,ICollection<string> wordpath, out bool docfilefound, out int count)
        {
            count = 0;
            int temp = count;
            docfilefound = false;
            var directoryNode = new TreeNode(directoryInfo.FullName);
            foreach (DirectoryInfo directory in directoryInfo.GetDirectories())
            {
                try
                {
                    bool tester;

                    TreeNode tempNode = CreateDirectoryNode(directory,wordpath, out tester, out count);
                    temp = temp + count;
                    docfilefound = docfilefound || tester;
                    if (tester)
                    {
                        directoryNode.Nodes.Add(tempNode);
                    }
                }

                catch (UnauthorizedAccessException)
                {
                }
                catch (AccessViolationException)
                {
                }
                    // ReSharper disable once EmptyGeneralCatchClause
                catch (Exception)
                {
                }
            }
            count = temp;
            foreach (var file in from file in directoryInfo.GetFiles()
                let extension = Path.GetExtension(file.FullName)
                where extension == ".doc" || extension == ".docx"
                select file)
            {
                count = count + 1;
                directoryNode.Nodes.Add(new TreeNode(file.FullName));
                docfilefound = true;
                wordpath.Add(file.FullName); 
            }

            return directoryNode;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            _checkedNodes = new List<TreeNode>();
            OpenCheckedNodes(treeView1.Nodes);
            OpenFiles();
        }

        private void DocSearch_Load(object sender, EventArgs e)
        {
            label3.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {  _wordfilepathholder.Clear();
            folderBrowserDialog1.Description =
                @"Select the directory that you wish to use as the default search directory";
            folderBrowserDialog1.ShowNewFolderButton = false;

            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result != DialogResult.OK) return;
            _folderName = folderBrowserDialog1.SelectedPath;
            _docfilespicker.RunWorkerAsync();
        }

        private void FileDialogueWork()
        {
            int count;

            SetLabelVisible();
            ListDirectory(_folderName, _wordfilepathholder, out count);

            SetLabelForFiles(count.ToString(CultureInfo.InvariantCulture));
        }

        private void AddNodesToTreeView(TreeNode tree)
        {
            if (treeView1.InvokeRequired)
            {
                var d = new TreeViewAdd(AddNodesToTreeView);
                Invoke(d, new object[] {tree});
            }
            else
            {
                treeView1.Nodes.Clear();
                treeView1.Nodes.Add(tree);
            }
        }


        private void SetLabelVisible()
        {
            if (label3.InvokeRequired)
            {
                var d = new SetLabelVisibleCallBack(SetLabelVisible);
                Invoke(d);
            }
            else
            {
                label3.Visible = true;
            }
        }

        private void SetLabelforWordFileCount(int i)
        {
            if (label4.InvokeRequired)
            {
                var d = new Label4TextCallback(SetLabelforWordFileCount);
                Invoke(d, new object[] {i});
            }
            else
            {
                int x = Int32.Parse(label4.Text);
                label4.Text = (x + i).ToString(CultureInfo.InvariantCulture);
            }
        }

        private void SetLabelForFiles(string text)
        {
            if (label1.InvokeRequired)
            {
                var d = new SetText1TextCallback(SetLabelForFiles);
                Invoke(d, new object[] {text});
            }
            else
            {
                label1.Text = text;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = string.Empty;
            label4.Text = string.Empty;
            label5.Visible = false;
        }

        private void textBox1_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            string[] temp = e.LinkText.Split('#');

            Process.Start(temp[1]);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            _worker.CancelAsync();
        }

        private delegate void Label4TextCallback(int i);

        private delegate void ProgressCallBack(int value);

        private delegate void ProgressVisibleCallBack();

        private delegate void SetLabelVisibleCallBack();

        private delegate void SetText1TextCallback(string text);

        private delegate void SetTextCallback(RichTextBoxEx text, Textboxexhelper helper);

        private delegate void TreeViewAdd(TreeNode x);
    }
}