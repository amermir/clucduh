using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop;

namespace Clucduh
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Attribute GetNextAttribute(DataTable dt)
        {
            Attribute[] list = new Attribute[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                Attribute attr = new Attribute();
                attr.Name = dt.Columns[i].ColumnName;
                attr.List = new Dictionary<string, int>();
                list[i] = attr;
                foreach (DataRow dr in dt.Rows)
                {
                    string key = dr[dt.Columns[i]].ToString();
                    if (attr.List.ContainsKey(key))
                        attr.List[key]++;
                    else
                        attr.List[key] = 1;
                }
            }
            double min = double.MaxValue;
            Attribute ret = null;
            foreach (var item in list)
            {
                if (item.GetEPen() < min)
                {
                    min = item.GetEPen();
                    ret = item;
                }
            }
            return ret;
        }

     private void BuildTree(DataTable dt, Node node, int stop)
     {
         if ((dt.Columns.Count < 1) || (dt.Rows.Count <= stop))
             return;
         Attribute attr = GetNextAttribute(dt);
         node.Name = attr.Name;
         node.Records = dt;
         node.Items = new Node[attr.GetNV()];
         int i = -1;
         foreach (string key in attr.List.Keys)
         {
             i++;
             node.Items[i] = new Node();
             node.Items[i].Parent = node;
             node.Items[i].ParentValue = key;
             string expr = string.Format("[{0}]='{1}'", attr.Name, key);
             DataTable nextdt = dt.Select(expr).CopyToDataTable();
             nextdt.Columns.Remove(nextdt.Columns[attr.Name]);
             node.Items[i].Records = nextdt;
             BuildTree(nextdt, node.Items[i], stop);
         }
     }

        private void DrawTree(TreeNode tn, Node Tree)
        {
            if (Tree == null) return;
            Node parent = Tree.Parent;
            string text = "";
            if (Tree.Parent == null)
                text = "Total:  ";
            Node prev = Tree;
            while (parent != null)
            {
                text = string.Format("{0} ({1}) -- ", parent.Name, prev.ParentValue) + text;
                parent = parent.Parent;
                prev = prev.Parent;
            }
            if (text.Length > 5)
                text = text.Substring(0, text.Length - 3);
            tn.Text = text + ":    " + Tree.Records.Rows.Count + " Records \r\n";
            if (Tree.Items == null)
                return;
            for (int i = 0; i < Tree.Items.Length; i++)
            {
                TreeNode childnode = new TreeNode();
                tn.Nodes.Add(childnode);
                childnode.Tag = Tree.Items[i];
                DrawTree(childnode, Tree.Items[i]);
            }
        }

        Node root = null;

        private void ProcessFile(string FileName)
        {
            string path = FileName;
            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);
            string sql = @"SELECT * FROM [" + fileName + "]";
            OleDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly + ";Extended Properties=\"Text;HDR=Yes\"");
            OleDbCommand cmd = new OleDbCommand(sql, connection);
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
            DataTable dt = new DataTable();
            dt.Locale = CultureInfo.CurrentCulture;
            adapter.Fill(dt);
            int stop = int.Parse(txtStop.Text);
            if (chkStop.Checked)
                stop = (int)(dt.Rows.Count * stop / 100.0);
            Node tree = new Node();
            root = tree;
            BuildTree(dt, tree, stop);
            TreeNode tn = new TreeNode();
            tn.Tag = root;
            DrawTree(tn, tree);
            treeView1.Nodes.Add(tn);
            return;
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            string file = openFileDialog1.FileName;
            ProcessFile(file);
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TreeNode tn = treeView1.SelectedNode;
            if (tn == null) return;
            if (tn.Tag == null) return;
            Node n = (tn.Tag as Node);
            string filter = "(1 = 1) ";
            List<string> cols = new List<string>();
            while (n.Parent != null)
            {
                cols.Add(n.Parent.Name);
                filter += string.Format("and [{0}] = '{1}' ", n.Parent.Name, n.ParentValue);
                n = n.Parent;
            }
            dataGridView1.DataSource = root.Records.Select(filter).CopyToDataTable();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    foreach (string s in cols)
                    {
                        if (s == root.Records.Columns[i].ColumnName)
                            row.Cells[i].Style.BackColor = Color.LightGray;
                    }
                }
            }
        }

        private void ExportTree(Node Tree, Microsoft.Office.Interop.Excel._Workbook WorkBook)
        {
            if (Tree == null) return;
            Node parent = Tree.Parent;
            string text = "";
            if (Tree.Parent == null)
                text = "Total: ";
            Node prev = Tree;
            string filter = "(1 = 1) ";
            List<string> cols = new List<string>();
            while (parent != null)
            {
                cols.Add(parent.Name);
                filter += string.Format("and {0} = '{1}' ", parent.Name, prev.ParentValue);
                text = string.Format("{0} ({1}) -- ", parent.Name, prev.ParentValue) + text;
                parent = parent.Parent;
                prev = prev.Parent;
            }
            if (text.Length > 5)
                text = text.Substring(0, text.Length - 3);
            DataTable dt = root.Records.Select(filter).CopyToDataTable();
            if (Tree.Items == null)
            {
                Microsoft.Office.Interop.Excel.Worksheet sh = (Microsoft.Office.Interop.Excel.Worksheet)WorkBook.Worksheets.Add();
                sh.Name = "Group " + group++.ToString();
                sh.Cells[1, 1] = text;
                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {
                    sh.Cells[2, i] = dt.Columns[i - 1].ColumnName;
                    foreach (string s in cols)
                    {
                        if (s == dt.Columns[i-1].ColumnName)
                        {
                            char c =(char)( 'A' + i - 1);
                            string range = string.Format("{0}2:{0}{1}", c, dt.Rows.Count + 1);
                            sh.get_Range(range, Missing.Value).Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray;
                        }
                    }
                }
                for (int i = 0; i < dt.Rows.Count - 1; i++)
                    for (int j = 0; j < dt.Columns.Count; j++)
                        sh.Cells[i + 3, j + 1] = dt.Rows[i][j].ToString();
                return;
            }
            for (int i = 0; i < Tree.Items.Length; i++)
            {
                TreeNode childnode = new TreeNode();
                childnode.Tag = Tree.Items[i];
                ExportTree(Tree.Items[i], WorkBook);
            }
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            openFileDialog1.ShowDialog();
            treeView1.ExpandAll();
        }

        int group = 0;

        private void btnExport_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet sh = workbook.Worksheets[1];
            group = 1;
            ExportTree(root, workbook);
            app.DisplayAlerts = false;
            sh.Delete();
            app.DisplayAlerts = true;
            app.Visible = true;
        }

        public class Node
        {
            public Node()
            {

            }

            string name;
            Node[] items;
            Node parent;
            DataTable records;
            string parentValue;

            public string Name { get => name; set => name = value; }
            public Node[] Items { get => items; set => items = value; }
            public Node Parent { get => parent; set => parent = value; }
            public DataTable Records { get => records; set => records = value; }
            public string ParentValue { get => parentValue; set => parentValue = value; }
        }

        public class Attribute
        {
            public Attribute()
            { }

            private Dictionary<string, int> dict = new Dictionary<string, int>();

            public Dictionary<string, int> List { get => dict; set => dict = value; }
            public string Name { get => name; set => name = value; }

            private string name;

            public int GetN()
            {
                int ret = 0;
                foreach (var d in dict)
                {
                    ret += d.Value;
                }
                return ret;
            }

            public int GetNV()
            {
                int ret = dict.Count;
                return ret;
            }

            public double GetCAen()
            {
                double ret = GetN() / (double)GetNV();
                return ret;
            }

            public double GetEPen()
            {
                double ret = 0;
                foreach (var d in dict)
                {
                    ret += Math.Abs(GetCAen() - d.Value);
                }
                return ret;
            }
        }

    }
}


