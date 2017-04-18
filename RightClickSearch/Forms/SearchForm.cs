using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace RightClickSearch
{
    public partial class SearchForm : Form
    {
        private string _filePath = null;
        private string _folderPath = null;

        public SearchForm(string arg = "")
        {
            InitializeComponent();
            IsFileOrFolder(arg);
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            if (_filePath != null)
            {
                this.label1.Text = "Search in file: " + _filePath;
            }
            else if (_folderPath != null)
            {
                this.label1.Text = "Search in file: " + _folderPath;
            }
            else
            {
                this.label1.Text = "Search entire computer.";
            }

            this.Text = "PowerSearch By MDS v" + Application.ProductVersion;
        }

        private void button1_Click(object sender, System.EventArgs e)
        {

        }

        private void IsFileOrFolder(string path)
        {
            try
            {
                if (File.Exists(path) == true)
                {
                    _filePath = path;
                }
                else if (Directory.Exists(path) == true)
                {
                    _folderPath = path;
                }
            }
            catch
            { }
        }
    }
}
