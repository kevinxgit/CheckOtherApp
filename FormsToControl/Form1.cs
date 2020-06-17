using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormsToControl
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            FillDataGrid();
        }
        
        private void FillDataGrid()
        {
            var list = new List<datasource>();
            list.Add(new datasource(1, "abcde", "john", "smith"));
            list.Add(new datasource(2, "fghi", "pete", "smith"));
            list.Add(new datasource(3, "bbbb", "james", "blue"));
            list.Add(new datasource(4, "dddd", "lou", "ackman"));

            this.dataGridView1.DataSource = list;
        }
    }

    public class datasource
    {
        public datasource(int id, string user, string firstname, string lastname)
        {
            Id = id;
            User = user;
            FirstName = firstname;
            LastName = lastname;
        }

        public int Id { get; set; }

        public string User { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }
    }
}
