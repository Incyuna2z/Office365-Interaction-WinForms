using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Office365_Interaction_Winforms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            var contactsClient = await O365Help.CreateSharePointClientAsync("MyFiles");
            var contacts = await O365Help.getMyFiles(contactsClient);
            foreach (var contact in contacts)
            {
                var item = new ListViewItem(contact.Name);
                listView1.Items.Add(item);
            }
        }
    }
}
