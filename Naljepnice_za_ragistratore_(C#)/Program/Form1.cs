using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Naljepnice_za_ragistratore
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Glavni i jedini gumb ovog prozora...
        // Otvara drugi prozor koji sadrži obrazac za popunjavanje...
        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            f2.ShowDialog();
        }
    }
}
/*  

 Produced by:
          -- Severin Knežević --  
    Email: knezevicseverin@gmail.com

 */