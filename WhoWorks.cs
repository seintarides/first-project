using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication14
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        class employment {
            int id;
        
        
        }


        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            
            
           

        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
           
            string id;
            if (e.KeyCode == Keys.Enter)
            {
                
                e.Handled = true;
                
                id = textBox1.Text;
                if (id== "221")
                {
                    MessageBox.Show("vrethike o ipallilos");

                }
                textBox1.Text = String.Empty; 
                
            }
            
        }
        
    }
    
}
