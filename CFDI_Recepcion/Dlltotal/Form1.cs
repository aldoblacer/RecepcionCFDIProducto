﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using VerificadorAF;

namespace Dlltotal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Verificador VAF = new Verificador();
            VAF.Proceso();
            MessageBox.Show("Proceso Listo");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Verificador VAF = new Verificador();
            VAF.Correo();
            MessageBox.Show("Listo");
        }
    }
}
