using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HYSYS;
using Microsoft.VisualBasic;
using System.Resources;
using System.Diagnostics;

namespace HYSYS_RTO
{
    public partial class Main_RTO : Form
    {
        HysysContral.HysysControlClass MyHysysControlClass = new HysysContral.HysysControlClass();
        PublicClass.ConsoleClass MyConsoleClass = new PublicClass.ConsoleClass();

        public Main_RTO()
        {
            InitializeComponent();
        }

        private void Main_RTO_Load(object sender, EventArgs e)
        {
            //FEED
            textBox_Feed_Temp.Text = "10";
            textBox_Feed_Press.Text = "140";
            textBox_Feed_MassFlow.Text = "3700";
            textBox_FeedCom_THF.Text = "0.44";
            textBox_FeedCom_Toluene.Text = "0.56";

            //T100            
            textBox_PCond.Text = "103";
            textBox_PReb.Text = "107";
            textBox_Con_DeltP.Text = "0.0000";
            textBox_Reb_DeltP.Text = "0.0000";
            textBox_T100_Specification_THF.Text = "0.9950";
            textBox_T100_Specification_Tol.Text = "0.9400";

            //Open Console
            MyConsoleClass.OpenConsole();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PlantApplication HysysApp = MyHysysControlClass.Hysys_Connect();
            dynamic HysysSimulation = Microsoft.VisualBasic.Interaction.GetObject("C:/ZY/05HYSYS/hysys_optimization_thf.hsc");
            HysysSimulation.visible = true;

            MyHysysControlClass.ReadHysysData(groupBox_FEED.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_OVHD.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_BTTM.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_COND.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_REB.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_T100.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_NumStages.Controls);
        }

        private void button_Close_Click(object sender, EventArgs e)
        {
            Process[] P = Process.GetProcessesByName("AspenHysys");
            P[0].Kill();
        }

        private void Main_RTO_FormClosed(object sender, FormClosedEventArgs e)
        {
            MyConsoleClass.CloseConsole();
        }

        private void button_OnHold_Click(object sender, EventArgs e)
        {
            PlantApplication HysysApp = MyHysysControlClass.Hysys_Connect();
            dynamic HysysSimulation = HysysApp.ActiveDocument;
            HysysSimulation.solver.CanSolve = false;
        }

        private void button_Active_Click(object sender, EventArgs e)
        {
            PlantApplication HysysApp = MyHysysControlClass.Hysys_Connect();
            dynamic HysysSimulation = HysysApp.ActiveDocument;
            HysysSimulation.solver.CanSolve = true;

            MyHysysControlClass.WriteHysysData(groupBox_FEED.Controls);

            MyHysysControlClass.ReadHysysData(groupBox_FEED.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_OVHD.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_BTTM.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_COND.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_REB.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_T100.Controls);
            MyHysysControlClass.ReadHysysData(groupBox_NumStages.Controls);
        }
    }
}
