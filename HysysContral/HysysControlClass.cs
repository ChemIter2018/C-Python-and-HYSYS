using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HYSYS;
using Microsoft.VisualBasic;
using System.Resources;
using System.Windows.Forms;

namespace HYSYS_RTO.HysysContral
{
    class HysysControlClass
    {
        public PlantApplication Hysys_Connect()
        {
            dynamic HysysApp = Microsoft.VisualBasic.Interaction.CreateObject("HYSYS.Application");
            //dynamic HysysSimulation = Microsoft.VisualBasic.Interaction.GetObject("C:/ZY/05HYSYS/hysys_optimization_thf.hsc");

            HysysApp.visible = true;
            //HysysSimulation.visible = true;

            return HysysApp;
        }

        public void ReadHysysData(Control.ControlCollection Con)
        {
            PlantApplication HysysApp = Hysys_Connect();
            dynamic HysysSimulationCase = HysysApp.ActiveDocument;
            dynamic FEED = HysysSimulationCase.Flowsheet.MaterialStreams.Item("FEED");
            dynamic Ovhd_Prod = HysysSimulationCase.Flowsheet.MaterialStreams.Item("OVHD_PROD");
            dynamic Bttm_Prod = HysysSimulationCase.Flowsheet.MaterialStreams.Item("BTTM_PROD");
            dynamic T100 = HysysSimulationCase.Flowsheet.Operations.Item("T-100");
            dynamic specTHF = T100.ColumnFlowsheet.Specifications.Item("THF_PURITY_SPEC");
            dynamic specToluene = T100.ColumnFlowsheet.Specifications.Item("TOLUENE_PURITY_SPEC");
            dynamic mainTower = T100.ColumnFlowsheet.Operations.Item("Main Tower");
            dynamic Condenser = T100.ColumnFlowsheet.Operations.Item("Condenser");
            dynamic Reboiler = T100.ColumnFlowsheet.Operations.Item("Reboiler");

            //MessageBox.Show(Ovhd_Prod.TemperatureValue.ToString());

            foreach (Control C in Con)
            {
                if (C.GetType().Name == "TextBox")
                {
                    if (((TextBox)C).Name == "textBox_OvhdProd_Temp")
                        ((TextBox)C).Text = Ovhd_Prod.TemperatureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_OvhdProd_Press")
                        ((TextBox)C).Text = Ovhd_Prod.PressureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_OvhdProd_MassFlow")
                        ((TextBox)C).Text = (Ovhd_Prod.MassFlowValue*3600).ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_OvhdProd_THFCom")
                        ((TextBox)C).Text = Ovhd_Prod.ComponentMassFractionValue[0].ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_OvhdProd_TOLCom")
                        ((TextBox)C).Text = Ovhd_Prod.ComponentMassFractionValue[1].ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_BttmProd_Temp")
                        ((TextBox)C).Text = Bttm_Prod.TemperatureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_BttmProd_Press")
                        ((TextBox)C).Text = Bttm_Prod.PressureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_BttmProd_MassFlow")
                        ((TextBox)C).Text = (Bttm_Prod.MassFlowValue*3600).ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_BttmProd_THFCom")
                        ((TextBox)C).Text = Bttm_Prod.ComponentMassFractionValue[0].ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_BttmProd_TOLCom")
                        ((TextBox)C).Text = Bttm_Prod.ComponentMassFractionValue[1].ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_Feed_InletStage")
                        ((TextBox)C).Text = mainTower.FeedStages.names[0].ToString();
                    if (((TextBox)C).Name == "textBox_T100_Cond_Power")
                        ((TextBox)C).Text = Condenser.HeatFlowValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_Reb_Power")
                        ((TextBox)C).Text = Reboiler.HeatFlowValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_RefluxRatio")
                        ((TextBox)C).Text = T100.ColumnFlowsheet.RefluxRatio.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_TopStageTemp")
                        ((TextBox)C).Text = mainTower.TopStageTemperatureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_TopStagePress")
                        ((TextBox)C).Text = mainTower.TopStagePressureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_BtmStageTemp")
                        ((TextBox)C).Text = mainTower.BottomStageTemperatureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_T100_BtmStagePress")
                        ((TextBox)C).Text = mainTower.BottomStagePressureValue.ToString("#0.0000");
                    if (((TextBox)C).Name == "textBox_NumStages")
                        ((TextBox)C).Text = mainTower.NumberOfStages.ToString();

                }
                    
            }
        }

        public void WriteHysysData(Control.ControlCollection Con)
        {
            PlantApplication HysysApp = Hysys_Connect();
            dynamic HysysSimulationCase = HysysApp.ActiveDocument;
            dynamic FEED = HysysSimulationCase.Flowsheet.MaterialStreams.Item("FEED");
            dynamic Ovhd_Prod = HysysSimulationCase.Flowsheet.MaterialStreams.Item("OVHD_PROD");
            dynamic Bttm_Prod = HysysSimulationCase.Flowsheet.MaterialStreams.Item("BTTM_PROD");
            dynamic T100 = HysysSimulationCase.Flowsheet.Operations.Item("T-100");
            dynamic specTHF = T100.ColumnFlowsheet.Specifications.Item("THF_PURITY_SPEC");
            dynamic specToluene = T100.ColumnFlowsheet.Specifications.Item("TOLUENE_PURITY_SPEC");
            dynamic mainTower = T100.ColumnFlowsheet.Operations.Item("Main Tower");
            dynamic Condenser = T100.ColumnFlowsheet.Operations.Item("Condenser");
            dynamic Reboiler = T100.ColumnFlowsheet.Operations.Item("Reboiler");

            foreach (Control C in Con)
            {
                if (C.GetType().Name == "TextBox")
                {
                    if (((TextBox)C).Name == "textBox_Feed_MassFlow")
                        FEED.MassFlowValue = double.Parse(((TextBox)C).Text)/3600;
                }
            }
        }
    }
}
