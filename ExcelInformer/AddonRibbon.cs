
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Reactive.Linq;
using System.Threading.Tasks;
using ExcelInformer.Properties;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Windows.Forms;
using ExcelInformer.Clients;
using Microsoft.Win32;

namespace ExcelInformer
{

    public partial class AddonRibbon
    {

        RegistryKey key = null;

        private void EiRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");

                Globals.ThisAddIn.Application.ScreenUpdating = true;

                safeSwitch.Checked = bool.Parse(key.GetValueOrDefault("SafeMode", safeSwitch.Checked)+"");
                startRowB.Text =            key.GetValueOrDefault("StartRow", startRowB.Text)+"";
                endRowB.Text =              key.GetValueOrDefault("EndRow", endRowB.Text)+"";
                startColumnB.Text =         key.GetValueOrDefault("StartColumn", startColumnB.Text) + "";
                offsetValue.Text =          key.GetValueOrDefault("OffsetValue", offsetValue.Text) + "";
                offsetTime.Text =           key.GetValueOrDefault("OffsetTime", offsetTime.Text) + "";
                key.Close();
            } catch (Exception ex) { MessageBox.Show(ex + ""); }
        }

        Microsoft.Office.Tools.Excel.Workbook workbook;
        bool scan = false;
        TimeSpan t20s = new TimeSpan(0,0,20);

        #region SCANS

        List<ChannelClient> channelClients;

        private async void Scan(bool onlyOne, bool safeMode, uint startRow = 1, uint endRow = 2000, uint startColumn = 1)
        {
            workbook = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook);

            channelClients = new List<ChannelClient>();
            #region Initializing
            updIndication(1);
            scan = true;

            string data = "data-ei";
            string exampleNID = "opc.tcp://192.168.120.85:4863*ns=1;s=t|OS_OS-A::P1B3_D1/AIB.Pv";
            Worksheet dataSheet = null;
            #endregion Initializing

            #region Create sheet if it doesnt exist
            try
            {
                dataSheet = workbook.Worksheets[data];
            }
            catch (Exception e)
            {
                dataSheet = workbook.Worksheets.Add();
                dataSheet.Name = data;
                dataSheet.Cells[1, 1] = exampleNID;
            }
            #endregion Create sheet if it doesnt exist
            
            Range dataRange = dataSheet.Range[dataSheet.Cells[startRow, startColumn], dataSheet.Cells[endRow, startColumn]];
            while (scan)
            {
                scannedCells.Label = "0";
                label1.Label = "0";
                DateTime beginTime = DateTime.Now;
                scanStart.Label = beginTime + "";
                for (int i = 1; i <= dataRange.Count; i++)
                {
                    try
                    {
                        if (!scan) break;
                        Range dataItem = dataRange.Item[i];
                        scannedCells.Label = int.Parse(scannedCells.Label)+1+"";
                        if (dataItem.Value2 + "" == "" || !(dataItem.Value2 + "").Contains("*")) continue; // skip if empty or incompatible

                        string url = (dataItem.Value2 + "").Split('*')[0];
                        string address = (dataItem.Value2 + "").Split('*')[1];

                        var findClient = channelClients.Find(a => a.Url == url);

                        if (findClient != null && !findClient.Connected())
                        {
                            channelClients.Remove(findClient);
                            findClient = null;
                        }

                        if (findClient == null)
                        { // client doesnt exist
                            await AddClient(url);
                        }

                        // find client again to use as argument
                        findClient = channelClients.Find(a => a.Url == url);

                        if (!safeMode)
                            modifyCell(dataItem, address, findClient);
                        else
                            await modifyCell(dataItem, address, findClient);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                    }
                }

                DateTime endTime = DateTime.Now;
                scanEnd.Label = endTime + "";
                scanTime.Label = (DateTime.Now - beginTime + "").Substring(0, 11);

                try
                {
                    dataRange.Item[1].Offset[0, uint.Parse(offsetTime.Text) + 1].Value2 = DateTime.Now + ""; // last scan time
                } catch(Exception e) { Console.WriteLine(e); }
                if (onlyOne) { break; }

                await Task.Delay(1000); // arbitrary delay
            }
            scan = false; 
            updIndication(0);
        }

        private async Task AddClient(string url)
        {
            IChannelClient newChannelClient = getChannelClientFromUrl(url);
            ChannelClient newClient = new ChannelClient(newChannelClient);
            await newClient.Connect(url);
            channelClients.Add(newClient);
        }

        private IChannelClient getChannelClientFromUrl(string url)
        {
            if (url.Contains("opc.tcp")) return new OPCuaClient();
            if (url.Contains("mongodb")) return new MongoDBClient();
            return null;
        }

        private async Task modifyCell(Range dataItem, string address, ChannelClient channel)
        {
            if (address == null || address == "") return;
            string readResult = "";
            try
            {
                readResult = await channel.Read(address);
            }
            catch (Exception e)
            {
                dataItem.Offset[0, uint.Parse(offsetValue.Text)].Value2 = e;
            }
            var offsetValueParsed = uint.Parse(offsetValue.Text);
            if (dataItem.Offset[0, offsetValueParsed].Value2 + "" != readResult)
            {
                if (!int.TryParse(label1.Label, out _)) label1.Label = "0";
                label1.Label = (int.Parse(label1.Label) + 1) + "";
                dataItem.Offset[0, offsetValueParsed].Value2 = readResult+"";
                dataItem.Offset[0, uint.Parse(offsetTime.Text)].Value2 = DateTime.Now + "";
            }
        }

        #endregion SCANS

        #region UI EVENTS

        private void startBtn_Click(object sender, RibbonControlEventArgs e)
        {
            doScan(false);
        }
        private void stopBtn_Click(object sender, RibbonControlEventArgs e)
        {
            scan = false;
            label1.Label = "0";
        }

        private void updOnceBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                doScan(true);
            } catch(Exception er)
            {
                MessageBox.Show(er.StackTrace);
            }
        }

        private void safeSwitch_Click(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            key.SetValue("SafeMode", safeSwitch.Checked);
            key.Close();
            
        }

        private void offsetTime_TextChanged(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            if (!uint.TryParse(offsetTime.Text, out uint val) || val <= 1) offsetTime.Text = "2";
            key.SetValue("OffsetTime", offsetTime.Text); //
            key.Close();
        }

        private void offsetValue_TextChanged(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            if (!uint.TryParse(offsetValue.Text, out uint val) || val <= 1) offsetValue.Text = "1";
            key.SetValue("OffsetValue", offsetValue.Text);
            key.Close();
        }

        private void startRowB_TextChanged(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            if (!uint.TryParse(startRowB.Text, out uint val) || val <= 1) startRowB.Text = "1";
            key.SetValue("StartRow", startRowB.Text);
            key.Close();
        }

        private void endRowB_TextChanged(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            if (!uint.TryParse(endRowB.Text, out uint val) || val <= 1) endRowB.Text = "2000";
            key.SetValue("EndRow", endRowB.Text);
            key.Close();
        }

        private void startColumnB_TextChanged(object sender, RibbonControlEventArgs e)
        {
            key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\ExcelInformer\defaults");
            if (!uint.TryParse(startColumnB.Text, out uint val) || val <= 1) startColumnB.Text = "1";
            key.SetValue("StartColumn", startColumnB.Text);
            key.Close();
        }

        private void helpButton_Click(object sender, RibbonControlEventArgs e)
        {
            Help helpForm = new Help();
            helpForm.Show();
        }

        #endregion UI EVENTS

        #region GENERAL

        private void doScan(bool onlyOne)
        {
            uint startRow;
            uint endRow;
            uint startColumn;

            if (!uint.TryParse(startRowB.Text, out startRow) || startRow < 1) startRow = 1;
            if (!uint.TryParse(endRowB.Text, out endRow) || endRow < 1) endRow = 2000;
            if (!uint.TryParse(startColumnB.Text, out startColumn) || startColumn < 1 ) startColumn = 1;

            if (endRow < startRow) { endRow = startRow + 1; MessageBox.Show("Invalid end/start row");  };

            if (!scan && !safeSwitch.Checked)
                Scan(onlyOne, false, startRow, endRow, startColumn);
            else
                Scan(onlyOne, true, startRow, endRow, startColumn);
        }

        private void updIndication(byte state)
        {
            switch (state)
            {
                case 2:
                    indicator.Image = Resources.indication_orange;
                    indicator.Label = "Idle";
                    break;
                case 1:
                    indicator.Image = Resources.indication_green;
                    indicator.Label = "Running";
                    break;
                default:
                    indicator.Image = Resources.indication_red;
                    indicator.Label = "Not Running";
                    break;
            }
        }



        #endregion GENERAL
    }
}
