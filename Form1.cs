using System;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using Makaretu.Dns;
using MinimalisticTelnet;
using System.Drawing;

namespace Discovery_Exporter
{
    public partial class Form1 : Form
    {
        //Set Listview Groups
        private ListViewGroup cs10pGrp = new ListViewGroup("CS10P", HorizontalAlignment.Left);
        private ListViewGroup cs7pGrp = new ListViewGroup("CS7P", HorizontalAlignment.Left);
        private ListViewGroup cs10Grp = new ListViewGroup("CS10", HorizontalAlignment.Left);
        private ListViewGroup cs7Grp = new ListViewGroup("CS7", HorizontalAlignment.Left);
        private ListViewGroup cs119Grp = new ListViewGroup("CS119", HorizontalAlignment.Left);
        private ListViewGroup cs118Grp = new ListViewGroup("CS118", HorizontalAlignment.Left);
        private ListViewGroup cs10nGrp = new ListViewGroup("CS10N", HorizontalAlignment.Left);
        private ListViewGroup vgtGrp = new ListViewGroup("VGT", HorizontalAlignment.Left);

        //Font Embed
        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        private static extern IntPtr AddFontMemResourceEx(IntPtr pbFont, uint cbFont, IntPtr pdv, [System.Runtime.InteropServices.In] ref uint pcFonts);
        private PrivateFontCollection fonts = new PrivateFontCollection();

        Font myFont;

        public Form1()
        {
            InitializeComponent();
            discoverDevices();
            //Initialize Listview Groups
            listView.BeginUpdate();
            listView.Groups.Add(cs10pGrp);
            listView.Groups.Add(cs7pGrp);
            listView.Groups.Add(cs10Grp);
            listView.Groups.Add(cs7Grp);
            listView.Groups.Add(cs119Grp);
            listView.Groups.Add(cs118Grp);
            listView.Groups.Add(cs10nGrp);
            listView.Groups.Add(vgtGrp);
            listView.EndUpdate();
            //Font Load
            byte[] fontData = Properties.Resources.RobotoLight;
            IntPtr fontPtr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(fontData.Length);
            System.Runtime.InteropServices.Marshal.Copy(fontData, 0, fontPtr, fontData.Length);
            uint dummy = 0;
            fonts.AddMemoryFont(fontPtr, Properties.Resources.RobotoLight.Length);
            AddFontMemResourceEx(fontPtr, (uint)Properties.Resources.RobotoLight.Length, IntPtr.Zero, ref dummy);
            System.Runtime.InteropServices.Marshal.FreeCoTaskMem(fontPtr);
            myFont = new Font(fonts.Families[0], 12.0F);

        }
        //Font Set
        private void Form1_Load_1(object sender, EventArgs e)
        {
            listView.Font = myFont;
            refreshButton.Font = myFont;
            exportButton.Font = myFont;
            logButton.Font = myFont;
            serializeButton.Font = myFont;
            resetButton.Font = myFont;
        }
        //Discovery
        public void discoverDevices()
        {
            var mdns = new MulticastService();
            var sd = new ServiceDiscovery(mdns);
            mdns.NetworkInterfaceDiscovered += (s, e) =>
            {
                //foreach (var nic in e.NetworkInterfaces)
                //{
                    //listView.Items.Add(nic.Name);
                    //listView.Items.Add(nic.Description);
                //}
                // Ask for the name of all services.
                sd.QueryAllServices();
            };

            sd.ServiceDiscovered += (s, serviceName) =>
            {
                // Ask for the name of instances of the service.
                mdns.SendQuery(serviceName, type: DnsType.PTR);
            };

            sd.ServiceInstanceDiscovered += (s, e) =>
            {
                // Ask for the service instance details.
                mdns.SendQuery(e.ServiceInstanceName, type: DnsType.SRV);
            };

            mdns.AnswerReceived += (s, e) =>
            {
                // Is this an answer to a service instance details?
                var servers = e.Message.Answers.OfType<SRVRecord>();
                foreach (var server in servers)
                {
                    // Ask for the host IP addresses.
                    mdns.SendQuery(server.Target, type: DnsType.A);
                }

                // Is this an answer to host addresses?
                var addresses = e.Message.Answers.OfType<AddressRecord>();
                string cs10p = "CS10p-";
                string cs7p = "CS7p-";
                string cs10 = "CS10-";
                string cs7 = "CS7-";
                string cs10n = "CS10n";
                string cs119 = "CS119-";
                string cs118 = "CS118-";
                string vgt = "VGt";
                foreach (var address in addresses)
                {
                    string addressName = (address.Name).ToString();
                    string addressIP = (address.Address).ToString();
                    bool vgtBool = addressName.Contains(vgt);
                    bool cs10pBool = addressName.Contains(cs10p);
                    bool cs7pBool = addressName.Contains(cs7p);
                    bool cs10Bool = addressName.Contains(cs10);
                    bool cs7Bool = addressName.Contains(cs7);
                    bool cs10nBool = addressName.Contains(cs10n);
                    bool cs119Bool = addressName.Contains(cs119);
                    bool cs118Bool = addressName.Contains(cs118);

                    if (vgtBool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, vgtGrp)).SubItems.Add(addressIP);
                    }
                    else if (cs7pBool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs7pGrp)).SubItems.Add(addressIP);
                    }
                    else if (cs10Bool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs10Grp)).SubItems.Add(addressIP);
                    }
                    else if (cs7Bool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs7Grp)).SubItems.Add(addressIP);
                    }
                    else if (cs10nBool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs10nGrp)).SubItems.Add(addressIP);
                    }
                    else if (cs119Bool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs119Grp)).SubItems.Add(addressIP);
                    }
                    else if (cs118Bool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs118Grp)).SubItems.Add(addressIP);
                    }
                    else if (cs10pBool)
                    {
                        listView.Items.Add(new ListViewItem(addressName, cs10pGrp)).SubItems.Add(addressIP);
                    }
                }
            };

            try
            {
                mdns.Start();
            }
            finally
            {
                System.Threading.Thread.Sleep(1000);
                mdns.Stop();
                mdns.Dispose();
                sd.Dispose();
            }
        }
        //Check LED
       private void exportButton_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count != 0)
            {
                try
                {
                    string ip = listView.SelectedItems[0].SubItems[1].Text;
                    string fileName = listView.SelectedItems[0].Text;
                    TelnetConnection tc = new TelnetConnection(ip, 23);
                    System.Threading.Thread.Sleep(1000);
                    if (tc.IsConnected != false)
                    {
                        tc.WriteLine("logo_led on");
                        MessageBox.Show("Verify LED is working!");
                        tc.WriteLine("logo_led off");
                        tc.WriteLine("close");
                    }
                    else
                    {
                        MessageBox.Show("Unable to connect!");
                    }
                }
                catch
                {
                    MessageBox.Show("Unable to connect!");
                }
            }
            else
            {
                MessageBox.Show("Select a device to check LED!");
            }
        }
        //Refresh List
        private void refreshButton_Click(object sender, EventArgs e)
        {
            listView.Items.Clear();
            listView.Refresh();
            discoverDevices();  
        }
        //Check Fans
        private void logButton_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count != 0)
            {
                try
                {
                    string ip = listView.SelectedItems[0].SubItems[1].Text;
                    string fileName = listView.SelectedItems[0].Text;
                    TelnetConnection tc = new TelnetConnection(ip, 23);
                    System.Threading.Thread.Sleep(1000);
                    if (tc.IsConnected != false)
                    {
                        tc.WriteLine("setfanspeed 0 manual");
                        tc.WriteLine("setfanspeed 0 100");
                        MessageBox.Show("Verify both fans are working and not noisy!");
                        tc.WriteLine("setfanspeed 0 auto");
                        tc.WriteLine("close");
                    }
                    else
                    {
                        MessageBox.Show("Unable to connect!");
                    }
                }
                catch 
                {
                    MessageBox.Show("Unable to connect!");
                }
            }
            else
            {
                MessageBox.Show("Select a device to check fans!");
            }
        }
        //Serialize Window
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(36, 36, 372, 13);
            textBox.SetBounds(36, 70, 200, 20);
            buttonOk.SetBounds(36, 120, 80, 40);
            buttonCancel.SetBounds(150, 120, 80, 40);

            label.AutoSize = true;
            form.ClientSize = new Size(400, 200);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;

            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();

            value = textBox.Text;
            return dialogResult;
        }
        //Serialize
        private void serializeButton_Click(object sender, EventArgs e)
        {
            string value = "";
            if (InputBox("Serialize", "Enter last significant non-zero digits to serialize device", ref value) == DialogResult.OK)
            {
                if (listView.SelectedItems.Count != 0)
                {
                    try
                    {
                        string ip = listView.SelectedItems[0].SubItems[1].Text;
                        string fileName = listView.SelectedItems[0].Text;
                        TelnetConnection tc = new TelnetConnection(ip, 23);
                        System.Threading.Thread.Sleep(1000);
                        if (tc.IsConnected != false)
                        {
                            tc.WriteLine("serialize " + value);
                            MessageBox.Show("Reset device to complete serialization!");
                            tc.WriteLine("close");
                        }
                        else
                        {
                            MessageBox.Show("Unable to connect!");
                        }
                    }
                    catch
                    {
                        MessageBox.Show("Unable to connect!");
                    }
                }
                else
                {
                    MessageBox.Show("Select a device to serialize!");
                }
            }
        }
        //Reset
        private void resetButton_Click(object sender, EventArgs e)
        {
            if (listView.SelectedItems.Count != 0)
            {
                try
                {
                    string ip = listView.SelectedItems[0].SubItems[1].Text;
                    string fileName = listView.SelectedItems[0].Text;
                    TelnetConnection tc = new TelnetConnection(ip, 23);
                    System.Threading.Thread.Sleep(1000);
                    if (tc.IsConnected != false)
                    {
                        tc.WriteLine("NVfactdflt");
                        MessageBox.Show("Device is being reset!");
                        listView.Items.Clear();
                        listView.Refresh();
                    }
                    else
                    {
                        MessageBox.Show("Unable to connect!");
                    }
                }
                catch
                {
                    MessageBox.Show("Unable to connect!");
                }
            }
            else
            {
                MessageBox.Show("Select a device to reset!");
            }
        }

        
    }
}
