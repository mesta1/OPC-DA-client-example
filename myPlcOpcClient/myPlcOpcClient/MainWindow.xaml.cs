using System.Collections.Generic;
using System.Windows;
using System;
using Opc.Da;
using System.Windows.Threading;
using System.Windows.Media;
using System.Threading;

namespace myPlcOpcClient
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        #region private fields

        private bool motorActive;
        private int motorSpeed;
        private bool autManSwitch;

        DispatcherTimer tmr = new DispatcherTimer();

        #endregion

        #region OPC private fields
                
        private Server server;
        private OpcCom.Factory fact = new OpcCom.Factory();

        private Subscription groupRead;        
        private SubscriptionState groupState;

        private Subscription groupWrite;
        private SubscriptionState groupStateWrite;

        private List<Item> itemsList = new List<Item>();

        #endregion

        #region Constructor

        public MainWindow()
        {
            InitializeComponent();

            tmr.Interval = TimeSpan.FromMilliseconds(100);
            tmr.Tick += new EventHandler(tmr_Tick);
            tmr.Start();

            ConnectToOpcServer();
        }

        #endregion

        #region GUI Update

        void tmr_Tick(object sender, EventArgs e)
        {
            if (motorActive == true)
                ledMotor.Fill = new SolidColorBrush(Colors.Green);
            else
                ledMotor.Fill = new SolidColorBrush(Colors.Gray);

            lblSpeed.Content = motorSpeed;
            btnAutMan.IsChecked = autManSwitch;
        }

        #endregion

        #region OPC Connection and Data Updated callback

        private void ConnectToOpcServer()
        {
            // 1st: Create a server object and connect to the RSLinx OPC Server
            try
            {
                server = new Opc.Da.Server(fact, null);
                server.Url = new Opc.URL("opcda://localhost/RSLinx OPC Server");

                //2nd: Connect to the created server
                server.Connect();

                //Read group subscription
                groupState = new Opc.Da.SubscriptionState();
                groupState.Name = "myReadGroup";
                groupState.UpdateRate = 200;
                groupState.Active = true;
                //Read group creation
                groupRead = (Opc.Da.Subscription)server.CreateSubscription(groupState);
                groupRead.DataChanged += new Opc.Da.DataChangedEventHandler(groupRead_DataChanged);

                Item item = new Item();
                item.ItemName = "[MYPLC]N7:0";
                itemsList.Add(item);

                item = new Item();
                item.ItemName = "[MYPLC]O:0/0";
                itemsList.Add(item);

                item = new Item();
                item.ItemName = "[MYPLC]B3:0/3";
                itemsList.Add(item);

                groupRead.AddItems(itemsList.ToArray());

                groupStateWrite = new Opc.Da.SubscriptionState();
                groupStateWrite.Name = "myWriteGroup";
                groupStateWrite.Active = false;
                groupWrite = (Opc.Da.Subscription)server.CreateSubscription(groupStateWrite);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }        

        void groupRead_DataChanged(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            foreach (ItemValueResult itemValue in values) 
            {
                switch (itemValue.ItemName) 
                {
                    case "[MYPLC]N7:0":
                        motorSpeed = Convert.ToInt32(itemValue.Value);
                        break;

                    case "[MYPLC]O:0/0":
                        motorActive = Convert.ToBoolean(itemValue.Value);
                        break;

                    case "[MYPLC]B3:0/3":
                        autManSwitch = Convert.ToBoolean(itemValue.Value);
                        break;
                }
            }
        }

        #endregion       

        #region Write Methods  

        private void WriteData(string itemName, int value)
        {
            groupWrite.RemoveItems(groupWrite.Items);
            List<Item> writeList = new List<Item>();
            List<ItemValue> valueList = new List<ItemValue>();

            Item itemToWrite = new Item();            
            itemToWrite.ItemName = itemName;
            ItemValue itemValue = new ItemValue(itemToWrite);
            itemValue.Value = value;

            writeList.Add(itemToWrite);
            valueList.Add(itemValue);
            //IMPORTANT:
            //#1: assign the item to the group so the items gets a ServerHandle
            groupWrite.AddItems(writeList.ToArray());
            // #2: assign the server handle to the ItemValue
            for (int i = 0; i < valueList.Count; i++ )
                valueList[i].ServerHandle = groupWrite.Items[i].ServerHandle;
            // #3: write
            groupWrite.Write(valueList.ToArray());
        }        


        private const int ON = 1;
        private const int OFF = 0;
        private void WritePushButton(string itemName)
        {
            groupWrite.RemoveItems(groupWrite.Items);
            List<Item> writeList = new List<Item>();
            List<ItemValue> valueList = new List<ItemValue>();

            Item itemToWrite = new Item();
            itemToWrite.ItemName = itemName;
            ItemValue itemValue = new ItemValue(itemToWrite);
            itemValue.Value = ON;

            writeList.Add(itemToWrite);
            valueList.Add(itemValue);
            //IMPORTANT:
            //#1: assign the item to the group so the items gets a ServerHandle
            groupWrite.AddItems(writeList.ToArray());
            // #2: assign the server handle to the ItemValue
            for (int i = 0; i < groupWrite.Items.Length; i++)
                valueList[i].ServerHandle = groupWrite.Items[i].ServerHandle;
            // #3: now write
            groupWrite.Write(valueList.ToArray());

            Thread.Sleep(200);

            itemValue.Value = OFF;

            writeList.Add(itemToWrite);
            valueList.Add(itemValue);
            //IMPORTANT:
            //#1: assign the item to the group so the items gets a ServerHandle
            groupWrite.AddItems(writeList.ToArray());
            // #2: assign the server handle to the ItemValue
            for (int i = 0; i < valueList.Count; i++)
                valueList[i].ServerHandle = groupWrite.Items[i].ServerHandle;
            // #3: now write
            groupWrite.Write(valueList.ToArray());
        }

        #endregion

        #region Callbacks

        private void btnAutMan_Click(object sender, RoutedEventArgs e)
        {
            int value = Convert.ToInt32(!btnAutMan.IsChecked);
            WriteData("[MYPLC]B3:0/3", value);
            e.Handled = true;
        }

        private void btnStartMotor_Click(object sender, RoutedEventArgs e)
        {
            WritePushButton("[MYPLC]B3:0/0");           
        }        

        private void btnStoptMotor_Click(object sender, RoutedEventArgs e)
        {
            WritePushButton("[MYPLC]B3:0/1");
        }

        private void Button_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            WriteData("[MYPLC]B3:0/4", ON);
        }

        private void Button_PreviewMouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            WriteData("[MYPLC]B3:0/4", OFF);
        }

        #endregion

    }
}
