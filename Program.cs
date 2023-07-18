using ConsoleOpc;
using OPCAutomation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

public class Program
{
    private static List<Item> OPCList=new List<Item>();
    static List<string> ItemIDs = new List<string>();
    static Array ItemID ;
    static Array ClientHandle;
    private static void Main(string[] args)
    {
        string hostName = string.Empty;
        string strHostIP = string.Empty;
        //获取本地计算机IP,计算机名称
        IPHostEntry IPHost = Dns.Resolve(Environment.MachineName);
        if (IPHost.AddressList.Length > 0)
        {
           strHostIP = IPHost.AddressList[0].ToString();
           hostName=Dns.GetHostEntry(strHostIP).HostName;
        }
        else
        {
            return;
        }

        //Console.WriteLine("Hello, World!");
        OPCServer server = new OPCServer();
        //server.Connect("KEPware.KEPServerEx.V6", "127.0.0.1");
        object serverList = server.GetOPCServers(hostName);//获取节点的opc服务
        foreach (object item in (Array)serverList) {
            server.Connect(item.ToString(), strHostIP);//"127.0.0.1"
        }


        if (server.ServerState == (int)OPCServerState.OPCRunning)
        {
            Console.WriteLine("已连接到：{0}", server.ServerName);
        }
        else
        {
            //这里你可以根据返回的状态来自定义显示信息，请查看自动化接口API文档
            Console.WriteLine("状态：{0}", server.ServerState.ToString());
        }

        OPCGroups groups;//组集合
        OPCGroup group;//组
        //OPCItems items;
        //OPCItem item;

        groups = server.OPCGroups;  //拿到组jih
        groups.DefaultGroupIsActive = true; //设置组集合默认为激活状态
        groups.DefaultGroupDeadband = 0;    //设置死区
        groups.DefaultGroupUpdateRate = 200;//设置更新频率

        OPCBrowser opcBrower=  server.CreateBrowser();
        //展开分支
        opcBrower.ShowBranches();
        //展开叶子
        opcBrower.ShowLeafs(true);

        List<string>  lstItem=new List<string>();
        
        foreach (string p in opcBrower)
        {
            lstItem.Add(p.ToString());   //将获取到的变量标签存入控件
        }

        //初始化OPCGroup
        OPCGroups KepGroups = server.OPCGroups;
        KepGroups.DefaultGroupDeadband = 0;   //死区值，设为0时，服务器端该组内任何数据变化都通知组
        KepGroups.DefaultGroupIsActive = true;

        group = KepGroups.Add("test");
        group.IsSubscribed = true; //是否为订阅
        group.UpdateRate = 200;    //刷新频率
     //   group.DataChange += Group_DataChange; //组内数据变化的回调函数 数据变化时，自动触发函数
        group.AsyncReadComplete += Group_AsyncReadComplete;  //异步读取完成回调 
        group.AsyncWriteComplete += Group_AsyncWriteComplete;  //异步写入完成回调
        group.AsyncCancelComplete += Group_AsyncCancelComplete; //异步取消读取、写入回调



          OPCList.Add(new Item()
          {
              ItemID = "通道 1.设备 1.标记 1"
          });
        //模拟器示例.函数.Ramp4
         OPCList.Add(new Item()
         {
             ItemID = "模拟器示例.函数.Ramp4"//此值kepserver不能修改
         });
        
        int count = OPCList.Count;

        List<int> ClientHandles = new List<int>();
        ItemIDs.Add("0");
        ClientHandles.Add(0);

        for (int i = 0; i < count; i++)
        {
            ItemIDs.Add(OPCList[i].ItemID);
            ClientHandles.Add(i+1);

        }
        //集合转换成array
        ItemID = ItemIDs.ToArray();
        ClientHandle = ClientHandles.ToArray();

        Array SeverHandles;
        Array Errors;
        ////Values.SetValue(222,0);
        //Values.SetValue(333,1);
       // Array Rvalues = Values.ToArray();
        int Tid = 0;
        int Cid = 0;


        group.OPCItems.AddItems(count, ref ItemID, ref ClientHandle, out SeverHandles, out Errors);
        
        //读取 一直
        // while (true)
        // {
        //     group.AsyncRead(count, ref SeverHandles,out  Errors, Tid, out Cid);
        // }
       // OPCItem bItem=group.OPCItems.GetOPCItem(int.Parse(SeverHandles.GetValue(2).ToString()));
        int[] temp = new int[2] { 0, int.Parse(SeverHandles.GetValue(1).ToString()) };
        Array serverHandles = (Array)temp;
        
        object[] valueTemp = new object[2] { "", "700" };
        Array values = (Array)valueTemp;
        
        group.AsyncWrite(1,ref serverHandles,ref values,out Errors,Tid,out Cid );//读取   
        group.AsyncRead(count, ref SeverHandles,out  Errors, Tid, out Cid);//写入
       // group.AsyncCancel(Cid);//取消
        Console.ReadKey();
    }

    private static void Group_AsyncCancelComplete(int CancelID)
    {
        Console.WriteLine("AsyncCancel");
    }

    private static void Group_AsyncWriteComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array Errors)
    {
        Console.WriteLine("AsyncWrite");
    }

    private static void Group_AsyncReadComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps, ref Array Errors)
    {
        Console.WriteLine("AsyncReadComplete");
        GetData(NumItems, ClientHandles, ItemValues, Qualities, TimeStamps);
    }

    private static void Group_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
    {
        Console.WriteLine("DataChange");
        //为了测试，所以加了控制台的输出，来查看事物ID号
        //Console.WriteLine("********"+TransactionID.ToString()+"*********");
        GetData(NumItems, ClientHandles, ItemValues, Qualities, TimeStamps);
    }

    private static void GetData(int NumItems,  Array ClientHandles,  Array ItemValues,  Array Qualities,  Array TimeStamps)
    {
        for (int i = 1; i <= NumItems; i++)
        {
            string a = ItemValues.GetValue(i).ToString();
            int clientHandle = Convert.ToInt32(ClientHandles.GetValue(i));
            for (int j = 0;j < OPCList.Count;j++)
            {
                if(j+1==clientHandle)
                {
                    OPCList[j].ItemValue = a;
                    OPCList[j].Quanlity = Qualities.GetValue(i).ToString();
                    OPCList[j].UpdateTime = TimeStamps.GetValue(i).ToString();
                }
            }

            OPCList.ForEach(item =>
            {
                Console.WriteLine($"{item.ItemID}-{item.ItemValue}-{item.Quanlity}-{item.UpdateTime}:");
            });
        
            // string b = Qualities.GetValue(i).ToString();
            // string c = TimeStamps.GetValue(i).ToString();
        }
    }
}