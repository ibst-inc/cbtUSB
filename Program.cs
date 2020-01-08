using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Management; // need to add System.Management to your project references.
using System.Runtime.InteropServices;
using System.IO;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using System.Data;
using System.Data.SqlClient;

namespace cbtUSB
{


    class Program
    {


        static int Main(string[] args)
        {
            String USBDRV = "";
            String strConn, sbtdir;
            MySqlConnection conn;
            int argsCount = args.Length;
            DBHelper db = new DBHelper();
            AppSettingsReader ar = new AppSettingsReader();
            String Server = (String)ar.GetValue("DBServerIP", typeof(String));
            String DB = (String)ar.GetValue("dbname", typeof(String));
            strConn = "Server="+Server+";Database="+DB+";Uid=root;Pwd=apmsetup;CharSet=utf8;";

            conn = new MySqlConnection(strConn);
            conn.Open();
            DriveInfo[] allDrives = DriveInfo.GetDrives();
            //    var drives = DriveInfo.GetDrives()
            //   .Where(drive => drive.IsReady && drive.DriveType == DriveType.Removable);
            foreach (DriveInfo d in allDrives)
            {
                if ((d.IsReady == true) && (d.DriveType == DriveType.Removable))
                {
                    USBDRV = d.Name;

                }
            }
            if (args[0].Equals("export"))
            {
                DataSet ds = new DataSet();
                DataSet ds2 = new DataSet();
                DataSet ds3 = new DataSet();

                string sql = "select * from sr_evaluate_tool_item where sei_id = '" + args[1] + "'  order by sti_cate_name, sti_order ASC, ABS(id) ASC";
 
                sbtdir = (String)ar.GetValue("sbtdir", typeof(String)) + "\\mv31\\toolbox\\data\\";
                MySqlDataAdapter adpt = new MySqlDataAdapter(sql, conn);
                adpt.Fill(ds, "sr_evaluate_tool_item");
                String JSONresult, JSONresult2, JSONresult3;
                JSONresult = JsonConvert.SerializeObject(ds.Tables[0]);
                sql = "SELECT * FROM sr_evaluate_plan where sei_id = " + args[1];
                adpt = new MySqlDataAdapter(sql, conn);
                adpt.Fill(ds2, "sr_evaluate_plan");
                JSONresult2 = JsonConvert.SerializeObject(ds2.Tables[0]);

                sql = String.Format(@"SELECT * FROM toolbox.sr_evaluate_tool_item_result
                                        where sei_id = {0}" 
                    , args[1]);
                adpt = new MySqlDataAdapter(sql, conn);
                adpt.Fill(ds3, "sr_evaluate_tool_item_result");
                JSONresult3 = JsonConvert.SerializeObject(ds3.Tables[0]);
                //json export to file
                System.IO.File.WriteAllText(USBDRV + "export.json", JSONresult);
                System.IO.File.WriteAllText(USBDRV + "export_header.json", JSONresult2);
                System.IO.File.WriteAllText(USBDRV + "export_result.json", JSONresult3);
                System.IO.Directory.CreateDirectory(USBDRV + "resdata");

                if (ds.Tables.Count > 0)
                {
                    foreach (DataRow r in ds.Tables[0].Rows)
                    {
                        if (!String.IsNullOrEmpty(r["bf_file"].ToString()))

                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["bf_file"].ToString(), USBDRV + "resdata\\" + r["bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["bf_file2"].ToString()))

                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["bf_file2"].ToString(), USBDRV + "resdata\\" + r["bf_file2"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti1_bf_file"].ToString()))

                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti1_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti1_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti2_bf_file"].ToString()))
                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti2_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti2_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti_3_bf_file"].ToString()))
                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti_3_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti_3_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti4_bf_file"].ToString()))

                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti4_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti4_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti5_bf_file"].ToString()))
                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti5_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti5_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                        if (!String.IsNullOrEmpty(r["sti6_bf_file"].ToString()))
                        {
                            try
                            {
                                System.IO.File.Copy(sbtdir + r["sti6_bf_file"].ToString(), USBDRV + "resdata\\" + r["sti6_bf_file"].ToString(), true);
                            }
                            catch
                            {
                            }
                        }
                    }
                }
                //화일 카피시작 
                /*  foreach (var usbDevice in usbDevices)
                  {
                      Console.WriteLine("Device ID: {0}, PNP Device ID: {1}, Description: {2}",

                         usbDevice.DeviceID, usbDevice.PnpDeviceID, usbDevice.Description);


                  }*/
            }
            else
            {
                string text = System.IO.File.ReadAllText(USBDRV + "export.json");
                string text2 = System.IO.File.ReadAllText(USBDRV + "export_header.json");
                //connsql = new SqlConnection(strConn);
                int msei_id;
                DataTable dt = JsonConvert.DeserializeObject<DataTable>(text);
                DataTable dt2 = JsonConvert.DeserializeObject<DataTable>(text2);
                //String Qstr= DBHelper.BulkInsert(ref dt, "sr_evaluate_tool_item");
                ar = new AppSettingsReader();
                sbtdir = (String)ar.GetValue("sbtdir", typeof(String)) + "\\mv31\\toolbox\\data\\";
                //connsql.Open();
                MySqlCommand scom2 = new MySqlCommand();
                scom2.Connection = conn;
                scom2.CommandText = " select max(sei_id) from sr_evaluate_plan ";

                msei_id = (Int32)scom2.ExecuteScalar();
                msei_id++;
                foreach (DataRow r in dt2.Rows)
                {
                    int rows;
                    MySqlCommand scom = new MySqlCommand();
                    scom.Connection = conn;
                    scom.CommandText = @"INSERT INTO `toolbox`.`sr_evaluate_plan`
                                        (`sei_id`,
                                        `sei_major`,
                                        `set_id`,
                                        `sei_gubun`,
                                        `sei_term`,
                                        `sei_start`,
                                        `sei_end`,
                                        `sei_place`,
                                        `sei_evaluator`,
                                        `sei_students`,
                                        `sei_status`,
                                        `sei_dur`)
                                        VALUES
                                        (@sei_id,
                                        @sei_major,
                                        @set_id,
                                        @sei_gubun,
                                        @sei_term,
                                        @sei_start,
                                        @sei_end,
                                        @sei_place,
                                        @sei_evaluator,
                                        @sei_students,
                                        @sei_status,
                                        @sei_dur);";
                    MySqlParameter pparam1 = new MySqlParameter();
                    pparam1.ParameterName = "@sei_id";
                    pparam1.Value = msei_id;
                    scom.Parameters.Add(pparam1);
                    MySqlParameter pparam2 = new MySqlParameter();
                    pparam2.ParameterName = "@sei_major";
                    pparam2.Value = r["sei_major"].ToString();
                    scom.Parameters.Add(pparam2);
                    MySqlParameter pparam3 = new MySqlParameter();
                    pparam3.ParameterName = "@set_id";
                    pparam3.Value = r["set_id"].ToString();
                    scom.Parameters.Add(pparam3);
                    MySqlParameter pparam4 = new MySqlParameter();
                    pparam4.ParameterName = "@sei_gubun";
                    pparam4.Value = r["sei_gubun"].ToString();
                    scom.Parameters.Add(pparam4);
                    MySqlParameter pparam5 = new MySqlParameter();
                    pparam5.ParameterName = "@sei_term";
                    pparam5.Value = r["sei_term"].ToString();
                    scom.Parameters.Add(pparam5);
                    MySqlParameter pparam6 = new MySqlParameter();
                    pparam6.ParameterName = "@sei_start";
                    pparam6.Value = r["sei_start"].ToString();
                    scom.Parameters.Add(pparam6);
                    MySqlParameter pparam7 = new MySqlParameter();
                    pparam7.ParameterName = "@sei_end";
                    pparam7.Value = r["sei_end"].ToString();
                    scom.Parameters.Add(pparam7);
                    MySqlParameter pparam8 = new MySqlParameter();
                    pparam8.ParameterName = "@sei_place";
                    pparam8.Value = r["sei_place"].ToString();
                    scom.Parameters.Add(pparam8);
                    MySqlParameter pparam9 = new MySqlParameter();
                    pparam9.ParameterName = "@sei_evaluator";
                    pparam9.Value = r["sei_evaluator"].ToString();
                    scom.Parameters.Add(pparam9);
                    MySqlParameter pparam10 = new MySqlParameter();
                    pparam10.ParameterName = "@sei_students";
                    pparam10.Value = r["sei_students"].ToString();
                    scom.Parameters.Add(pparam10);
                    MySqlParameter pparam11 = new MySqlParameter();
                    pparam11.ParameterName = "@sei_status";
                    pparam11.Value = r["sei_status"].ToString();
                    scom.Parameters.Add(pparam11);
                    MySqlParameter pparam12 = new MySqlParameter();
                    pparam12.ParameterName = "@sei_dur";
                    pparam12.Value = r["sei_dur"].ToString();
                    scom.Parameters.Add(pparam12);
                    rows = scom.ExecuteNonQuery();
                }
                foreach (DataRow r in dt.Rows)
                {
                    int rows;
                    MySqlCommand scom = new MySqlCommand();
                    scom.Connection = conn;
                    scom.CommandText = @"INSERT INTO `sr_evaluate_tool_item` 
                        ( `sti_type` ,
                        `sei_id` ,
                        `set_id` ,
                        `sti_name` ,
                        `sti_content` ,
                        `sti_file1` ,
                        `bf_file` ,
                        `sti_po` ,
                        `sti_score` ,
                        `sti_column` ,
                        `sti_cate_name` ,
                        `sti_1` ,
                        `sti_2` ,
                        `sti_3` ,
                        `sti_4` ,
                        `sti_5` ,
                        `sti_6` ,
                        `sti_7` ,
                        `sti_8` ,
                        `sti_9` ,
                        `sti_10` ,
                        `sti_answer` ,
                        `sti_test_type` ,
                        `sti_file2` ,
                        `bf_file2` ,
                        `sti1_file` ,
                        `sti1_bf_file` ,
                        `sti2_file` ,
                        `sti2_bf_file` ,
                        `sti3_file` ,
                        `sti_3_bf_file` ,
                        `sti4_file` ,
                        `sti4_bf_file` ,
                        `sti5_file` ,
                        `sti5_bf_file` ,
                        `sti6_file` ,
                        `sti6_bf_file` ,
                        `sti_order`
                        ) 
                        VALUES (
                         @sti_type ,
                        @sei_id ,
                        @set_id ,
                        @sti_name ,
                        @sti_content ,
                        @sti_file1 ,
                        @bf_file ,
                        @sti_po ,
                        @sti_score ,
                        @sti_column ,
                        @sti_cate_name ,
                        @sti_1 ,
                        @sti_2 ,
                        @sti_3 ,
                        @sti_4 ,
                        @sti_5 ,
                        @sti_6 ,
                        @sti_7 ,
                        @sti_8 ,
                        @sti_9 ,
                        @sti_10 ,
                        @sti_answer ,
                        @sti_test_type ,
                        @sti_file2 ,
                        @bf_file2 ,
                        @sti1_file ,
                        @sti1_bf_file ,
                        @sti2_file ,
                        @sti2_bf_file ,
                        @sti3_file ,
                        @sti_3_bf_file ,
                        @sti4_file ,
                        @sti4_bf_file ,
                        @sti5_file ,
                        @sti5_bf_file ,
                        @sti6_file ,
                        @sti6_bf_file ,
                        @sti_order)";
                    MySqlParameter pparam1 = new MySqlParameter();
                    pparam1.ParameterName = "@sti_type";
                    pparam1.Value = r["sti_type"].ToString();
                    scom.Parameters.Add(pparam1);
                    MySqlParameter pparam2 = new MySqlParameter();
                    pparam2.ParameterName = "@sei_id";
                    pparam2.Value = msei_id;
                    scom.Parameters.Add(pparam2);
                    MySqlParameter pparam3 = new MySqlParameter();
                    pparam3.ParameterName = "@set_id";
                    pparam3.Value = r["set_id"].ToString();
                    scom.Parameters.Add(pparam3);
                    MySqlParameter pparam4 = new MySqlParameter();
                    pparam4.ParameterName = "@sti_name";
                    pparam4.Value = r["sti_name"].ToString();
                    scom.Parameters.Add(pparam4);
                    MySqlParameter pparam5 = new MySqlParameter();
                    pparam5.ParameterName = "@sti_content";
                    pparam5.Value = r["sti_content"].ToString();
                    scom.Parameters.Add(pparam5);
                    MySqlParameter pparam6 = new MySqlParameter();
                    pparam6.ParameterName = "@sti_file1";
                    pparam6.Value = r["sti_file1"].ToString();
                    scom.Parameters.Add(pparam6);
                    MySqlParameter pparam7 = new MySqlParameter();
                    pparam7.ParameterName = "@bf_file";
                    pparam7.Value = r["bf_file"].ToString();
                    scom.Parameters.Add(pparam7);
                    MySqlParameter pparam8 = new MySqlParameter();
                    pparam8.ParameterName = "@sti_po";
                    pparam8.Value = r["sti_po"].ToString();
                    scom.Parameters.Add(pparam8);
                    MySqlParameter pparam9 = new MySqlParameter();
                    pparam9.ParameterName = "@sti_score";
                    pparam9.Value = r["sti_score"].ToString();
                    scom.Parameters.Add(pparam9);
                    MySqlParameter pparam10 = new MySqlParameter();
                    pparam10.ParameterName = "@sti_column";
                    pparam10.Value = r["sti_column"].ToString();
                    scom.Parameters.Add(pparam10);
                    MySqlParameter pparam11 = new MySqlParameter();
                    pparam11.ParameterName = "@sti_cate_name";
                    pparam11.Value = r["sti_cate_name"].ToString();
                    scom.Parameters.Add(pparam11);
                    MySqlParameter pparam12 = new MySqlParameter();
                    pparam12.ParameterName = "@sti_1";
                    pparam12.Value = r["sti_1"].ToString();
                    scom.Parameters.Add(pparam12);
                    MySqlParameter pparam13 = new MySqlParameter();
                    pparam13.ParameterName = "@sti_2";
                    pparam13.Value = r["sti_2"].ToString();
                    scom.Parameters.Add(pparam13);
                    MySqlParameter pparam14 = new MySqlParameter();
                    pparam14.ParameterName = "@sti_3";
                    pparam14.Value = r["sti_3"].ToString();
                    scom.Parameters.Add(pparam14);
                    MySqlParameter pparam15 = new MySqlParameter();
                    pparam15.ParameterName = "@sti_4";
                    pparam15.Value = r["sti_4"].ToString();
                    scom.Parameters.Add(pparam15);
                    MySqlParameter pparam16 = new MySqlParameter();
                    pparam16.ParameterName = "@sti_5";
                    pparam16.Value = r["sti_5"].ToString();
                    scom.Parameters.Add(pparam16);
                    MySqlParameter pparam17 = new MySqlParameter();
                    pparam17.ParameterName = "@sti_6";
                    pparam17.Value = r["sti_6"].ToString();
                    scom.Parameters.Add(pparam17);
                    MySqlParameter pparam18 = new MySqlParameter();
                    pparam18.ParameterName = "@sti_7";
                    pparam18.Value = r["sti_7"].ToString();
                    scom.Parameters.Add(pparam18);
                    MySqlParameter pparam19 = new MySqlParameter();
                    pparam19.ParameterName = "@sti_8";
                    pparam19.Value = r["sti_8"].ToString();
                    scom.Parameters.Add(pparam19);
                    MySqlParameter pparam20 = new MySqlParameter();
                    pparam20.ParameterName = "@sti_9";
                    pparam20.Value = r["sti_9"].ToString();
                    scom.Parameters.Add(pparam20);
                    MySqlParameter pparam21 = new MySqlParameter();
                    pparam21.ParameterName = "@sti_10";
                    pparam21.Value = r["sti_10"].ToString();
                    scom.Parameters.Add(pparam21);
                    MySqlParameter pparam22 = new MySqlParameter();
                    pparam22.ParameterName = "@sti_answer";
                    pparam22.Value = r["sti_answer"].ToString();
                    scom.Parameters.Add(pparam22);
                    MySqlParameter pparam23 = new MySqlParameter();
                    pparam23.ParameterName = "@sti_test_type";
                    pparam23.Value = r["sti_test_type"].ToString();
                    scom.Parameters.Add(pparam23);
                    MySqlParameter pparam24 = new MySqlParameter();
                    pparam24.ParameterName = "@sti_file2";
                    pparam24.Value = r["sti_file2"].ToString();
                    scom.Parameters.Add(pparam24);
                    MySqlParameter pparam25 = new MySqlParameter();
                    pparam25.ParameterName = "@bf_file2";
                    pparam25.Value = r["bf_file2"].ToString();
                    scom.Parameters.Add(pparam25);
                    MySqlParameter pparam26 = new MySqlParameter();
                    pparam26.ParameterName = "@sti1_file";
                    pparam26.Value = r["sti1_file"].ToString();
                    scom.Parameters.Add(pparam26);
                    MySqlParameter pparam27 = new MySqlParameter();
                    pparam27.ParameterName = "@sti1_bf_file";
                    pparam27.Value = r["sti1_bf_file"].ToString();
                    scom.Parameters.Add(pparam27);
                    MySqlParameter pparam28 = new MySqlParameter();
                    pparam28.ParameterName = "@sti2_file";
                    pparam28.Value = r["sti2_file"].ToString();
                    scom.Parameters.Add(pparam28);
                    MySqlParameter pparam29 = new MySqlParameter();
                    pparam29.ParameterName = "@sti2_bf_file";
                    pparam29.Value = r["sti2_bf_file"].ToString();
                    scom.Parameters.Add(pparam29);
                    MySqlParameter pparam30 = new MySqlParameter();
                    pparam30.ParameterName = "@sti3_file";
                    pparam30.Value = r["sti3_file"].ToString();
                    scom.Parameters.Add(pparam30);
                    MySqlParameter pparam31 = new MySqlParameter();
                    pparam31.ParameterName = "@sti_3_bf_file";
                    pparam31.Value = r["sti_3_bf_file"].ToString();
                    scom.Parameters.Add(pparam31);
                    MySqlParameter pparam32 = new MySqlParameter();
                    pparam32.ParameterName = "@sti4_file";
                    pparam32.Value = r["sti4_file"].ToString();
                    scom.Parameters.Add(pparam32);
                    MySqlParameter pparam33 = new MySqlParameter();
                    pparam33.ParameterName = "@sti4_bf_file";
                    pparam33.Value = r["sti4_bf_file"].ToString();
                    scom.Parameters.Add(pparam33);
                    MySqlParameter pparam34 = new MySqlParameter();
                    pparam34.ParameterName = "@sti5_file";
                    pparam34.Value = r["sti5_file"].ToString();
                    scom.Parameters.Add(pparam34);
                    MySqlParameter pparam35 = new MySqlParameter();
                    pparam35.ParameterName = "@sti5_bf_file";
                    pparam35.Value = r["sti5_bf_file"].ToString();
                    scom.Parameters.Add(pparam35);
                    MySqlParameter pparam36 = new MySqlParameter();
                    pparam36.ParameterName = "@sti6_file";
                    pparam36.Value = r["sti6_file"].ToString();
                    scom.Parameters.Add(pparam36);
                    MySqlParameter pparam37 = new MySqlParameter();
                    pparam37.ParameterName = "@sti6_bf_file";
                    pparam37.Value = r["sti6_bf_file"].ToString();
                    scom.Parameters.Add(pparam37);
                    MySqlParameter pparam38 = new MySqlParameter();
                    pparam38.ParameterName = "@sti_order";
                    pparam38.Value = r["sti_order"].ToString();
                    scom.Parameters.Add(pparam38);

                    rows = scom.ExecuteNonQuery();
                    if (!String.IsNullOrEmpty(r["bf_file"].ToString()))

                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["bf_file"].ToString(), sbtdir + r["bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["bf_file2"].ToString()))

                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["bf_file2"].ToString(), sbtdir + r["bf_file2"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti1_bf_file"].ToString()))

                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti1_bf_file"].ToString(), sbtdir + r["sti1_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti2_bf_file"].ToString()))
                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti2_bf_file"].ToString(), sbtdir + r["sti2_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti_3_bf_file"].ToString()))
                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti_3_bf_file"].ToString(), sbtdir + r["sti_3_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti4_bf_file"].ToString()))

                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti4_bf_file"].ToString(), sbtdir + r["sti4_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti5_bf_file"].ToString()))
                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti5_bf_file"].ToString(), sbtdir + r["sti5_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }
                    if (!String.IsNullOrEmpty(r["sti6_bf_file"].ToString()))
                    {
                        try
                        {
                            System.IO.File.Copy(USBDRV + "resdata\\" + r["sti6_bf_file"].ToString(), sbtdir + r["sti6_bf_file"].ToString(), true);
                        }
                        catch
                        {
                        }
                    }

                }
            }
            return 0;
        }
        static List<USBDeviceInfo> GetUSBDevices()

        {

            List<USBDeviceInfo> devices = new List<USBDeviceInfo>();



            ManagementObjectCollection collection;

            using (var searcher = new ManagementObjectSearcher(@"Select * From Win32_USBHub"))

                collection = searcher.Get();



            foreach (var device in collection)

            {

                devices.Add(new USBDeviceInfo(

                    (string)device.GetPropertyValue("DeviceID"),

                    (string)device.GetPropertyValue("PNPDeviceID"),

                    (string)device.GetPropertyValue("Description")

                    ));

            }



            collection.Dispose();



            return devices;

        }
    }



    class USBDeviceInfo

    {

        public USBDeviceInfo(string deviceID, string pnpDeivceID, string description)

        {

            this.DeviceID = deviceID;

            this.PnpDeviceID = pnpDeivceID;

            this.Description = description;

        }



        public string DeviceID

        {

            get;

            private set;

        }



        public string PnpDeviceID

        {

            get;

            private set;

        }



        public string Description

        {

            get;

            private set;

        }

    }
    class DBHelper
    {
        /// <summary>
        /// Creates a multivalue insert for MySQL from a given DataTable
        /// </summary>
        /// <param name="table">reference to the Datatable we're building our String on</param>
        /// <param name="table_name">name of the table the insert is created for</param>
        /// <returns>Multivalue insert String</returns>
        public static String BulkInsert(ref DataTable table, String table_name)
        {
            try
            {
                StringBuilder queryBuilder = new StringBuilder();
                DateTime dt;

                queryBuilder.AppendFormat("INSERT INTO `{0}` (", table_name);

                // more than 1 column required and 1 or more rows
                if (table.Columns.Count > 1 && table.Rows.Count > 0)
                {
                    // build all columns
                    queryBuilder.AppendFormat("`{0}`", table.Columns[0].ColumnName);

                    if (table.Columns.Count > 1)
                    {
                        for (int i = 1; i < table.Columns.Count; i++)
                        {
                            queryBuilder.AppendFormat(", `{0}` ", table.Columns[i].ColumnName);
                        }
                    }

                    queryBuilder.AppendFormat(") VALUES (", table_name);

                    // build all values for the first row
                    // escape String & Datetime values!
                    if (table.Columns[0].DataType == typeof(String))
                    {
                        queryBuilder.AppendFormat("'{0}'", MySqlHelper.EscapeString(table.Rows[0][table.Columns[0].ColumnName].ToString()));
                    }
                    else if (table.Columns[0].DataType == typeof(DateTime))
                    {
                        dt = (DateTime)table.Rows[0][table.Columns[0].ColumnName];
                        queryBuilder.AppendFormat("'{0}'", dt.ToString("yyyy-MM-dd HH:mm:ss"));
                    }
                    else if (table.Columns[0].DataType == typeof(Int32))
                    {
                        queryBuilder.AppendFormat("{0}", table.Rows[0].Field<Int32?>(table.Columns[0].ColumnName) ?? 0);
                    }
                    else
                    {
                        queryBuilder.AppendFormat(", {0}", table.Rows[0][table.Columns[0].ColumnName].ToString());
                    }

                    for (int i = 1; i < table.Columns.Count; i++)
                    {
                        // escape String & Datetime values!
                        if (table.Columns[i].DataType == typeof(String))
                        {
                            queryBuilder.AppendFormat(", '{0}'", MySqlHelper.EscapeString(table.Rows[0][table.Columns[i].ColumnName].ToString()));
                        }
                        else if (table.Columns[i].DataType == typeof(DateTime))
                        {
                            dt = (DateTime)table.Rows[0][table.Columns[i].ColumnName];
                            queryBuilder.AppendFormat(", '{0}'", dt.ToString("yyyy-MM-dd HH:mm:ss"));

                        }
                        else if (table.Columns[i].DataType == typeof(Int32))
                        {
                            queryBuilder.AppendFormat(", {0}", table.Rows[0].Field<Int32?>(table.Columns[i].ColumnName) ?? 0);
                        }
                        else
                        {
                            queryBuilder.AppendFormat(", {0}", table.Rows[0][table.Columns[i].ColumnName].ToString());
                        }
                    }

                    queryBuilder.Append(")");
                    queryBuilder.AppendLine();

                    // build all values all remaining rows
                    if (table.Rows.Count > 1)
                    {
                        // iterate over the rows
                        for (int row = 1; row < table.Rows.Count; row++)
                        {
                            // open value block
                            queryBuilder.Append(", (");

                            // escape String & Datetime values!
                            if (table.Columns[0].DataType == typeof(String))
                            {
                                queryBuilder.AppendFormat("'{0}'", MySqlHelper.EscapeString(table.Rows[row][table.Columns[0].ColumnName].ToString()));
                            }
                            else if (table.Columns[0].DataType == typeof(DateTime))
                            {
                                dt = (DateTime)table.Rows[row][table.Columns[0].ColumnName];
                                queryBuilder.AppendFormat("'{0}'", dt.ToString("yyyy-MM-dd HH:mm:ss"));
                            }
                            else if (table.Columns[0].DataType == typeof(Int32))
                            {
                                queryBuilder.AppendFormat("{0}", table.Rows[row].Field<Int32?>(table.Columns[0].ColumnName) ?? 0);
                            }
                            else
                            {
                                queryBuilder.AppendFormat(", {0}", table.Rows[row][table.Columns[0].ColumnName].ToString());
                            }

                            for (int col = 1; col < table.Columns.Count; col++)
                            {
                                // escape String & Datetime values!
                                if (table.Columns[col].DataType == typeof(String))
                                {
                                    queryBuilder.AppendFormat(", '{0}'", MySqlHelper.EscapeString(table.Rows[row][table.Columns[col].ColumnName].ToString()));
                                }
                                else if (table.Columns[col].DataType == typeof(DateTime))
                                {
                                    dt = (DateTime)table.Rows[row][table.Columns[col].ColumnName];
                                    queryBuilder.AppendFormat(", '{0}'", dt.ToString("yyyy-MM-dd HH:mm:ss"));
                                }
                                else if (table.Columns[col].DataType == typeof(Int32))
                                {
                                    queryBuilder.AppendFormat(", {0}", table.Rows[row].Field<Int32?>(table.Columns[col].ColumnName) ?? 0);
                                }
                                else
                                {
                                    queryBuilder.AppendFormat(", {0}", table.Rows[row][table.Columns[col].ColumnName].ToString());
                                }
                            } // end for (int i = 1; i < table.Columns.Count; i++)

                            // close value block
                            queryBuilder.Append(")");
                            queryBuilder.AppendLine();

                        } // end for (int r = 1; r < table.Rows.Count; r++)

                        // sql delimiter =)
                        queryBuilder.Append(";");

                    } // end if (table.Rows.Count > 1)

                    return queryBuilder.ToString();
                }
                else
                {
                    return "";
                } // end if(table.Columns.Count > 1 && table.Rows.Count > 0)
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return "";
            }
        }
    }
}

