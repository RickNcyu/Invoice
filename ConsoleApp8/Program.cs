using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using ExcelDataReader;
using System.Text.RegularExpressions;

namespace ConsoleApp8
{  
    class Program
    {   
        static void Main(string[] args)
        {
            DataSet result;
            //

            string fileName = @"./110全年各站配號.xlsx";
            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);

            /////////////////////////////////////////////////////////////

            /////////////////////////////////////////////////////////////
            string path = @"D:/各站發票";
            System.IO.Directory.CreateDirectory(path);
            
            /////////////////////////////////////////////////////////////


            //xlsx讀取
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fs))
            {
                //
   
                //忽略第一行資料 exceldatasetconfiguration
                result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                int Colcount = result.Tables[0].Columns.Count;
                string[] name =  new string[] {"KD","LY","NT","QN","SH","UC"};
                int Rowcount = result.Tables[0].Rows.Count;
                
                for (int i = 0; i < Rowcount-1; i++)
                {
                    string filename;
                    filename = @"D:/各站發票/" + result.Tables[0].Rows[i][0].ToString();
                    System.IO.Directory.CreateDirectory(filename); ;

                    string output,outputtitle;
                    string outfileName = filename+ @"/invnoapply.csv";
                    FileStream fs2 = new FileStream(outfileName, FileMode.Create, FileAccess.Write);
                    using (StreamWriter sw = new StreamWriter(fs2, Encoding.Default))
                    {
                        outputtitle = "營業人統編"+","+"發票類別代號"+","+"發票類別"+","+"發票期別" + "," +"發票字軌名稱"+"," + "發票起號"+"," + "發票迄號";
                        sw.WriteLine(outputtitle);

                        int count = 1;
                        for (int q = 0; q < 6; q++)
                        {
                            output = result.Tables[0].Rows[i][1].ToString() + "," + "7" + "," + "一般稅額計算" + "," + "110/" + count.ToString("00")+ " ~ "+"110/" + ((count+1).ToString("00")) + ","+name[q]+","+ result.Tables[0].Rows[i][5].ToString()+","+ result.Tables[0].Rows[i][6].ToString();
                            sw.WriteLine(output);
                            count=count+2;
                        }   
                    }

                }
            }
            
        }
    }
}
