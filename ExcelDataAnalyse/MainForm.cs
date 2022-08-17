using ExcelDataReader;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDataAnalyse
{
    public partial class MainForm : Form
    {
        private string outFilePath = "";

        public MainForm()
        {
            InitializeComponent();
            //textBox_zongchengji.AutoSize = false;
            //textBox_zongchengji.Height = 35;
            button6.Enabled = false;
            string path = Application.StartupPath + "\\报表模板.xlsx";
            Console.WriteLine(">>>>path:" + path);
            textBox_baobiaomoban.Text = path;

            outFilePath= Application.StartupPath + "\\报表\\";

            if (!Directory.Exists(outFilePath)) {
                Directory.CreateDirectory(outFilePath);
            }
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "请选择总成绩表文件";
            openFileDialog.FileName = "总成绩表.xlsx";
      

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
             
                //得到打开的文件路径（包括文件名）
               String path = openFileDialog.FileName.ToString();
                textBox_zongchengji.Text = path;
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "请选择科目满分表文件";
            openFileDialog.FileName = "科目满分表.xlsx";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                //得到打开的文件路径（包括文件名）
                String path = openFileDialog.FileName.ToString();
                textBox_kemumanfen.Text = path;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "请选择初期人数表文件";
            openFileDialog.FileName = "期初人数表.xlsx";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                //得到打开的文件路径（包括文件名）
                String path = openFileDialog.FileName.ToString();
                textBox_chuqirenshu.Text = path;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "请选择任课教师名单表文件";
            openFileDialog.FileName = "任课教师表.xlsx";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                //得到打开的文件路径（包括文件名）
                String path = openFileDialog.FileName.ToString();
                textBox_renkejiaoshi.Text = path;
            }
        }


        private void button6_Click(object sender, EventArgs e)
        {
            openFileDialog.Title = "请选择报表模板表文件";
            openFileDialog.FileName = "报表模板.xlsx";


            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {

                //得到打开的文件路径（包括文件名）
                String path = openFileDialog.FileName.ToString();
                textBox_baobiaomoban.Text = path;
            }
        }

        private DataTable UpdateDataTable(DataTable argDataTable, String columnName)
        {
            DataTable dtResult = new DataTable();
            //克隆表结构
            dtResult = argDataTable.Clone();
            foreach (DataColumn col in dtResult.Columns)
            {
                if (col.ColumnName == columnName)
                {
                    //修改列类型
                    col.DataType = typeof(Decimal);
                }
            }
            foreach (DataRow row in argDataTable.Rows)
            {
                DataRow newDtRow = dtResult.NewRow();
                foreach (DataColumn column in argDataTable.Columns)
                {
                    if (column.ColumnName == columnName)
                    {
                        newDtRow[column.ColumnName] = Convert.ToDecimal(row[column.ColumnName]);
                    }
                    else
                    {
                        newDtRow[column.ColumnName] = row[column.ColumnName];
                    }
                }
                dtResult.Rows.Add(newDtRow);
            }
            return dtResult;
        }


        private DataTable UpdateDataTableInt(DataTable argDataTable, string[] columnName)
        {
            DataTable dtResult = new DataTable();
            //克隆表结构
            dtResult = argDataTable.Clone();
            foreach (DataColumn col in dtResult.Columns)
            {
                if ( columnName.Contains(col.ColumnName))
                {
                    //修改列类型
                    col.DataType = typeof(Int32);
                }
            }
            foreach (DataRow row in argDataTable.Rows)
            {
                
                DataRow newDtRow = dtResult.NewRow();
                foreach (DataColumn column in argDataTable.Columns)
                {
                    if (  columnName.Contains(column.ColumnName))
                    {
                        try
                        {
                            newDtRow[column.ColumnName] = Convert.ToInt32(row[column.ColumnName]);
                        }
                        catch (Exception e) {
                            
                        }
                       
                    }
                    else
                    {
                        newDtRow[column.ColumnName] = row[column.ColumnName];
                    }
                }
                dtResult.Rows.Add(newDtRow);
            }
            return dtResult;
        }


        private void button5_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(outFilePath))
            {
                Directory.CreateDirectory(outFilePath);
            }
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;

            if (textBox_zongchengji.Text == ""
                || textBox_chuqirenshu.Text == ""
                || textBox_kemumanfen.Text == ""
                || textBox_renkejiaoshi.Text == "") {
                MessageBox.Show("请先选择以上四个数据表的路径");
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                //button6.Enabled = true;

                return;
            }
            try
            {


                int canKaoTotal = 0;//参考总人数，按照班级-人数存储
                DataTable chuqirenshuTable = null;
                DataTable manfenTable = null;
                DataTableCollection renkejiaoshiTables = null;
                DataTableCollection zongchengjiTables = null;
                DataTable xuekeTable = null;//当前统计学科成绩表--总成绩表里的页

                DataTable baobiaomobanTable = null;//输出报表模板Table

                //0.报表模板表
                using (var stream = File.Open(textBox_baobiaomoban.Text, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        baobiaomobanTable = result.Tables[0];
                    }
                }

                //1.期初人数表
                using (var stream = File.Open(textBox_chuqirenshu.Text, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        chuqirenshuTable = result.Tables[0];
                        chuqirenshuTable.Rows[0].Delete();
                        chuqirenshuTable.AcceptChanges();
                    }
                }

                //2.满分表
                using (var stream = File.Open(textBox_kemumanfen.Text, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {

                        var result = reader.AsDataSet();
                        manfenTable = result.Tables[0];
                        manfenTable.Rows[0].Delete();
                        manfenTable.AcceptChanges();
                    }
                }

                //3.任课教师表
                using (var stream = File.Open(textBox_renkejiaoshi.Text, FileMode.Open, FileAccess.Read))
                {

                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        renkejiaoshiTables = result.Tables;
                    }
                }

                //4.总成绩表
                using (var stream = File.Open(textBox_zongchengji.Text, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();
                        zongchengjiTables = result.Tables;
                    }
                }


                //根据给定的任课教师表 循环生成各科目报表
                foreach (DataTable dt_renkejiaoshi in renkejiaoshiTables) {

                    //该科目所有任课老师
                    dt_renkejiaoshi.Rows[0].Delete();
                    dt_renkejiaoshi.AcceptChanges();

                    DataTable outTable = baobiaomobanTable.Clone();

                    string tableName = dt_renkejiaoshi.TableName;
                    string nianjixueke = "";
                    xuekeTable = null;

                    foreach (DataTable tb in zongchengjiTables) {
                        if (tableName == tb.TableName) {
                            xuekeTable = tb;
                            //去除前三行头
                            xuekeTable.Rows[0].Delete();
                            xuekeTable.Rows[1].Delete();
                            xuekeTable.Rows[2].Delete();
                            xuekeTable.AcceptChanges();
                            break;
                        }
                    }

                    decimal manfen = 0;
                
                    //当前学科满分值
                    foreach (DataRow row in manfenTable.Rows) {
                        if (Convert.ToString(row[0]) == tableName) {
                            manfen = Convert.ToDecimal(row[1]);
                            Console.WriteLine(">>>当前学科:"+ Convert.ToString(row[0]) + ">>>满分：" +row[1]);
                            break;
                        }
                    }

                    decimal jige = Decimal.Multiply(manfen, new Decimal(0.6));
                    decimal youxiu = Decimal.Multiply(manfen, new Decimal(0.85));
                    Console.WriteLine(">>>当前学科及格分数:" + jige + ">>>优秀分数：" + youxiu);

                    canKaoTotal = xuekeTable.Rows.Count;//参考人数
                    Console.WriteLine(">>>参考总人数:" + canKaoTotal);

                    DataTable xuekeTableTmep = UpdateDataTable(xuekeTable, "Column7"); //第八列 分数

                    ////单科成绩按照班级分组统计
                    //var query = from t in xuekeTableTmep.AsEnumerable()
                    //            group t by t.Field<string>("Column3") into m //第四列 班级
                    //            select new
                    //            {
                    //                name = m.Key,
                    //                score = m.Sum(n =>  n.Field<decimal>("Column7")) //第八列 分数
                    //            };

                    int row_index = 0;//统计表行号
                    var paiming_cankao = new SortedList<decimal, int>();
                    var paiming_qichu = new SortedList<decimal, int>();
                    var paiming_cankao_xuexiao= new SortedList<decimal, int>();
                    var paiming_qichu_xuexiao = new SortedList<decimal, int>();


                    //遍历任课教师表，统计每一个任课老师的成绩
                    foreach (DataRow row_renkejiaoshi in dt_renkejiaoshi.Rows) {
                        int jigeTotal = 0;
                        int youxiuTotal = 0;
                        int cankaorenshu = 0;
                        decimal scoreTotal = 0;
                        decimal pingjunfen_cankao = 0;
                        decimal jigelv_cankao = 0;
                        decimal youxiulv_cankao = 0;
                        decimal pingjunfen_qichu = 0;
                        decimal jigelv_qichu = 0;
                        decimal youxiulv_qichu = 0;
                        string renkejiaoshi = Convert.ToString(row_renkejiaoshi["Column0"]);
                        string renkebanji = Convert.ToString(row_renkejiaoshi["Column3"]);
                        nianjixueke= Convert.ToString(row_renkejiaoshi["Column1"]);//年级学科

                        int qichuTotal = 0;

                        //计算每个任课教师的期初人数
                        foreach (DataRow row_qichu in chuqirenshuTable.Rows)
                        {

                            if (Convert.ToString(row_qichu[0])!="" && renkebanji.Contains(Convert.ToString(row_qichu[0])))
                            {
                                qichuTotal += Convert.ToInt32(row_qichu[1]);
                            }
                        }

                        //遍历单科成绩表进行统计及格率、优秀率
                        foreach (DataRow row_xuekechengji in xuekeTable.Rows)
                        {
                            if (renkebanji.Contains(Convert.ToString(row_xuekechengji["Column3"])))
                            {
                                cankaorenshu++;
                                decimal score = Convert.ToDecimal(row_xuekechengji["Column7"]);
                                scoreTotal += score;
                                //0.6及格，0.85优秀
                                if (score >= jige)
                                {
                                    jigeTotal++;
                                }

                                if (score >= youxiu)
                                {
                                    youxiuTotal++;
                                }
                            }

                        }
                        pingjunfen_cankao = Decimal.Divide(scoreTotal, new Decimal(cankaorenshu));
                        pingjunfen_cankao = Decimal.Round(pingjunfen_cankao, 2);

                        pingjunfen_qichu = Decimal.Divide(scoreTotal, new Decimal(qichuTotal));
                        pingjunfen_qichu = Decimal.Round(pingjunfen_qichu, 2);

                        jigelv_cankao = (Decimal.Divide(new Decimal(jigeTotal), new Decimal(cankaorenshu)));
                        jigelv_cankao =  Decimal.Round(jigelv_cankao * 100, 2) ;

                        jigelv_qichu = Decimal.Divide(jigeTotal, new Decimal(qichuTotal));
                        jigelv_qichu =  Decimal.Round(jigelv_qichu * 100, 2);

                        youxiulv_cankao = (Decimal.Divide(new Decimal(youxiuTotal), new Decimal(cankaorenshu)));
                        youxiulv_cankao = Decimal.Round(youxiulv_cankao * 100, 2);

                        youxiulv_qichu = (Decimal.Divide(new Decimal(youxiuTotal), new Decimal(qichuTotal)));
                        youxiulv_qichu =  Decimal.Round(youxiulv_qichu * 100, 2);

                        //Console.WriteLine("教师:" + renkejiaoshi+",参考人数："+ cankaorenshu + ",总分：" + scoreTotal + ",按参考算平均分"+ pingjunfen_cankao + ",及格人数:" + jigeTotal+ ",按参考算及格率"+ jigelv_cankao * 100 + "%,优秀人数:" + youxiuTotal+ ",按参考算优秀率(%)"+ youxiulv_cankao * 100);

                        DataRow row = outTable.NewRow();//报表新增当前教师的统计记录
                        row[0] = renkejiaoshi;//任课教师
                        row[1] = row_renkejiaoshi[1];//学科年级
                        row[2] = row_renkejiaoshi[2];//学校
                        row[3] = row_renkejiaoshi[3];//班级
                        row[4] = Convert.ToInt32( cankaorenshu);//参考人数
                        row[5] = Convert.ToInt32(scoreTotal);//总分
                        row[6] = pingjunfen_cankao;//按参考算平均分
                        row[7] = Convert.ToInt32(jigeTotal);//及格人数
                        row[8] = jigelv_cankao;//按参考算及格率(%)
                        row[9] = Convert.ToInt32(youxiuTotal);//优秀人数
                        row[10] = youxiulv_cankao;//按参考算优秀率(%)
                        row[11] = pingjunfen_cankao+jigelv_cankao+youxiulv_cankao;//按参考算三率和
                        row[12] = 0;//按参考算排名
                        row[13] = Convert.ToInt32(qichuTotal);//期初人数
                        row[14] = pingjunfen_qichu;//按期初算平均分
                        row[15] = jigelv_qichu;//按期初算及格率(%)
                        row[16] = youxiulv_qichu;//按期初算优秀率(%)
                        row[17] = pingjunfen_qichu+ jigelv_qichu+ youxiulv_qichu;//按期初算三率和
                        row[18] = 0;//按期初算排名
                        outTable.Rows.Add(row);
                        paiming_cankao.Add(pingjunfen_cankao + jigelv_cankao + youxiulv_cankao,row_index);
                        paiming_qichu.Add(pingjunfen_qichu + jigelv_qichu + youxiulv_qichu, row_index);
                        row_index++;

                    }

                    int i = 0;
                    foreach (KeyValuePair<decimal, int> kvp in paiming_cankao.Reverse()) {
                        i++;
                        outTable.Rows[Convert.ToInt32(kvp.Value)]["Column12"] = i;
                    }

                    i = 0;
                    foreach (KeyValuePair<decimal, int> kvp in paiming_qichu.Reverse())
                    {
                        i++;
                        outTable.Rows[Convert.ToInt32(kvp.Value)]["Column18"] = i;
                    }
                    DataTable outTableTmp = UpdateDataTableInt(outTable, new string[] { "Column4", "Column5", "Column7", "Column9", "Column13" }); //第八列 分数

                    //按照学校分组统计
                    var query = from t in outTableTmp.AsEnumerable()
                                group t by t.Field<string>("Column2") into m //第3列 学校
                                select new
                                {
                                    name = m.Key,
                                    //xuexiao = m.Field<string>("Column1"),
                                    cankaorenshu_cankao_xuexiao = m.Sum(n => n.Field<Int32>("Column4")), //第5列 总人数
                                    scoreTotal_cankao_xuexiao = m.Sum(n => n.Field<Int32>("Column5")), //第6列 总分
                                    jigeTotal_cankao_xuexiao = m.Sum(n => n.Field<Int32>("Column7")), //第8列 及格人数
                                    youxiulv_cankao_xuexiao = m.Sum(n => n.Field<Int32>("Column9")), //第10列 优秀人数
                                    cankaorenshu_qichu_xuexiao = m.Sum(n => n.Field<Int32>("Column13")) //第14列 期初总人数
                                };
                    if (query.ToList().Count > 0) {

                        query.ToList().ForEach(q => {
                            decimal pingjunfen_cankao = Decimal.Divide(q.scoreTotal_cankao_xuexiao, new Decimal(q.cankaorenshu_cankao_xuexiao));
                            pingjunfen_cankao = Decimal.Round(pingjunfen_cankao, 2);

                            decimal pingjunfen_qichu = Decimal.Divide(q.scoreTotal_cankao_xuexiao, new Decimal(q.cankaorenshu_qichu_xuexiao));
                            pingjunfen_qichu = Decimal.Round(pingjunfen_qichu, 2);

                            decimal jigelv_cankao = (Decimal.Divide(new Decimal(q.jigeTotal_cankao_xuexiao), new Decimal(q.cankaorenshu_cankao_xuexiao)));
                            jigelv_cankao = Decimal.Round(jigelv_cankao * 100, 2);

                            decimal jigelv_qichu = Decimal.Divide(q.jigeTotal_cankao_xuexiao, new Decimal(q.cankaorenshu_qichu_xuexiao));
                            jigelv_qichu = Decimal.Round(jigelv_qichu * 100, 2);

                            decimal youxiulv_cankao = (Decimal.Divide(new Decimal(q.youxiulv_cankao_xuexiao), new Decimal(q.cankaorenshu_cankao_xuexiao)));
                            youxiulv_cankao = Decimal.Round(youxiulv_cankao * 100, 2);

                            decimal youxiulv_qichu = (Decimal.Divide(new Decimal(q.youxiulv_cankao_xuexiao), new Decimal(q.cankaorenshu_qichu_xuexiao)));
                            youxiulv_qichu = Decimal.Round(youxiulv_qichu * 100, 2);

                            DataRow row = outTableTmp.NewRow();//报表新增当前教师的统计记录
                            row[0] = q.name+"合计";//任课教师
                            row[1] = nianjixueke;//学科年级
                            row[2] = q.name;//学校
                            row[3] = "-";//班级
                            row[4] = q.cankaorenshu_cankao_xuexiao;//参考人数
                            row[5] = q.scoreTotal_cankao_xuexiao;//总分
                            row[6] = pingjunfen_cankao;//按参考算平均分
                            row[7] = q.jigeTotal_cankao_xuexiao;//及格人数
                            row[8] = jigelv_cankao;//按参考算及格率(%)
                            row[9] = q.youxiulv_cankao_xuexiao;//优秀人数
                            row[10] = youxiulv_cankao;//按参考算优秀率(%)
                            row[11] = pingjunfen_cankao + jigelv_cankao + youxiulv_cankao;//按参考算三率和
                            row[12] = 0;//按参考算排名
                            row[13] = q.cankaorenshu_qichu_xuexiao;//期初人数
                            row[14] = pingjunfen_qichu;//按期初算平均分
                            row[15] = jigelv_qichu;//按期初算及格率(%)
                            row[16] = youxiulv_qichu;//按期初算优秀率(%)
                            row[17] = pingjunfen_qichu + jigelv_qichu + youxiulv_qichu;//按期初算三率和
                            row[18] = 0;//按期初算排名
                            outTableTmp.Rows.Add(row.ItemArray);
                            paiming_cankao_xuexiao.Add(pingjunfen_cankao + jigelv_cankao + youxiulv_cankao, row_index);
                            paiming_qichu_xuexiao.Add(pingjunfen_qichu + jigelv_qichu + youxiulv_qichu, row_index);
                            row_index++;
                        });

                    }

                    i = 0;
                    foreach (KeyValuePair<decimal, int> kvp in paiming_cankao_xuexiao.Reverse())
                    {
                        i++;
                        outTableTmp.Rows[Convert.ToInt32(kvp.Value)]["Column12"] = i;
                    }

                    i = 0;
                    foreach (KeyValuePair<decimal, int> kvp in paiming_qichu_xuexiao.Reverse())
                    {
                        i++;
                        outTableTmp.Rows[Convert.ToInt32(kvp.Value)]["Column18"] = i;
                    }
                    ExportToExcel(outTableTmp, outFilePath+ nianjixueke + ".xlsx",textBox_baobiaomoban.Text);
                }


            }
            catch (Exception ee)
            {
                Console.WriteLine( ee.StackTrace);
            }
            finally {
                button1.Enabled = true;
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                //button6.Enabled = true;
            }
           


        }



        /// <summary>
        /// 判断是否为兼容模式
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static bool GetIsCompatible(string filePath)
        {
            return filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase);
        }



        /// <summary>
        /// 创建工作薄
        /// </summary>
        /// <param name="isCompatible"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isCompatible,string templeteFile)
        {
            using (FileStream fs = new FileStream(templeteFile, FileMode.Open, FileAccess.Read))
            {
                if (isCompatible)
                {
                    return new HSSFWorkbook(fs);
                }
                else
                {
                    return new XSSFWorkbook(fs);
                }

            }
 
        }

        /// <summary>
        /// 创建工作薄(依据文件流)
        /// </summary>
        /// <param name="isCompatible"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isCompatible, dynamic stream)
        {
            if (isCompatible)
            {
                return new HSSFWorkbook(stream);
            }
            else
            {
                return new XSSFWorkbook(stream);
            }
        }


        /// <summary>
        /// 由DataTable导出Excel
        /// </summary>
        /// <param name="sourceTable">要导出数据的DataTable</param>
        /// <returns>Excel工作表</returns>
        public static string ExportToExcel(DataTable sourceTable, string filePath, string templeteFile, string sheetName = "result")
        {
            if (sourceTable.Rows.Count <= 0) return null;

           
            if (string.IsNullOrEmpty(templeteFile)) return null;

            bool isCompatible = GetIsCompatible(templeteFile);

            IWorkbook workbook = CreateWorkbook(isCompatible,templeteFile);


            ISheet sheet = workbook.GetSheet("Sheet1");
            //IRow headerRow = sheet.CreateRow(0);
            //// handling header.
            //foreach (DataColumn column in sourceTable.Columns)
            //{
            //    ICell headerCell = headerRow.CreateCell(column.Ordinal);
            //    headerCell.SetCellValue(column.ColumnName);
            //}

            // handling value.
            int rowIndex = 2;

            foreach (DataRow row in sourceTable.Rows)
            {
                IRow dataRow = sheet.GetRow(rowIndex);
                //foreach (DataColumn column in sourceTable.Columns)
                //{
                //    dataRow.GetCell(column.Ordinal).SetCellValue((row[column] ?? "").ToString());
                //}

                //IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in sourceTable.Columns)
                {
                    dataRow.GetCell(column.Ordinal).SetCellValue((row[column] ?? "").ToString());
                }

                rowIndex++;
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                ms.Position = 0;

                using (FileStream fs1 = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    byte[] data = ms.ToArray();
                    fs1.Write(data, 0, data.Length);
                    fs1.Flush();
                    fs1.Dispose();

                    data = null;
                }

                sheet = null;
                workbook = null;

            }

            //FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            //workbook.Write(fs);
            //fs.Dispose();

            //sheet = null;
            ////headerRow = null
            //workbook = null;
            //workbookTmp = null;

            return filePath;
        }


    }
}
