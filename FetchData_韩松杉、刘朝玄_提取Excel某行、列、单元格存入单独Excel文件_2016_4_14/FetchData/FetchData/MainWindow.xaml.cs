using System;
using System.Windows;
using System.Windows.Data;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;

namespace FetchData
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ExcelFile excelFile;
        Excel.Application excelApp;

        public MainWindow()
        {
            //变量初始化
            excelFile = new ExcelFile();
            excelApp = new Excel.Application();
            InitializeComponent();
        }

        #region 数据及逻辑等处理方法
        /// <summary>
        /// 将string类型加入到ListBox中的异步处理方法
        /// </summary>
        /// <param name="add">将要加入的字符串</param>
        public void AddList(string add)
        {
            if (add == "End")
            {
                add = "已生成： " + excelFile.fileFolderPath + "\\result.xlsx";
            }
            //采用invoke的方式告知List需要执行的操作（即添加add这个字符串作为一个Item），并下拉List
            list_workstate.Dispatcher.Invoke(new System.Action(() =>
            {
                list_workstate.Items.Add(add);
                list_workstate.ScrollIntoView(add);
            }));
        }

        /// <summary>
        /// 根据工作状态来设置控件是否可用（设置isEnanbled属性）
        /// </summary>
        /// <param name="working">true表示工作状态，false表示未在工作的状态</param>
        public void SetViewState(bool working)
        {
            this.Dispatcher.Invoke(new System.Action(() =>
            {
                //采用invoke的方式告知控件需要执行的操作，防止跨线程访问出错
                radio_file_all.IsEnabled = !working;
                box_start_date.IsEnabled = !working;
                box_end_date.IsEnabled = !working;
                radio_row.IsEnabled = !working;
            }));
        }

        /// <summary>
        /// 将整形变量转为Excel文档使用的26进制表示
        /// </summary>
        /// <param name="num">将要转化的数字</param>
        /// <returns>以字符串形式存储的Excel所用的26进制符号</returns>
        public string IntToString26(int num)
        {
            string ret = "";
            int temp_int;
            char temp_byte;
            while (num > 26)
            {
                //每次对num取余，分别处理每一位
                temp_int = num % 26;

                //同上，思路一样，每次都根据字符'A'的差值来推算
                temp_byte = (char)(temp_int + 'A' - 1);
                ret = temp_byte + ret;
                if (num > 26)
                {
                    //由于是整形，因此可以直接除
                    num = num / 26;
                }
            }
            //num小于26时跳出，此时还需要处理个位
            ret = (char)(num + 'A' - 1) + ret;
            return ret;
        }

        /// <summary>
        /// 将Excel文档使用的26进制表示转为整形变量
        /// </summary>
        /// <param name="num">将要转化的26进制表示</param>
        /// <returns>整形数字</returns>
        public int String26ToInt(string num)
        {
            if (num == "")
            {
                //输入为空的情况
                return -1;
            }
            int result = 0;
            int temp = 0;
            int count = 0;
            while (count < num.Length)
            {
                //对num的每一位转化为数字的处理
                temp = num[count] - 'A' + 1;
                if (temp < 1 || temp > 26)
                {
                    //如果该字符不在'A'到'Z'之间，则返回错误
                    return -1;
                }
                result = result * 26 + temp;
                count++;
            }
            return result;
        }

        /// <summary>
        /// 将string类型输入的日期向后推一天，并输出
        /// </summary>
        /// <param name="str">将要推算的日期</param>
        /// <returns>向后推算后的日期</returns>
        public string AddDate(string str)
        {
            int year, month, day;
            string yearStr, monthStr, dayStr;
            //通过Substing来对输入字符串进行截取操作
            year = Int32.Parse(str.Substring(0, 4));
            month = Int32.Parse(str.Substring(4, 2));
            day = Int32.Parse(str.Substring(6, 2));

            //先对31号做特殊处理
            if (day == 31)
            {
                if (month == 12)
                {
                    year++;
                    month = 1;
                    day = 1;
                }
                else
                {
                    day = 1;
                    month++;
                }
            }
            //对小月的30号做特殊处理
            else if (day == 30)
            {
                if (month == 4 && month == 6 && month == 9 && month == 11)
                {
                    month++;
                    day = 1;
                }
                else
                {
                    day++;
                }
            }
            //对二月做特殊处理
            else if (month == 2)
            {
                if (day == 29)
                {
                    day = 1;
                    month++;
                }
                else if (day == 28)
                {
                    if (year % 4 == 0)
                    {
                        day++;
                    }
                    else
                    {
                        day = 1;
                        month++;
                    }
                }
                else
                {
                    day++;
                }
            }
            else
            {
                day++;
            }


            yearStr = year.ToString();
            //分别为月份为一位数或是两位数，判断是否补“0”
            if (month <= 9)
            {
                monthStr = "0" + month.ToString();
            }
            else
            {
                monthStr = month.ToString();
            }
            //分别为天数为一位数或是两位数，判断是否补“0”
            if (day <= 9)
            {
                dayStr = "0" + day.ToString();
            }
            else
            {
                dayStr = day.ToString();
            }
            //组合并返回
            str = yearStr + monthStr + dayStr;
            return str;
        }

        /// <summary>
        /// 数据处理线程，用于读取一批Excel文件并按需求写入目标Excel文件
        /// </summary>
        public void MakeExcelFileThread()
        {
            //---------预处理部分-------
            //检查文件是否存在，如果存在，则建立一个新的文件
            int index = 1;  //重名文件的起始值
            while (File.Exists(excelFile.saveFileName))
            {
                //不断检查下一个编号的文件是否已存在，直至不重名
                excelFile.saveFileName = excelFile.fileFolderPath + "\\" + excelFile.fileNameWithoutExtension + "_" + index + excelFile.extension;
                index++;
            }
            try
            {
                //创建Excel工作表
                excelFile.workbook = excelApp.Workbooks.Add(true);
                //将Worksheet名字改为设定的名字
                excelFile.workbook.Worksheets["Sheet1"].Name = excelFile.saveSheetName;
                //另存为
                excelFile.workbook.SaveAs(excelFile.saveFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                AddList("已新建：" + excelFile.saveFileName + " 文件");
            }
            catch
            {
                AddList("创建Excel文件：" + excelFile.saveFileName + " 失败！");
                SetViewState(false);
                return;
            }
            //---------局部变量的定义部分-------
            Workbook saveworkbook, tempWorkBook;      //保存至文件的Workbook
            Worksheet saveworksheet, tempWorkSheet;    //保存至文件的Worksheet
            //按文件名打开将要存储的Workbook和Worksheet，默认表名为“Sheet1”
            //saveworkbook = excelApp.Workbooks.Open(excelFile.saveFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            saveworkbook = excelFile.workbook;
            saveworksheet = (Excel.Worksheet)saveworkbook.Worksheets[excelFile.saveSheetName];
            //存储所用的临时变量
            int completeFileCount =0;   //已处理的文件数量
            string rangeStart;  //存储Range的开始位置
            string rangeEnd;    //存储Range的结束位置
            string tempFileName = "";
            int tempLength = 0;

            //---------处理Excel文件部分-------
            //分情况将目录下的Excel文件的指定内容，写入到某文件中
            //源文件名已经保存在excelFile.fileNames[]数组中了，数量为excelFile.fileCount
            //目标文件的文件名为excelFile.saveFileName

            switch (excelFile.workMode)
            {
                case ExcelFile.WorkMode.ROW:
                    {
                        //取某行存储
                        while (completeFileCount < excelFile.fileCount)
                        {

                            try
                            {
                                if (excelFile.readAllFiles)
                                {
                                    //剔除文件名为非数字的文件（选择“所有文件”时需要判定）
                                    //如果文件名非数字，则会出现异常，进入catch{}
                                    Int32.Parse(Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]));
                                }
                                tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[completeFileCount], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets[excelFile.tempSheetName];
                                tempFileName = Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]);

                                //第一列存储文件名
                                saveworksheet.Cells[(completeFileCount + 1), 1] = tempFileName;

                                //计算数据存储位置
                                //存储Range开始列号为："B + (已完成文件个数 + 1)"
                                rangeStart = "B" + (completeFileCount + 1);
                                //数据源文件的数据个数
                                tempLength = String26ToInt(excelFile.row_endCol) - String26ToInt(excelFile.row_startCol) + 1;
                                //存储Range结束位置为：" (长度+1后转为26进制) + (已完成文件个数 + 1)"
                                rangeEnd = IntToString26(tempLength + 1) + (completeFileCount + 1);

                                //读出数据源文件的Range并赋给存储Range
                                saveworksheet.Range[rangeStart, rangeEnd].Value = tempWorkSheet.Range[excelFile.row_startCol + excelFile.row_rowNumber, excelFile.row_endCol + excelFile.row_rowNumber].Value;
                                tempWorkBook.Close();
                                AddList("已提取" + excelFile.fileNames[completeFileCount]);
                            }
                            catch
                            {
                                completeFileCount++;
                                continue;
                            }
                            completeFileCount++;
                        }
                        break;
                    }
                case ExcelFile.WorkMode.COL:
                    {
                        while (completeFileCount < excelFile.fileCount)
                        {
                            try
                            {
                                if (excelFile.readAllFiles)
                                {
                                    //剔除文件名为非数字的文件（选择“所有文件”时需要判定）
                                    //如果文件名非数字，则会出现异常，进入catch{}
                                    Int32.Parse(Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]));
                                }
                                tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[completeFileCount], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets[excelFile.tempSheetName];
                                tempFileName = Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]);
                                //第一行全为文件名
                                saveworksheet.Cells[1, (completeFileCount + 1)] = tempFileName;

                                //计算数据存储位置
                                //存储Range开始列号为："(已完成文件个数 + 1) + 2"
                                rangeStart = IntToString26(completeFileCount + 1) + "2";
                                //数据源文件的数据个数
                                tempLength = Int32.Parse(excelFile.col_endRow) - Int32.Parse(excelFile.col_startRow) + 1;
                                //存储Range结束位置为：" (已完成文件个数 + 1 后转为26进制) + (长度 + 1) "
                                rangeEnd = IntToString26(completeFileCount + 1) + (tempLength + 1);

                                //读出数据源文件的Range并赋给存储Range
                                saveworksheet.Range[rangeStart, rangeEnd].Value = tempWorkSheet.Range[excelFile.col_colNumber + excelFile.col_startRow, excelFile.col_colNumber + excelFile.col_endRow].Value;
                                tempWorkBook.Close();
                                AddList("已提取" + excelFile.fileNames[completeFileCount]);
                            }
                            catch
                            {
                                completeFileCount++;
                                continue;
                            }
                            completeFileCount++;
                        }
                        break;
                    }
                case ExcelFile.WorkMode.SINGLE:
                    {
                        while (completeFileCount < excelFile.fileCount)
                        {
                            try
                            {
                                if (excelFile.readAllFiles)
                                {
                                    //剔除文件名为非数字的文件（选择“所有文件”时需要判定）
                                    //如果文件名非数字，则会出现异常，进入catch{}
                                    Int32.Parse(Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]));
                                }
                                tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[completeFileCount], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                                tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets[excelFile.tempSheetName];
                                tempFileName = Path.GetFileNameWithoutExtension(excelFile.fileNames[completeFileCount]);
                                //第一列存储文件名
                                saveworksheet.Cells[(completeFileCount + 1), 1] = tempFileName;
                                //逐一将数据源文件的每个Cell对应到存储Excel中去
                                saveworksheet.Cells[(completeFileCount + 1), 2] = tempWorkSheet.Cells[excelFile.single_rowNumber, excelFile.single_colNumber];
                                tempWorkBook.Close();
                                AddList("已提取" + excelFile.fileNames[completeFileCount]);
                            }
                            catch
                            {
                                completeFileCount++;
                                continue;
                            }
                            completeFileCount++;
                        }
                        break;
                    }
                case ExcelFile.WorkMode.NONE:
                    {
                        AddList("非法的工作状态！");
                        break;
                    }
            }
            //---------后续处理部分-------
            saveworkbook.Save();
            saveworkbook.Close();
            excelApp.Workbooks.Close();
            AddList("完成！");
            SetViewState(false);
            /**
                        //if ((bool)radio_row.IsChecked)
                        //{            //取某行存储
                        //    while (completeFileCount < excelFile.fileCount)
                        //    {
                        //        tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[filenameSquence], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //        tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets["Sheet1"];
                        //        saveworksheet.Cells[saveWorkSheetrow, 1] = excelFile.fileNames[filenameSquence];//第一列全为文件名，第二列才开始存数据
                        //        //新文件存储的开始列号为B，结束列号为用户输入的结束列-开始列号然后加上1
                        //        rangeStart = "B" + IntToString26(saveWorkSheetrow);
                        //        rangeEnd = IntToString26(String26ToInt(excelFile.row_endCol) - String26ToInt(box_row_startCol.Text) + 1) + IntToString26(saveWorkSheetrow);
                        //        saveworksheet.Range[rangeStart, rangeEnd].Value = tempWorkSheet.Range[box_row_startCol.Text + box_row_rowNumber.Text, excelFile.row_endCol + box_row_rowNumber.Text];
                        //        tempWorkBook.Close();
                        //        saveWorkSheetrow++;
                        //        completeFileCount++;
                        //        filenameSquence++;

                        //    }
                        //}
                        //else if ((bool)radio_col.IsChecked)
                        //{           //取某列存储
                        //    while (completeFileCount < excelFile.fileCount)
                        //    {
                        //        tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[filenameSquence], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //        tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets["Sheet1"];
                        //        saveworksheet.Cells[1, saveWorkSheetcol] = excelFile.fileNames[filenameSquence];//第一行全为文件名，第二行才开始存数据
                        //        //新文件存储的开始行号为2（第一行存了文件名），结束行号为用户输入的结束行-开始行然后加上1
                        //        rangeStart = IntToString26(saveWorkSheetcol) + "2";
                        //        rangeEnd = IntToString26(saveWorkSheetcol) + Convert.ToString(Convert.ToInt32(box_col_endRow.Text) - Convert.ToInt32(box_col_startRow.Text) + 1);
                        //        saveworksheet.Range[rangeStart,rangeEnd].Value= tempWorkSheet.Range[box_col_rowNumber.Text + box_col_startRow.Text, box_col_rowNumber.Text + box_col_endRow.Text];
                        //        tempWorkBook.Close();
                        //        saveWorkSheetcol++;
                        //        completeFileCount++;
                        //        filenameSquence++;

                        //    }
                        //}
                        //else if ((bool)radio_single.IsChecked)
                        //{
                        //    while (completeFileCount < excelFile.fileCount)
                        //    {
                        //        tempWorkBook = excelApp.Workbooks.Open(excelFile.fileNames[filenameSquence], Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //        tempWorkSheet = (Excel.Worksheet)tempWorkBook.Worksheets["Sheet1"];
                        //        saveworksheet.Cells[saveWorkSheetrow, 1] = excelFile.fileNames[filenameSquence];
                        //        saveworksheet.Cells[saveWorkSheetrow, 2] = tempWorkSheet.Cells[excelFile.single_rowNumber, box_single_colNumber];
                        //        tempWorkBook.Close();
                        //        saveWorkSheetrow++;
                        //        completeFileCount++;
                        //        filenameSquence++;

                        //    }
                        //    saveworkbook.Save();

                        //}

*/


        }
        #endregion



        #region 控件事件响应的处理方法
        /// <summary>
        /// 点击“浏览”的处理方法
        /// </summary>
        /// <param name="sender">事件产生者</param>
        /// <param name="e">事件参数，表示事件类型</param>
        private void button_file_browse_Click(object sender, RoutedEventArgs e)
        {
            //调用ChooseFilePath，若返回正确值，则将文件路径赋给文件路径Textbox
            if (excelFile.ChooseFilePath())
            {
                box_file_path.Text = excelFile.fileFolderPath;
            }
        }

        /// <summary>
        /// 点击“开始”的处理方法
        /// </summary>
        /// <param name="sender">事件产生者</param>
        /// <param name="e">事件参数，表示事件类型</param>
        private void button_start_Click(object sender, RoutedEventArgs e)
        {
            SetViewState(true);
            int startDate = 0;
            int endDate = 0;
            //判断文件路径是否已经确定
            if (excelFile.fileFolderPath == "")
            {
                AddList("非法的文件路径！");
                SetViewState(false);
                return;
            }
            #region 判断输入的值是否都合法
            excelFile.readAllFiles = true;
            //如果选中“自定义日期”
            if ((bool)radio_file_any.IsChecked)
            {
                //判断开始时间与结束日期是否正确填写
                try
                {
                    if (box_start_date.Text.Length != 8 || box_end_date.Text.Length != 8)
                    {
                        //如果起始日期或结束日期长度不对，则跳出
                        AddList("请检查开始日期和结束日期的填写！");
                        SetViewState(false);
                        return;
                    }
                    startDate = Int32.Parse(box_start_date.Text);
                    endDate = Int32.Parse(box_end_date.Text);
                    if (startDate > endDate)
                    {
                        //如果起始日期比结束日期大，则跳出
                        AddList("请检查开始日期和结束日期的填写！");
                        SetViewState(false);
                        return;
                    }                    
                }
                catch
                {
                    //如果填入的字符串非纯数字，则跳出
                    AddList("请检查开始日期和结束日期的填写！");
                    SetViewState(false);
                    return;
                }
                excelFile.readAllFiles = false;
            }
            //检查目标文件名是否为空
            if (box_save_filename.Text == "")
            {
                AddList("请检查“保存文件名”输入！");
                SetViewState(false);
                return;
            }
            //检查源文件表名和目标文件存储表名是否为空
            if (box_temp_sheetName.Text == "")
            {
                AddList("请检查“源文件表名”输入！");
                SetViewState(false);
                return;
            }
            if (box_save_sheetName.Text == "")
            {
                AddList("请检查“存储表名”输入！");
                SetViewState(false);
                return;
            }

            //如果选中“某行”
            if ((bool)radio_row.IsChecked)
            {
                try
                {
                    Int32.Parse(box_row_rowNumber.Text);
                }
                catch
                {
                    AddList("请检查“行号”输入！");
                    SetViewState(false);
                    return;
                }
                if (String26ToInt(box_row_startCol.Text) == -1)
                {
                    AddList("请检查“起始列”输入！");
                    SetViewState(false);
                    return;
                }
                if (String26ToInt(box_row_endCol.Text) == -1)
                {
                    AddList("请检查“终止列”输入！");
                    SetViewState(false);
                    return;
                }
                excelFile.workMode = ExcelFile.WorkMode.ROW;
                excelFile.row_rowNumber = box_row_rowNumber.Text;
                excelFile.row_startCol = box_row_startCol.Text;
                excelFile.row_endCol = box_row_endCol.Text;
            }
            //如果选中“某列”
            if ((bool)radio_col.IsChecked)
            {
                if (String26ToInt(box_col_colNumber.Text) == -1)
                {
                    AddList("请检查“列号”输入！");
                    SetViewState(false);
                    return;
                }
                try
                {
                    Int32.Parse(box_col_startRow.Text);
                }
                catch
                {
                    AddList("请检查“起始行”输入！");
                    SetViewState(false);
                    return;
                }
                try
                {
                    Int32.Parse(box_col_endRow.Text);
                }
                catch
                {
                    AddList("请检查“终止行”输入！");
                    SetViewState(false);
                    return;
                }
                excelFile.workMode = ExcelFile.WorkMode.COL;
                excelFile.col_colNumber = box_col_colNumber.Text;
                excelFile.col_startRow = box_col_startRow.Text;
                excelFile.col_endRow = box_col_endRow.Text;
            }
            //如果选中“某单元格”
            if ((bool)radio_single.IsChecked)
            {
                if (String26ToInt(box_single_colNumber.Text) == -1)
                {
                    AddList("请检查“列号”输入！");
                    SetViewState(false);
                    return;
                }
                try
                {
                    Int32.Parse(box_single_rowNumber.Text);
                }
                catch
                {
                    AddList("请检查“行号”输入！");
                    SetViewState(false);
                    return;
                }
                excelFile.workMode = ExcelFile.WorkMode.SINGLE;
                excelFile.single_rowNumber = box_single_rowNumber.Text;
                excelFile.single_colNumber = box_single_colNumber.Text;
            }
            excelFile.fileNameWithoutExtension = box_save_filename.Text;
            excelFile.saveFileName = excelFile.fileFolderPath + "\\" + box_save_filename.Text + ".xlsx";
            excelFile.tempSheetName = box_temp_sheetName.Text;
            excelFile.saveSheetName = box_save_sheetName.Text;
            /**   本部分功能已经弃用
                        //如果选中“自定义范围”
                        if ((bool)radio_any.IsChecked)
                        {
                            try
                            {
                                Int32.Parse(box_any_startRow.Text);
                            }
                            catch
                            {
                                AddList("请检查“起始行”输入！");
                                return;
                            }
                            try
                            {
                                Int32.Parse(box_any_endRow.Text);
                            }
                            catch
                            {
                                AddList("请检查“终止行”输入！");
                                return;
                            }

                            if (String26ToInt(box_any_startCol.Text) == -1)
                            {
                                AddList("请检查“起始列”输入！");
                                return;
                            }
                            if (String26ToInt(box_any_endCol.Text) == -1)
                            {
                                AddList("请检查“终止列”输入！");
                                return;
                            }
                        }
             */
            #endregion


            #region 检查是要打开目录下所有文件还是某日期范围内的文件，并分别计算文件数量
            excelFile.fileCount = 0;
            excelFile.fileNames = new string[10000];
            if ((bool)radio_file_all.IsChecked)
            {
                excelFile.fileNames = Directory.GetFiles(excelFile.fileFolderPath, "*", SearchOption.TopDirectoryOnly);
                excelFile.fileCount = excelFile.fileNames.Length;
            }
            else if ((bool)radio_file_any.IsChecked)
            {
                //计数
                try
                {
                    int end = Int32.Parse(box_end_date.Text);
                    string currentDate = box_start_date.Text;
                    string currentFileName;
                    //循环检查文件是否存在，直至到结束日期
                    while (Int32.Parse(currentDate) <= endDate)
                    {
                        currentFileName = excelFile.fileFolderPath + "\\" + currentDate + ".xlsx";
                        if (File.Exists(currentFileName))
                        {
                            //文件存在则数量+1,并将文件存入数组中
                            excelFile.fileNames[excelFile.fileCount] = currentFileName;
                            excelFile.fileCount++;
                        }
                        //日期后推一天
                        currentDate = AddDate(currentDate);
                    }
                }
                catch
                {
                    AddList("文件计数错误！");
                    SetViewState(false);
                    return;
                }
            }
            #endregion

            //建立处理线程并执行
            Thread makeExcelFileThread = new Thread(MakeExcelFileThread);
            makeExcelFileThread.Start();
        }
        #endregion
    }

    /// <summary>
    /// 记录Excel文件信息，提供一些处理方法的类
    /// </summary>
    public class ExcelFile
    {
        public string fileFolderPath;   //文件路径（不含文件名）
        public string saveFileName;     //将储存的目标文件名（含路径）
        public string fileNameWithoutExtension;

        public string[] fileNames;      //存储来源文件名的数组
        public int fileCount;           //文件名数组中存储的文件名数量

        public Workbook workbook;   //工作表变量

        public enum WorkMode : int
        {
            NONE = 0,
            COL = 1,
            ROW = 2,
            SINGLE = 3
        }; //提取文件内容的方式（按行、列、单元格）
        public WorkMode workMode = new WorkMode();  
        public bool readAllFiles;       //是否读取所有文件

        //分别为记录控件上的参数的变量
        public string tempSheetName;
        public string saveSheetName;
        
        public string row_rowNumber;
        public string row_startCol;
        public string row_endCol;
        public string col_colNumber;
        public string col_startRow;
        public string col_endRow;
        public string single_rowNumber;
        public string single_colNumber;


        public string extension = ".xlsx";
        public FolderBrowserDialog fileBrowser;     //选择文件目录的对话框，从WinForm里面引用的

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelFile()
        {
            fileFolderPath = "";
            fileBrowser = new FolderBrowserDialog();
            fileNames = new string[10000];
            fileCount = 0;
            workMode = WorkMode.NONE;
        }

        /// <summary>
        /// 选择文件目录
        /// </summary>
        /// <returns>true表示有选择某个文件目录并记录，false表示未选择</returns>
        public bool ChooseFilePath()
        {
            DialogResult result = fileBrowser.ShowDialog();
            //如果对话框以点击“选择该文件夹”为 结束，而不是关闭
            if (result == DialogResult.OK)
            {
                //保存现在选择的文件路径
                fileFolderPath = fileBrowser.SelectedPath;
                return true;
            }
            return false;
        }
    }

    //这个bool型反转的值转换器暂时没用到
    /// <summary>
    /// 值转换器，用于控件元素的值绑定，bool型反转
    /// </summary>
    public class BoolToOppositeBoolConverter : IValueConverter
    {
        /// <summary>
        /// 正向转换bool值
        /// </summary>
        /// <param name="value">转换输入值</param>
        /// <param name="targetType">输出类型</param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return !(bool)value;
        }
        /// <summary>
        /// 反向转换bool值
        /// </summary>
        /// <param name="value">转换输入</param>
        /// <param name="targetType">输出类型</param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return !(bool)value;
        }
    }

    /// <summary>
    /// 值转换器，用于控件元素的值绑定，将bool型与Visibility关联
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// 正向转换
        /// </summary>
        /// <param name="value">转换输入</param>
        /// <param name="targetType">输出类型</param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((bool)value)
            {
                //如果输入true，则返回“可见”的Visibility
                return Visibility.Visible;
            }
            else
            {
                //如果输入false，则返回“隐藏”的Visibility
                return Visibility.Hidden;
            }
        }
        /// <summary>
        /// 反向转换
        /// </summary>
        /// <param name="value">转换输入</param>
        /// <param name="targetType">输出类型</param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            if ((Visibility)value == Visibility.Visible)
            {
                //如果输入“可见”的Visibility，则返回true
                return true;
            }
            else
            {
                //如果输入“不可见”的Visibility，则返回false
                return false;
            }
        }
    }

}
