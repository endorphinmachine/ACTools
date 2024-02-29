using NPOI.SS.Formula.Functions;
using NPOI.XWPF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AcceptionTools
{
    public partial class ACTools : Form
    {

        public List<Build> builds = new List<Build>();
        public string wordTemplatePath = @"核实概况模板.docx";
        public string wordSavePath = string.Empty;
        public string noMatchItem = string.Empty;

        public ACTools()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void excelInputButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = ".xlsx文件|*.xlsx";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelPathBox.Text = openFileDialog.FileName;
                }
            }

            if (File.Exists(ExcelPathBox.Text))
            {
                builds.Clear();

                List<string> buildNameList = new List<string>();

                using (var excelPackage = new ExcelPackage(new FileInfo(ExcelPathBox.Text)))
                {
                    // 读取工程编号
                    var excelSheet = excelPackage.Workbook.Worksheets[0];

                    string projNumberText = excelSheet.Cells[2, 1].Text.Substring(5);

                    ProjNumber.Text = "工程编号：" + projNumberText;

                    // 读取建设项目名称及信息, 按建筑项目名称所在列划分
                    List<int> MergedRow = FindMergedRow(excelSheet);

                    foreach (int RowIndex in MergedRow)
                    {
                        var buildNameText = excelSheet.Cells[RowIndex, 1].Text.Trim();

                        if (NonEmptyCellCount(excelSheet, RowIndex) == 1)
                        {
                            Build build = new Build
                            {
                                Name = buildNameText
                            };

                            buildNameList.Add(buildNameText);

                            // 读取建设项目对应的各项规划条件核实信息
                            int startRow = RowIndex;

                            while (NonEmptyCellCount(excelSheet, startRow) != 0 && excelSheet.Cells[startRow, 1].Text != "地上∑")
                            {
                                startRow += 1;
                            }

                            int endRow = startRow + 1;

                            while (NonEmptyCellCount(excelSheet, endRow) != 1 && excelSheet.Cells[endRow, 1].Text != "外墙饰面面积")
                            {
                                endRow += 1;
                            }

                            build.Area = GetBuildInfo(excelSheet, startRow, endRow, 1, 10);
                            build.Function = GetBuildInfo(excelSheet, startRow, endRow - 3, 5, 6);
                            build.Public = GetBuildInfo(excelSheet, startRow, endRow - 3, 7, 8);
                            build.Other = GetBuildInfo(excelSheet, startRow, endRow - 3, 9, 10);
                            build.FPOArea = GetBuildInfo(excelSheet, startRow, endRow - 3, 1, 10);
                            build.Balcony = excelSheet.Cells[endRow - 2, 4].Text.Trim();
                            build.House = excelSheet.Cells[endRow - 2, 11].Text.Trim();
                            // build.CheckInfo = GetCheckInfo(build.Area, build.Balcony, build.House);
                            builds.Add(build);
                        }
                    }
                }

                // 删除非建筑项目名称，绑定下拉列表
                buildNameList.RemoveAll(string.IsNullOrEmpty);

                buildNameBox.DataSource = buildNameList;

                buildNameBox.SelectedIndex = 0;

            }
        }

        private void wordInputButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = ".docx文件|*.docx";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    WordPathBox.Text = openFileDialog.FileName;
                }
            }
        }


        public static List<int> FindMergedRow(ExcelWorksheet worksheet)
        {
            List<int> MergedRow = new List<int>();

            int totalColumns = 11;

            // 遍历所有合并单元格
            foreach (var mergeAddress in worksheet.MergedCells)
            {
                // 解析合并单元格的地址
                var CellAddress = worksheet.Cells[mergeAddress];

                // 如果合并单元格的长度等于表格的总列数，则记录其列数
                if (CellAddress.Columns == totalColumns)
                {
                    MergedRow.Add(CellAddress.Start.Row);
                }
            }
            MergedRow.Sort();
            MergedRow.RemoveAt(0);
            MergedRow.RemoveAt(MergedRow.Count - 1);

            return MergedRow;
        }


        public static string ReplacePath(string originalPath, string buildName)
        {
            string fileDir = Path.GetDirectoryName(originalPath);

            string newPath = Path.Combine(fileDir, buildName + "-核实概况.docx");

            return newPath;
        }


        public static int NonEmptyCellCount(ExcelWorksheet worksheet, int rowIndex)
        {
            var row = worksheet.Cells[rowIndex, 1, rowIndex, worksheet.Dimension.End.Column];

            return row.Count(cell => !string.IsNullOrEmpty(cell.Text));
        }


        private static Dictionary<string, string> GetBuildInfo(ExcelWorksheet worksheet, int startRow, int endRow, int startCol, int endCol)
        {
            List<string> buildInfoList = new List<string>();

            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startCol; j <= endCol; j++)
                {
                    var cell = worksheet.Cells[i, j].Text.Trim();

                    buildInfoList.Add(cell);
                }
            }
            buildInfoList.RemoveAll(x => x == "");

            // 各项目数据存放字典
            Dictionary<string, string> buildInfo = new Dictionary<string, string>();

            for (int i = 0; i < buildInfoList.Count - 1; i++)
            {
                string key = buildInfoList[i];

                string value = buildInfoList[i + 1];

                if (Regex.IsMatch(key, @"[\u4e00-\u9fa5]") && Regex.IsMatch(value, @"\d"))
                {
                    buildInfo.Add(key, value);
                }
            }
            return buildInfo;
        }

        private void WriteTabel(Build build, int mode)
        {
            string wordInputPath = string.Empty;

            switch (mode)
            {
                // 批量生成
                case 0:
                    wordInputPath = wordTemplatePath;
                    wordSavePath = ReplacePath(ExcelPathBox.Text, build.Name);
                    break;
                // 分项生成
                case 1:
                    wordInputPath = WordPathBox.Text;
                    wordSavePath = wordInputPath;
                    break;
            }

            using (var fs = new System.IO.FileStream(wordInputPath, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                XWPFDocument wordDoc = new XWPFDocument(fs);

                // 2张表格：核实概况表，成果汇总表
                var tableA = wordDoc.Tables[0];

                var tableB = wordDoc.Tables[1];

                // 写入核实槪况表 tableA

                // 建设项目名称
                WriteCell(tableA, 1, 1, build.Name);

                // 工程编号
                WriteCell(tableA, 8, 1, ProjNumber.Text);

                // 基底面积（m2）
                build.Area.TryGetValue("基底面积", out string baseArea);
                WriteCell(tableA, 12, 3, baseArea);

                // 计算容积率面积（m2）
                build.Area.TryGetValue("计算容积率面积", out string calcuArea);
                WriteCell(tableA, 13, 3, calcuArea);

                // 阳台面积（m2）
                WriteCell(tableA, 14, 3, build.Balcony);

                // 住宅户数
                WriteCell(tableA, 16, 3, build.House);

                //外墙饰面建筑面积（不取整）
                build.Area.TryGetValue("外墙饰面面积", out string SMMJ);
                Console.WriteLine(SMMJ);
                var SMMJCell = tableA.GetRow(17).GetCell(2);
                SMMJCell.RemoveParagraph(0);
                SMMJCell.SetText(SMMJ.Replace("平方米", ""));

                //计算容积率饰面面积（不取整）
                build.Area.TryGetValue("计算容积率饰面面积", out string JRSM);
                var JRSMCell = tableA.GetRow(17).GetCell(4);
                JRSMCell.RemoveParagraph(0);
                JRSMCell.SetText(JRSM);

                //饰面厚（不取整）
                build.Area.TryGetValue("饰面厚度", out string SMHD);
                var SMHDCell = tableA.GetRow(17).GetCell(6);
                SMHDCell.RemoveParagraph(0);
                SMHDCell.SetText(SMHD);

                //写入成果汇总表 

                //总建筑面积(m2) 地上面积(m2）地下面积(m2)
                build.Area.TryGetValue("总面积∑", out string ZMJStr);
                double ZMJ = StrTrans(ZMJStr);

                build.Area.TryGetValue("地上∑", out string DSMJStr);
                double DSMJ = StrTrans(DSMJStr);

                build.Area.TryGetValue("地下∑", out string DXMJStr);
                double DXMJ = StrTrans(DXMJStr);

                // 处理面积闭合差
                // 地上面积为0，地下面积不为0
                if (DSMJ == 0 && DXMJ != 0)
                {
                    //直接调整字典中的值
                    DiffAdj(build.FPOArea, DXMJ);
                    
                    WriteCell(tableB, 1, 4, ZMJ.ToString());
                    WriteCell(tableB, 2, 5, DSMJ.ToString());
                    WriteCell(tableB, 3, 5, DXMJ.ToString());
                }
                // 地上面积不为0， 地下面积为0
                else if (DSMJ != 0 && DXMJ == 0)
                {
                    DiffAdj(build.FPOArea, DSMJ);
                    WriteCell(tableB, 1, 4, ZMJ.ToString());
                    WriteCell(tableB, 2, 5, DSMJ.ToString());
                    WriteCell(tableB, 3, 5, DXMJ.ToString());
                }
                // 地上、地下面积均不为0
                else if (DSMJ != 0 && DXMJ != 0)
                {
                    DiffAdj(build.FPOArea, ZMJ);
                    WriteCell(tableB, 1, 4, ZMJ.ToString());
                    WriteCell(tableB, 2, 5, DSMJ.ToString());
                    WriteCell(tableB, 3, 5, DXMJ.ToString());
                }

                // 主要功能
                int FRow = 0;
                if (FindCellIndex(tableB, "主要功能") != null)
                {
                    FRow = FindCellIndex(tableB, "主要功能").Item1 + 2;
                }
                else
                {
                    FRow = FindCellIndex(tableB, "主").Item1 + 2;
                }

                foreach (var kvp in build.Function)
                {
                    WriteCell(tableB, FRow, 1, kvp.Key);
                    WriteCell(tableB, FRow, 3, kvp.Value);
                    FRow += 1;
                }

                // 公共服务设施
                int PRow = 0;
                if (FindCellIndex(tableB, "公共服务设施") != null)
                {
                    PRow = FindCellIndex(tableB, "公共服务设施").Item1 + 2;
                }
                else
                {
                    PRow = FindCellIndex(tableB, "公").Item1 + 2;
                }

                foreach (var kvp in build.Public)
                {
                    WriteCell(tableB, PRow, 1, kvp.Key);
                    WriteCell(tableB, PRow, 3, kvp.Value);
                    PRow += 1;
                }
                // 车库配建&其他功能
                foreach (var kvp in build.Other)
                {
                    Tuple<int, int> CellIndex = FindCellIndex(tableB, kvp.Key);

                    // 如果excel表格有匹配项，则直接写入word表格
                    if (CellIndex != null)
                    {
                        WriteCell(tableB, CellIndex.Item1, CellIndex.Item2 + 2, kvp.Value);
                    }
                    // 如果excel表格无匹配项，则写入word表格新增行
                    else if (kvp.Key == "地上机动车库")
                    {
                        Tuple<int, int> RefIndex = FindCellIndex(tableB, "地上汽车库");
                        WriteCell(tableB, RefIndex.Item1, RefIndex.Item2 + 2, kvp.Value);
                    }

                    else if (kvp.Key == "地下机动车库")
                    {
                        Tuple<int, int> RefIndex = FindCellIndex(tableB, "地下汽车库");
                        WriteCell(tableB, RefIndex.Item1, RefIndex.Item2 + 2, kvp.Value);
                    }

                    else if (kvp.Key == "其它")
                    {
                        Tuple<int, int> RefIndex = FindCellIndex(tableB, "其他");
                        WriteCell(tableB, RefIndex.Item1, RefIndex.Item2 + 2, kvp.Value);
                    }

                    else if (kvp.Key == "天面梯屋及机房")
                    {
                        Tuple<int, int> RefIndex = FindCellIndex(tableB, "屋顶梯屋及电梯机房");
                        WriteCell(tableB, RefIndex.Item1, RefIndex.Item2 + 2, kvp.Value);
                    }

                    else 
                    {
                        noMatchItem += kvp.Key + "\n";
                    }
                }
                using (var newfs = new System.IO.FileStream(wordSavePath, System.IO.FileMode.Create, System.IO.FileAccess.Write))
                {
                    wordDoc.Write(newfs);
                }
            }
        }

        private void DiffAdj(Dictionary<string, string> dic, double orgSum)
        {

            int orgIntSum = (int)Math.Round(orgSum, 0);
            int sum = 0;
            double max = 0.00;
            double min = 0.00;
            string maxKey = string.Empty;
            string minKey = string.Empty;

            foreach (var kvp in dic)
            {
                if (kvp.Key != "总面积∑" && kvp.Key != "地上∑" && kvp.Key != "地下∑")
                {

                    double doubleVal = StrTrans(kvp.Value);
                    int intVal = (int)Math.Round(doubleVal, 0);
                    // 取整后增量
                    double diff = doubleVal - intVal;
                    if (diff >= 0 && diff > max)
                    {
                        max = diff;
                        maxKey = kvp.Key;
                    }
                    if (diff < 0 && diff < min)
                    {
                        min = diff;
                        minKey = kvp.Key;
                    }
                    sum += intVal;
                }
            }
                
            if (sum > orgIntSum)
            {
                double candiVal = StrTrans(dic[minKey]);
                dic[minKey] = (candiVal - 1).ToString();
                if (minKey.Contains("地下"))
                {
                    double adjValue = StrTrans(dic["地下∑"]) - 1;
                    dic["地下∑"] = adjValue.ToString();
                }
                else
                {
                    double adjValue = StrTrans(dic["地上∑"]) - 1;
                    dic["地上∑"] = adjValue.ToString();
                }
            }
            else if (sum < orgIntSum)
            {
                double candiVal = StrTrans(dic[maxKey]);
                dic[maxKey] = (candiVal + 1).ToString();
                if (maxKey.Contains("地下"))
                {
                    double adjValue = StrTrans(dic["地下∑"]) + 1;
                    dic["地下∑"] = adjValue.ToString();
                }
                else
                {
                    double adjValue = StrTrans(dic["地上∑"]) + 1;
                    dic["地上∑"] = adjValue.ToString();
                }
            } 
        }

        private double StrTrans(string input)
        {
            if (input != null)
            {
                double.TryParse(input, out double output);
                return output;
            }
            return 0.00;
        }

        private void WriteCell(XWPFTable table, int row, int col, string input)
        {
            var cell = table.GetRow(row).GetCell(col);
            cell.RemoveParagraph(0);

            if (double.TryParse(input, out double output))
            {
                cell.SetText(Math.Round(output, 0).ToString());
            }
            else 
            {
                cell.SetText(input);
            }
        }
        

        private void generateButton_Click(object sender, EventArgs e)
        {
            string buildMessage = string.Empty;
            noMatchItem = string.Empty;
            if (!File.Exists(ExcelPathBox.Text))
            {
                MessageBox.Show("未找到excel文件，路径错误");
            }
            else if (!File.Exists(wordTemplatePath))
            {
                MessageBox.Show("未找到模板文件，路径错误");
            }
            else
            {
                foreach (Build build in builds)
                {
                    // 数据写入
                    WriteTabel(build, 0);
                    buildMessage += build.Name + "\n";
                }
                if (noMatchItem == string.Empty)
                {
                    string path = Path.GetDirectoryName(wordSavePath);
                    MessageBox.Show(buildMessage + "文件保存在:" + path,
                    "核实概况信息写入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("请检查源文件，以下项目字段未匹配:" + noMatchItem,
                    "注意！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private static Tuple<int, int> FindCellIndex(XWPFTable table, string targetValue)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var row = table.GetRow(rowIndex);

                for (int columnIndex = 0; columnIndex < row.GetTableCells().Count; columnIndex++)
                {
                    // 获取当前单元格的值
                    string cellValue = row.GetCell(columnIndex).GetText();

                    // 检查是否与给定值相同
                    if (cellValue == targetValue)
                    {
                        return Tuple.Create(rowIndex, columnIndex);
                    }
                }
            }
            return null;
        }

        private void writeButton_Click(object sender, EventArgs e)
        {
            if (File.Exists(ExcelPathBox.Text) && File.Exists(WordPathBox.Text))
            {
                if (noMatchItem =="")
                {
                    string targetBuildName = buildNameBox.SelectedValue.ToString();
                    Build targetBuild = builds.FirstOrDefault(build => build.Name == targetBuildName);
                    WriteTabel(targetBuild, 1);
                    string path = Path.GetDirectoryName(wordSavePath);
                    MessageBox.Show(targetBuildName + "\n文件已保存在" + wordSavePath,
                    "核实概况信息写入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                 else
                {
                    MessageBox.Show("请检查源文件，以下项目字段未匹配:" + noMatchItem,
                    "注意！", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("未找到 .xlsx|.docx 文件，路径错误");
            }
        }

        // 检核各项核实概况信息是否正确
        private void checkButton_Click(object sender, EventArgs e)
        {
            if (File.Exists(ExcelPathBox.Text) && File.Exists(WordPathBox.Text))
            {
                string targetBuildName = buildNameBox.SelectedValue.ToString();
                Build targetBuild = builds.FirstOrDefault(build => build.Name == targetBuildName);
                CheckInfo(targetBuild);
            }
            else
            {
                MessageBox.Show("未找到 .xlsx|.docx 文件，路径错误");
            }
        }

        private static Dictionary<string, string> GetCheckInfo(Dictionary<string, string> buildInfo, string balcony, string house)
        {
            buildInfo.Add("阳台面积：", balcony);
            buildInfo.Add("住宅户数：", house);
            return null;
        }

        public void CheckInfo(Build build)
        {
            Build wordBuildinfo = new Build();

            using (var fs = new System.IO.FileStream(WordPathBox.Text, System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                XWPFDocument wordDoc = new XWPFDocument(fs);

                var tableA = wordDoc.Tables[0];
                var tableB = wordDoc.Tables[1];
            }
        }

    }


    public class Build
    {
        public string Name;

        public string Balcony;

        public string House;
        public Dictionary<string, string> Area { get; set; }
        public Dictionary<string, string> Function { get; set; }
        public Dictionary<string, string> Public { get; set; }
        public Dictionary<string, string> Other { get; set; }
        public Dictionary<string, string> FPOArea { get; set; }
        public Dictionary<string, string> CheckInfo { get; set; }

    }
}
