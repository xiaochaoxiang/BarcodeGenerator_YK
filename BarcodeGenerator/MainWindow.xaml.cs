using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using BarcodeGenerator.Model;
using BarcodeProcessing;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Brushes = System.Windows.Media.Brushes;
using Newtonsoft.Json;
using Clipboard = System.Windows.Clipboard;
using DataGrid = System.Windows.Controls.DataGrid;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using TextBox = System.Windows.Controls.TextBox;

namespace BarcodeGenerator
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private const string Version = " 1.0.0.0628";
        private const int ConcentrationLength = 10;
        private FileSystemWatcher _watcher;
        private List<string> _files = new List<string>();
        private object _obj = new object();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            Loaded += MainWindow_Loaded;
            //this.Title += Version;
            //_watcher = new FileSystemWatcher(App.WatchFilePath);
            //_watcher.Created += WatcherFile_Created;
            //_watcher.Changed += WatcherFile_Created;
            //_watcher.IncludeSubdirectories = true;
            //_watcher.Filter = "*.json";
            //_watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size;
            //_watcher.EnableRaisingEvents = true;
            //Task.Factory.StartNew(HandleFile);
        }

        private void HandleFile()
        {
            while (true)
            {
                if (_files.Count > 0)
                {
                    try
                    {
                        var file = _files.First();
                        lock (_obj)
                        {
                            _files.RemoveAt(0);
                        }

                        using (FileStream fs = new FileStream(file, FileMode.Open))
                        {
                            using (StreamReader sr = new StreamReader(fs))
                            {
                                var jsonText = sr.ReadToEnd();
                                var entity = JsonConvert.DeserializeObject<JsonEntity>(jsonText);
                                var code = GetIntegrationBarcodeFromJson(entity);
                                if (code != null)
                                {
                                    var jsonFile = file.Substring(0, file.Length - 5) + ".png";
                                    BarcodeHelper.Encode_DM(code, 8).Save(jsonFile);
                                    Dispatcher.InvokeAsync(() =>
                                    {
                                        JsonError.Items.Insert(0,
                                            string.Format("{0:yyyy-MM-dd HH:mm:ss}: 二维码生成成功[{1}]", DateTime.Now, jsonFile));
                                    });
                                }
                                else
                                {
                                    Dispatcher.InvokeAsync(() =>
                                    {
                                        JsonError.Items.Insert(0,
                                            string.Format("{0:yyyy-MM-dd HH:mm:ss}: Json文件数据错误", DateTime.Now));
                                    });
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Dispatcher.InvokeAsync(() =>
                        {
                            JsonError.Items.Insert(0,
                                string.Format("{0:yyyy-MM-dd HH:mm:ss}: Json文件格式错误,{1}", DateTime.Now, ex.Message));
                        });
                    }
                }
                else
                {
                    Thread.Sleep(1000);
                }
            }
        }

        private void WatcherFile_Created(object sender, FileSystemEventArgs e)
        {
            if (!e.Name.Contains("-") && !_files.Contains(e.FullPath))
            {
                lock (_obj)
                {
                    _files.Add(e.FullPath);
                }
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            dpCurveLotNo.SelectedDate = DateTime.Now;
            dpExpirationDate.SelectedDate = DateTime.Now.AddYears(1);
            for (int i = 2; i <= 15; i++)
            {
                cboSize.Items.Add(i);
            }

            for (int i = 1; i <= 10; i++)
            {
                cboSpacer.Items.Add(i);
            }

            tbPath.Text = App.WatchFilePath;
            InitProductInfo();
            for (int i = 0; i < 9; i++)
            {
                _points.Add(new CurvePointModel { Concentration = "0", Response = "0" });
            }

            TbLeftDays.Text = (License.Status.Evaluation_Time - License.Status.Evaluation_Time_Current).ToString();
        }

        private void ButtonCreateBarcode_Click(object sender, RoutedEventArgs e)
        {
            switch (tabBarcode.SelectedIndex)
            {
                case 0:
                    CreateIntegrationBarcode();
                    break;
                case 2:
                    CreateSuperFlexReagentBarcode();
                    break;
            }
        }

        private TextBox _errorTextBox;
        private bool CheckInput()
        {
            if (_errorTextBox != null)
            {
                _errorTextBox.Background = Brushes.White;
                _errorTextBox = null;
            }

            if (dpExpirationDate.SelectedDate == null)
            {
                MessageBox.Show(GetStringByKey("LangErrorExpirationDateNull"));
                return false;
            }

            if (dpCurveLotNo.SelectedDate == null)
            {
                MessageBox.Show(GetStringByKey("LangErrorCurveNo"));
                return false;
            }

            long temp;
            float fTemp;
            if (txtItemNo.Text.Trim().Length != 4)
            {
                _errorTextBox = txtItemNo;
                MessageBox.Show(GetStringByKey("LangErrorItemNoLength"));
                return false;
            }

            Regex regex = new Regex(@"^[0-9][0-9a-zA-Z][0-9]{2}$");
            var result = regex.IsMatch(txtItemNo.Text.Trim());
            if (!result)
            {
                _errorTextBox = txtItemNo;
                MessageBox.Show(GetStringByKey("LangErrorItemNo"));
                return false;
            }


            if (!long.TryParse(txtLotNo.Text, out temp) || txtLotNo.Text.Trim().Length > 10)
            {
                _errorTextBox = txtLotNo;
                MessageBox.Show(GetStringByKey("LangErrorLotNo"));
                return false;
            }
            if (txtLotNo.Text.Trim().Length < 10)
            {
                txtLotNo.Text = txtLotNo.Text.Trim().PadLeft(10, '0');
            }

            int index = 0;
            foreach (var point in Points)
            {
                if (!long.TryParse(point.Response, out temp) || temp < 0 || point.Response.Trim().Length > 8)
                {
                    MessageBox.Show(string.Format(GetStringByKey("LangErrorResponse"), index + 1));
                    return false;
                }
                if (!float.TryParse(point.Concentration, out fTemp) || fTemp < 0 || point.Concentration.Trim().Length > ConcentrationLength)
                {
                    MessageBox.Show(string.Format(GetStringByKey("LangErrorConcentration"), index + 1));
                    return false;
                }
                index++;
            }

            if (!long.TryParse(txtCaliLotNo1.Text, out temp) || txtCaliLotNo1.Text.Trim().Length > 10)
            {
                _errorTextBox = txtCaliLotNo1;
                MessageBox.Show(GetStringByKey("LangErrorLotNoS1"));
                return false;
            }
            if (txtCaliLotNo1.Text.Trim().Length < 10)
            {
                txtCaliLotNo1.Text = txtCaliLotNo1.Text.Trim().PadLeft(10, '0');
            }
            if (!long.TryParse(txtCaliLotNo2.Text, out temp) || txtCaliLotNo2.Text.Trim().Length > 10)
            {
                _errorTextBox = txtCaliLotNo2;
                MessageBox.Show(GetStringByKey("LangErrorLotNoS2"));
                return false;
            }
            if (txtCaliLotNo2.Text.Trim().Length < 10)
            {
                txtCaliLotNo2.Text = txtCaliLotNo2.Text.Trim().PadLeft(10, '0');
            }
            if (!float.TryParse(txtCaliConc1.Text, out fTemp) || txtCaliConc1.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = txtCaliConc1;
                MessageBox.Show(GetStringByKey("LangErrorConcS1"));
                return false;
            }
            if (!float.TryParse(txtCaliConc2.Text, out fTemp) || txtCaliConc2.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = txtCaliConc2;
                MessageBox.Show(GetStringByKey("LangErrorConcS2"));
                return false;
            }
            if (!long.TryParse(txtReagentLotNo.Text, out temp) || txtReagentLotNo.Text.Trim().Length > 10)
            {
                _errorTextBox = txtReagentLotNo;
                MessageBox.Show(GetStringByKey("LangErrorReagentLotNo"));
                return false;
            }
            if (txtReagentLotNo.Text.Trim().Length < 10)
            {
                txtReagentLotNo.Text = txtReagentLotNo.Text.Trim().PadLeft(10, '0');
            }

            if (dpExpirationDate.SelectedDate < DateTime.Today)
            {
                MessageBox.Show(GetStringByKey("LangErrorExpirationDate"));
                return false;
            }


            if (!long.TryParse(tbLotNoC1.Text, out temp) || tbLotNoC1.Text.Trim().Length > 10)
            {
                _errorTextBox = tbLotNoC1;
                MessageBox.Show(GetStringByKey("LangErrorLotNoC1"));
                return false;
            }
            if (tbLotNoC1.Text.Trim().Length < 10)
            {
                tbLotNoC1.Text = tbLotNoC1.Text.Trim().PadLeft(10, '0');
            }         
            if (!float.TryParse(tbTargetValueC1.Text, out fTemp) || tbTargetValueC1.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbTargetValueC1;
                MessageBox.Show(GetStringByKey("LangErrorTargetValueC1"));
                return false;
            }
            if (!float.TryParse(tbLowerLimitC1.Text, out fTemp) || tbLowerLimitC1.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbLowerLimitC1;
                MessageBox.Show(GetStringByKey("LangErrorLowerLimitC1"));
                return false;
            }
            if (!float.TryParse(tbUpperLimitC1.Text, out fTemp) || tbUpperLimitC1.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbUpperLimitC1;
                MessageBox.Show(GetStringByKey("LangErrorUpperLimitC1"));
                return false;
            }

            if (!long.TryParse(tbLotNoC2.Text, out temp) || tbLotNoC2.Text.Trim().Length > 10)
            {
                _errorTextBox = tbLotNoC2;
                MessageBox.Show(GetStringByKey("LangErrorLotNoC2"));
                return false;
            }
            if (tbLotNoC2.Text.Trim().Length < 10)
            {
                tbLotNoC2.Text = tbLotNoC2.Text.Trim().PadLeft(10, '0');
            }
            if (!float.TryParse(tbTargetValueC2.Text, out fTemp) || tbTargetValueC2.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbTargetValueC2;
                MessageBox.Show(GetStringByKey("LangErrorTargetValueC2"));
                return false;
            }
            if (!float.TryParse(tbLowerLimitC2.Text, out fTemp) || tbLowerLimitC2.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbLowerLimitC2;
                MessageBox.Show(GetStringByKey("LangErrorLowerLimitC2"));
                return false;
            }
            if (!float.TryParse(tbUpperLimitC2.Text, out fTemp) || tbUpperLimitC2.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = tbUpperLimitC2;
                MessageBox.Show(GetStringByKey("LangErrorUpperLimitC2"));
                return false;
            }

            if (!float.TryParse(txtUncertain1.Text, out fTemp) || txtUncertain1.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = txtUncertain1;
                MessageBox.Show(GetStringByKey("LangErrorUncertain1"));
                return false;
            }

            if (!float.TryParse(txtUncertain2.Text, out fTemp) || txtUncertain2.Text.Trim().Length > ConcentrationLength)
            {
                _errorTextBox = txtUncertain2;
                MessageBox.Show(GetStringByKey("LangErrorUncertain2"));
                return false;
            }

            regex = new Regex(@"^[0-9a-fA-F]{1}$");
            result = regex.IsMatch(TbVersion.Text.Trim());
            if (!result)
            {
                _errorTextBox = TbVersion;
                MessageBox.Show(GetStringByKey("LangErrorVersion"));
                return false;
            }

            //if (_generateType == 1)
            {
                if (!long.TryParse(tbLotNoC3.Text, out temp) || tbLotNoC3.Text.Trim().Length > 10)
                {
                    _errorTextBox = tbLotNoC3;
                    MessageBox.Show(GetStringByKey("LangErrorLotNoC3"));
                    return false;
                }

                if (tbLotNoC3.Text.Trim().Length < 10)
                {
                    tbLotNoC3.Text = tbLotNoC3.Text.Trim().PadLeft(10, '0');
                }

                if (!float.TryParse(tbTargetValueC3.Text, out fTemp) ||
                    tbTargetValueC3.Text.Trim().Length > ConcentrationLength)
                {
                    _errorTextBox = tbTargetValueC3;
                    MessageBox.Show(GetStringByKey("LangErrorTargetValueC3"));
                    return false;
                }

                if (!float.TryParse(tbLowerLimitC3.Text, out fTemp) || tbLowerLimitC3.Text.Trim().Length >
                                                                    ConcentrationLength)
                {
                    _errorTextBox = tbLowerLimitC3;
                    MessageBox.Show(GetStringByKey("LangErrorLowerLimitC3"));
                    return false;
                }

                if (!float.TryParse(tbUpperLimitC3.Text, out fTemp) || tbUpperLimitC3.Text.Trim().Length >
                                                                    ConcentrationLength)
                {
                    _errorTextBox = tbUpperLimitC3;
                    MessageBox.Show(GetStringByKey("LangErrorUpperLimitC3"));
                    return false;
                }
            }

            return true;
        }

        private void CreateIntegrationBarcode()
        {
            try
            {
                if (CheckInput())
                {
                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "Excel file(*.xls)|*.xls|Text file(*.txt)|*.txt",
                        InitialDirectory = Directory.GetCurrentDirectory()
                    };
                    if (sfd.ShowDialog() == true)
                    {
                        var code = GenerateIntegrationBarcode();

                        var ext = sfd.FileName.Substring(sfd.FileName.Length - 3, 3);
                        if (ext == "xls")
                        {
                            ExcelHelper.ExportBarcodeList(sfd.FileName, new List<string> {code}, "IntegrationBarcode");
                        }
                        else
                        {
                            ExportToTxt(sfd.FileName, new List<string> {code});
                        }

                        MessageBox.Show(GetStringByKey("LangBarcodeGenerateSuccess"));
                    }
                }
                else
                {
                    if (_errorTextBox != null)
                    {
                        _errorTextBox.Background = Brushes.Red;
                        _errorTextBox.Focus();
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private string GetIntegrationBarcodeFromJson(JsonEntity entity)
        {
            var itemNo = entity.Project.PadLeft(4, '0');
            StringBuilder sbContent = new StringBuilder();
            sbContent.Append(itemNo);
            sbContent.Append(entity.Lot.PadLeft(10, '0'));
            sbContent.Append(entity.StdCalibrateLot.PadLeft(8, '0'));

            //发光值，8个字符
            sbContent.Append(entity.StdItemALight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemBLight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemCLight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemDLight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemELight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemFLight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemGLight.PadLeft(8, '0'));
            sbContent.Append(entity.StdItemHLight.PadLeft(8, '0'));

            //浓度
            sbContent.Append(entity.StdItemAdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemBdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemCdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemDdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemEdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemFdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemGdensity.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.StdItemHdensity.PadLeft(ConcentrationLength, '0'));

            sbContent.Append(entity.S1Lot.PadLeft(10, '0'));
            sbContent.Append(entity.S2Lot.PadLeft(10, '0'));
            sbContent.Append(entity.S1Density.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.S1Uncertain.PadLeft(10, '0'));
            sbContent.Append(entity.S2Density.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.S2Uncertain.PadLeft(10, '0'));

            //批号
            sbContent.Append(entity.C1Lot.PadLeft(10, '0'));
            sbContent.Append(entity.C2Lot.PadLeft(10, '0'));
            //靶值+下限+上限
            sbContent.Append(entity.C1Target.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.C1DensityDownLimit.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.C1DensityUpLimit.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.C2Target.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.C2DensityDownLimit.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.C2DensityUpLimit.PadLeft(ConcentrationLength, '0'));
            sbContent.Append(entity.ReagentSnr.PadLeft(10, '0'));
            sbContent.Append(entity.ExpiredDate.PadLeft(8, '0'));
            sbContent.Append(entity.StdCalibrateRev);

            if (itemNo.StartsWith("1"))
            {
                sbContent.Append(SuperFlexReserved);
            }
            else
            {
                sbContent.Append(entity.C3Lot.PadLeft(10, '0'));
                sbContent.Append(entity.C3Target.PadLeft(ConcentrationLength, '0'));
                sbContent.Append(entity.C3DensityDownLimit.PadLeft(ConcentrationLength, '0'));
                sbContent.Append(entity.C3DensityUpLimit.PadLeft(ConcentrationLength, '0'));
                sbContent.Append(entity.ProductSpecification);
                sbContent.Append(entity.CapacityOfReagent);
                sbContent.Append(SuperEfficReserved);
            }

            var code = BarcodeEncodeDecodeHelper.EncodeIntegrationBarcode(sbContent.ToString());
            return code;
        }

        private string GenerateIntegrationBarcode()
        {
            StringBuilder sbContent = new StringBuilder();
            sbContent.Append(txtItemNo.Text.Trim().PadLeft(4, '0'));
            sbContent.Append(txtLotNo.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(dpCurveLotNo.SelectedDate.Value.ToString("yyyyMMdd"));
   //         for(int i=0;i< Points.Count-1;i++)
			//{
   //             sbContent.Append(Points[i].Response.Trim().PadLeft(8, '0'));
   //         }
   //         for(int i=0;i<Points.Count-1;i++)
			//{
   //             sbContent.Append(Points[i].Concentration.Trim().PadLeft(ConcentrationLength, '0'));
   //         }
            foreach (var point in Points)
            {
                sbContent.Append(point.Response.Trim().PadLeft(8, '0'));
            }
            foreach (var point in Points)
            {
                sbContent.Append(point.Concentration.Trim().PadLeft(ConcentrationLength, '0'));
            }

            sbContent.Append(txtCaliLotNo1.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(txtCaliLotNo2.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(txtCaliConc1.Text.Trim().PadLeft(ConcentrationLength, '0'));
            sbContent.Append(txtUncertain1.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(txtCaliConc2.Text.Trim().PadLeft(ConcentrationLength, '0'));
            sbContent.Append(txtUncertain2.Text.Trim().PadLeft(10, '0'));

            //批号
            sbContent.Append(tbLotNoC1.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(tbLotNoC2.Text.Trim().PadLeft(10, '0'));
            //靶值+下限+上限
            sbContent.Append(tbTargetValueC1.Text.Trim().PadLeft(ConcentrationLength, '0'));
            sbContent.Append(tbLowerLimitC1.Text.Trim().PadLeft(ConcentrationLength, '0'));
            sbContent.Append(tbUpperLimitC1.Text.Trim().PadLeft(ConcentrationLength, '0'));

            sbContent.Append(tbTargetValueC2.Text.Trim().PadLeft(ConcentrationLength, '0'));
            sbContent.Append(tbLowerLimitC2.Text.Trim().PadLeft(ConcentrationLength, '0'));   
            sbContent.Append(tbUpperLimitC2.Text.Trim().PadLeft(ConcentrationLength, '0'));

            sbContent.Append(txtReagentLotNo.Text.Trim().PadLeft(10, '0'));
            sbContent.Append(dpExpirationDate.SelectedDate.Value.ToString("yyyyMMdd"));
            sbContent.Append(TbVersion.Text.Trim());

            //if (_generateType == 0)
            //{
            //    sbContent.Append(SuperFlexReserved);
            //}
            //else
            {
                sbContent.Append(tbLotNoC3.Text.Trim().PadLeft(10, '0'));
                sbContent.Append(tbTargetValueC3.Text.Trim().PadLeft(ConcentrationLength, '0'));
                sbContent.Append(tbLowerLimitC3.Text.Trim().PadLeft(ConcentrationLength, '0'));
                sbContent.Append(tbUpperLimitC3.Text.Trim().PadLeft(ConcentrationLength, '0'));

                sbContent.Append(cboProductSpecification.SelectedIndex);
                sbContent.Append(cboCapacityOfReagent.SelectedIndex);
                sbContent.Append(SuperEfficReserved);
                //9
                var point = Points.LastOrDefault();
                sbContent.Append(point.Response.Trim().PadLeft(8, '0'));
                sbContent.Append(point.Concentration.Trim().PadLeft(ConcentrationLength, '0'));
                //c4
                sbContent.Append(tbLotNoC4.Text.Trim().PadLeft(10, '0'));
				sbContent.Append(tbTargetValueC4.Text.Trim().PadLeft(ConcentrationLength, '0'));
				sbContent.Append(tbLowerLimitC4.Text.Trim().PadLeft(ConcentrationLength, '0'));
				sbContent.Append(tbUpperLimitC4.Text.Trim().PadLeft(ConcentrationLength, '0'));
            }

            var code = BarcodeEncodeDecodeHelper.EncodeIntegrationBarcode(sbContent.ToString());
            return code;
        }

        /// <summary>
        /// 成品规格代码
        /// </summary>
        private string _productSpecificationCode;

        /// <summary>
        /// 试剂装量代码
        /// </summary>
        private string _capacityOfReagentCode;

        /// <summary>
        /// 预留19位
        /// </summary>
        private const string SuperEfficReserved = "0000000000000000000";

        /// <summary>
        /// 预留13位
        /// </summary>
        private const string SuperFlexReserved = "0000000000000";

        private void CreateSuperFlexReagentBarcode()
        {
            try
            {
                long temp;
                if (!long.TryParse(txtLotNoR.Text, out temp) || txtLotNoR.Text.Trim().Length != 10)
                {
                    MessageBox.Show(GetStringByKey("LangErrorReagentLotNo1"));
                    return;
                }

                int startIndex;
                if (!int.TryParse(txtStartIndexR.Text, out startIndex) || startIndex <= 0 || startIndex > 999999)
                {
                    MessageBox.Show(GetStringByKey("LangErrorStartIndex"));
                    return;
                }

                int num;
                if (!int.TryParse(txtNumR.Text, out num) || num <= 0 || num > 999999)
                {
                    MessageBox.Show(GetStringByKey("LangErrorNum"));
                    return;
                }

                if (startIndex + num > 1000000)
                {
                    MessageBox.Show(GetStringByKey("LangErrorTotal"));
                    return;
                }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel(*.xls)|*.xls|Text file(*.txt)|*.txt",
                    InitialDirectory = Directory.GetCurrentDirectory()
                };
                if (sfd.ShowDialog() == true)
                {
                    var fileName = sfd.FileName;
                    var lotNo = txtLotNoR.Text.Trim().PadLeft(10, '0');
                    Task.Factory.StartNew(() =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            gridMain.IsEnabled = false;
                            tbHint.Text = GetStringByKey("LangGeneratingBarcode");
                        });
                        List<string> barcodeList = new List<string>();
                        for (int i = startIndex; i < startIndex + num; i++)
                        {
                            var tempCode = string.Format("{0}{1}", lotNo, i.ToString().PadLeft(6, '0'));
                            tempCode = BarcodeEncodeDecodeHelper.EncodeReagentBarcode(tempCode);
                            barcodeList.Add(tempCode);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            tbHint.Text = GetStringByKey("LangExporting");
                        });

                        var ext = fileName.Substring(fileName.Length - 3, 3);
                        if (ext == "xls")
                        {
                            ExcelHelper.ExportBarcodeList(fileName, barcodeList, "ReagentBarcode");
                        }
                        else
                        {
                            ExportToTxt(fileName, barcodeList);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            gridMain.IsEnabled = true;
                            tbHint.Text = GetStringByKey("LangBarcodeGenerateSuccess");
                            GC.Collect();
                            MessageBox.Show(GetStringByKey("LangBarcodeGenerateSuccess"));
                        });
                    });
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void ButtonExit_OnClick(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void ExportToTxt(string fileName, List<string> barcodeList)
        {
            FileStream fs = new FileStream(fileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            for (int i = 0; i < barcodeList.Count; i++)
            {
                sw.Write(barcodeList[i]);
                if (i != barcodeList.Count - 1)
                {
                    sw.Write("\r\n");
                }
            }
            sw.Flush();
            sw.Close();
            fs.Close();
        }

        private void ButtonParseBarcode_OnClick(object sender, RoutedEventArgs e)
        {
            string lotNo = null, serialNo = null;
            if (GetReagentBarcode(tbReagentBarcode.Text.Trim(), ref lotNo, ref serialNo))
            {
                tbLotNo.Text = lotNo;
                tbSerialNo.Text = serialNo;
            }
            else
            {
                MessageBox.Show(GetStringByKey("LangErrorBarcodeParse"));
            }
        }

        private void ButtonClear_OnClick(object sender, RoutedEventArgs e)
        {
            tbReagentBarcode.Text = string.Empty;
            tbLotNo.Text = string.Empty;
            tbSerialNo.Text = string.Empty;
            tbReagentBarcode.Focus();
        }

        private bool GetReagentBarcode(string barcode, ref string lotNo, ref string serialNo)
        {
            if (string.IsNullOrEmpty(barcode))
                return false;
            if (barcode.Length != 16)
                return false;
            barcode = BarcodeEncodeDecodeHelper.DecodeReagentBarcode(barcode);
            lotNo = barcode.Substring(0, 10);
            serialNo = barcode.Substring(10, 6);
            return true;
        }

        private void ButtonParseIntegBarcode_OnClick(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(pTbBarcode.Text))
            {
                ParseBarcode(pTbBarcode.Text.Trim());
            }
        }

        private void ParseBarcode(string code)
        {
            var barcode = BarcodeEncodeDecodeHelper.DecodeIntegrationBarcode(code);
            if (barcode != null && barcode.Length == BarcodeEncodeDecodeHelper.SuperEfficBarcodeLength)
            {
                pTbItemNo.Text = barcode.Substring(0, 4).ToUpper();
                pTbLotNo.Text = barcode.Substring(4, 10);
                pTbCurveNo.Text = barcode.Substring(14, 8);

                int index = 22;
                //发光值
                for (int i = 0; i < 8; i++)
                {
                    (pSpResponse.Children[i + 1] as TextBox).Text =
                        Convert.ToInt32(barcode.Substring(index, 8)).ToString();
                    index += 8;
                }

                //浓度
                for (int i = 0; i < 8; i++)
                {
                    (pSpConc.Children[i + 1] as TextBox).Text =
                        GetDisplayText(barcode.Substring(index, ConcentrationLength));
                    index += ConcentrationLength;
                }

                pTbLot1.Text = barcode.Substring(index, 10);
                index += 10;
                pTbLot2.Text = barcode.Substring(index, 10);
                index += 10;
                pTbConc1.Text = GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                pTbUncertain1.Text = GetDisplayText(barcode.Substring(index, ConcentrationLength));
                //pTbUncertain1.Text = Convert.ToSingle(barcode.Substring(index, ConcentrationLength)).ToString();
                index += ConcentrationLength;
                pTbConc2.Text = GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                pTbUncertain2.Text = GetDisplayText(barcode.Substring(index, ConcentrationLength));
                //pTbUncertain2.Text = Convert.ToSingle(barcode.Substring(index, ConcentrationLength)).ToString();
                index += ConcentrationLength;
                //qc
                ptbLotNoC1.Text = barcode.Substring(index, 10);
                index += 10;
                ptbLotNoC2.Text = barcode.Substring(index, 10);
                index += 10;
                ptbTargetValueC1.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                ptbLowerLimitC1.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                ptbUpperLimitC1.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                ptbTargetValueC2.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                ptbLowerLimitC2.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;
                ptbUpperLimitC2.Text =
                    GetDisplayText(barcode.Substring(index, ConcentrationLength));
                index += ConcentrationLength;

                pTbReagentLotNo.Text = barcode.Substring(index, 10);
                index += 10;
                tbExpirationDate.Text = barcode.Substring(index, 8);
                index += 8;
                var version = barcode.Substring(index, 1);
                index += 1;
                pTbVersion.Text = version.Replace('.', 'A').ToUpper();

                //if (barcode.Length == BarcodeEncodeDecodeHelper.SuperEfficBarcodeLength)
                {
                    //_parseType = 1;
                    spEfficModel.Visibility = Visibility.Visible;
                    gpEfficModel.Visibility = Visibility.Visible;

                    //qc3
                    ptbLotNoC3.Text = barcode.Substring(index, 10);
                    index += 10;
                    ptbTargetValueC3.Text =
                        GetDisplayText(barcode.Substring(index, ConcentrationLength));
                    index += ConcentrationLength;
                    ptbLowerLimitC3.Text =
                        GetDisplayText(barcode.Substring(index, ConcentrationLength));
                    index += ConcentrationLength;
                    ptbUpperLimitC3.Text =
                        GetDisplayText(barcode.Substring(index, ConcentrationLength));
                    index += ConcentrationLength;

                    var strProductSpecificationCode = barcode.Substring(index, 1);
                    index += 1;
                    var strCapacityOfReagentCode = barcode.Substring(index, 1);
                    index += 1;

                    if (int.TryParse(strProductSpecificationCode, out var iProductSpecification))
                    {
                        var model = _models.FirstOrDefault(m => m.Code == iProductSpecification);
                        if (model != null)
                        {
                            pTbProductSpecification.Text = model.Message;
                        }
                    }

                    if (int.TryParse(strCapacityOfReagentCode, out var iCapacityOfReagentCode))
                    {
                        var model = _models.FirstOrDefault(m => m.Code == iCapacityOfReagentCode);
                        if (model != null)
                        {
                            pTbCapacityOfReagent.Text = model.Message;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show(GetStringByKey("LangErrorBarcodeParse"));
            }
        }

        private void ButtonSavePng_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (CheckInput())
                {
                    SaveFileDialog sfd = new SaveFileDialog
                    {
                        Filter = "PNG(*.png)|*.png"
                    };
                    if (sfd.ShowDialog() == true)
                    {
                        var code = GenerateIntegrationBarcode();
                        BarcodeHelper.Encode_DM(code, cboSize.SelectedIndex + 2, cboSpacer.SelectedIndex + 1)
                            .Save(sfd.FileName);
                        MessageBox.Show(GetStringByKey("LangSaveSuccess"));
                    }
                }
                else
                {
                    if (_errorTextBox != null)
                    {
                        _errorTextBox.Background = Brushes.Red;
                        _errorTextBox.Focus();
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void ButtonSelectPng_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog ofd = new OpenFileDialog {Filter = "PNG|*.png"};
                if (ofd.ShowDialog() == true)
                {
                    ParseBarcode(pTbBarcode.Text = BarcodeHelper.Decode(new Bitmap(ofd.FileName)));
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void ButtonClearAll_OnClick(object sender, RoutedEventArgs e)
        {
            pTbBarcode.Text = string.Empty;
            pTbItemNo.Text = string.Empty;
            pTbLotNo.Text = string.Empty;
            pTbCurveNo.Text = string.Empty;
            
            foreach (var child in pSpResponse.Children)
            {
                var tb = child as TextBox;
                if (tb != null)
                {
                    tb.Text = string.Empty;
                }
            }

            foreach (var child in pSpConc.Children)
            {
                var tb = child as TextBox;
                if (tb != null)
                {
                    tb.Text = string.Empty;
                }
            }

            pTbConc1.Text = string.Empty;
            pTbLot1.Text = string.Empty;
            pTbConc2.Text = string.Empty;
            pTbLot2.Text = string.Empty;
            pTbReagentLotNo.Text = string.Empty;
            tbExpirationDate.Text = string.Empty;

            ptbLotNoC1.Text = string.Empty;
            ptbTargetValueC1.Text = string.Empty;
            ptbLowerLimitC1.Text = string.Empty;
            ptbUpperLimitC1.Text = string.Empty;
            ptbLotNoC2.Text = string.Empty;
            ptbTargetValueC2.Text = string.Empty;
            ptbLowerLimitC2.Text = string.Empty;
            ptbUpperLimitC2.Text = string.Empty;
            pTbProductSpecification.Text = string.Empty;
            pTbCapacityOfReagent.Text = string.Empty;

            ptbLotNoC3.Text = string.Empty;
            ptbTargetValueC3.Text = string.Empty;
            ptbLowerLimitC3.Text = string.Empty;
            ptbUpperLimitC3.Text = string.Empty;
            pTbUncertain1.Text = string.Empty;
            pTbUncertain2.Text = string.Empty;
        }

        private void CboSize_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int size = 80 + cboSize.SelectedIndex*40;
            int margin = (cboSpacer.SelectedIndex + 1)*2;
            var trueSize = size + margin;
            tbSize.Text = string.Format("{0}*{1}", trueSize, trueSize);
        }

        private readonly ShowConcentrationFormat _picWindow = new ShowConcentrationFormat();
        private void Image_PreviewMouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _picWindow.WindowState = WindowState.Normal;
            _picWindow.Show();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (_picWindow != null)
            {
                _picWindow.IsCloseApp = true;
                _picWindow.Close();
            }
        }

        private string GetStringByKey(string key)
        {
            string result = string.Empty;
            try
            {
                result = FindResource(key).ToString();
            }
            catch (Exception exception)
            {
                result = string.Empty;
            }

            return result;
        }

        private static readonly List<string> ExportName = new List<string>
        {
            "项目号", "成品批号", "失效日期", "工作校准品批号", "浓度1", "浓度2", "浓度3", "浓度4", "浓度5", "浓度6", "浓度7", "浓度8",
            "发光值1", "发光值2", "发光值3", "发光值4", "发光值5", "发光值6", "发光值7", "发光值8", "S1批号", "S1浓度", "S1不确定度", "S2批号", "S2浓度",
            "S2不确定度", "C1批号", "C1靶值", "C1浓度下限", "C1浓度上限", "C2批号", "C2靶值", "C2浓度下限", "C2浓度上限", "试剂批号","企标版本号","C3批号", "C3靶值", "C3浓度下限", "C3浓度上限","成品规格","试剂装量"
        };

        private static readonly List<string> ExportNameEn = new List<string>
        {
            "ItemNo", "LotNo", "ExpirationDate", "CurveLotNo", "Concentration1", "Concentration2", "Concentration3",
            "Concentration4", "Concentration5", "Concentration6", "Concentration7", "Concentration8",
            "Response1", "Response2", "Response3", "Response4", "Response5", "Response6", "Response7", "Response8",
            "S1 LotNo", "S1 Concentration", "S1 Uncertain", "S2 LotNo", "S2 Concentration",
            "S2 Uncertain", "C1 LotNo", "C1 TargetValue", "C1 ConcentrationLowerLimit", "C1 ConcentrationUpperLimit",
            "C2 LotNo", "C2 TargetValue", "C2 ConcentrationLowerLimit", "C2 ConcentrationUpperLimit", "ReagentLotNo","Version","C3 LotNo", "C3 TargetValue", "C3 ConcentrationLowerLimit", "C3 ConcentrationUpperLimit","Product specification","Capacity of reagent"
        };

        private void Button_Click_Export(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel file(*.xls)|*.xls",
                    InitialDirectory = Directory.GetCurrentDirectory()
                };
                if (sfd.ShowDialog() == true)
                {
                    FileStream fs = File.Create(sfd.FileName);
                    IWorkbook xssfworkbook = new XSSFWorkbook();
                    var sheet = xssfworkbook.CreateSheet("sheet0");

                    int count = 0;
                    if (_parseType == 0)
                    {
                        count = 36;
                    }
                    else if (_parseType == 1)
                    {
                        count = 42;
                    }
                    if (App.IsEnglish)
                    {
                        for (int i = 0; i < count; i++)
                        {
                            GetCell(sheet, i, 0).SetCellValue(ExportNameEn[i]);
                        }
                    }
                    else
                    {
                        for (int i = 0; i < count; i++)
                        {
                            GetCell(sheet, i, 0).SetCellValue(ExportName[i]);
                        }
                    }

                    GetCell(sheet, 0, 1).SetCellValue(pTbItemNo.Text);
                    GetCell(sheet, 1, 1).SetCellValue(pTbLotNo.Text);
                    GetCell(sheet, 2, 1).SetCellValue(tbExpirationDate.Text);
                    GetCell(sheet, 3, 1).SetCellValue(pTbCurveNo.Text);
                    for (int i = 0; i < pSpConc.Children.Count-1; i++)
                    {
                        GetCell(sheet, i + 4, 1).SetCellValue((pSpConc.Children[i + 1] as TextBox).Text);
                    }

                    for (int i = 0; i < pSpConc.Children.Count-1; i++)
                    {
                        GetCell(sheet, i + 12, 1).SetCellValue((pSpResponse.Children[i + 1] as TextBox).Text);
                    }

                    GetCell(sheet, 20, 1).SetCellValue(pTbLot1.Text);
                    GetCell(sheet, 21, 1).SetCellValue(pTbConc1.Text);
                    GetCell(sheet, 22, 1).SetCellValue(pTbUncertain1.Text);
                    GetCell(sheet, 23, 1).SetCellValue(pTbLot2.Text);
                    GetCell(sheet, 24, 1).SetCellValue(pTbConc2.Text);
                    GetCell(sheet, 25, 1).SetCellValue(pTbUncertain2.Text);
                    GetCell(sheet, 26, 1).SetCellValue(ptbLotNoC1.Text);
                    GetCell(sheet, 27, 1).SetCellValue(ptbTargetValueC1.Text);
                    GetCell(sheet, 28, 1).SetCellValue(ptbLowerLimitC1.Text);
                    GetCell(sheet, 29, 1).SetCellValue(ptbUpperLimitC1.Text);
                    GetCell(sheet, 30, 1).SetCellValue(ptbLotNoC2.Text);
                    GetCell(sheet, 31, 1).SetCellValue(ptbTargetValueC2.Text);
                    GetCell(sheet, 32, 1).SetCellValue(ptbLowerLimitC2.Text);
                    GetCell(sheet, 33, 1).SetCellValue(ptbUpperLimitC2.Text);
                    GetCell(sheet, 34, 1).SetCellValue(pTbReagentLotNo.Text);
                    GetCell(sheet, 35, 1).SetCellValue(pTbVersion.Text);
                    if (_parseType == 1)
                    {
                        GetCell(sheet, 36, 1).SetCellValue(ptbLotNoC3.Text);
                        GetCell(sheet, 37, 1).SetCellValue(ptbTargetValueC3.Text);
                        GetCell(sheet, 38, 1).SetCellValue(ptbLowerLimitC3.Text);
                        GetCell(sheet, 39, 1).SetCellValue(ptbUpperLimitC3.Text);
                        GetCell(sheet, 40, 1).SetCellValue(pTbProductSpecification.Text);
                        GetCell(sheet, 41, 1).SetCellValue(pTbCapacityOfReagent.Text);
                    }

                    xssfworkbook.Write(fs);
                    fs.Close();

                    MessageBox.Show(GetStringByKey("LangExportSuccess"));
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private static XSSFCell GetCell(ISheet sheet, int row, int col)
        {
            var tempRow = sheet.GetRow(row) as XSSFRow ?? sheet.CreateRow(row) as XSSFRow;
            var cell = tempRow.GetCell(col) as XSSFCell ?? tempRow.CreateCell(col) as XSSFCell;
            return cell;
        }

        private void Button_Click_ChoosePath(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tbPath.Text = App.WatchFilePath = fbd.SelectedPath;
                _watcher.Path = App.WatchFilePath;
            }
        }

        private List<SpecificationModel> _models = new List<SpecificationModel>();
        private List<ProductNumberModel> _products = new List<ProductNumberModel>();
        private void InitProductInfo()
        {
            _models.Clear();
            var file = AppDomain.CurrentDomain.BaseDirectory + @"Data\Specifications.txt";
            if (File.Exists(file))
            {
                StreamReader sr = new StreamReader(file, Encoding.UTF8);
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!line.StartsWith(";"))
                    {
                        var temp = line.Split(new[] {",", "\t"}, StringSplitOptions.None);
                        if (temp.Length == 2 && int.TryParse(temp[0], out var code) && code >= 0)
                        {
                            _models.Add(new SpecificationModel {Code = code, Message = temp[1]});
                        }
                    }
                }
            }
            else
            {
                _models.Add(new SpecificationModel {Code = 0, Message = "50"});
                _models.Add(new SpecificationModel {Code = 1, Message = "100"});
                _models.Add(new SpecificationModel {Code = 2, Message = "200"});
            }

            _products.Clear();
            file = AppDomain.CurrentDomain.BaseDirectory + @"Data\Products.txt";
            if (File.Exists(file))
            {
                StreamReader sr = new StreamReader(file, Encoding.UTF8);
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!line.StartsWith(";"))
                    {
                        var tmp = line.Split(new[] {",", "\t"}, StringSplitOptions.None);
                        if (tmp.Length >= 3)
                        {
                            var product = new ProductNumberModel
                                {ProductNumber = tmp[0], ProductName = tmp[1], ItemNo = tmp[2]};
                            if (tmp.Length >= 4)
                            {
                                if (int.TryParse(tmp[3], out var code1) && code1 >= 0)
                                {
                                    product.ProductSpecificationCode = code1;
                                }

                                if (tmp.Length >= 5 && int.TryParse(tmp[4], out var code2) && code2 >= 0)
                                {
                                    product.CapacityOfReagentCode = code2;
                                }
                            }

                            _products.Add(product);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 平台（生成二维码），0-SuperFlex，1-SuperEffic
        /// </summary>
        private int _generateType = 1;
        //private void cboProducts_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    if (e.AddedItems[0] is ProductNumberModel model)
        //    {
        //        txtItemNo.Text = model.ItemNo;
        //        if (model.ItemNo.StartsWith("1"))//SuperFlex
        //        {
        //            _generateType = 0;
        //            spEfficInfo.Visibility = Visibility.Collapsed;
        //            gpEfficInfo.Visibility = Visibility.Collapsed;
        //        }
        //        else if(model.ItemNo.StartsWith("2"))//SuperEffic
        //        {
        //            _generateType = 1;
        //            spEfficInfo.Visibility = Visibility.Visible;
        //            gpEfficInfo.Visibility = Visibility.Visible;
        //            _productSpecificationCode = model.ProductSpecificationCode.ToString();
        //            _capacityOfReagentCode = model.CapacityOfReagentCode.ToString();

        //            txtProductSpecification.Text =
        //                _models.FirstOrDefault(m => m.Code == model.ProductSpecificationCode).Message;
        //            txtCapacityOfReagent.Text = _models.FirstOrDefault(m => m.Code == model.CapacityOfReagentCode).Message;
        //        }
        //    }
        //}

        private ObservableCollection<CurvePointModel> _points = new ObservableCollection<CurvePointModel>();

        public ObservableCollection<CurvePointModel> Points => _points;

        private void DataGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (sender is DataGrid grid)
            {
                if (e.Key == Key.V && e.KeyboardDevice.Modifiers == ModifierKeys.Control)
                {
                    var totalText = Clipboard.GetText();
                    if (!string.IsNullOrEmpty(totalText))
                    {
                        try
                        {
                            var rowDatas = totalText.Split(new[] { "\r\n" }, StringSplitOptions.None).ToList();
                            rowDatas.RemoveAll(string.IsNullOrWhiteSpace);

                            var row = Points.IndexOf(grid.CurrentItem as CurvePointModel);
                            var col = grid.SelectedCells[0].Column.DisplayIndex;

                            int rowCount = 8;
                            if (rowDatas.Count + row <= 8)
                            {
                                rowCount = rowDatas.Count + row;
                            }

                            int rowIndex = 0;
                            if (col == 0)
                            {
                                for (int i = row; i < rowCount; i++)
                                {
                                    var tmpData = rowDatas[rowIndex++]
                                        .Split(new[] { "\t" }, StringSplitOptions.None)
                                        .ToList();
                                    if (tmpData.Count == 0)
                                    {

                                    }
                                    else if (tmpData.Count == 1)
                                    {
                                        Points[i].Concentration = tmpData[0];
                                    }
                                    else
                                    {
                                        Points[i].Concentration = tmpData[0];
                                        Points[i].Response = tmpData[1];
                                    }
                                }
                            }
                            else if (col == 1)
                            {
                                for (int i = row; i < rowCount; i++)
                                {
                                    var tmpData = rowDatas[rowIndex++]
                                        .Split(new[] { "\t" }, StringSplitOptions.None)
                                        .ToList();
                                    if (tmpData.Count == 0)
                                    {

                                    }
                                    else if (tmpData.Count >= 1)
                                    {
                                        Points[i].Response = tmpData[0];
                                    }
                                }
                            }

                            grid.Items.Refresh();
                        }
                        catch (Exception exception)
                        {
                            MessageBox.Show(exception.Message);
                        }
                    }
                }
                else if (e.Key == Key.Delete)
                {
                    for (int i = 0; i < grid.SelectedCells.Count; i++)
                    {
                        if (grid.SelectedCells[i].Item is CurvePointModel data)
                        {
                            if (grid.SelectedCells[i].Column.DisplayIndex == 0)
                            {
                                data.Concentration = "0";
                            }
                            else if (grid.SelectedCells[i].Column.DisplayIndex == 1)
                            {
                                data.Response = "0";
                            }
                        }
                    }
                    grid.Items.Refresh();
                }
            }
        }

        /// <summary>
        /// 平台（解析二维码），0-SuperFlex，1-SuperEffic
        /// </summary>
        private int _parseType = 1;

        private void ButtonCreateSuperFlexReagentBarcode_OnClick(object sender, RoutedEventArgs e)
        {
            CreateSuperFlexReagentBarcode();
        }

        private void ButtonCreateSuperEfficReagentBarcode_OnClick(object sender, RoutedEventArgs e)
        {
            CreateSuperEfficReagentBarcode();
        }

        private void ButtonParseBarcodeEffic_OnClick(object sender, RoutedEventArgs e)
        {
            var barcode = tbReagentBarcodeEffic.Text.Trim();
            if (string.IsNullOrEmpty(barcode)|| barcode.Length != 18)
            {
                MessageBox.Show(GetStringByKey("LangErrorBarcodeParse"));
                return;
            }

            tbItemNoEffic.Text = barcode.Substring(0, 4);
            tbLotNoEffic.Text = barcode.Substring(4, 10);
            tbSerialNoEffic.Text = barcode.Substring(14, 4);
        }

        private void ButtonClearEfficInfo_OnClick(object sender, RoutedEventArgs e)
        {
            tbReagentBarcodeEffic.Text = string.Empty;
            tbLotNoEffic.Text = string.Empty;
            tbSerialNoEffic.Text = string.Empty;
            tbReagentBarcodeEffic.Focus();
        }

        private void CboProductsReagent_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.AddedItems[0] is ProductNumberModel model)
            {
                txtItemNoReagent.Text = model.ItemNo;
                if (model.ItemNo.StartsWith("1")) //SuperFlex
                {
                    spReagentSuperFlex.Visibility = Visibility.Visible;
                    spReagentSuperEffic.Visibility = Visibility.Collapsed;
                    tbSuperFlexPlatform.Visibility = Visibility.Visible;
                    tbSuperEfficPlatform.Visibility = Visibility.Collapsed;
                }
                else if (model.ItemNo.StartsWith("2"))//SuperEffic
                {
                    spReagentSuperFlex.Visibility = Visibility.Collapsed;
                    spReagentSuperEffic.Visibility = Visibility.Visible;
                    tbSuperFlexPlatform.Visibility = Visibility.Collapsed;
                    tbSuperEfficPlatform.Visibility = Visibility.Visible;
                }
            }
        }

        private void CreateSuperEfficReagentBarcode()
        {
            try
            {
                if (string.IsNullOrEmpty(txtItemNoReagent.Text))
                {
                    MessageBox.Show(GetStringByKey("LangErrorReagentGenerateItemNo"));
                    return;
                }

                long temp;
                if (!long.TryParse(txtLotNoREffic.Text, out temp) || txtLotNoREffic.Text.Trim().Length != 10)
                {
                    MessageBox.Show(GetStringByKey("LangErrorReagentLotNo1"));
                    return;
                }

                int startIndex;
                if (!int.TryParse(txtStartIndexREffic.Text, out startIndex) || startIndex <= 0 || startIndex > 9999)
                {
                    MessageBox.Show(GetStringByKey("LangErrorStartIndex"));
                    return;
                }

                int num;
                if (!int.TryParse(txtNumREffic.Text, out num) || num <= 0 || num > 9999)
                {
                    MessageBox.Show(GetStringByKey("LangErrorNum"));
                    return;
                }

                if (startIndex + num > 10000)
                {
                    MessageBox.Show(GetStringByKey("LangErrorTotal"));
                    return;
                }

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "Excel(*.xls)|*.xls|Text file(*.txt)|*.txt",
                    InitialDirectory = Directory.GetCurrentDirectory()
                };
                if (sfd.ShowDialog() == true)
                {
                    var fileName = sfd.FileName;
                    var lotNo = txtLotNoREffic.Text.Trim().PadLeft(10, '0');
                    var itemNo = txtItemNoReagent.Text.Trim();
                    Task.Factory.StartNew(() =>
                    {
                        Dispatcher.Invoke(() =>
                        {
                            gridMain.IsEnabled = false;
                            tbHint.Text = GetStringByKey("LangGeneratingBarcode");
                        });
                        List<string> barcodeList = new List<string>();
                        for (int i = startIndex; i < startIndex + num; i++)
                        {
                            var tempCode = string.Format("{0}{1}{2}", itemNo, lotNo, i.ToString().PadLeft(4, '0'));
                            barcodeList.Add(tempCode);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            tbHint.Text = GetStringByKey("LangExporting");
                        });

                        var ext = fileName.Substring(fileName.Length - 3, 3);
                        if (ext == "xls")
                        {
                            ExcelHelper.ExportBarcodeList(fileName, barcodeList, "ReagentBarcode");
                        }
                        else
                        {
                            ExportToTxt(fileName, barcodeList);
                        }

                        Dispatcher.Invoke(() =>
                        {
                            gridMain.IsEnabled = true;
                            tbHint.Text = GetStringByKey("LangBarcodeGenerateSuccess");
                            GC.Collect();
                            MessageBox.Show(GetStringByKey("LangBarcodeGenerateSuccess"));
                        });
                    });
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void Generate2DCodeProductNo_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (e.Source is TextBox tb && !string.IsNullOrEmpty(tb.Text))
            {
                var product = _products.FirstOrDefault(p => p.ProductNumber == tb.Text.Trim());
                if (product != null)
                {
                    txtProductName.Text = product.ProductName;
                    txtItemNo.Text = product.ItemNo;

                    {
                        spEfficInfo.Visibility = Visibility.Visible;
                        gpEfficInfo.Visibility = Visibility.Visible;
                        cboProductSpecification.SelectedIndex = product.ProductSpecificationCode;
                        cboCapacityOfReagent.SelectedIndex = product.CapacityOfReagentCode;
                    }
                }
                else
                {
                    txtProductName.Text = string.Empty;
                    txtItemNo.Text = string.Empty;
                }
            }
        }

        private void GenerateReagentCodeProductNo_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (e.Source is TextBox tb && !string.IsNullOrEmpty(tb.Text))
            {
                var product = _products.FirstOrDefault(p => p.ProductNumber == tb.Text.Trim());
                if (product != null)
                {
                    txtProductNameReagent.Text = product.ProductName;
                    txtItemNoReagent.Text = product.ItemNo;
                    if (product.ItemNo.StartsWith("1")) //SuperFlex
                    {
                        spReagentSuperFlex.Visibility = Visibility.Visible;
                        spReagentSuperEffic.Visibility = Visibility.Collapsed;
                        tbSuperFlexPlatform.Visibility = Visibility.Visible;
                        tbSuperEfficPlatform.Visibility = Visibility.Collapsed;
                    }
                    else if (product.ItemNo.StartsWith("2"))//SuperEffic
                    {
                        spReagentSuperFlex.Visibility = Visibility.Collapsed;
                        spReagentSuperEffic.Visibility = Visibility.Visible;
                        tbSuperFlexPlatform.Visibility = Visibility.Collapsed;
                        tbSuperEfficPlatform.Visibility = Visibility.Visible;
                    }
                }
                else
                {
                    txtProductNameReagent.Text = string.Empty;
                    txtItemNoReagent.Text = string.Empty;
                }
            }
        }

        private string GetDisplayText(string result)
        {
            var d = Convert.ToDouble(result);
            if (App.Flag == 0)
            {
                int digitNum = ConcentrationFormatting.GetDigitNum(d);
                return d.ToString("F" + digitNum);
            }

            return d.ToString();
        }

        private bool _autoGenerateComponentNo = true;

        private void CheckBox_Click(object sender, RoutedEventArgs e)
        {
            _autoGenerateComponentNo = (bool) chkAutoGenerateComponentNo.IsChecked;
        }

        private void txtLotNo_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (e.Source is TextBox tb)
            {
                if (tb.Text.Length > 8)
                {
                    return;
                }

                if (_autoGenerateComponentNo)
                {
                    var lotNo = tb.Text.PadLeft(8, '0');
                    txtLotNo.Text = txtReagentLotNo.Text = string.Format("00{0}", lotNo);
                    txtCaliLotNo1.Text = string.Format("0{0}1", lotNo);
                    txtCaliLotNo2.Text = string.Format("0{0}2", lotNo);
                    tbLotNoC1.Text = string.Format("0{0}3", lotNo);
                    tbLotNoC2.Text = string.Format("0{0}4", lotNo);
                    if (!_noC3)
                        tbLotNoC3.Text = string.Format("0{0}5", lotNo);
                }
            }
        }

        private bool _noC3 = true;
        private void ChkNoC3_OnClick(object sender, RoutedEventArgs e)
        {
            _noC3 = (bool) chkNoC3.IsChecked;
            gpEfficInfo.IsEnabled = !_noC3;
            if (_noC3)
            {
                tbLotNoC3.Text = "0000000000";
                tbTargetValueC3.Text = "0.0";
                tbLowerLimitC3.Text = "0.0";
                tbUpperLimitC3.Text = "0.0";
            }
        }
        private bool _noC4 = true;
        private void chkNoC4_Click(object sender, RoutedEventArgs e)
		{
            _noC4 = (bool)chkNoC4.IsChecked;
            gpEfficInfo4.IsEnabled = !_noC4;
            if (_noC4)
            {
                tbLotNoC4.Text = "0000000000";
                tbTargetValueC4.Text = "0.0";
                tbLowerLimitC4.Text = "0.0";
                tbUpperLimitC4.Text = "0.0";
            }
        }
	}
}
