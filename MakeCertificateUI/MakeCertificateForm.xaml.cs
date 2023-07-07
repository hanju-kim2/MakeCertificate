using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Threading;
using System.Windows.Media;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Windows.Threading;
using Microsoft.FSharp.Collections;

// 개발용 @"C:\Users\kimha\Desktop\MakeCertificate\configData"
// 배포용 @".\configData"
// 회사용 @"C:\Users\kimhanju\Desktop\MakeCertificate\configData"

namespace MakeCertificateUI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MakeCertificateForm : Window
    {
        string pathValue;
        string certificateFileName;
        string certifacateNum;
        string photoFileName;
        string Date;
        string warnings;
        string productName;
        string configRoot;
        List<string> photoWarning;
        List<string> photoListForWarning_KSC_9815_9547;
        List<string> photoListForWarning_KSC_9832_9835;
        List<string> photoListForWarning_KSC_9814_1_9814_2;
        List<string> photoListForWarning_KSX_3124_3126;
        List<string> photoListForWarning_KSC_9610_6_1_3;
        List<string> photoListForWarning_KSC_9610_6_2_4;
        List<string> photoListForWarning_KSX_3124;
        List<string> photoListForWarning_KSX_3124_3125;
        List<string> photoListForWarning_KSX_3143_KSC_9814_2;
        List<string> 이하여백File;
        List<List<string>> 이하여백candidiate_KSC_9815_9547;
        List<List<string>> 이하여백candidiate_KSC_9832_9835;
        List<List<string>> 이하여백candidiate_KSC_9814_1_9814_2;
        List<List<string>> 이하여백candidiate_KSX_3124_3126;
        List<List<string>> 이하여백candidiate_KSC_9610_6_1_3;
        List<List<string>> 이하여백candidiate_KSC_9610_6_2_4;
        List<List<string>> 이하여백candidiate_KSX_3124;
        List<List<string>> 이하여백candidiate_KSX_3124_3125;
        List<List<string>> 이하여백candidiate_KSX_3143_KSC_9814_2;

        public MakeCertificateForm()
        {
            InitializeComponent();
            configRoot = @".\configData";
            pathValue = "";
            certificateFileName = "";
            certifacateNum = "";
            photoFileName = "";
            warnings = "";
            productName = "";
            Date = "";
            photoWarning = new List<string>();
            이하여백File = new List<string>();
        }

        private void KSX_3124_3126()
        {
            if (CheckDate("KSX_3124_3126") == false)
                return;

            photoListForWarning_KSX_3124_3126 = new List<string>(File.ReadAllLines(configRoot + @"\KSX_3124_3126\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSX_3124_3126\이하여백.txt").ToList();
            이하여백candidiate_KSX_3124_3126 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSX_3124_3126.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSX_3124_3126"); //excel의 생성까지 담당
            if (Processing("KSX_3124_3126", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSX_3124_3126");

            MakeProductPhotoFile("KSX_3124_3126");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSC_9832_9835()
        {
            if (CheckDate("KSC_9832_9835") == false)
                return;

            photoListForWarning_KSC_9832_9835 = new List<string>(File.ReadAllLines(configRoot + @"\KSC_9832_9835\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSC_9832_9835\이하여백.txt").ToList();
            이하여백candidiate_KSC_9832_9835 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSC_9832_9835.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSC_9832_9835");
            if (Processing("KSC_9832_9835", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSC_9832_9835");

            MakeProductPhotoFile("KSC_9832_9835");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSC_9814_1_9814_2()
        {
            if (CheckDate("KSC_9814_1_9814_2") == false)
                return;

            photoListForWarning_KSC_9814_1_9814_2 = new List<string>(File.ReadAllLines(configRoot + @"\KSC_9814_1_9814_2\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSC_9814_1_9814_2\이하여백.txt").ToList();
            이하여백candidiate_KSC_9814_1_9814_2 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSC_9814_1_9814_2.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSC_9814_1_9814_2");
            if (Processing("KSC_9814_1_9814_2", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSC_9814_1_9814_2");

            MakeProductPhotoFile("KSC_9814_1_9814_2");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSC_9815_9547()
        {
            if (CheckDate("KSC_9815_9547") == false)
                return;

            photoListForWarning_KSC_9815_9547 = new List<string>(File.ReadAllLines(configRoot + @"\KSC_9815_9547\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSC_9815_9547\이하여백.txt").ToList();
            이하여백candidiate_KSC_9815_9547 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSC_9815_9547.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSC_9815_9547");
            if (Processing("KSC_9815_9547", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSC_9815_9547");

            MakeProductPhotoFile("KSC_9815_9547");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSC_9610_6_1_3()
        {
            if (CheckDate("KSC_9610_6_1_3") == false)
                return;

            photoListForWarning_KSC_9610_6_1_3 = new List<string>(File.ReadAllLines(configRoot + @"\KSC_9610_6_1_3\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSC_9610_6_1_3\이하여백.txt").ToList();
            이하여백candidiate_KSC_9610_6_1_3 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSC_9610_6_1_3.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSC_9610_6_1_3");
            if (Processing("KSC_9610_6_1_3", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSC_9610_6_1_3");

            MakeProductPhotoFile("KSC_9610_6_1_3");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSC_9610_6_2_4()
        {
            if (CheckDate("KSC_9610_6_2_4") == false)
                return;

            photoListForWarning_KSC_9610_6_2_4 = new List<string>(File.ReadAllLines(configRoot + @"\KSC_9610_6_2_4\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSC_9610_6_2_4\이하여백.txt").ToList();
            이하여백candidiate_KSC_9610_6_2_4 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSC_9610_6_2_4.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSC_9610_6_2_4");
            if (Processing("KSC_9610_6_2_4", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSC_9610_6_2_4");

            MakeProductPhotoFile("KSC_9610_6_2_4");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSX_3124()
        {
            if (CheckDate("KSX_3124") == false)
                return;

            photoListForWarning_KSX_3124 = new List<string>(File.ReadAllLines(configRoot + @"\KSX_3124\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSX_3124\이하여백.txt").ToList();
            이하여백candidiate_KSX_3124 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSX_3124.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSX_3124");
            if (Processing("KSX_3124", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSX_3124");

            MakeProductPhotoFile("KSX_3124");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSX_3124_3125()
        {
            if (CheckDate("KSX_3124_3125") == false)
                return;

            photoListForWarning_KSX_3124_3125 = new List<string>(File.ReadAllLines(configRoot + @"\KSX_3124_3125\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSX_3124_3125\이하여백.txt").ToList();
            이하여백candidiate_KSX_3124_3125 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSX_3124_3125.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSX_3124_3125");
            if (Processing("KSX_3124_3125", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSX_3124_3125");

            MakeProductPhotoFile("KSX_3124_3125");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void KSX_3143_KSC_9814_2()
        {
            if (CheckDate("KSX_3143_KSC_9814_2") == false)
                return;

            photoListForWarning_KSX_3143_KSC_9814_2 = new List<string>(File.ReadAllLines(configRoot + @"\KSX_3143_KSC_9814_2\PhotoFileName.txt"));
            이하여백File = File.ReadAllLines(configRoot + @"\KSX_3143_KSC_9814_2\이하여백.txt").ToList();
            이하여백candidiate_KSX_3143_KSC_9814_2 = new List<List<string>>();
            foreach (var line in 이하여백File)
            {
                이하여백candidiate_KSX_3143_KSC_9814_2.Add(line.Split(',').ToList());
            }

            ExcelLibrary.ExcelRecord excel = InsertNumberAndDate("KSX_3143_KSC_9814_2");
            if (Processing("KSX_3143_KSC_9814_2", excel) == false)
                return; // compareForErrorProcessing()에서 걸리는 경우
            ExcelScenario.closeExcel(excel);

            PrintPhotoWarning("KSX_3143_KSC_9814_2");

            MakeProductPhotoFile("KSX_3143_KSC_9814_2");

            progressBar.Value = 100;
            MessageBox.Show("완료");
        }

        private void makeCertificateButton_Click(object sender, RoutedEventArgs e)
        {
            photoWarning.Clear();
            warnings = "";
            progressBar.Value = 0;
            progress_text.Foreground = Brushes.Black;
            warningBox.Text = warnings;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            if (selectForm.SelectedIndex == -1)
            {
                MessageBox.Show("규격을 선택하세요");
            }
            else if (selectForm.SelectedIndex == 0) // 조명기기
            {
                try
                {
                    KSC_9815_9547();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 1) // 멀티미디어 기기
            {
                try
                {
                    KSC_9832_9835();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 2) // 가전기기
            {
                try
                {
                    KSC_9814_1_9814_2();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 3) // 무선데이터통신시스템용
            {
                try
                {
                    KSX_3124_3126();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 4) // 주거 환경에서 사용되는 기기
            {
                try
                {
                    KSC_9610_6_1_3();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 5) // 산업 환경에서 사용되는 기기
            {
                try
                {
                    KSC_9610_6_2_4();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 6) // 무선기기
            {
                try
                {
                    KSX_3124();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 7) // 특정소출력 무선기기
            {
                try
                {
                    KSX_3124_3125();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
            else if (selectForm.SelectedIndex == 8) // 가정용 무선전력 전송기기
            {
                try
                {
                    KSX_3143_KSC_9814_2();
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
        }


        private void selectPathButton_Click(object sender, RoutedEventArgs e)
        {
            // CommonOpenFileDialog 클래스 생성
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            string[] shortPath;
            // 처음 보여줄 폴더 설정(안해도 됨) 
            //dialog.InitialDirectory = ""; 
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                pathValue = dialog.FileName;
                shortPath = pathValue.Split('\\'); // text field 크기 상 일부만 보여줌
                folderPath.Text = shortPath[shortPath.Length - 2] + "\\" + shortPath[shortPath.Length - 1];

                try
                {
                    List<string> infoList = new List<string>();
                    if (selectForm.SelectedIndex == 0)
                    {
                        infoList = OutlineInfo("KSC_9815_9547");
                    }
                    else if(selectForm.SelectedIndex == 1)
                    {
                        infoList = OutlineInfo("KSC_9832_9835");
                    }
                    else if(selectForm.SelectedIndex == 2)
                    {
                        infoList = OutlineInfo("KSC_9814_1_9814_2");
                    }
                    else if(selectForm.SelectedIndex == 3)
                    {
                        infoList = OutlineInfo("KSX_3124_3126");
                    }
                    else if (selectForm.SelectedIndex == 4)
                    {
                        infoList = OutlineInfo("KSC_9610_6_1_3");
                    }
                    else if (selectForm.SelectedIndex == 5)
                    {
                        infoList = OutlineInfo("KSC_9610_6_2_4");
                    }
                    else if (selectForm.SelectedIndex == 6)
                    {
                        infoList = OutlineInfo("KSX_3124");
                    }
                    else if (selectForm.SelectedIndex == 7)
                    {
                        infoList = OutlineInfo("KSX_3124_3125");
                    }
                    else if (selectForm.SelectedIndex == 8)
                    {
                        infoList = OutlineInfo("KSX_3143_KSC_9814_2");
                    }

                    company.Text = infoList[0];
                    manufacturer.Text = infoList[1];
                    country.Text = infoList[2];
                    productName = infoList[3];
                }
                catch (Exception error)
                {
                    MessageBox.Show("에러 발생: " + error.Message);
                }
            }
        }

        private void ProgressBar_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e) { } // 필요없는 함수인데 없으면 에러남

        
        private ExcelLibrary.ExcelRecord InsertNumberAndDate(string module)
        {
            certificateFileName = certificateName.Text;
            certifacateNum = certificateNumber.Text;
            photoFileName = photoName.Text;
            Date = DateTime.Parse(issueDate.Text).ToString("yyyy년 MM월 dd일");
            ExcelLibrary.ExcelRecord excel = ExcelScenario.openExcel(configRoot + @"\" + module + @"\" + module + "_Result폼.xlsx", module, pathValue + @"\04_report\" + certificateFileName + ".xlsx");
            ExcelScenario.insertStringProcessing(excel, "J8", certifacateNum);
            ExcelScenario.insertStringProcessing(excel, "M66", certifacateNum);
            ExcelScenario.insertStringProcessing(excel, "A44", Date);
            ExcelScenario.insertStringProcessing(excel, "B66", Date);
            return excel;
        }
        
        // module은 해당 함수 이름
        private bool CheckDate(string module)
        {
            List<string> dateList = ExcelScenario.getStringsIn측정지시서Processing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", ListModule.OfSeq(new List<string>{ "G4", "B5", "G5" })).ToList();
            DateTime receiptDate = DateTime.Parse(dateList[0]);
            DateTime testStartDate = DateTime.Parse(dateList[1]);
            DateTime testEndDate = DateTime.Parse(dateList[2]);
            Date = issueDate.Text;
            DateTime dateTime = DateTime.Parse(Date);

            if (DateTime.Compare(receiptDate, testStartDate) > 0 || DateTime.Compare(testStartDate, testEndDate) > 0 || DateTime.Compare(testEndDate, dateTime) > 0)
            {
                MessageBox.Show("접수일, 시험일, 발급일이 어긋납니다.");
                return false;
            }
            return true;
        }
        
        private bool Processing(string module, ExcelLibrary.ExcelRecord excel)
        {
            var errorMessages = ExcelScenario.compareForErrorProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\ErrorConfig.xml", excel);
            foreach (var error in errorMessages)
            {
                if (error.logType == ExcelScenario.LogType.Stop) // {Success: 성공, Warning: 빈값, Exception: 텍스트}의 경우는 성공 처리함
                {
                    ExcelScenario.closeExcel(excel);
                    MessageBox.Show(error.message);
                    return false;
                }
            }
            progressBar.Value = 10;
            progress_text.Foreground = Brushes.White;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.tableProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\CellConfig.xml", excel);
            progressBar.Value = 20;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.cellCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\CellConfig.xml", excel);
            progressBar.Value = 30;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.checkBoxCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\ControlConfig.xml", excel);
            ExcelScenario.groupCheckBoxCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\ControlConfig.xml", excel);
            progressBar.Value = 40;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.oddFooterProcessing(certifacateNum, excel);
            progressBar.Value = 50;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            Write이하여백(module, excel);
            progressBar.Value = 60;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            photoWarning.AddRange(ExcelScenario.photoCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\PhotoConfig.xml", excel));
            progressBar.Value = 70;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.excelPhotoCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\PhotoConfig.xml", excel);
            progressBar.Value = 80;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            ExcelScenario.deletePageProcessing(configRoot + @"\" + module + @"\CellConfig.xml", 55, excel);
            progressBar.Value = 90;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);

            return true;
        }

        private void PrintPhotoWarning(string module)
        {
            List<string> printWarning = new List<string>();
            if (module == "KSX_3124_3126")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSX_3124_3126.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1","").Replace(".jpg",""));
                    }
                }
            }
            else if(module == "KSC_9832_9835")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSC_9832_9835.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if(module == "KSC_9814_1_9814_2")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSC_9814_1_9814_2.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if(module == "KSC_9815_9547")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSC_9815_9547.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if (module == "KSC_9610_6_1_3")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSC_9610_6_1_3.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if (module == "KSC_9610_6_2_4")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSC_9610_6_2_4.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if (module == "KSX_3124")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSX_3124.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if (module == "KSX_3124_3125")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSX_3124_3125.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }
            else if (module == "KSX_3143_KSC_9814_2")
            {
                foreach (string warning in photoWarning)
                {
                    if (photoListForWarning_KSX_3143_KSC_9814_2.Contains(warning)) // +1이 없는 경우만 체크
                    {
                        printWarning.Add(warning.Replace("+1", "").Replace(".jpg", ""));
                    }
                }
            }

            foreach (string warning in printWarning)
            {
                warnings += warning + " 사진이 존재하지 않습니다.\n";
            }
            warningBox.Text = warnings;
            this.Dispatcher.Invoke((ThreadStart)(() => { }), DispatcherPriority.ApplicationIdle);
        }

        private void MakeProductPhotoFile(string module)
        {
            ExcelLibrary.ExcelRecord excel = ExcelScenario.openExcel(configRoot + @"\제품사진폼.xlsx", "Sheet1", pathValue + @"\04_report\" + photoFileName + ".xlsx");
            ExcelScenario.insertStringProcessing(excel, "B3", "모델명 : " + productName);
            ExcelScenario.photoCopyProcessing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", configRoot + @"\" + module + @"\ProductPhotoConfig.xml", excel);
            ExcelScenario.closeExcel(excel);
        }

        private List<string> OutlineInfo(string module)
        {
            return ExcelScenario.getStringsIn측정지시서Processing(pathValue, configRoot + @"\" + module + @"\InputFileConfig.xml", ListModule.OfSeq(new List<string>{ "B8", "B15", "G15", "B24" })).ToList();
        }

        private void Write이하여백(string module, ExcelLibrary.ExcelRecord excel)
        {
            switch(module)
            {
                case "KSX_3124_3126":
                    foreach (List<string> chapter in 이하여백candidiate_KSX_3124_3126)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSC_9815_9547":
                    foreach (List<string> chapter in 이하여백candidiate_KSC_9815_9547)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSC_9832_9835":
                    foreach (List<string> chapter in 이하여백candidiate_KSC_9832_9835)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSC_9814_1_9814_2":
                    foreach (List<string> chapter in 이하여백candidiate_KSC_9814_1_9814_2)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSC_9610_6_1_3":
                    foreach (List<string> chapter in 이하여백candidiate_KSC_9610_6_1_3)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSC_9610_6_2_4":
                    foreach (List<string> chapter in 이하여백candidiate_KSC_9610_6_2_4)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSX_3124":
                    foreach (List<string> chapter in 이하여백candidiate_KSX_3124)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSX_3124_3125":
                    foreach (List<string> chapter in 이하여백candidiate_KSX_3124_3125)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
                case "KSX_3143_KSC_9814_2":
                    foreach (List<string> chapter in 이하여백candidiate_KSX_3143_KSC_9814_2)
                    {
                        foreach (string cell in chapter)
                        {
                            if (ExcelScenario.getStringProcessing(excel, cell).Trim() == "")
                            {
                                ExcelScenario.insertStringProcessing(excel, cell, "- 이하여백 -");
                                break;
                            }
                        }
                    }
                    break;
            }
        }

    }
}
