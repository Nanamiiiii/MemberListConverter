using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

using Microsoft.Win32;

namespace MemberListConverter
{
    struct Member
    {
        public string handle;
        public string name_kana;
        public string name;
        public string payment;
        public string role;
        public string studentNumber;
        public string univName;
    }

    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Refs_Click(object sender, RoutedEventArgs e)
        {
            // 実行ディレクトリの取得
            string path = Directory.GetCurrentDirectory();

            // ファイル選択ダイアログ
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "CSVファイル (*.csv)|*.csv|すべてのファイル (*.*)|*.*",
                Title = "元データを選択",
                InitialDirectory = path
            };
            if (dialog.ShowDialog() == true)
            {
                // 選択ファイルのパスを格納
                src_path.Text = dialog.FileName;
            }
        }

        private void Run_Convert(object sender, RoutedEventArgs e)
        {
            // 必要な変数を取得
            string[] additional_member = additinalMem.Text.Split(',');
            string source_path = src_path.Text;

            // ファイル存在確認
            if (!System.IO.File.Exists(source_path))
            {
                MessageBox.Show("ファイルが存在しません", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (System.IO.File.Exists("converted_member.csv"))
            {
                MessageBoxResult res = MessageBox.Show("converted_member.csv は存在します．上書きしますか？", "Warning", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (res == MessageBoxResult.No)
                {
                    return;
                }
            }

            // ソースファイル読み込み
            string del_regx = "(\"$|^\")";
            StreamReader reader = new StreamReader(source_path, Encoding.GetEncoding("UTF-8"));
            StreamWriter writer = new StreamWriter("converted_member.csv", false, Encoding.GetEncoding("UTF-8"));
            reader.ReadLine();
            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                var line_spl = Regex.Split(Regex.Replace(line, del_regx, ""), "\",\"");
                // MessageBox.Show(line);
                try
                {
                    Member mMember = new Member
                    {
                        handle = line_spl[2],
                        name_kana = fmtName(line_spl[7]),
                        name = fmtName(line_spl[6]),
                        payment = line_spl[3],
                        role = line_spl[5],
                        studentNumber = fmtStudentNumber(line_spl[12]).Substring(0, 8),
                        univName = line_spl[9]
                    };
                    // 抽出
                    if (mMember.univName == "早稲田大学" && (mMember.payment == "YES" || additional_member.Contains(mMember.handle)))
                    {
                        writer.WriteLine("{0}\t{1}\t{2}", mMember.studentNumber, mMember.name, mMember.name_kana);
                    }
                }
                catch (System.ArgumentOutOfRangeException exc)
                {
                    // 配列外アクセス（登録情報がそもそもたりてない）は無視
                    continue;
                }
            }
            writer.Close();
            reader.Close();
            MessageBox.Show("処理が完了しました");
        }

        private string fmtName(string name)
        {
            return Regex.Replace(name, @"\s", "　");
        }

        private string fmtStudentNumber(string name)
        {
            return Regex.Replace(Regex.Replace(Microsoft.VisualBasic.Strings.StrConv(name, VbStrConv.Narrow), "-[0-9]", "").ToUpper(), @"(\s|　)", "");
        }

    }
}
