using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.Net;
using System.Text.RegularExpressions;
using System.IO;

namespace PdfPrint
{
    public partial class frmPdfPrint : Form
    {
        public frmPdfPrint()
        {
            InitializeComponent();
        }

        /// <summary>
        /// ダウンロード処理
        /// </summary>
        /// <param name="ReqNo"></param>
        /// <param name="fileName"></param>
        private void DoSendRequest(out string fileName)
        {
            fileName = null;

            string[] cmd = System.Environment.GetCommandLineArgs();

            if (cmd == null)
            {
                return;
            }
            string url = cmd[1];

            //---------------------------
            // サーバーにクエリを投げる
            //---------------------------
            WebResponse response = null;
            try
            {
                HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
                request.Method = "GET";
                request.ContentType = "text/xml";
                response = request.GetResponse();
            }
            catch (WebException ex)
            {
                response = ex.Response;

            }

            // ヘッダ情報からファイル名の取得
            string header = response.Headers["Content-Disposition"];

            // 初期値
            fileName = getFileName("download.pdf", header);

            // ファイルのダウンロード
            using (Stream st = response.GetResponseStream())
            using (FileStream fs = new FileStream(fileName, FileMode.Create)) {
                Byte[] buf = new Byte[response.ContentLength];
                int count = 0;
                do {
                    count = st.Read(buf, 0, buf.Length);
                    fs.Write(buf, 0, count);
                } while (count != 0);
            }
            response.Close();
        }

        /// <summary>
        /// ファイル名取得
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        private string getFileName(string defFile, string content)
        {
            string filename = defFile;

            if (!String.IsNullOrEmpty(content))
            {
                // filename=＜ファイル名＞の抜き出し
                Regex re = new Regex(@"
                filename\s*=\s*
                (?:
                    ""(?<filename>[^""]*)""
                    |
                    (?<filename>[^;]*)
                )
                ", RegexOptions.IgnoreCase
                    | RegexOptions.IgnorePatternWhitespace);

                Match m = re.Match(content);
                if (m.Success)
                {
                    filename = HttpUtility.UrlDecode(m.Groups["filename"].Value);
                }
            }

            // Tempフォルダの取得

            return filename;
        }

        /// <summary>
        /// 印刷処理
        /// </summary>
        /// <param name="filename"></param>
        private void DoPrintPdf(string filename)
        {
            string strRegPath;
            Microsoft.Win32.RegistryKey rKey;

            // キーを取得(最初にAcrobat,だめならAdobeReader)
            strRegPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\Acrobat.exe";
            rKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(strRegPath);

            if (rKey == null)
            {
                strRegPath = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\AcroRd32.exe";
                rKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(strRegPath);
            }

            // 値(exeのパス)を取得(既定の値の場合は空文字指定)
            string location;

            try
            {
                // 値(exeのパス)を取得(既定の値の場合は空文字指定)
                location = rKey.GetValue("").ToString();
            } catch (NullReferenceException ex)
            {
                throw new ApplicationException("AcrobatもしくはAdobeReaderがインストールされていないため、PDFファイルの印刷ができません。");
            } finally
            {
                // 開いたレジストリキーを閉じる
                rKey.Close();
            }
 
            // ===Acrobatを起動し印刷===
            System.Diagnostics.Process pro  = new System.Diagnostics.Process();

            // .Net的書き方(C#でも可能な書き方)
            // Acrobatのフルパス設定
            pro.StartInfo.FileName = location;
            pro.StartInfo.Verb = "open";
            // Acrobatのコマンドライン引数設定
            pro.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            pro.StartInfo.Arguments = " /n /p /h " + filename;
            // プロセスを新しいWindowで起動
            pro.StartInfo.CreateNoWindow = true;
            // プロセス起動
            pro.Start();
            //プロセス終了
            pro.WaitForExit(5000);
            pro.Kill();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fname = null;
            DoSendRequest(out fname);
            DoPrintPdf(fname);
        }

    }
}