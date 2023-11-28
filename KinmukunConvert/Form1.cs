using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

using System.Drawing.Imaging;

using ZXing;
using System.Diagnostics;

namespace KinmukunConvert
{
    public partial class Form1 : Form
    {
        private Button[] buttons;



        // 日時の各値を表示します
        //int year = dNow.Year;              // 現在の年を返します
        string month ;             // 現在の月を返します
        string day;                 // 現在の日を返します
        string hour ;               // 現在の時を返します
        string minute ;           // 現在の分を返します

        string currentsnum;

        string myDir = Directory.GetCurrentDirectory();

        Button currentbutton;
        int shokuinsu;
        int byocount=0;
        string barcode,barcodetmp;

        //yarinaosi no tame

        string ptext;
        string psrcfile;
        int pi;
        string ppsrcfile, ptxt;//rsq.csv you

        public Form1()
        {
            InitializeComponent();
            // この2行はデザイナーで設定する
            this.KeyPreview = true;
            this.KeyPress += Form1_KeyPress;
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // イベントが未処理でTextBoxにフォーカスがなく、入力文字がa-zの場合
            // if (!e.Handled && !comboBox1.Focused && 'a' <= e.KeyChar && e.KeyChar <= 'z')
            if (!e.Handled)
            {         
              //labelBarcodetest.Focus();
              //labelBarcodetest.Text = labelBarcodetest.Text + (e.KeyChar.ToString() + "(" + (Convert.ToInt32(e.KeyChar)).ToString() + "):");
                
                if (Convert.ToInt32(e.KeyChar)==13)
                {
                 
                    barcode = barcodetmp.Trim();
                    barcodetmp = "";
                    labelBarcodetest.Text = barcode; //'9017' 17ban       9997 yarinasoi  9990 tuika shusei

                    // MessageBox.Show(barcode.Substring(2, 2));

                    if (barcode.Substring(0,3)=="999")
                    {
                        if (barcode.Substring(3, 1) == "0")  //tuika shusei
                        {
                            checkShusei.Checked=true;
                        }
                        else  //yarinaosi jisso sakiokuri
                        {
                            yarinaosi();
                           
                        }

                    }
                    else
                    {
                        // string snum = Regex.Replace(((Button)sender).Text, @"[^0-9]", "");
                        int bnum = int.Parse(barcode.Substring(2, 2)) ;
                        string snum = bnum.ToString().Trim();

                        //Bsentaku(snum, sender);
                        // this.buttons[i]
                        Bsentaku(snum, this.buttons[bnum-1]);

                    }

             

                }
                else
                {
                    barcodetmp = barcodetmp + e.KeyChar.ToString();
                }


       
                
                e.Handled = true;
            }



        }

        private void button1_Click(object sender, EventArgs e)
        {

            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "default.csv";
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            //ofd.InitialDirectory = @"C:\";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "ファイル(*.csv;*.csv)|csv.*|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 2;
            //タイトルを設定する
            ofd.Title = "開くファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                //Console.WriteLine(ofd.FileName);
                textBoxSrc.Text = ofd.FileName;

                Properties.Settings.Default.txtSrc = textBoxSrc.Text;

                Properties.Settings.Default.Save();

            }
          



        }

        private void button2_Click(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "default.csv";
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            //ofd.InitialDirectory = @"C:\";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "ファイル(*.csv;*.csv)|csv.*|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 2;
            //タイトルを設定する
            ofd.Title = "開くファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                //Console.WriteLine(ofd.FileName);
                textBoxDst.Text = ofd.FileName;
                Properties.Settings.Default.txtDst = textBoxDst.Text;

                Properties.Settings.Default.Save();

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            


            // 現在の日時を取得します
            DateTime dNow = System.DateTime.Now;

            // 日時の各値を表示します
            //int year = dNow.Year;              // 現在の年を返します
            month = Right("00" + dNow.Month.ToString(), 2);             // 現在の月を返します
             day = dNow.Day.ToString();                 // 現在の日を返します
             hour = Right("00" + dNow.Hour.ToString(), 2);               // 現在の時を返します
             minute = Right("00" + dNow.Minute.ToString(), 2);           // 現在の分を返します

            textBoxSrc.Text = Properties.Settings.Default.txtSrc;
            textBoxDst.Text=Properties.Settings.Default.txtDst  ;
            textBoxSort.Text = Properties.Settings.Default.txtSort;
            textBoxGakkoMei.Text=Properties.Settings.Default.txtGakkoMei ;
            textBoxCal.Text=Properties.Settings.Default.txtCal  ;
            textBoxHandleNum.Text=Properties.Settings.Default.txtHandleNum;
            textBoxHandleNum2.Text=Properties.Settings.Default.txtHandleNum2;



            shuttaikinKirikae(dNow);


            try
            {
                string text2 = System.IO.File.ReadAllText(textBoxSort.Text, Encoding.GetEncoding("Shift_JIS"));
                List<string> lst = text2.Split('\n').ToList();
                foreach (string tx in lst)
                {
                    listBox1.Items.Add(tx);
                }
            }
            catch(Exception ex)
            {

            }

            //***********date
            labelDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            //*****************name load
            string text="";
            try
            {
                 text = System.IO.File.ReadAllText(textBoxSrc.Text, Encoding.GetEncoding("Shift_JIS"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新元職員データの欄に、職員データ.csvの場所を指定してください");
            }
            List<string> stL = new List<string>();

                stL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();

                // List<string> lst = stL.OrderBy(st => stoi(st.Split(',')[2])).ToList();

                String[] sname = stL.ToArray(); //name       
                shokuinsu = sname.Length;
           
            //*******************:hyoji
            string srcdir="";
            try
            {
                srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            }
            catch(Exception ex)
            {

            }
               

                this.buttons = new Button[100];
                for (int i = 0; i < sname.Length-1; i++)
                {
                    //ボタンコントロールのインスタンス作成
                    this.buttons[i] = new Button();

                    //プロパティ設定
                    this.buttons[i].Name = "btn" + i.ToString();
                    this.buttons[i].Text = (i + 1).ToString() + sname[i];
                    this.buttons[i].Top = (i % 10) * 35 + 70;
                    this.buttons[i].Left = (int)(i / 10) * 111;
                    this.buttons[i].Height = 33;
                    this.buttons[i].Width = 110;
                    float newSize = 9.0f;
                    this.buttons[i].Font = new Font("ＭＳ ゴシック", newSize);
                    this.buttons[i].Click += buttonNoShow_Click;
                    this.buttons[i].MouseMove += button_Hover;
                    this.buttons[i].MouseLeave += button_Leave;
                this.buttons[i].MouseDown +=button_MouseDown;


                //color file huruitoki del
                //  delscolor(srcdir + "\\backup\\" + (i + 1).ToString().Trim() + ".txt");

                try
                {
                    colorRead(srcdir, i);
                }catch(Exception ex)
                {

                }
                                

                    //コントロールをフォームに追加
                    tabPage1.Controls.Add(this.buttons[i]);
                }
           


            //shokika folder sakusei etc
            Init();           

            if (textBoxKaisiHun.Text == "" || textBoxKaisiJi.Text == "" || textBoxShuryoHun.Text == "" || textBoxShuryoJi.Text == "")
            {
                MessageBox.Show("勤務開始時刻と勤務終了時刻を設定してください（8:30と17:00の形式で)");
            }


          //  srcdir = Path.GetDirectoryName(textBoxSrc.Text);
         //   int fileCount = Directory.GetFiles(srcdir + "\\" + comboBox1.Text + "\\xls").Length;
            List<int> ban = Enumerable.Range(1, 100).ToList();
            foreach (int bi in ban)
            {
                comboBox2.Items.Add(bi.ToString().Trim());
                comboBox3.Items.Add(bi.ToString().Trim());
                comboBox4.Items.Add(bi.ToString().Trim());
            }
            int currentYear = DateTime.Now.Year;

            comboBox5.Items.Add((currentYear-1).ToString());
            comboBox5.Items.Add(currentYear.ToString());
            comboBox5.Items.Add((currentYear+1).ToString());

            comboBox5.Text = currentYear.ToString();

        }


        private void delscolor(string sfile)
        {
            if (File.Exists(sfile))
            {
               // MessageBox.Show(sfile + "=sfile  " + File.GetCreationTime(sfile).ToString() + " " + DateTime.Now.Day.ToString());

                if (File.GetCreationTime(sfile).Day < DateTime.Now.Day)
                {
                    File.Delete(sfile);
                }
            }

         

        }

        private void Init()
        {
            
            try
            {
                 string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
                // MessageBox.Show(srcdir);
                string dirmonth = srcdir + "\\" + month + "\\data";
                CreateDir(dirmonth);
                CreateDir(srcdir + "\\" + month + "\\xls");
                CreateDir(srcdir + "\\maindata");
                CreateDir(srcdir + "\\backup");

                string kisodata = srcdir + "\\maindata\\基礎データ.csv";
                //MessageBox.Show(File.Exists(kisodata).ToString());

                if (File.Exists(kisodata))
                {

                    string kisodt = File.ReadAllText(kisodata, Encoding.GetEncoding("Shift_JIS"));
                    string[] val = kisodt.Split(',');
                    string[] kaisiji = val[2].Replace("\"", "").Split(':');
                    string[] shuryoji = val[3].Replace("\"", "").Split(':');
                    textBoxKaisiJi.Text = kaisiji[0];
                    textBoxKaisiHun.Text = kaisiji[1];
                    textBoxShuryoJi.Text = shuryoji[0];
                    textBoxShuryoHun.Text = shuryoji[1];
                }
            }
            catch(Exception ex)
            {

            }
           
          
           

        }

        private void CreateDir(string dirmonth)
        {
            if (System.IO.Directory.Exists(@dirmonth))
            {
                // MessageBox.Show("sampleフォルダは存在します");
            }
            else
            {
                // MessageBox.Show("sampleフォルダは存在しません");

                Directory.CreateDirectory(@dirmonth);
            }

        }


        private int sTohun(string hm)
        {
       

            int HH = stoi(hm.Split(':')[0]);
            int MM = stoi(hm.Split(':')[1]);

            int hun = HH * 60 + MM;

            return hun;
        }

        private void button_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Left:
                  //  MessageBox.Show("マウスの左ボタンが押されました。");
                    break;
                case MouseButtons.Middle:
                  //  MessageBox.Show("マウスの中央ボタンが押されました。");
                    break;
                case MouseButtons.Right:
                    //  MessageBox.Show("マウスの右ボタンが押されました。");
                    checkShusei.Checked=true;
                    buttonNoShow_Click( sender,e);
                    break;
            }

        }



        private void buttonD_Click(object sender,EventArgs e)
        { 
      

        }

        private void button_Hover(object sender, EventArgs e)
        {
            //if (!checkShusei.Checked)
            //{
              //  labelCname.Text = ((Button)sender).Text;
            // }

            string snum = Regex.Replace(((Button)sender).Text, @"[^0-9]", "");
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);


            //labelName.Text = snamer(stoi(snum) - 1);
            string line = "";
            string KJt="00", KHt="00", SJt="00", SHt="00";

            try {
                line = dtYomi(srcdir, snum, month, day);
                string[] linea = line.Split(',');

                 KJt = linea[1].Split(':')[0].Replace("\"", "");
                 KHt = linea[1].Split(':')[1].Replace("\"", "");
                 SJt = linea[2].Split(':')[0].Replace("\"", "");
                 SHt = linea[2].Split(':')[1].Replace("\"", "");
            }
            catch(Exception ex)
            {

            }

            

            labelCname.Text = ((Button)sender).Text+" " + KJt + ":" + KHt + "~" + SJt + ":" + SHt;
            //snamer(stoi(snum) - 1)+" "+KJt+":"+KHt+"~"+SJt+":"+SHt;

            //  labelYobi.Text = "(" + dateTimePicker1.Value.ToString("ddd") + ")";



        }

        private void button_Leave(object sender,EventArgs e)
        {
            labelCname.Text = "";
        }

        private void buttonNoShow_Click(object sender, EventArgs e)
        {


            string snum = Regex.Replace(((Button)sender).Text, @"[^0-9]", "");

            Bsentaku(snum, sender);
            // this.buttons[i]


        }

        private void Bsentaku(string snum,object sender)
        {
            currentsnum = snum;
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            currentbutton = (Button)sender;

            if (checkShusei.Checked)//shusei
            {

                ((Button)sender).BackColor = Color.LightGray;
                panel1.Visible = true;
                labelName.Text = snamer(stoi(snum) - 1);


                string line = dtYomi(srcdir, snum, month, day);
                string[] linea = line.Split(',');

                comboBoxKJ.Text = linea[1].Split(':')[0].Replace("\"", "");
                comboBoxKH.Text = linea[1].Split(':')[1].Replace("\"", "");
                comboBoxSJ.Text = linea[2].Split(':')[0].Replace("\"", "");
                comboBoxSH.Text = linea[2].Split(':')[1].Replace("\"", "");

                dateTimePicker1.Value = DateTime.Now;

                labelYobi.Text = "(" + dateTimePicker1.Value.ToString("ddd") + ")";

                return;

            }

            Task<string> sname = dtHozon(srcdir, snum, radioS.Checked, month, day, hour, minute);

            //******************iro henko & hozon
            string scolor = "";
            System.Media.SoundPlayer player = null;



            //  MessageBox.Show(currentbutton.ToString());

            // if (((Button)sender).BackColor == Color.Green)
            if (radioT.Checked)
            {
                ((Button)sender).BackColor = Color.Red;
                scolor = "Red";
                player = new System.Media.SoundPlayer(myDir + "\\taikin.wav");
                //非同期再生する
                player.Play();

            }
            // else if(((Button)sender).BackColor == Color.Red)
            else if (radioS.Checked)
            {
                ((Button)sender).BackColor = Color.LightGreen;
                scolor = "Green";
                player = new System.Media.SoundPlayer(myDir + "\\shukkin.wav");
                //非同期再生する
                player.Play();
            }

            player.Dispose();

            //  System.IO.File.WriteAllText(srcdir + "\\backup\\" +snum+ ".txt", scolor,Encoding.GetEncoding("Shift_JIS"));

            //  MessageBox.Show(dtL[0]+dtL[1]);
        }


        private string henkoVal(int tuki,int hi,string df)
        {
            //txtBoxCal 読み取り
            string[] gyo = textBoxCal.Text.Split('\n');
            //MessageBox.Show(gyo[0].Trim().Split(',')[0]+"="+gyo[0].Trim().Split(',')[1]);

            string hv = "";
            if (gyo.Length == 0) //txtBxocalがないとき
            {
                hv = df;
            }
            else
            {
                foreach (string md in gyo)//textboxCalの全行をチェック
                {
                    if (md.Trim().Split(',')[0].Split('/')[0] == tuki.ToString() && md.Trim().Split(',')[0].Split('/')[1] == hi.ToString())
                    {
                        // MessageBox.Show(md.Trim());
                        try
                        {
                            hv = md.Trim().Split(',')[1];
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error!:「年間カレンダーと異なる出勤日の設定」の各行にカンマ（,)を必ず入れてください");
                            //  Application.Exit();
                        }

                        break;
                    }
                    else
                    {
                        hv = df;
                    }
                }
            }
           
            
            return hv;
        }


        private void ExistCheckAndCreateDataFile(string srcdir,string snum,string monthx)
        {
            CreateDir(srcdir + "\\" + monthx + "\\data"); //月のフォルダないときは作成

            string srcfile = srcdir + "\\" + monthx + "\\data\\" + monthx + " " + snum.Trim() + ".csv";

            if (!File.Exists(srcfile)) //最初にファイルがあるかチェック　ないときは、作成
            {
                string alltxt = "";

                //16,"hg"
                //" 1","08:10","09:00",20
                //" 2*","00:00","00:00",0
                //" 31","00:00","00:00",0

               




                foreach(int idx in Enumerable.Range(0, 32))
                {
                    if (idx == 0)
                    {
                        alltxt = alltxt + snum.Trim() + "," + snamer(stoi(snum)-1) + "\r\n";
                    }
                    else
                    {
                        string astr = "";
                        DateTime dn=DateTime.Now;
                     

                        string dyobi = "";
                        DateTime dt;
                        try
                        {
                            dt = new DateTime(dn.Year, dn.Month, idx, 12, 0, 0, 0);
                            dyobi = dt.ToString("ddd");
                        }
                        catch(Exception ex)
                        {

                        }
                       if( dyobi== "土" || dyobi=="日")
                        {
                            astr = "*";
                        }

                        //MessageBox.Show(idx.ToString()+henkoVal(dn.Month, idx, astr));
                       
                        alltxt = alltxt + "\" "+idx.ToString().Trim()+henkoVal(dn.Month, idx, astr) + "\",\"00:00\",\"00:00\",0"+"\r\n";
                    }
                }

                System.IO.File.WriteAllText(srcfile, alltxt, Encoding.GetEncoding("Shift_JIS"));
            }
        

        }

        async private Task<string> dtHozon(string srcdir,string snum,bool shukkinflg,string monthx,string dayx,string hourx,string minutex) {
            //shokuin bango shutoku
     

            for(int i=0; i < shokuinsu-1; i++)
            {
                ExistCheckAndCreateDataFile(srcdir, (i+1).ToString(), monthx);// zenin shori sitahougii
            }
          

            string srcfile = srcdir + "\\" + monthx + "\\data\\" + monthx + " " + snum.Trim() + ".csv";

            string text = System.IO.File.ReadAllText(srcfile, Encoding.GetEncoding("Shift_JIS"));


            String[] dtL;
            dtL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Where(st => !String.IsNullOrEmpty(st)).ToArray();


            //sigyo shuryo jikoku load

            string text2 = System.IO.File.ReadAllText(srcdir + "\\maindata\\基礎データ.csv", Encoding.GetEncoding("Shift_JIS"));
            String[] dtL2;
            dtL2 = text2.Split(',');

            int sigyo = sTohun(dtL2[2].Replace("\"", ""));
            int shugyo = sTohun(dtL2[3].Replace("\"", ""));

            string alltxt = text;
            string backup = "";
            string tmp = "";
            string sname = "";
            foreach (string sline in dtL)
            {
                string stmp = "";
                string[] val = sline.Split(','); //val[1]:shukkin    val[2]:taikin
                string cjikoku = hourx + ":" + minutex;
                //int icjikoku = sTohun(cjikoku);

          
                if (stoi(val[0].Replace("\"", "").Trim().Replace("*","")) == stoi(dayx)&&val[0].Contains("\""))
                {
                   // MessageBox.Show(val[0].Replace("\"", "").Trim().Replace("*", "") + "=" + stoi(day).ToString());
                    int sa = 0;
                    int sa1 = 0;
                    int sa2 = 0;

                   // if (radioS.Checked) //shukkin
                   if(shukkinflg)
                    {
                        val[1] = "\"" + cjikoku + "\"";
                    }
                    else
                    {
                        val[2] = "\"" + cjikoku + "\"";
                    }

                   // MessageBox.Show(val[1].Trim()+ val[2].Trim());

                    if (val[1].Replace("\"", "") == "00:00" || val[2].Replace("\"", "") == "00:00")
                    {
                        sa = 0;
                      
                    }
                    else
                    {
                       

                        sa1 = sigyo - sTohun(val[1].Replace("\"", ""));
                        sa2 = sTohun(val[2].Replace("\"", "")) - shugyo;

                        if (sa1 < 0)
                        {
                            sa1 = 0;
                        }

                        if (sa2 < 0)
                        {
                            sa2 = 0;
                        }


                        sa = sa1 + sa2;
                    }


                    stmp = val[0] + "," + val[1] + "," + val[2] + "," + sa.ToString().Trim();


                    //kokode  shutaikin jikoku irohe hanei
                    //string snum,bool shukkinflg,string monthx,string dayx,string hourx,string minutex

                    shokikaIro(int.Parse(monthx), int.Parse(dayx), int.Parse(snum));  //その日の最初の人が初期化する　前日の職員番号と最終退勤時刻、今朝の職員番号と最初の出勤時刻も記録

                    henkoIro(int.Parse(monthx),int.Parse(dayx),int.Parse(snum),val[1],val[2]);

                    backup = stmp;
                }
                else
                {
                    stmp = sline;
                    string[] val2 = sline.Split(',');
                    if (!val2[0].Contains("\"")) //sento pass
                    {
                        sname = val2[1];

                    }
                }

                tmp = tmp + stmp + "\r\n";
            }
            alltxt = tmp;

            //MessageBox.Show(alltxt);
            //MessageBox.Show(srcfile);
            System.IO.File.WriteAllText(srcfile, alltxt, Encoding.GetEncoding("Shift_JIS"));

            //yarinasoi no tame
            ptext = text;
            psrcfile = srcfile;

              //backup
            System.IO.File.WriteAllText(srcdir + "\\backup\\" + monthx + "_"+dayx+"_"+ snum.Trim() + ".bak", sname+","+backup, Encoding.GetEncoding("Shift_JIS"));
            await Task.Delay(int.Parse(textBoxWaitTime.Text));
            //kinmukun irohyoji hanei
            KinmukunSosa ks = new KinmukunSosa(this); // 自フォームへの参照を渡す
            
            for (int ii = 0; ii < int.Parse(comboBoxKurikaesiKaisu.Text); ii++)
            {
                ks.testc(false, int.Parse(textBoxWaitTime.Text)); //falseでリストボックスに表示しない
               
            }
            
           


            return sname;
        }

        private string snamer(int bango)
        {
            string textx = System.IO.File.ReadAllText(textBoxSrc.Text, Encoding.GetEncoding("Shift_JIS"));
            List<string> stL = new List<string>();

            stL = ((IEnumerable<string>)textx.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();

            String[] sname = stL.ToArray(); //name   

            return sname[bango];
        }



        private string dtYomi(string srcdir, string snum,string monthx,string dayx)
        {
            string srcfile = srcdir + "\\" + monthx + "\\data\\" + monthx + " " + snum.Trim() + ".csv";

            string text = System.IO.File.ReadAllText(srcfile, Encoding.GetEncoding("Shift_JIS"));
            String[] dtL;
            dtL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Where(st => !String.IsNullOrEmpty(st)).ToArray();

            string stmp = "";
            foreach (string sline in dtL)
            {     
                string[] val = sline.Split(',');  
                // MessageBox.Show(val[0]+"="+ stoi(day).ToString());
                if (stoi(val[0].Replace("\"", "").Replace("*","").Trim()) == stoi(dayx) && val[0].Contains("\""))
                {
                    stmp = sline;
                }         
            }
            return stmp;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "default.csv";
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            //ofd.InitialDirectory = @"C:\";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter = "ファイル(*.csv;*.csv)|csv.*|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに選択されるものを指定する
            //2番目の「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 2;
            //タイトルを設定する
            ofd.Title = "開くファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //OKボタンがクリックされたとき、選択されたファイル名を表示する
                //Console.WriteLine(ofd.FileName);
                textBoxSort.Text = ofd.FileName;
                Properties.Settings.Default.txtSort                    = textBoxSort.Text;

                Properties.Settings.Default.Save();

             


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string text = System.IO.File.ReadAllText(textBoxSort.Text,Encoding.GetEncoding("Shift_JIS"));
            List<string> lst = text.Split('\n').ToList();
            foreach(string tx in lst)
            {
                listBox1.Items.Add(tx);
            }
        }


        private int stoi(string st)
        {
            int result;

            bool isParsed = int.TryParse( st, out result);
            if (isParsed)
            {
                return result;
            }
            else
            {
                return 0;
            }
            
        }

        public static string Right(string str, int len)
        {
            if (len < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (str == null)
            {
                return "";
            }
            if (str.Length <= len)
            {
                return str;
            }
            return str.Substring(str.Length - len, len);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string text = System.IO.File.ReadAllText(textBoxSort.Text, Encoding.GetEncoding("Shift_JIS"));
            List<string> stL = new List<string>();

            stL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();

           // List<string> lst = stL.OrderBy(st => stoi(st.Split(',')[2])).ToList();

            List<string> lst2 = stL.OrderBy(st => stoi(st.Split(',')[2])).Select(st=> st.Split(',')[0]).ToList(); //name
            List<string> lst3 = stL.OrderBy(st => stoi(st.Split(',')[2])).Select(st => st.Split(',')[2]+"."+ st.Split(',')[0]).ToList(); //num+name


            string textTmp="";
            int cc = 0;
            foreach (string tx in lst2)
            {
                if (cc != 0)
                {
                    listBox2.Items.Add(tx);
                    textTmp = textTmp + tx + "\r\n";
                }
                cc++;
            } 

            System.IO.File.WriteAllText(@textBoxDst.Text, textTmp, Encoding.GetEncoding("Shift_JIS"));

            string textTmp2 = "";
            int cc2 = 0;
            foreach (string tx in lst3)
            {
                if (cc2 != 0)
                {
                    textTmp2 = textTmp2 + tx + "\r\n";
                }
                cc2++;
                
               
            }
            System.IO.File.WriteAllText(@textBoxDst.Text+"x", textTmp2, Encoding.GetEncoding("Shift_JIS"));

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBoxTuki.Text == "")
            {
                MessageBox.Show("並び替えしたい月を入力してください");
                return;
            }

            string text = System.IO.File.ReadAllText(textBoxSort.Text, Encoding.GetEncoding("Shift_JIS"));
            List<string> stL = new List<string>();
            stL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();

            List<string> lst2 = stL.OrderBy(st => stoi(st.Split(',')[2])).ToList(); //name
         
            int cc = 0;
            foreach (string txs in lst2)
            {
                if (cc != 0)
                {
                    string srcn = txs.Split(',')[1];
                    string dstn = txs.Split(',')[2];
                    string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
                    string dstdir = Path.GetDirectoryName(textBoxDst.Text);

                    string tukis = Right("00"+textBoxTuki.Text,2);

                    string srcfile = srcdir + "\\" + tukis + "\\data\\" + tukis + " " + srcn + ".csv";
                    string dstfile = dstdir + "\\" + tukis + "\\data\\" + tukis + " " + dstn + ".csv";

                    // MessageBox.Show(srcfile);
                    // MessageBox.Show(dstfile);
                    try
                    {
                        System.IO.File.Copy(@srcfile, @dstfile, true);
                    }catch(Exception ex)
                    {

                    }

                     
                }
                cc++;
            }

            MessageBox.Show("変換修了");


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            labelDate.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm");

            // 現在の日時を取得します
            DateTime dNow2 = System.DateTime.Now;
            // 日時の各値を表示します
            //int year = dNow.Year;              // 現在の年を返します
            month = Right("00" + dNow2.Month.ToString(), 2);             // 現在の月を返します
            day = dNow2.Day.ToString();                 // 現在の日を返します
            hour = Right("00" + dNow2.Hour.ToString(), 2);               // 現在の時を返します
            minute = Right("00" + dNow2.Minute.ToString(), 2);


            shuttaikinKirikae(dNow2);



            //colorupdate per minute
            byocount++;
            if (byocount > 5)
            {
                
                byocount = 0;
                string text = "";
                try
                {
                    text = System.IO.File.ReadAllText(textBoxSrc.Text, Encoding.GetEncoding("Shift_JIS"));
                    List<string> stL = new List<string>();
                    stL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();
                    //*******************:hyoji
                    string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
                    for (int i = 0; i < stL.Count-1; i++)
                    {
                        colorRead(srcdir, i);
                      //  System.IO.File.WriteAllText(srcdir + "\\backup\\" + (i + 1).ToString() + ".txt", scolor, Encoding.GetEncoding("Shift_JIS"));

                    }
                }
                catch (Exception ex)
                {
                   // MessageBox.Show("cupdate:error="+ex.Message.ToString());
                }
            }

        }

        private void shuttaikinKirikae(DateTime dNow2)
        {
            //*******shuttakin kirikae
            DateTime dt = new DateTime(dNow2.Year, dNow2.Month, dNow2.Day, 12, 0, 0, 0);
            if (dNow2 > dt) //afternoon
            {
                radioT.Checked = true;
            }
            else
            {
                radioS.Checked = true;
            }
        }

        private void colorRead(string srcdir,int i)
        {
            string line = dtYomi(srcdir, (i+1).ToString(), Right("00" + DateTime.Now.Month.ToString(), 2), DateTime.Now.Day.ToString());
            string[] linea = line.Split(',');
            string hours, minutes, hourt, minutet;
            hours = linea[1].Split(':')[0].Replace("\"", "");
            minutes = linea[1].Split(':')[1].Replace("\"", "");
            hourt = linea[2].Split(':')[0].Replace("\"", "");
            minutet = linea[2].Split(':')[1].Replace("\"", "");

            //***********button color   jikoku de handan
            Button btn = this.buttons[i];

            if (hours == "00" && minutes == "00" && hourt == "00" && minutet == "00")
            {
                btn.BackColor = Color.LightGray;

            }
            else if ((hours != "00" || minutes != "00") && hourt == "00" && minutet == "00")
            {
                btn.BackColor = Color.LightGreen;

            }
            else if ((hourt != "00" || minutet != "00"))
            {
                btn.BackColor = Color.Red;

            }
        }

      

        private bool JikokuCheck()
        {
            String KJ = comboBoxKJ.Text;
            String KH = comboBoxKH.Text;
            String SJ = comboBoxSJ.Text;
            String SH = comboBoxSH.Text;
            if (IsTwoDigitNumber(KJ)&& IsTwoDigitNumber(KH) && IsTwoDigitNumber(SJ) && IsTwoDigitNumber(SH))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static bool IsTwoDigitNumber(string input)
        {
            // 文字列が2文字で、かつそれが数字であることをチェックします
            return input.Length == 2 && int.TryParse(input, out int number);
        }


        private void button7_Click(object sender, EventArgs e)
        {

            if (JikokuCheck()) 
            {

            }
            else
            {
                MessageBox.Show("すべて２けたの数字にしてから保存してください。（時刻の数字は選択により入力してください）");
                return;
            }
                

            checkShusei.Checked = false;

            string monthd = dateTimePicker1.Value.ToString("MM");
            string dayd = dateTimePicker1.Value.ToString("dd");

            //MessageBox.Show(monthd + dayd);
           // string snum = Regex.Replace(((Button)sender).Text, @"[^0-9]", "");
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);

            string hours =comboBoxKJ.Text;
            string minutes =comboBoxKH.Text;
            string hourt =comboBoxSJ.Text;
            string minutet =comboBoxSH.Text;

            dtHozon(srcdir, currentsnum,true, monthd, dayd,hours,minutes);
            dtHozon(srcdir, currentsnum, false, monthd, dayd, hourt, minutet);

            //***********button color   jikoku de handan             
            string scolor = "";
            if (hours=="00"&&minutes=="00"&&hourt=="00"&&minutet=="00"&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.LightGray;
                scolor="";
            }
            else if ((hours!="00"||minutes!="00")&&hourt=="00"&&minutet=="00"&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.LightGreen;
                scolor="Green";
            }else if ((hourt!="00"||minutet!="00")&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.Red;
                scolor="Red";
            } 

           //  System.IO.File.WriteAllText(srcdir + "\\backup\\" +currentsnum+ ".txt", scolor,Encoding.GetEncoding("Shift_JIS"));



            panel1.Visible = false;
        }

        private void textBoxKaisiJi_TextChanged(object sender, EventArgs e)
        {
           // button8.Visible = true;
        }

        private void textBoxShuryoJi_TextChanged(object sender, EventArgs e)
        {
           // button8.Visible = true;
        }

        private void textBoxKaisiHun_TextChanged(object sender, EventArgs e)
        {
           // button8.Visible = true;
        }

        private void textBoxShuryoHun_TextChanged(object sender, EventArgs e)
        {
           // button8.Visible = true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
          

    

        }

        private void textBoxKaisiJi_MouseDown(object sender, MouseEventArgs e)
        {
            button8.Visible = true;
        }

        private void textBoxKaisiHun_MouseDown(object sender, MouseEventArgs e)
        {
            button8.Visible = true;
        }

        private void textBoxShuryoJi_MouseDown(object sender, MouseEventArgs e)
        {
            button8.Visible = true;
        }

        private void textBoxShuryoHun_MouseDown(object sender, MouseEventArgs e)
        {
            button8.Visible = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click_1(object sender, EventArgs e)
        {

            Properties.Settings.Default.txtGakkoMei = textBoxGakkoMei.Text;
            Properties.Settings.Default.txtCal = textBoxCal.Text;
            Properties.Settings.Default.txtHandleNum=textBoxHandleNum.Text;
            Properties.Settings.Default.txtHandleNum2=textBoxHandleNum2.Text;

            Properties.Settings.Default.Save();


            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string kisodata = srcdir + "\\maindata\\基礎データ.csv";

            //"【鶴岡養護学校】",0,"8:30","17:00","8:30"
            string kaisiji = textBoxKaisiJi.Text + ":" + textBoxKaisiHun.Text;
            string shuryoji = textBoxShuryoJi.Text + ":" + textBoxShuryoHun.Text;
            string kisodt = "\"【" + textBoxGakkoMei.Text + "】\",0,\"" + kaisiji + "\",\"" + shuryoji + "\",\"" + kaisiji + "\"";

            File.WriteAllText(kisodata, kisodt, Encoding.GetEncoding("Shift_JIS"));

            // error check
            string[] gyo = textBoxCal.Text.Split('\n');
              string hv = "";
            foreach (string md in gyo)
            {
                if (!md.Contains(',')&&md!="")
                {
                    MessageBox.Show(""+md+"にカンマ（,)を入れてください");
                    Application.Exit();
                }

            }


            }

        private void radioS_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button9_Click_1(object sender, EventArgs e)
        {
            checkShusei.Checked = false;

            string monthd = dateTimePicker1.Value.ToString("MM");
            string dayd = dateTimePicker1.Value.ToString("dd");

            //MessageBox.Show(monthd + dayd);
            // string snum = Regex.Replace(((Button)sender).Text, @"[^0-9]", "");
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);

            string hours = comboBoxKJ.Text;
            string minutes = comboBoxKH.Text;
            string hourt = comboBoxSJ.Text;
            string minutet = comboBoxSH.Text;

          //  dtHozon(srcdir, currentsnum, true, monthd, dayd, hours, minutes);
          //  dtHozon(srcdir, currentsnum, false, monthd, dayd, hourt, minutet);

            //***********button color   jikoku de handan             
            string scolor = "";
            if (hours=="00"&&minutes=="00"&&hourt=="00"&&minutet=="00"&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.LightGray;
                scolor="";
            }
            else if ((hours!="00"||minutes!="00")&&hourt=="00"&&minutet=="00"&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.LightGreen;
                scolor="Green";
            }
            else if ((hourt!="00"||minutet!="00")&&dateTimePicker1.Value.Date==DateTime.Now.Date)
            {
                currentbutton.BackColor = Color.Red;
                scolor="Red";
            }

            panel1.Visible = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
           dateTimePicker1.Value= dateTimePicker1.Value.AddDays(1);

            hidukehenko(currentsnum,Right("00"+dateTimePicker1.Value.Month.ToString(),2),dateTimePicker1.Value.Day.ToString());


        }

        private void button11_Click(object sender, EventArgs e)
        {
            dateTimePicker1.Value = dateTimePicker1.Value.AddDays(-1);

            hidukehenko(currentsnum, Right("00"+dateTimePicker1.Value.Month.ToString(),2), dateTimePicker1.Value.Day.ToString());
        }


        private void hidukehenko(string snumt,string montht,string dayt)
        {
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);


                panel1.Visible = true;
                labelName.Text = snamer(stoi(snumt) - 1);

            //  MessageBox.Show(snumt);
            try
            {
                string line = dtYomi(srcdir, snumt, montht, dayt);
                string[] linea = line.Split(',');

                comboBoxKJ.Text = linea[1].Split(':')[0].Replace("\"", "");
                comboBoxKH.Text = linea[1].Split(':')[1].Replace("\"", "");
                comboBoxSJ.Text = linea[2].Split(':')[0].Replace("\"", "");
                comboBoxSH.Text = linea[2].Split(':')[1].Replace("\"", "");
            }
            catch(Exception ex)
            {
                MessageBox.Show("その日付のデータは未登録です");
            }
              

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            hidukehenko(currentsnum, Right("00" + dateTimePicker1.Value.Month.ToString(), 2), dateTimePicker1.Value.Day.ToString());


            labelYobi.Text = "("+dateTimePicker1.Value.ToString("ddd")+")";
        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }


        private void shokikaIro(int tuki, int hi, int shokuinbango)
        {
            List<string> lst = new List<string>();
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string srcfile = srcdir + "\\rsq.csv";
            try
            {
                string txt = System.IO.File.ReadAllText(srcfile, Encoding.GetEncoding("Shift_JIS"));
                lst = ((IEnumerable<string>)txt.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString() + ":rsq.csvデータがありません");
            }

            List<(int,string)> saishoNum = lst.Where(xs => xs.Split(',').Length == 3).Select((xs,index) =>(index, xs.Split(',')[2].Replace("\"", "") )).ToList() ;

            int pjikan = 0;
            int pindex = 0;
            string ptime = "";
            foreach((int,string) tpl in saishoNum)
            {
                string stime = tpl.Item2;
                int jikan = int.Parse(stime.Replace(":", ""));

                if (jikan > pjikan)
                {
                    pjikan = jikan;
                    ptime = stime;
                    pindex = tpl.Item1+1;
                }
            }
            string saishuNumJikoku = pindex.ToString()+",\""+ptime+"\"";


            string alltxt = "";
            int sbango = 0;
            
            //一行目の日付の読み取り
            string today = lst[0].Split(',')[0].Replace("\"", "");  //"2022-04-14" keisiki
            int gtuki = int.Parse(today.Split('-')[1]);
            int ghi = int.Parse(today.Split('-')[2]);

            //今日の日付
            DateTime dnow =  DateTime.Now;
           
            int ntuki = dnow.Month;
            int nhi = dnow.Day;
            string ndate = "\""+dnow.Year.ToString() + "-" + ntuki.ToString("00")+"-"+nhi.ToString("00")+"\"";
           // MessageBox.Show(ndate+","+lst[0].Split(',')[1]);

            string shukkinjikoku = "\"00:00\"";
            string taikinjikoku = "\"00:00\"";

            int saishugyo = lst.Count-1;


            if (ntuki != gtuki || nhi != ghi)  //rsq.csvが今日の日付でないとき、初期化する
            {
                string xgyo;
                foreach (string gyo in lst)
                {
                    xgyo = "";
                    
                    if (sbango == 0)
                    {
                        xgyo = ndate + "," + lst[0].Split(',')[1];
                    }

           
                    if (sbango > 0 && sbango < saishugyo)
                    {
                        //出勤、退勤の状況   1,"08:03","00:00"keisiki
                        string irobango = "0";
                        xgyo = irobango + "," + shukkinjikoku + "," + taikinjikoku;
                    }

                    if (sbango == saishugyo)
                    {
                        xgyo = saishuNumJikoku+","+shokuinbango.ToString()+",\""+dnow.Hour.ToString("00")+":"+dnow.Minute.ToString("00")+"\"";
                      //  MessageBox.Show(xgyo);
                    }

                    alltxt = alltxt + xgyo + "\r\n";
                    sbango++;
                }

                if (alltxt != "")
                {
                     System.IO.File.WriteAllText(srcfile, alltxt, Encoding.GetEncoding("Shift_JIS"));
                  //  MessageBox.Show(alltxt);
                }
            }

          

        }





        private void henkoIro(int tuki,int hi,int shokuinbango,string shukkinjikoku,string taikinjikoku)
        {
            List<string> lst = new List<string>();

            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string srcfile = srcdir+"\\rsq.csv";

            

            try
            {
                string txt = System.IO.File.ReadAllText(srcfile, Encoding.GetEncoding("Shift_JIS"));
                lst = ((IEnumerable<string>)txt.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();

                ptxt = txt;
                ppsrcfile = srcfile;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString()+":rsq.csvデータがありません");

            }

            string alltxt = "";
            int sbango = 0;
            string today = "";
           // int gnen = 0;
            int gtuki = 0;
            int ghi=0;
            foreach(string gyo in lst)
            {
                if (sbango==0)
                {
                    today=gyo.Split(',')[0].Replace("\"","");  //"2022-04-14" keisiki
                                                               //日付一致チェック
                    // gnen = int.Parse(today.Split('-')[0]);
                     gtuki = int.Parse(today.Split('-')[1]);
                     ghi = int.Parse(today.Split('-')[2]);
                }

                string xgyo = gyo;
                if (sbango==shokuinbango &&  gtuki==tuki && ghi==hi)
                {
                    //出勤、退勤の状況   1,"08:03","00:00"keisiki
                    string irobango = "0";
                    if (shukkinjikoku=="\"00:00\""&&taikinjikoku=="\"00:00\"")
                    {
                        irobango="0";
                    }

                    if (shukkinjikoku!="\"00:00\""&&taikinjikoku=="\"00:00\"")
                    {
                        irobango="1";
                    }

                    if (shukkinjikoku!="\"00:00\""&&taikinjikoku!="\"00:00\"")
                    {
                        irobango="2";
                    }

                    xgyo=irobango+","+shukkinjikoku+","+ taikinjikoku  ;
                }       

                alltxt=alltxt+xgyo+"\r\n";
               
                sbango++;
            }

            if (alltxt!="")
            {
                System.IO.File.WriteAllText(srcfile, alltxt, Encoding.GetEncoding("Shift_JIS"));
            }
            
        }




        private void button12_Click(object sender, EventArgs e)
        {
            KinmukunSosa ks = new KinmukunSosa(this); // 自フォームへの参照を渡す
            ks.testc(true,1000); //trueでリストボックスに表示もする
        }

        private void button13_Click(object sender, EventArgs e)
        {
            KinmukunSosa ks = new KinmukunSosa(this); // 自フォームへの参照を渡す
            ks.testhani(int.Parse(textBoxKaisiKinmukun.Text),int.Parse(textBoxShuryoKinmukun.Text)); //trueでリストボックスに表示もする
        }

        private void checkShusei_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            panel2.BackgroundImageLayout = ImageLayout.Zoom;
            try
            {
                panel2.BackgroundImage = System.Drawing.Image.FromFile(myDir + "\\henko.jpg");
            }
            catch(Exception ex)
            {
                MessageBox.Show("henko.jpg file not found");
            }
            

            panel2.Visible = true;

        }

        private void button16_Click(object sender, EventArgs e)
        {
            yarinaosi();
        }

        private Image BarcodeImage()
        {
            var bacodeWriter = new BarcodeWriter();
            // バーコードの種類
            //bacodeWriter.Format = BarcodeFormat.CODE_128;
            bacodeWriter.Format = BarcodeFormat.CODE_39;
            // サイズ
            bacodeWriter.Options.Height = 45;
            bacodeWriter.Options.Width = 240;
            // バーコード左右の余白
            bacodeWriter.Options.Margin = 40;
            // バーコードのみ表示するか
            // falseにするとテキストも表示する
            bacodeWriter.Options.PureBarcode = true;
            // pictureBoxBarcode.Image = bacodeWriter.Write("Test0123");
            //using System.Drawing;////////////////////////////////////////////////////////
            //描画先とするImageオブジェクトを作成する
            Bitmap canvas = new Bitmap(pictureBoxBarcode.Width, pictureBoxBarcode.Height);
            //ImageオブジェクトのGraphicsオブジェクトを作成する
            Graphics g = Graphics.FromImage(canvas);
            //画像ファイルを読み込んで、Imageオブジェクトとして取得する
            // Image img = Image.FromFile(@"C:\test\1.bmp");
            //画像をcanvasの座標(20, 10)の位置に描画する


            StringFormat stringFormat = new StringFormat();
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;
            // Draw the text and the surrounding rectangle.

            Font font1 = new Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point);

            // Create a StringFormat object with the each line of text, and the block
            // of text centered on the page.               
            stringFormat.Alignment = StringAlignment.Center;
            stringFormat.LineAlignment = StringAlignment.Center;
            // Draw the text and the surrounding rectangle.


            string text = "";
            try
            {
                text = System.IO.File.ReadAllText(textBoxSrc.Text, Encoding.GetEncoding("Shift_JIS"));
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新元職員データの欄に、職員データ.csvの場所を指定してください");
            }
            List<string> stL = new List<string>();
            stL = ((IEnumerable<string>)text.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None)).Select(st => st.Trim()).Where(st => !String.IsNullOrEmpty(st)).ToList<string>();


            int ox = -20;
            int oy = 0;
            int retusu = 4;
            int retukankaku = 220;
            int gyosu = 10;
            int gyokankaku = 110;

            int bango = 0;
            int iwidth = 200;
            int iheight = 30;

            foreach (string name in stL)
            {
                bango++;
                string pages = comboBoxPage.Text;

                if (bango >= 1 + (int.Parse(pages) - 1) * gyosu * retusu && bango <= int.Parse(pages) * gyosu * retusu)
                {
                    int dx = (bango - 1) % retusu;
                    int dy = (bango - 1 - (int.Parse(pages) - 1) * gyosu * retusu) / retusu;
                    Rectangle rect1 = new Rectangle(ox + dx * retukankaku, oy + dy * gyokankaku , 200, 40);

                    g.DrawString(string.Format("{0:D3}", bango) + "." + name, font1, Brushes.Black, rect1, stringFormat);
                    Image img = bacodeWriter.Write("9" + string.Format("{0:D3}", bango));
                    g.DrawImage(img, ox + dx * retukankaku, oy + dy * gyokankaku+40, img.Width,img.Height);

                }


            }

            //torikesi
            Rectangle rect2 = new Rectangle(ox, oy + 1100, 200, 40);
            g.DrawString("修正", font1, Brushes.Black, rect2, stringFormat);
            Image img2 = bacodeWriter.Write("9990");
            g.DrawImage(img2, ox, oy + 1100 + 40, img2.Width, img2.Height);

            //shusei
            Rectangle rect3 = new Rectangle(ox + 400, oy + 1100, 200, 40);
            g.DrawString("やり直し", font1, Brushes.Black, rect3, stringFormat);
            Image img3 = bacodeWriter.Write("9997");
            g.DrawImage(img3, ox + 400, oy + 1100 + 40, img3.Width, img3.Height);

            //date meisho
            Rectangle rect4 = new Rectangle(ox+50 , oy + 1210, 500, 40);
            g.DrawString(DateTime.Now.ToString()+"" + stL[stL.Count - 1], font1, Brushes.Black, rect4, stringFormat);


            font1.Dispose();
            //Imageオブジェクトのリソースを解放する
     
            //Graphicsオブジェクトのリソースを解放する
            g.Dispose();
            //PictureBox1に表示する
            // pictureBoxBarcode.Image = canvas;
            return canvas;


            // バーコードBitmapを作成
            //  using (var bitmap = bacodeWriter.Write("Test0123"))
            // {
            // 画像として保存
            //    var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Barcode.png");
            //   bitmap.Save(filePath, ImageFormat.Png);
            // }

        }


        private void button17_Click(object sender, EventArgs e)
        {


            pictureBoxBarcode.Image = BarcodeImage();
            //PrintDocumentオブジェクトの作成
            System.Drawing.Printing.PrintDocument pd =
                new System.Drawing.Printing.PrintDocument();
            //PrintPageイベントハンドラの追加
            pd.PrintPage +=
                new System.Drawing.Printing.PrintPageEventHandler(pd_PrintPage);

            //PrintDialogクラスの作成
            PrintDialog pdlg = new PrintDialog();

            pd.DefaultPageSettings.Margins.Left = 25;
            pd.DefaultPageSettings.Margins.Right = 25;
            pd.DefaultPageSettings.Margins.Top = 50;
            pd.DefaultPageSettings.Margins.Bottom = 50;


            //PrintDocumentを指定
            pdlg.Document = pd;


 


            //印刷の選択ダイアログを表示する
            if (pdlg.ShowDialog() == DialogResult.OK)
            {
                //OKがクリックされた時は印刷する
                pd.Print();
            }

        }

        private void pd_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            //Image img2 = canvas;
            Image img2 = BarcodeImage();
            //画像を描画する
            e.Graphics.DrawImage(img2, e.MarginBounds);
            //次のページがないことを通知する
            e.HasMorePages = false;
            //後始末をする
            img2.Dispose();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                this.KeyPreview = false;
                this.KeyPress -= Form1_KeyPress;

            }
            else
            {
                this.KeyPreview = true;
                this.KeyPress += Form1_KeyPress;
            }

        }

        async private void button18_Click(object sender, EventArgs e)
        {
            if(comboBox1.Text=="")
            {
                MessageBox.Show("月を選択してください");

            }
            else
            {
                string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
                int fileCount = Directory.GetFiles(srcdir + "\\" + comboBox1.Text + "\\xls").Length;

                List<int> ban = Enumerable.Range(1, fileCount).ToList();

                // var lst =  ban.Select(ii =>  hukugen(comboBox1.Text, ii.ToString().Trim()));

                var tasks = ban.Select(async ii => await hukugenAsync(comboBox1.Text, ii.ToString().Trim()));
                await Task.WhenAll(tasks);

                MessageBox.Show("復元が終了しました。");
            }

           

        }

        private async Task<int> hukugenAsync(string tuki,string bango)
        {
            label29.Text = bango + "個のデータを処理中";
            // 非同期処理の途中でExcelの操作があるため、COMオブジェクトを非同期で操作することは難しいですが、
            // 例えばファイルI/O操作などは非同期で行えます

            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string name="";
            // データを格納するためのリストまたは配列を作成
            List<List<string>> data = new List<List<string>>();

            await Task.Run(() =>
            {
                string srcDt = srcdir + "\\" + tuki + "\\xls\\" + tuki + " " + bango + ".xls";
                string dstDt = srcdir + "\\" + tuki + "\\data\\" + tuki.Trim() + " " + bango.Trim() + ".csv";

                // Excelアプリケーションオブジェクトを作成
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(srcDt);
                // ワークシートを取得
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // シートのインデックスは1から始まります


            

                name = worksheet.Cells[2, 7].Value2.ToString();

                // データを読み取るループ
                int row = 1;

                List<int> numbers = Enumerable.Range(1, 6).ToList();            

                while (worksheet.Cells[row, 1].Value2 != null)
                {
                    List<string> cellValueL = new List<string>();
                    foreach (int ii in numbers)
                    {
                        try
                        {
                            string wkst = worksheet.Cells[row, ii].Value2.ToString();

                            if (ii > 2)
                            {
                                // TimeSpan.FromDays(numericValue)
                                cellValueL.Add(TimeSpan.FromDays(double.Parse(wkst)).ToString());
                            }
                            else
                            {

                                cellValueL.Add(wkst);
                            }

                        }
                        catch (Exception ee)
                        {

                        }
                    }

                    data.Add(cellValueL);
                    row++;
                }

                // ワークブックとExcelアプリケーションをクローズ
                workbook.Close(false);
                excelApp.Quit();

                // リソースの解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            });

            // 非同期でファイル書き込みを実行
            //string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string alltxt = await ProcessExcelDataAsync(srcdir, tuki, bango,name,data);

            // 非同期でファイル書き込みが完了したことを示す結果を返す
            return 0;


            return 0;
        }

        private async Task<string> ProcessExcelDataAsync(string srcdir, string tuki, string bango,string name,List<List<string>> data)
        {
        
            string alltxt = bango + ",\"" + name + "\"\r\n";
            // 他の非同期処理などを実行
                     
            // データを使って何かを行う
            foreach (List<string> item in data)
            {
                int ci = 0;
                string num = "";
                bool ast = false;
                string start = "\"00:00\"";
                string end = "\"00:00\"";
                string jikan = "0";
                foreach (string item2 in item)
                {
                    ci++;

                    switch (ci)
                    {
                        case 1:
                            if (int.TryParse(item2, out int numericValue))
                            {
                                //  Console.WriteLine(numericValue);
                                if (numericValue < 32)
                                {
                                    num = "\" " + numericValue.ToString();
                                }
                                else
                                {
                                    num = "0";
                                }

                            }
                            else
                            {
                                // Console.WriteLine("文字列を整数に変換できませんでした。");
                                num = "0";
                            }
                            break;
                        case 2:
                            if (item2.Trim() == "1" || item2.Trim() == "7")
                            {
                                ast = true;
                            }
                            break;
                        case 3:
                            if (item2 == "1.00:00:00")
                            {
                                start = "\"00:00\"";
                            }
                            else
                            {
                                start = "\"" + item2.Substring(0, 5) + "\"";
                            }

                            break;
                        case 4:
                            end = "\"" + item2.Substring(0, 5) + "\"";
                            break;
                        case 5:
                            string[] dt = item2.Split(':');
                            jikan = (int.Parse(dt[0]) * 60 + int.Parse(dt[1])).ToString();
                            break;
                    }
                    Console.Write(item2 + ",");
                }

                if (num != "0")
                {
                    if (ast)
                    {
                        alltxt = alltxt + num + "*\"," + start + "," + end + "," + jikan + "\r\n";
                    }
                    else
                    {
                        alltxt = alltxt + num + "\"," + start + "," + end + "," + jikan + "\r\n";
                    }
                }
                Console.WriteLine("");
            }

            string dstDt = Path.Combine(srcdir, tuki, "data", tuki.Trim() + " " + bango.Trim() + ".csv");

            await Task.Run(() =>
            {
                // ファイル書き込みを非同期で行う
                System.IO.File.WriteAllText(dstDt, alltxt, Encoding.GetEncoding("Shift_JIS"));
            });

            // 非同期処理が完了したことを示す結果を返す
            return alltxt;
        }

        async private void button19_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" || comboBox2.Text=="")
            {
                MessageBox.Show("月や番号を選択してください");
            }
            else
            {
                // var tasks = ban.Select(async ii => await hukugenAsync(comboBox1.Text, ii.ToString().Trim()));
                //  await Task.WhenAll(tasks);

                var dd = hukugenAsync(comboBox1.Text, comboBox2.Text.Trim());

                MessageBox.Show("復元が終了しました。");
            }

             
        }

        async private void button20_Click(object sender, EventArgs e)
        {
            
            
            if (comboBox1.Text == "" || comboBox3.Text=="" || comboBox4.Text=="")
            {
                MessageBox.Show("月を選択、開始、終了番号を入力してください");

            }
            else
            {
              //  string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
              //  int fileCount = Directory.GetFiles(srcdir + "\\" + comboBox1.Text + "\\xls").Length;

                List<int> ban = Enumerable.Range(int.Parse(comboBox3.Text), int.Parse(comboBox4.Text)).ToList();

                // var lst =  ban.Select(ii =>  hukugen(comboBox1.Text, ii.ToString().Trim()));

                var tasks = ban.Select(async ii => await hukugenDummyAsync(comboBox1.Text, ii.ToString().Trim()));
                await Task.WhenAll(tasks);

                MessageBox.Show("復元が終了しました。");
            }
        }

        private void yarinaosi()
        {
            if (ptext != "")  //yarinasoi
            {
                System.IO.File.WriteAllText(psrcfile, ptext, Encoding.GetEncoding("Shift_JIS"));
                System.IO.File.WriteAllText(ppsrcfile, ptxt, Encoding.GetEncoding("Shift_JIS")); //rsq.csv hukkatu
                for (int i = 0; i < shokuinsu - 1; i++)
                {

                    colorRead(Path.GetDirectoryName(textBoxSrc.Text), i);
                }
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            panel2.Visible = false;
        }


        private async Task<int> hukugenDummyAsync(string tuki, string bango)
        {
            label29.Text = bango + "個のデータを処理中";

            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string name = "*";
            string alltxt = await ProcessExcelDataDummyAsync(srcdir, tuki, bango, name);
            return 0;

        }

        private void button21_Click(object sender, EventArgs e)
        {
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            string dirmonth = srcdir + "\\" + month + "\\data";
            textBoxRsq.Text = "";

            string tdata = "";

            try
            {
                using (StreamReader sr = new StreamReader(srcdir+"\\rsq.csv"))
                {
                    string line;
                    sr.ReadLine();
                    // ファイルから1行ずつ読み込む
                    while ((line = sr.ReadLine()) != null)
                    {
                        // Console.WriteLine(line); // 読み込んだ行をコンソールに表示（ここでは例として表示していますが、必要に応じて他の処理を行うことができます）

                     //   textBoxRsq.Text = textBoxRsq.Text+line+"\n";
                     tdata=tdata+line + "\n";                       

                    }
                }



            }
            catch (Exception ex)
            {
                textBoxRsq.Text="Error:ファイルを読み込めませんでした。" ;
            }


            List<String> tdt = tdata.Split('\n').ToList();       
            if (tdt.Any())
            {
                tdt.RemoveAt(tdt.Count - 1);  //末尾削除
            }

            List<(int,String,String)> erdt = tdt.Select((dx,index) => (num:index,fst:dx.Split(',')[1] ,snd:dx.Split(',')[2])).Where(dt2 => !corcheck(dt2.num,dt2.fst,dt2.snd)).ToList();

            foreach( (int num,String fst,String snd) in erdt)
            {
                textBoxRsq.Text = textBoxRsq.Text + (num+1).ToString()+"行目の"+fst+":"+snd+"がエラーの原因と思われます。rsq.csvの該当箇所を修正してください。\n";

            }

            if (textBoxRsq.Text == "")
            {
                textBoxRsq.Text = "各行の時刻は問題ないようです。";
            }
        }

        private bool corcheck(int num,String fs,String sn)
        {
            try
            {
                String KJ = fs.Replace("\"", "").Split(':')[0];
                String KH = fs.Replace("\"", "").Split(':')[1];
                String SJ = sn.Replace("\"", "").Split(':')[0];
                String SH = sn.Replace("\"", "").Split(':')[1];


                if (IsTwoDigitNumber(KJ) && IsTwoDigitNumber(KH) && IsTwoDigitNumber(SJ) && IsTwoDigitNumber(SH))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                return true;
            }

        
        }

        private void button22_Click(object sender, EventArgs e)
        {
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
        
            textBoxDt.Text = "";

            string tdata = "";

            string thisMonth = DateTime.Now.ToString("MM");
            string dirmonth = srcdir + "\\" + thisMonth + "\\data";

           // string folderPath = @"C:\YourFolderPath"; // フォルダのパスを指定

            List<string> filesInFolder = GetAllFilesInFolder(dirmonth);

            List<(String,String,List<String>)> lld = new List<(String,String,List<String>)>();

            comboBox6.Items.Clear();
            foreach (string fileName in filesInFolder)
            {

                comboBox6.Items.Add(fileName);
                //  Console.WriteLine(fileName);
                List<String> ddx = new List<String>();
                String simei="";
                try
                {
                    Encoding sjisEnc = Encoding.GetEncoding("Shift-JIS");
                    using (StreamReader sr = new StreamReader(dirmonth + "\\"+fileName,sjisEnc))
                    {                       
                        string line;
                        simei=sr.ReadLine();
                        while ((line = sr.ReadLine()) != null)
                        {
                            ddx.Add(line);
                        }
                    }
                }
                catch (Exception ex)
                {
                    textBoxDt.Text = textBoxDt.Text+"Error:読み込めないファイル(" +fileName+")がありました。\n";
                }

                lld.Add((fileName,simei,ddx));
            }


            List<(String, String)> erdt = lld.Select(tpl => (fname: tpl.Item1, chushutu: chushutu(tpl.Item1,tpl.Item2,tpl.Item3))).ToList();

            foreach ((String,String)  tpl in erdt)
            {
              //  if (tpl.Item2 != "")
                //{
              //    textBoxDt.Text = textBoxDt.Text + tpl.Item1 + "の" + tpl.Item2 + "がエラーの原因と思われます。該当箇所を修正してください。\n";
                    //}
            }

            if (textBoxDt.Text == "")
            {
                textBoxDt.Text = "各職員の時刻データに問題ないようです。";
            }

        }

        private String chushutu(String fname,String simei,List<String> item2L)
        {

            //" 31","00:00","00:00",0
            List<(String, String, String)> tmpL = new List<(string, string, string)>();
            tmpL= item2L.Select(st => (st.Split(',')[0], st.Split(',')[1], st.Split(',')[2])).Where( dt=> !corcheck(0,dt.Item2,dt.Item3) ).ToList();


            foreach ((String num, String fst, String snd) in tmpL)
            {
                textBoxDt.Text = textBoxDt.Text + "ファイル名:"+fname+"("+simei+")の "+ num + "日の" + fst + ":" + snd + "がエラーの原因と思われます\n";

            }

         


            return "";
        }


        static List<string> GetAllFilesInFolder(string folderPath)
        {
            List<string> fileNames = new List<string>();

            try
            {
                // フォルダ内のすべてのファイルを取得
                string[] files = Directory.GetFiles(folderPath);

                // ファイル名をリストに追加
                foreach (string file in files)
                {
                    fileNames.Add(Path.GetFileName(file));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("エラー: " + ex.Message);
            }

            return fileNames;
        }

        private void button23_Click(object sender, EventArgs e)
        {
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            //   StreamReader sr = new StreamReader(srcdir + "\\rsq.csv")

            Process.Start("notepad.exe", srcdir + "\\rsq.csv");
        }

        private void button24_Click(object sender, EventArgs e)
        {
            string srcdir = Path.GetDirectoryName(textBoxSrc.Text);
            //   StreamReader sr = new StreamReader(srcdir + "\\rsq.csv")
            string thisMonth = DateTime.Now.ToString("MM");
            string dirmonth = srcdir + "\\" + thisMonth + "\\data";
            Process.Start("notepad.exe", dirmonth+"\\"+comboBox6.Text);
        }

        private int yobi(int nen,int tuki,int hi)
        {
            DateTime targetDate = new DateTime(nen, tuki, hi); // 例として2023年10月12日を使用

            // 曜日を数値で取得
            return  (int)targetDate.DayOfWeek;
            
        }

        private async Task<string> ProcessExcelDataDummyAsync(string srcdir, string tuki, string bango, string name)
        {

            int year = int.Parse(comboBox5.Text);
            int month = int.Parse(tuki);
            

            string alltxt = bango + ",\"" + name + "\"\r\n";
            List<int> hidukeL = Enumerable.Range(1, 31).ToList();
            foreach (int dy in hidukeL)
            {
            
                 string start = "\"00:00\"";
                string end = "\"00:00\"";
                string jikan = "0";
                string num = "\" "+dy.ToString().Trim();

      
                    if (yobi(year,month,dy)==0 || yobi(year, month, dy) == 6)
                    {
                        alltxt = alltxt + num + "*\"," + start + "," + end + "," + jikan + "\r\n";
                    }
                    else
                    {
                        alltxt = alltxt + num + "\"," + start + "," + end + "," + jikan + "\r\n";
                    }
             

            }
            string dstDt = Path.Combine(srcdir, tuki, "data", tuki.Trim() + " " + bango.Trim() + ".csv");
            await Task.Run(() =>
            {
                System.IO.File.WriteAllText(dstDt, alltxt, Encoding.GetEncoding("Shift_JIS"));
            });
            return alltxt;
        }


    }
}
