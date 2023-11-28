using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KinmukunConvert
{
    public class KinmukunSosa
    {

        public static Form1 fm1 =new Form1();
       

        public KinmukunSosa(Form1 ks)
        {
            fm1=ks;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);


        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int MK_LBUTTON = 0x0001;
        public static int GWL_STYLE = -16;

        public static int num = 0;
        public static int num2 = 0;
        public static bool listboxAddflg = false;
        public static int kaisix = 0;
        public static int shuryox = 0;

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hWnd, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32")]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);



        public void button1_Click(object sender, EventArgs e)
        {
   
        }


        public void testhani(int kaisi,int shuryo)
        {
            num=0;
            kaisix=kaisi;
            shuryox=shuryo;
           
            fm1.listBoxHandleNum.Items.Clear();
            listboxAddflg=false;

       
            IntPtr hWnd = FindWindow(null, "きんむくん メイン");
            if (hWnd != IntPtr.Zero)
            {
                //ウィンドウを作成したプロセスのIDを取得する
                int processId;
                GetWindowThreadProcessId(hWnd, out processId);
                //Processオブジェクトを作成する
                Process p = Process.GetProcessById(processId);

                //ウィンドウをアクティブにする
                SetForegroundWindow(p.MainWindowHandle);

                NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc3, new IntPtr(1));          



            }
        }




       async public void testc(bool lbaflg,int waitTime)
        {
            num=0;
            num2=0;
            fm1.listBoxHandleNum.Items.Clear();
            listboxAddflg=lbaflg;


            //タイトルが"無題 - メモ帳"のウィンドウを探す
            IntPtr hWnd = FindWindow(null, "きんむくん メイン");
            if (hWnd != IntPtr.Zero)
            {
                //ウィンドウを作成したプロセスのIDを取得する
                int processId;
                GetWindowThreadProcessId(hWnd, out processId);
                //Processオブジェクトを作成する
                Process p = Process.GetProcessById(processId);

         
             //   MessageBox.Show("プロセス名:" + p.ProcessName);

                //ウィンドウをアクティブにする
                SetForegroundWindow(p.MainWindowHandle);
                await Task.Delay(waitTime);

                NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc, new IntPtr(1));

                NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc2, new IntPtr(1));



            }

        }


        public class Window
        {
            public string ClassName;
            public string Title;
            public IntPtr hWnd;
            public int Style;
        }


        static bool EnumChildWindowsProc(IntPtr hWnd, IntPtr lParam)
        {
            var count = NativeMethods.GetWindowTextLength(hWnd);
            var sb = new StringBuilder(count + 1);
            var ret = NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);

            num=num+1;


            //  MessageBox.Show(num.ToString()+"番目："+hWnd.ToString()+":"+sb.ToString()+":"+sb.GetHashCode());
            if (listboxAddflg==true)
            {
                fm1.listBoxHandleNum.Items.Add(num.ToString()+"番目："+sb.ToString());
            }
            

            if (num==int.Parse(fm1.textBoxHandleNum.Text))  //番号指定 144
            {
                // マウスを押して放す
                SendMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0x000A000A);
                SendMessage(hWnd, WM_LBUTTONUP, 0x00000000, 0x000A000A);
            }


            // さらに自身の子ウインドウを列挙
            NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc, new IntPtr(lParam.ToInt32() + 1));

            return true;
        }

        static bool EnumChildWindowsProc2(IntPtr hWnd, IntPtr lParam)
        {
            var count = NativeMethods.GetWindowTextLength(hWnd);
            var sb = new StringBuilder(count + 1);
            var ret = NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);

            num2=num2+1;          

          
            if (num2==int.Parse(fm1.textBoxHandleNum2.Text))  //番号指定 144
            {
                // マウスを押して放す
                SendMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0x000A000A);
                SendMessage(hWnd, WM_LBUTTONUP, 0x00000000, 0x000A000A);
            }

           

            // さらに自身の子ウインドウを列挙
            NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc2, new IntPtr(lParam.ToInt32() + 1));

            return true;
        }

        static bool EnumChildWindowsProc3(IntPtr hWnd, IntPtr lParam)
        {
            var count = NativeMethods.GetWindowTextLength(hWnd);
            var sb = new StringBuilder(count + 1);
            var ret = NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
            num=num+1;
            if (num>=kaisix && num<=shuryox)  //番号指定 144? あたりをさがす
            {
                MessageBox.Show(num.ToString()+"番でチェックします");
                // マウスを押して放す
                SendMessage(hWnd, WM_LBUTTONDOWN, MK_LBUTTON, 0x000A000A);
                SendMessage(hWnd, WM_LBUTTONUP, 0x00000000, 0x000A000A);
            }
                   

            // さらに自身の子ウインドウを列挙
            NativeMethods.EnumChildWindows(hWnd, EnumChildWindowsProc3, new IntPtr(lParam.ToInt32() + 1));

            return true;
        }

        class NativeMethods
        {
            [return: MarshalAs(UnmanagedType.Bool)]
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool EnumWindows(
                EnumWindowsProc lpEnumFunc,
                IntPtr lParam
            );

            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool EnumChildWindows(
                IntPtr hWndParent,
                EnumWindowsProc lpEnumFunc,
                IntPtr lParam
            );

            [DllImport("user32.dll", SetLastError = true)]
            public static extern Int32 GetWindowTextLength(
                IntPtr hWnd
            );

            [DllImport("user32.dll", SetLastError = true)]
            public static extern Int32 GetWindowText(
                IntPtr hWnd,
                StringBuilder lpString,
                Int32 nMaxCount
            );
        }



        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr FindWindow(
            string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowThreadProcessId(
            IntPtr hWnd, out int lpdwProcessId);
    }
}
