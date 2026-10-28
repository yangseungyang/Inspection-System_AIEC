using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using COMCTRLEXLib;

namespace ComCtrlEX_CS
{
    public partial class MainFrm : Form
    {
        // 通信コントロールのインスタンス
        ComMng oComMng = new ComMng();
        // 通信中フラグ
        bool bConnect;

        public MainFrm()
        {
            InitializeComponent();
        }

        // ［ポート設定］ボタンクリック
        private void btn_PortNo_Click(object sender, EventArgs e)
        {
            short port;
            short iSync = (short)1;     // 同期型

            port = short.Parse(txt_PortNo.Text);
            port --;    // パラメータの引数は実ポート番号-1で指定。

            try
            {
                // 通信設定を起動
                oComMng.PortSetWakeUp(port, iSync);
            }
            catch (COMException oComException)
            {
                return;
            }
            catch (Exception oException)
            {
                return;
            }
        }

        // ［モニタ開始］ボタンクリック
        private void btn_Start_Click(object sender, EventArgs e)
        {
            short port;
            short iSync = (short)1;     // 同期型

            port = short.Parse(txt_PortNo.Text);
            port--;    // パラメータの引数は実ポート番号-1で指定。
            F3OpenInfo2 oInfo = new F3OpenInfo2();

            try
            {
                // 接続開始
                oComMng.Comopn2(port, out oInfo, iSync);

                // フォーム等のステータス変更
                btn_Start.Enabled = false;
                btn_End.Enabled = true;
                bConnect = true;

                // タイマの開始
                timer1.Enabled = true;
            }
            catch (COMException oComException)
            {
                return;
            }
            catch (Exception oException)
            {
                return;
            }
        }

        // ［モニタ終了］ボタンクリック
        private void btn_End_Click(object sender, EventArgs e)
        {
            short port;

            port = short.Parse(txt_PortNo.Text);
            port--;    // パラメータの引数は実ポート番号-1で指定。

            try
            {
                // 切断
                oComMng.Comcls(port);

                // フォーム等のステータス変更
                btn_Start.Enabled = true;
                btn_End.Enabled = false;
                bConnect = true;

                // タイマの開始
                timer1.Enabled = false;
            }
            catch (COMException oComException)
            {
                return;
            }
            catch (Exception oException)
            {
                return;
            }
        }

        // フォームロード処理
        private void MainFrm_Load(object sender, EventArgs e)
        {
            btn_End.Enabled = false;
            bConnect = false;
        }

        // タイマアップイベント
        private void timer1_Tick(object sender, EventArgs e)
        {
            int iPort = int.Parse(txt_PortNo.Text) - 1;

            try
            {
                short iGetData;     // モニタ結果
                // D0001の読み込み
                // [AccessType]デバイス種類（Dレジスタ＝４）
                // [Num]データ数
                // [Address]アドレス
                oComMng.DevrdWord((short)iPort, 4, 1, 1, out iGetData);

                // iGetDataをテキストボックスに格納
                txt_Data.Text = iGetData.ToString();
            }
            catch (COMException oComException)
            {
                return;
            }
            catch (Exception oException)
            {
                return;
            }
        }
    }
}