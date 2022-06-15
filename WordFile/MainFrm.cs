using CCWin;
using Word = Microsoft.Office.Interop.Word;

namespace WordFile
{
    public partial class MainFrm : Skin_Color
    {
        public MainFrm()
        {
            InitializeComponent();
        }

        private Word.Application G_WordApplication;//定义Word应用程序字段
        private object G_Missing = Type.Missing;//定义G_Missing字段并添加引用

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbdia = new FolderBrowserDialog();

            if (fbdia.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = fbdia.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*G_OpenFileDialog = new OpenFileDialog();//创建打开文件对话框对象
            G_OpenFileDialog.Filter = "*.doc|*.doc";//筛选文件
            G_OpenFileDialog.InitialDirectory = textBox1.Text;
            DialogResult P_DialogResult =//弹出打开文件对话框
                G_OpenFileDialog.ShowDialog();*/

            string dir1 = textBox1.Text;

            ThreadPool.QueueUserWorkItem(//开始线程池
                (o) =>//使用Lambda表达式
                {
                    DirectoryInfo dires = new DirectoryInfo(dir1);
                    FileInfo[] files = dires.GetFiles("*.doc");
                    G_WordApplication =//创建Word应用程序对象
                           new Word.Application();
                    foreach (FileInfo f in files)
                    {
                        object P_FilePath = dir1 + "\\\\" + f;//创建Object对象

                        try
                        {
                            Word.Document P_Document = G_WordApplication.Documents.Open(//打开Word文档
                             ref P_FilePath, ref G_Missing, ref G_Missing,
                             ref G_Missing, ref G_Missing, ref G_Missing,
                             ref G_Missing, ref G_Missing, ref G_Missing,
                             ref G_Missing, ref G_Missing, ref G_Missing,
                             ref G_Missing, ref G_Missing, ref G_Missing,
                             ref G_Missing);

                            Word.Range P_Range = //得到文档范围
                                P_Document.Range(ref G_Missing, ref G_Missing);
                            Word.Find P_Find = //得到Find对象
                                P_Range.Find;

                            if (textBox2.Text != "")
                            {
                                this.Invoke(//在窗体线程中执行
                                    (MethodInvoker)(() =>//使用Lambda表达式
                                    {
                                        P_Find.Text = textBox2.Text;//设置查找的文本
                                        P_Find.Replacement.Text = textBox3.Text;//设置替换的文本
                                    }));
                                object P_Replace = Word.WdReplace.wdReplaceAll;//定义替换方式对象
                                bool P_bl = P_Find.Execute(//开始替换
                                    ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref P_Replace,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing);
                            }

                            if (textBox4.Text != "")
                            {
                                this.Invoke(//在窗体线程中执行
                                    (MethodInvoker)(() =>//使用Lambda表达式
                                    {
                                        P_Find.Text = textBox4.Text;//设置查找的文本
                                        P_Find.Replacement.Text = textBox5.Text;//设置替换的文本
                                    }));
                                object P_Replace = Word.WdReplace.wdReplaceAll;//定义替换方式对象
                                bool P_bl = P_Find.Execute(//开始替换
                                    ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref P_Replace,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing);
                            }

                            if (textBox6.Text != "")
                            {
                                this.Invoke(//在窗体线程中执行
                                    (MethodInvoker)(() =>//使用Lambda表达式
                                    {
                                        P_Find.Text = textBox6.Text;//设置查找的文本
                                        P_Find.Replacement.Text = textBox7.Text;//设置替换的文本
                                    }));
                                object P_Replace = Word.WdReplace.wdReplaceAll;//定义替换方式对象
                                bool P_bl = P_Find.Execute(//开始替换
                                    ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref P_Replace,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing);
                            }

                            if (textBox8.Text != "")
                            {
                                this.Invoke(//在窗体线程中执行
                                    (MethodInvoker)(() =>//使用Lambda表达式
                                    {
                                        P_Find.Text = textBox8.Text;//设置查找的文本
                                        P_Find.Replacement.Text = textBox9.Text;//设置替换的文本
                                    }));
                                object P_Replace = Word.WdReplace.wdReplaceAll;//定义替换方式对象
                                bool P_bl = P_Find.Execute(//开始替换
                                    ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref P_Replace,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing);
                            }
                            if (textBox10.Text != "")
                            {
                                this.Invoke(//在窗体线程中执行
                                    (MethodInvoker)(() =>//使用Lambda表达式
                                    {
                                        P_Find.Text = textBox10.Text;//设置查找的文本
                                        P_Find.Replacement.Text = textBox11.Text;//设置替换的文本
                                    }));
                                object P_Replace = Word.WdReplace.wdReplaceAll;//定义替换方式对象
                                bool P_bl = P_Find.Execute(//开始替换
                                    ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref P_Replace,
                                    ref G_Missing, ref G_Missing, ref G_Missing, ref G_Missing);
                            }

                            G_WordApplication.Documents.Save(//保存文档
                                ref G_Missing, ref G_Missing);
                            ((Word._Document)P_Document).Close(//关闭文档
                                ref G_Missing, ref G_Missing, ref G_Missing);
                        }
                        catch (Exception g)
                        {
                        }
                    }
                    ((Word._Application)G_WordApplication).Quit(//退出Word应用程序
                                ref G_Missing, ref G_Missing, ref G_Missing);
                    MessageBox.Show("替换完成");
                });
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}