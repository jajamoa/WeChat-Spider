#define HISTORY

using FlaUI.Core.Input;
using FlaUI.Core.Shapes;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;
using System.Threading;

namespace UIAutomationTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //var app = FlaUI.Core.Application.Launch("notepad.exe");
            var app = FlaUI.Core.Application.Attach(11004);
            var pad = FlaUI.Core.Application.Attach(3128);
            var ans = FlaUI.Core.Application.Attach(8840);
            int endcount = 0;
            string prestr="";
            string ans1="";
            using (var automation = new UIA3Automation())
            {
                var window = app.GetMainWindow(automation);
                window.Focus();
                Console.WriteLine(window.ActualWidth);
                Console.WriteLine(window.ActualHeight);
                Console.WriteLine(window.FrameworkType);
                var mouseX = window.Properties.BoundingRectangle.Value.Left + 50;
                var mouseY = window.Properties.BoundingRectangle.Value.Top + 200;
                Mouse.Position = new Point(mouseX, mouseY);
                Mouse.Click(MouseButton.Left);
                var newwindow = app.GetAllTopLevelWindows(automation)[0];
                var padwindow = pad.GetAllTopLevelWindows(automation)[0];
                var answindow = ans.GetAllTopLevelWindows(automation)[0];

                //Excel
                XSSFWorkbook workBook = new XSSFWorkbook();  //实例化XSSF
                XSSFSheet sheet = (XSSFSheet)workBook.CreateSheet();  //创建一个sheet
                IRow frow0 = sheet.CreateRow(0);
                frow0.CreateCell(0).SetCellValue("序号");
                frow0.CreateCell(1).SetCellValue("标题");
                frow0.CreateCell(2).SetCellValue("来源");
                frow0.CreateCell(3).SetCellValue("时间");
                frow0.CreateCell(4).SetCellValue("阅读量");
                frow0.CreateCell(5).SetCellValue("字数");
                string saveFileName = "C:\\Users\\cityscience\\Desktop\\12.xlsx";
                int count = 0;
                int T = 20000;
                while (T--!=0)
                {
                    //高度更高的窗口为主窗口
                    if (newwindow.ActualHeight>window.ActualHeight)
                    {
                        var temp = newwindow;
                        newwindow = window;
                        window = temp;
                    }
                   

                    newwindow.Focus();
                    Thread.Sleep(100);
                    mouseX = newwindow.Properties.BoundingRectangle.Value.Left + 10;
                    mouseY = newwindow.Properties.BoundingRectangle.Value.Top + 10;
                    Mouse.Position = new Point(mouseX, mouseY);
                    Mouse.LeftClick();
                    Thread.Sleep(1500);
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_C);
                    padwindow.Focus();
                    mouseX = padwindow.Properties.BoundingRectangle.Value.Left + 10;
                    mouseY = padwindow.Properties.BoundingRectangle.Value.Top + 10;
                    Mouse.Position = new Point(mouseX, mouseY);
                    Mouse.Click(MouseButton.Left);
                    
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);
                    Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_S);
                    Thread.Sleep(1000);
                    string text = "";
                    StreamReader sr = new StreamReader("C:\\Users\\cityscience\\Desktop\\new.txt", Encoding.GetEncoding("GB2312"));
                    text = sr.ReadToEnd();
                    if (text == prestr) endcount++;
                    if (endcount == 10) break;
                    if (text != prestr && text.LastIndexOf("阅读") > 0)
                    {
                        endcount = 0;
                        //var textBox = padwindow.FindAllDescendants()[0].AsTextBox();
                        var ansBox = answindow.FindAllDescendants()[0].AsTextBox();
                        //var text = textBox.Text;               
                        IRow frow1 = sheet.CreateRow(++count);
                        int pos = text.IndexOf("\n");
                        ans1 = ans1 + "标题：" + text.Substring(0, pos-1);
                        frow1.CreateCell(1).SetCellValue(text.Substring(0, pos - 1));

                        string tempstr=text.Substring(pos+1, text.IndexOf("\n", pos+3)-pos-1);
                        if (tempstr.IndexOf("原创：") > 0) tempstr=tempstr.Remove(0, 6);
                        ans1 = ans1 + "\n来源：" + tempstr.Substring(0, tempstr.LastIndexOf(" "));
                        frow1.CreateCell(2).SetCellValue(tempstr.Substring(0, tempstr.LastIndexOf(" ")));

                        ans1 = ans1 + "\n时间：" + tempstr.Substring(tempstr.LastIndexOf(" "),tempstr.Length-tempstr.LastIndexOf(" "));
                        frow1.CreateCell(3).SetCellValue(tempstr.Substring(tempstr.LastIndexOf(" "), tempstr.Length - tempstr.LastIndexOf(" ")));

                        tempstr = text.Substring(text.LastIndexOf("阅读")+3, 5);
                        if (tempstr.IndexOf("在看") > 0) tempstr = tempstr.Substring(0, tempstr.IndexOf("在看") - 1);
                        if (tempstr.IndexOf("在") > 0) tempstr = tempstr.Substring(0, tempstr.IndexOf("在") - 1); 
                        ans1 = ans1 + "\n阅读 "+ tempstr + " 字数 " + text.Length.ToString() + "\n\n";
                        //写入excel
                        //新建一行
                        
                        frow1.CreateCell(0).SetCellValue(count);                        
                        frow1.CreateCell(4).SetCellValue(tempstr);
                        frow1.CreateCell(5).SetCellValue(text.Length.ToString());

                        //防Block
                        Random rd = new Random();
                        int Sleep_Time = (int)(rd.Next(5, 11) * 30 * 1000);                     
                        Thread.Sleep(Sleep_Time);

                        //实时保存
                        using (FileStream fs = new FileStream(saveFileName, FileMode.Create, FileAccess.Write))
                        {
                            workBook.Write(fs);  //写入文件
                            workBook.Close();  //关闭
                        }

                        answindow.Focus();
                        ansBox.Enter(ans1);
                        //prestr=text;
                    }
                    prestr = text;
                    sr.Close();
                    //Console.WriteLine(textBox.Name);
                    //new term
                    double scnum = -1.99;
                    window.Focus();
                    #if (HISTORY)
                        mouseX = window.Properties.BoundingRectangle.Value.Left + 40;
                        mouseY = window.Properties.BoundingRectangle.Value.Top + 50;
                        Mouse.Position = new Point(mouseX, mouseY);              
                        Mouse.Click(MouseButton.Left);
                        Thread.Sleep(1500);
                        scnum = -0.99;
                    #endif
                    mouseX = window.Properties.BoundingRectangle.Value.Left + 200;
                    mouseY = window.Properties.BoundingRectangle.Value.Top + 200;
                    Mouse.Position = new Point(mouseX, mouseY);
                    if (text.LastIndexOf("阅读") > 0) Mouse.Scroll(scnum);
                    else Mouse.Scroll(-0.05);
                    Thread.Sleep(700);
                    Mouse.Click(MouseButton.Left);
                    if (text.LastIndexOf("阅读") <= 0) Mouse.Click(MouseButton.Left);
                }
                //Console.WriteLine(window.Title);
                //Console.WriteLine(newwindow.Title);
                using (FileStream fs = new FileStream(saveFileName, FileMode.Create, FileAccess.Write))
                {
                    workBook.Write(fs);  //写入文件
                    workBook.Close();  //关闭
                }
                workBook.Close();
                Console.ReadKey();
            }
        }
    }
}
