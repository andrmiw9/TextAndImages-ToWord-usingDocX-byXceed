using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Net;


// TODO: Asking (and deleting if necessary) pics, that were saved, and picks, which were downloaded BEFORE changing h1 name (they have old name so far)
// TODO: Fix awful bug with Styles, god bless fucking Word
// TODO: add feature to ignore templates count, in order to quit whenever user consider exiting (yeah, infinite loop while user exits)
// TODO: mb some editing tools for nice look in Word will come in handy
// TODO: Searching and picking pic from Internet option
// TODO: перевод на русский мб


namespace WorkWithWord2_2_graphical_DocX_lib_
{
    public partial class Form1 : Form
    {
        public int picnumber = 0;
        public static int numt = -1;

        public Form1(int smth)                  // smth > 0, проверку сделали в Entering.cs
        {
            InitializeComponent();
            richTextBox1.Enabled = false;
            numt = smth;
            richTextBox1.Text += "Template number: " + (Entering.Templates - numt + 1) + "\n";
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.InitialImage = System.Drawing.Image.FromFile(initialpicfullpath);
            pictureBox2.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox2.InitialImage = System.Drawing.Image.FromFile(initialpicfullpath);
            using (StreamWriter str = new StreamWriter(picnumberpath, false, System.Text.Encoding.Default))
            {
                str.WriteLine(1);               // Обнуляем counter картинок для конкретного места
            }
        }
        private string initialpicfullpath = @"C:\Users\Sinitza\Documents\APlaces\InitialImage.jpg";
        private string picnumberpath = @"C:\Users\Sinitza\Documents\APlaces\Templates\number.txt";
        private string pictodownloadpath = @"C:\Users\Sinitza\Documents\APlaces\Photo\";
        private static bool completed = false;

        private void button1_Click(object sender, EventArgs e)  // Create and paste Template button
        {
            //pictureBox1.Image = System.Drawing.Image.FromFile(pictodownloadpath + "ex1.jpg");
            if (isImg1Seen && isImg2Seen)       // Проверка на то, что юзер увидел картинки
            {
                string h1 = textBox1.Text;
                string h2 = textBox2.Text;      // h1, h2, h3 - ненужные по сути переменные
                string h3 = textBox3.Text;
                if (numt <= -1)                 // Проверка на то, что кол-во template'ов <<адекватное>> (по сути не нужна, но пускай будет)
                {
                    MessageBox.Show("Errorrrrr");
                    Thread.Sleep(2000);
                    Environment.Exit(0);
                }
                else
                {
                    #region working but inactive check for h1&h2
                    /*
                    //Проверка для h1 и h2
                    if (string.IsNullOrWhiteSpace(h1))
                    {
                        smartswitcher += 1;
                    }
                    if (string.IsNullOrWhiteSpace(h2))
                    {
                        if (smartswitcher == 1)
                        {
                            Console.WriteLine("Error: Empty arguments!");                       //empty arg-s
                            Console.ReadKey();
                            Environment.Exit(0);
                        }
                        else                                
                        {
                            EditT(h1, default);             //1 arg(h1)
                            smartswitcher += 2;
                        }
                    }
                    if(smartswitcher == 1)
                    {
                        EditT(default, h2);                 //1 arg(h2)
                    }
                    if(smartswitcher == 0)
                    {
                        EditT(h1, h2);                      //normal arg-s(h1,h2)
                    }
                    */
                    #endregion
                    if (!String.IsNullOrWhiteSpace(LastIMG1Name) && !String.IsNullOrWhiteSpace(LastIMG2Name))   // Проверка на непустые названия картинок
                    {
                        Programm.EditT(LastIMG1Name, LastIMG2Name, h1, h2, h3);
                        Programm.AddT();
                        //richTextBox1.Text += "Template number " + (Entering.Templates - numt + 1) + " added! \n"; // not works
                        if (numt > 1)
                        {
                            this.Hide();    // это п*здец, но я заметил это уже слишком поздно... R.I.P.
                            Form1 a = new Form1(--numt);                            // Вызов новой формы
                            a.StartPosition = FormStartPosition.Manual;     // Код для вызова новой формы на месте старой
                            a.Location = Location;                          // Код для вызова новой формы на месте старой
                            a.Show();
                        }
                        else
                        {
                            MessageBox.Show("Last Template done, exiting...");      // Конец (Официальный Выход Из программы)
                            using (StreamWriter str = new StreamWriter(picnumberpath, false, System.Text.Encoding.Default))
                            {
                                str.WriteLine(1);                                   // Подчищаем хвосты
                            }
                            Thread.Sleep(1000);
                            Application.Exit();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Ошибка: путь к одной из двух активных картинок не распознан!");
                    }
                    //AddPic();
                    //AddT();
                }
                //pictureBox1.Load();
                //pictureBox1.Refresh();

                //pictureBox1.Image = System.Drawing.Image.FromFile(imgname);
                /*byte[] bmpData = new WebClient().DownloadData(imgadress);
                using (MemoryStream ms = new MemoryStream(bmpData))
                {
                    Bitmap bmp = (Bitmap)Bitmap.FromStream(ms);
                    bmp.Save("captcha.jpg");
                }*/
            }
            else
            {
                MessageBox.Show("Сначала рефрешните картинки");
            }
        }
        // Environment.Exit(0);
        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            //MessageBox.Show("Closed!");

            if (!completed && !String.IsNullOrWhiteSpace(textBox1.Text))       // Проверка
            {
                using (StreamReader sr = new StreamReader(picnumberpath))
                {
                    picnumber = int.Parse(sr.ReadToEnd());
                }
                if (picnumber != 1 && !(String.IsNullOrWhiteSpace(LastIMG1Name)))
                {
                    File.Delete(LastIMG1Name);
                }
                if (picnumber != 1 && !(String.IsNullOrWhiteSpace(LastIMG2Name)))
                {
                    File.Delete(LastIMG2Name);
                }
            }
            using (StreamWriter str = new StreamWriter(picnumberpath, false, System.Text.Encoding.Default))
            {
                str.WriteLine(1);
            }
            Application.Exit();
        }
        private string LastIMG1Name = "";
        private string LastIMG2Name = "";
        bool isImg1Seen = false;                // Для проверки на то, что юзер увидел обе картинки
        bool isImg2Seen = false;                // Фактически означает что у нас точно есть как минимум 2 картинки

        private void button2_Click(object sender, EventArgs e)  // refresh img1 (левую)
        {
            isImg1Seen = true;
            pictureBox1.Show();
            if (String.IsNullOrWhiteSpace(textBox1.Text))       // Проверка на непустое название h1
            {
                MessageBox.Show("Please enter h1 first");
            }
            else
            {
                using (StreamReader sr = new StreamReader(picnumberpath))
                {
                    picnumber = int.Parse(sr.ReadToEnd());
                }
                if (picnumber != 1 && !(String.IsNullOrWhiteSpace(LastIMG1Name)))
                {
                    File.Delete(LastIMG1Name);
                    picnumber--;
                }
                string imgadress = textBox4.Text;
                string imgname = pictodownloadpath + textBox1.Text + "_" + picnumber + ".jpg";
                try
                {
                    new WebClient().DownloadFile(imgadress, imgname);
                    pictureBox1.ImageLocation = (imgadress);
                    LastIMG1Name = imgname;
                    using (StreamWriter str = new StreamWriter(picnumberpath, false, System.Text.Encoding.Default))
                    {
                        str.WriteLine(picnumber + 1);
                    }
                }
                catch (Exception e1)
                {
                    MessageBox.Show("Exception: " + e1);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)  // refresh img2 (правую)
        {
            isImg2Seen = true;
            pictureBox2.Show();
            if (String.IsNullOrWhiteSpace(textBox1.Text))       // Проверка на непустое название h1
            {
                MessageBox.Show("Please enter h1 first");
            }
            else
            {
                using (StreamReader sr = new StreamReader(picnumberpath))
                {
                    picnumber = int.Parse(sr.ReadToEnd());
                }
                if (picnumber != 1 && !(String.IsNullOrWhiteSpace(LastIMG2Name)))
                {
                    File.Delete(LastIMG2Name);
                    picnumber--;
                }
                string imgadress = textBox5.Text;
                string imgname = pictodownloadpath + textBox1.Text + "_" + picnumber + ".jpg";
                try
                {
                    new WebClient().DownloadFile(imgadress, imgname);
                    pictureBox2.ImageLocation = (imgadress);
                    LastIMG2Name = imgname;
                    using (StreamWriter str = new StreamWriter(picnumberpath, false, System.Text.Encoding.Default))
                    {
                        str.WriteLine(picnumber + 1);
                    }
                }
                catch (Exception e1)
                {
                    MessageBox.Show("Exception: " + e1);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)  // Сохранить левую картинку
        {
            if (String.IsNullOrWhiteSpace(textBox1.Text))       // Проверка на непустое название h1
            {
                MessageBox.Show("Please enter h1 first");
            }
            else
            {
                if (isImg1Seen)
                {
                    LastIMG1Name = "";
                    richTextBox1.Text += "IMG 1 (on the left) will be saved \n";
                    pictureBox1.Hide();
                    textBox4.Text = "";
                }
                else
                {
                    MessageBox.Show("Please refresh img1 first");
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)  // Сохранить правую картинку
        {
            if (String.IsNullOrWhiteSpace(textBox1.Text))       // Проверка на непустое название h1
            {
                MessageBox.Show("Please enter h1 first");
            }
            else
            {
                if (isImg2Seen)
                {
                    LastIMG2Name = "";
                    richTextBox1.Text += "IMG 2 (on the right) will be saved \n";
                    pictureBox2.Hide();
                    textBox5.Text = "";
                }
                else
                {
                    MessageBox.Show("Please refresh img2 first");
                }
            }
        }
    }
}