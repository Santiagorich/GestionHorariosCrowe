using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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

namespace GestionHorariosCrowe
{
    public partial class Configuracion : Form
    {

        public Configuracion()
        {
        InitializeComponent();
            if (File.Exists("Config.txt"))
            {
                IEnumerable<String> lines = File.ReadLines("Config.txt");
                bunifuTextBox2.Text = lines.ElementAt(0);
                bunifuTextBox3.Text = lines.ElementAt(1);

            }
        }

        private void bunifuButton5_Click(object sender, EventArgs e)
        {
            this.Hide();

        }
        private bool mouseDown;
        private Point lastLocation;

        private void conf_MouseDown(object sender, MouseEventArgs e)
        {

            mouseDown = true;
            lastLocation = e.Location;
        }

        private void conf_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void conf_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
        private void bunifuButton3_Click(object sender, EventArgs e)
        {
            String text = "";
            text+=bunifuTextBox2.Text +"\n";
            text+=bunifuTextBox3.Text + "\n";
            if (Main.logged == false)
            {
                Main.driver.Navigate().GoToUrl("https://apps.moyal.com.uy/GestionProyectos/servlet/seg_login");
            var username = Main.driver.FindElement(By.Id("vUSUARIOID"));
            var password = Main.driver.FindElement(By.Id("vUSUARIOPASSWORD"));
            username.SendKeys(bunifuTextBox2.Text);
            password.SendKeys(bunifuTextBox3.Text);
            var buttonlog = Main.driver.FindElement(By.CssSelector("[name='BUTTON1']"));
            buttonlog.Click();
            try
            {
                
                    Main.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#IMAGETIEMPOS")));
                
            }
            catch (Exception a)
            {
                MessageBox.Show("Error al logearse");
            }
            }
            this.Hide();
            File.WriteAllText("Config.txt", text);
            
            
        }

        private void Configuracion_Load(object sender, EventArgs e)
        {
            if (File.Exists("Config.txt"))
            {
                IEnumerable<String> lines = File.ReadLines("Config.txt");
                if (lines.ElementAt(0) != "" && lines.ElementAt(1) != "")
                {


                    Main.usernameS = lines.ElementAt(0);
                    Main.passwordS = lines.ElementAt(1);



                }
            }
            else
            {
                var frm = new Main();

                frm.Show();
                frm.Focus();

            }
        }
    }
}
