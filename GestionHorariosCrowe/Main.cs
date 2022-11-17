using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GestionHorariosCrowe
{
    public partial class Main : Form
    {

       
        public static DateTime dt = DateTime.UtcNow;
        public static String today = dt.ToString("dd/MM/yy");
        public static String passwordS = "";
        public static String usernameS = "";
        public static Boolean logged = false;
        public static Boolean created = false;
        public static Boolean invisor = false;
        public static String createdurl;
        public static String oldtext;
        public static ChromeOptions ops = new ChromeOptions();
        public static List<string> allowedTypes2 = new List<string>();
        public static AutoCompleteStringCollection allowedTypes = new AutoCompleteStringCollection();
        public static IWebDriver driver;
        public static WebDriverWait wait;
        public static IJavaScriptExecutor js;
        public static ChromeDriverService service;
        public static Thread t;

        public void setup()
        {
            service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            //ops.AddArguments("headless");
            driver = new ChromeDriver(service, ops);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            js = (IJavaScriptExecutor)driver;
            t = new Thread(loading);
        }
        public Main()
        {
            InitializeComponent();
            setup();
            login();
            start();
        }
       public void start()
        {
            bunifuDatePicker1.Text = today;
            if (File.Exists("Horas.txt"))
            {
                IEnumerable<String> linesa = File.ReadLines("Horas.txt");
                bunifuTextBox2.Text = linesa.ElementAt(1);
                bunifuTextBox3.Text = linesa.ElementAt(2);
                bunifuTextBox4.Text = linesa.ElementAt(0);

            }
            if (!File.Exists("Proyectos.txt"))
            {
                File.WriteAllText("Proyectos.txt", "");
                getproyectos();
            }
            else
            {
                foreach (string line in File.ReadLines("Proyectos.txt"))
                {
                    allowedTypes.Add(line);
                }
                bunifuTextBox4.AutoCompleteCustomSource = allowedTypes;

            }
            registro();
            if (created == true)
            {
                showproyectos(createdurl);
            }
            if (!File.Exists("Tareas.txt"))
            {
                File.WriteAllText("Tareas.txt", "");
                gettareas(bunifuTextBox4.Text);

            }
            else
            {
                foreach (string line in File.ReadLines("Tareas.txt"))
                {
                    allowedTypes2.Add(line);
                }
                bunifuDropdown1.DataSource = allowedTypes2;

            }
            if (allowedTypes.Contains(bunifuTextBox4.Text) && bunifuTextBox4.Text != "")
            {
                gettareas(bunifuTextBox4.Text);
            }
            
        }

        private bool mouseDown;
        private Point lastLocation;
        public void loading()
        {
            t.Start();
        }
        private void Main_MouseDown(object sender, MouseEventArgs e)
        {

            mouseDown = true;
            lastLocation = e.Location;
        }

        private void Main_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {

                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }

        }

        private void Main_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
        public void login()
        {
            if (File.Exists("Config.txt"))
            {
                IEnumerable<String> lines = File.ReadLines("Config.txt");
                bunifuTextBox2.Text = lines.ElementAt(0);
                bunifuTextBox3.Text = lines.ElementAt(1);
                usernameS = lines.ElementAt(0);
                passwordS = lines.ElementAt(1);

            }
            driver.Navigate().GoToUrl("https://apps.moyal.com.uy/GestionProyectos/servlet/seg_login");
            var username = driver.FindElement(By.Id("vUSUARIOID"));
            var password = driver.FindElement(By.Id("vUSUARIOPASSWORD"));
            username.SendKeys(usernameS);
            password.SendKeys(passwordS);
            var buttonlog = driver.FindElement(By.CssSelector("[name='BUTTON1']"));
            buttonlog.Click();
            try
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#IMAGETIEMPOS")));
                logged = true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Error al logearse");
            }
            }

        public void registro()
        {
            bunifuDataGridView1.Rows.Clear();
            bunifuDataGridView1.Columns[0].HeaderText = "Fechas";
            bunifuDataGridView1.Columns[1].HeaderText = "Nombre";
            bunifuDataGridView1.Columns[2].HeaderText = "Horas";
            bunifuDataGridView1.Columns[3].HeaderText = "Link";
            bunifuDataGridView1.Columns[3].Visible = false;
            created = false;
            bunifuButton7.Visible = false;
            driver.Navigate().GoToUrl("https://apps.moyal.com.uy/GestionProyectos/servlet/gpgt_registrohoras_registro");
            String day = "";
            if (bunifuDatePicker1.Text != "")
            {
                day = bunifuDatePicker1.Text;
            }
            else
            {
                day = today;
            }
            var mes = new SelectElement(driver.FindElement(By.Id("vFMESACTIVIDAD")));
            mes.SelectByValue(day.Split('/')[1]);
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector(".GridOdd")));
            IList<IWebElement> grid = (IList<IWebElement>)driver.FindElements(By.Id("GridcabezalContainerDiv"));
            IList<IWebElement> rows = (IList<IWebElement>)grid[0].FindElements(By.CssSelector("tr"));
            //IEnumerable<String> rowsS = grid[0].GetAttribute("innerText").Split('\n');
            foreach (var row in rows.Skip(1))
            {
                IList<IWebElement> cols = (IList<IWebElement>)row.FindElements(By.CssSelector("a"));
                string[] split = row.GetAttribute("innerText").Trim().Split('\t');
                bunifuDataGridView1.Rows.Add(split[0], split[1], split[2],"", cols[1].GetAttribute("href"));
                if (split[0].Contains(day))
                {
                    created = true;
                    createdurl = cols[1].GetAttribute("href");
                }

            }
            if(created == false && day!=today)
            {
                bunifuLabel7.Text = "No se han encontrado registros para el " + day.Split('/')[0];
            }else if (created == false && day == today)
            {
                bunifuLabel7.Text = "No se han encontrado registros para hoy";
            }
            else
            {
                bunifuLabel7.Text = "";

            }
            /*foreach (var row in rows)
            {
                textgrid += row.GetAttribute("innerText") + "\n";
            }*/



        }
        public void showproyectos(string url)
        {

            bunifuDataGridView1.Rows.Clear();
            bunifuDataGridView1.Columns[0].HeaderText = "Proyecto";
            bunifuDataGridView1.Columns[1].HeaderText = "Tarea";
            bunifuDataGridView1.Columns[2].HeaderText = "Hora Inicio";
            bunifuDataGridView1.Columns[3].HeaderText = "Hora Final";
            bunifuDataGridView1.Columns[3].Visible = true;

            bunifuButton7.Visible = true;
            driver.Navigate().GoToUrl(url);
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#GridregistrosContainerTbl")));
            IWebElement table = driver.FindElement(By.CssSelector("#GridregistrosContainerTbl > tbody"));
            IList<IWebElement> rows = (IList<IWebElement>)table.FindElements(By.CssSelector("tr"));
            foreach (IWebElement row in rows)
            {
                IList<IWebElement> cols = row.FindElements(By.CssSelector("td"));
                SelectElement proyecselect = new SelectElement(cols[2].FindElement(By.CssSelector("select")));
                SelectElement tareaselect = new SelectElement(cols[3].FindElement(By.CssSelector("select")));
                IWebElement inin = cols[4].FindElement(By.CssSelector("input"));
                IWebElement infin = cols[5].FindElement(By.CssSelector("input"));

                bunifuDataGridView1.Rows.Add(proyecselect.SelectedOption.Text.Trim(), tareaselect.SelectedOption.Text.Trim(),inin.GetAttribute("value").Trim(), infin.GetAttribute("value").Trim());
            }
            SelectElement tareas = new SelectElement(driver.FindElement(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_TAREASNEGOCIOSID_0001")));
            foreach (var option2 in tareas.Options)
            {
                allowedTypes2.Add(option2.Text);
            }


        }
        public void getproyectos()
        {


            String day = "";
            if (bunifuDatePicker1.Text != "")
            {
                day = bunifuDatePicker1.Text;
            }
            else
            {
                day = today;
            }
            String urln = "https://apps.moyal.com.uy/GestionProyectos/servlet/gpgt_registrohoras_ingreso?INS,0,230,," + day.Split('/')[1] + "," + day.Split('/')[2];

            if (driver.Url != urln )
            {
                driver.Navigate().GoToUrl(urln);
            }
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#vDIACOMBO")));
                
                SelectElement dia = new SelectElement(driver.FindElement(By.CssSelector("#vDIACOMBO")));
                IWebElement buttonag = driver.FindElement(By.CssSelector("[name='NUEVAFILA']"));
                dia.SelectByValue(day.Split('/')[0].TrimStart('0'));
                buttonag.Click();
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_PROYECTOID_0001")));
                SelectElement proyecto = new SelectElement(driver.FindElement(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_PROYECTOID_0001")));
            allowedTypes.Clear();
            string text = "";
            foreach (var option in proyecto.Options){
                allowedTypes.Add(option.Text);
                text += option.Text + Environment.NewLine;
            }
     
                File.WriteAllText("Proyectos.txt", text);
      
            bunifuTextBox4.AutoCompleteCustomSource = allowedTypes;
            bunifuLabel3.Text = day;





        }


        public void gettareas(String proyectostr)
        {
            String day = "";
            if (bunifuDatePicker1.Text != "")
            {
                day = bunifuDatePicker1.Text;
            }
            else
            {
                day = today;
            }
            String urln = "https://apps.moyal.com.uy/GestionProyectos/servlet/gpgt_registrohoras_ingreso?INS,0,230,," + day.Split('/')[1] + "," + day.Split('/')[2];
            if (driver.Url != urln)
            {
                driver.Navigate().GoToUrl(urln);
            }
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#vDIACOMBO")));

            SelectElement dia = new SelectElement(driver.FindElement(By.CssSelector("#vDIACOMBO")));
            IWebElement buttonag = driver.FindElement(By.CssSelector("[name='NUEVAFILA']"));
            dia.SelectByValue(day.Split('/')[0].TrimStart('0'));
            buttonag.Click();
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_PROYECTOID_0001")));
            SelectElement proyecto = new SelectElement(driver.FindElement(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_PROYECTOID_0001")));
            SelectElement tarea = new SelectElement(driver.FindElement(By.CssSelector("#vGPGT_REGISTROHORAS_TAREAS_TAREASNEGOCIOSID_0001")));
            proyecto.SelectByText(proyectostr);
            allowedTypes2.Clear();
            string text = "";

            foreach (var option in tarea.Options)
            {
                allowedTypes2.Add(option.Text);
                text += option.Text + Environment.NewLine;

            }
            File.WriteAllText("Tareas.txt", text);
            bunifuDropdown1.DataSource = allowedTypes2;

        }

        public void agregarhora(String url,String horainiciostr, String horafinalstr, String proyectostr)
        {

            String day = "";
            if (bunifuDatePicker1.Text != "")
            {
                day = bunifuDatePicker1.Text;
            }
            else
            {
                day = today;
            }
            var match = Regex.Match(horainiciostr, "^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$", RegexOptions.IgnoreCase);
            var match2 = Regex.Match(horafinalstr, "^([0-1]?[0-9]|2[0-3]):[0-5][0-9]$", RegexOptions.IgnoreCase);

            if (!match.Success || !match2.Success)
            {
                MessageBox.Show("Hora invalida");
                return;
            }
            if (driver.Url != url)
            {
                driver.Navigate().GoToUrl(url);
            }
            if (url.Contains("INS,"))
            {
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#vDIACOMBO")));
                SelectElement dia = new SelectElement(driver.FindElement(By.CssSelector("#vDIACOMBO")));
                dia.SelectByValue(day.Split('/')[0].TrimStart('0'));
            }
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#GridregistrosContainerTbl")));
            IWebElement buttonag = driver.FindElement(By.CssSelector("[name='NUEVAFILA']"));
            buttonag.Click();
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("input[value='00:00']")));
            IWebElement tabla = driver.FindElement(By.CssSelector("#GridregistrosContainerTbl"));
            IList<IWebElement> tablerows = tabla.FindElements(By.CssSelector(".GridOdd"));
            IWebElement lastrow = tablerows[tablerows.Count - 1];
            IList<IWebElement> fields = lastrow.FindElements(By.CssSelector("td"));
            SelectElement proyecto = new SelectElement(fields[2].FindElement(By.CssSelector("select")));
            SelectElement tarea = new SelectElement(fields[3].FindElement(By.CssSelector("select")));
            proyecto.SelectByText(proyectostr);
            tarea.SelectByText("Desarrollo");
            IWebElement horainicio = fields[4].FindElement(By.CssSelector("input"));
            IWebElement horafinal = fields[5].FindElement(By.CssSelector("input"));
            setAttributeValue(horainicio, horainiciostr);
            setAttributeValue(horafinal, horafinalstr);
            IWebElement buttonconf = driver.FindElement(By.CssSelector("[name='CONFIRMAR']"));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(buttonconf));
            js.ExecuteScript("gx.evt.setGridEvt(32,null);gx.evt.execEvt('EENTER.',this)");
            bunifuTextBox2.Text = horafinalstr;
            bunifuTextBox3.Text = "00:00";

        }


        public void crear()
        {


            String day = "";
            if (bunifuDatePicker1.Text != "")
            {
                day = bunifuDatePicker1.Text;
            }
            else
            {
                day = today;
            }
            String text = "";
            text += bunifuTextBox4.Text + "\n";
            text += bunifuTextBox2.Text + "\n";
            text += bunifuTextBox3.Text + "\n";
            File.WriteAllText("Horas.txt", text);
            driver.Navigate().GoToUrl("https://apps.moyal.com.uy/GestionProyectos/servlet/gpgt_registrohoras_registro");
            var mes = new SelectElement(driver.FindElement(By.Id("vFMESACTIVIDAD")));
            mes.SelectByValue(day.Split('/')[1]);
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector("#GridcabezalContainerRow_0001")));
            IList<IWebElement> grid = (IList<IWebElement>)driver.FindElements(By.Id("GridcabezalContainerDiv"));
            IList<IWebElement> rows = (IList<IWebElement>)grid[0].FindElements(By.CssSelector("tr"));

            foreach (var row in rows)
            {
                IList<IWebElement> columns = row.FindElements(By.TagName("td"));

                    if (row.GetAttribute("innerText").Contains(day))
                    {
                        IWebElement delete = row.FindElement(By.CssSelector("[title='Eliminar']"));
                        driver.Navigate().GoToUrl(delete.FindElement(By.XPath("./..")).GetAttribute("href"));
                        js.ExecuteScript("gx.evt.setGridEvt(32,null);gx.evt.execEvt('EENTER.',this)");
                        break;
                    }   
            }
            agregarhora(("https://apps.moyal.com.uy/GestionProyectos/servlet/gpgt_registrohoras_ingreso?INS,0,230,," + day.Split('/')[1] + "," + day.Split('/')[2]), bunifuTextBox2.Text, bunifuTextBox3.Text, bunifuTextBox4.Text);
            //System.Threading.Thread.Sleep(1000);
            registro();



        }
        public void setAttributeValue(IWebElement elem, String value)
        {
            js.ExecuteScript("arguments[0].setAttribute(arguments[1],arguments[2])",
                elem, "value", value
            );
        }

        private void bunifuButton1_Click(object sender, EventArgs e)
        {
            var frm = new Configuracion();
            frm.Location = this.Location;
            frm.StartPosition = FormStartPosition.Manual;
            frm.FormClosing += delegate { this.Show(); };
            frm.Show();
        }

        private void bunifuButton4_Click(object sender, EventArgs e)
        {
            bunifuButton4.Text = "Cargando...";
            getproyectos();
            bunifuButton4.Text = "Refrescar";
        }

        private void bunifuButton5_Click(object sender, EventArgs e)
        {
            driver.Quit();
            Application.Exit();
        }

        private void bunifuButton2_Click(object sender, EventArgs e)
        {
            crear();

        }

        private void bunifuButton3_Click(object sender, EventArgs e)
        {
            registro();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void bunifuDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /*var frm = new Visor();
            frm.Location = this.Location;
            frm.StartPosition = FormStartPosition.Manual;
            frm.FormClosing += delegate { this.Show(); };
            frm.Show();*/
            showproyectos(bunifuDataGridView1.SelectedRows[0].Cells[4].Value.ToString());
            invisor = true;
        }

        private void bunifuButton6_Click(object sender, EventArgs e)
        {
            /*var frm = new Visor();
            frm.Location = this.Location;
            frm.StartPosition = FormStartPosition.Manual;
            frm.FormClosing += delegate { this.Show(); };
            frm.Show();*/
            registro();
        }

        private void bunifuLabel6_Click(object sender, EventArgs e)
        {

        }

        private void bunifuButton7_Click(object sender, EventArgs e)
        {
            invisor = false;
            registro();

        }

        private void bunifuCheckBox1_CheckedChanged(object sender, Bunifu.UI.WinForms.BunifuCheckBox.CheckedChangedEventArgs e)
        {
            if (bunifuCheckBox1.Checked == false && created == true)
            {
                bunifuButton3.Visible = true;
            }
            else
            {
                bunifuButton3.Visible = false;

            }
        }

        private void bunifuButton3_Click_1(object sender, EventArgs e)
        {
            String horafinalstr = bunifuTextBox3.Text;
            agregarhora(createdurl, bunifuTextBox2.Text, horafinalstr, bunifuTextBox4.Text);
            showproyectos(createdurl);

        }

        private void bunifuDatePicker1_ValueChanged(object sender, EventArgs e)
        {
            String oldcreatedurl = createdurl;
            if (bunifuDatePicker1.Text != "")
            {
                registro();
                if (oldcreatedurl != createdurl)
                {
                    showproyectos(createdurl);
                }
                
            }
        }

        private void bunifuButton6_Click_1(object sender, EventArgs e)
        {
            if (invisor == true)
            {

                showproyectos(createdurl);
            }
            else
            {
                registro();

            }
        }

        private void bunifuLabel8_Click(object sender, EventArgs e)
        {

        }

        private void bunifuTextBox4_TextChanged(object sender, EventArgs e)
        {
            if (allowedTypes.Contains(bunifuTextBox4.Text)&& bunifuTextBox4.Text != oldtext)
            {
                gettareas(bunifuTextBox4.Text);
                oldtext = bunifuTextBox4.Text;
            }
        }

        private void bunifuDropdown1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void bunifuLabel3_Click(object sender, EventArgs e)
        {

        }

        private void bunifuSeparator1_Click(object sender, EventArgs e)
        {

        }
    }
}
