using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using BrWo.Dominio;
using BrWo.Tools;
using DeathByCaptcha;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Exception = System.Exception;

namespace ConsigSP
{
    class Program
    {
        static public IWebDriver driver;
        static public int regraRobo;
        static string CPF;
        static string imgFileName;
        static Captcha captcha;
        static int maxErrosEmSequencia = 10;
        static DAL dal = new DAL();
        static Client client;
        private static string dbcUser;
        private static string dbcPws;
        static bool erroFatal = false;
        static bool finalizador = false;
        static bool captchaOk = false;
        static int qtdErros = 0;

        static void Main(string[] args)
        {
            do
            {
                erroFatal = false;
                try
                {
                    captchaOk = false;
                    Daileon();
                }
                catch
                {
                    try { driver.Quit(); } catch { }
                    if (finalizador == true)
                    {
                        return;
                    }
                }

            } while (true);
        }
        static void Daileon()
        {

            Settings settings = new Settings();
            var tag = DateTime.Now.Ticks.ToString();
            var usuario = settings.Get("Usuario");
            var regraRobo = int.Parse(settings.Get("Regra"));
            var senha = settings.Get("Senha");
            dbcUser = settings.Get("dbcUser");
            dbcPws = settings.Get("dbcPws");
            client = (Client)new SocketClient(dbcUser, dbcPws);
            maxErrosEmSequencia = int.Parse(settings.Get("MaxErrosEmSequencia"));
            Console.WriteLine($"\nVersão {settings.Get("Versao")}");
            Console.WriteLine($"Tag: {tag}");
            Console.WriteLine($"Regra: {regraRobo}\n");
            OutLine($"Saldo: {client.Balance / 100}");

            var imgFilePath = AppDomain.CurrentDomain.BaseDirectory + "img\\";
            if (!Directory.Exists(imgFilePath))
                Directory.CreateDirectory(imgFilePath);

            imgFileName = imgFilePath + tag + ".png";

            DAL dal = new DAL();
 
            BancoDeCPFs bancoDeCPFs = new BancoDeCPFs(tag);

            AbrirPagina();
            try
            {
                while (captchaOk == false)
                {
                    Login(usuario, senha);
                    IWebElement elementCaptcha = driver.FindElement(By.Id("cipCaptchaImg"));
                    IWebElement elementTxt = driver.FindElement(By.Id("captcha"));
                    IWebElement click = driver.FindElement(By.Id("idc"));
                    DeCaptcha(elementCaptcha, elementTxt, click);
                    if (captchaOk == false)
                        qtdErros++;
                    else
                        qtdErros = 0;
                    if (qtdErros >= maxErrosEmSequencia)
                    {
                        SairDoRobo(TipoErro.QuantidadeMaximaDeErrosEmSequencia);
                    }
                }
            }
            catch (Exception ex)
            {
                OutLine(ex.Message);
                Log.LogThis(ex.Message);
                SairDoRobo();
            }
            EscolhaMenu();

            //---------Pesquisa CPF---------//
            do
            {
                while (!erroFatal)
                {
                    try
                    {
                    Inicio:
                        ClienteRobo cliente = new ClienteRobo("ConsigSP");

                        CPF = bancoDeCPFs.FetchConsigSPCPF(regraRobo);
                        if (CPF == "")
                        {
                            Console.WriteLine($"Não há CPFs para processar na regra {regraRobo}");
                            Thread.Sleep(10000);
                            continue;
                        }
                        IWebElement elemento = driver.FindElement(By.Id("cpfServidor"));
                        elemento.Clear();
                        Thread.Sleep(3000);
                        elemento.Click();
                        elemento.SendKeys(BancoDeCPFs.FormataCPF(CPF));
                        driver.FindElement(By.XPath("//input[@value='Pesquisar']")).Click();
                        Thread.Sleep(3000);

                        try
                        {
                            TrataErrosCPF();
                        }
                        catch (Exception ex)
                        {
                            bancoDeCPFs.UpdateCPFConsigSPStatus(CPF, "ERRO", ex.Message);
                            continue;
                        }

                        Thread.Sleep(5000);
                        cliente.CPF = CPF;

                        if (!(IsElementPresent(By.XPath("//*[@id='divContentResultado']/span/div/div/div[2]/div/div[2]/span"))))
                        {
                            goto Inicio;
                        }
                        cliente.Nome = driver.FindElement(By.XPath("//*[@id='divContentResultado']/span/div/div/div[2]/div/div[2]/span")).Text;
                        cliente.Orgao = driver.FindElement(By.XPath("//*[@id='divContentResultado']/span/div/div/div[2]/div/div[3]/div/span")).Text;
                        cliente.Matricula = driver.FindElement(By.XPath("//*[@id='divContentResultado']/span/div/div/div[2]/div/div[3]/div[2]/span")).Text;

                        double margemBruta = 0;
                        Double.TryParse(driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/form/div[2]/div/div[4]/div[1]/div/span/div/div[3]/div[2]/div/div/span/div/table/tbody/tr/td[2]/span")).Text, out margemBruta);
                        cliente.MargemBruta = margemBruta;

                        double margemDisponivel = 0;
                        Double.TryParse(driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/form/div[2]/div/div[4]/div[1]/div/span/div/div[4]/div[2]/div/div/span/div/table/tbody/tr/td[2]/span")).Text, out margemDisponivel);
                        cliente.MargemDisponivel = margemDisponivel;

                        string scriptCliente = cliente.GeraScriptInsert();
                        dal.ExecutarScript(scriptCliente);
                    }
                    catch (Exception ex)
                    {
                        OutLine(ex.Message);
                        Log.LogThis(ex.Message);
                    }

                    bancoDeCPFs.UpdateCPFConsigSPStatus(CPF, "OK");
                }
            }
            while (!erroFatal);
            Thread.Sleep(5000);
            driver.Quit();
        }
        private static void MudaCaptcha()
        {
            for (var x = 0; x < 2; x++)
            {
                driver.FindElement(By.Id("ida")).Click();
            }
            Thread.Sleep(1000);
            driver.FindElement(By.Id("captcha")).Click();

        }
        private static void TrataErrosCPF()
        {
            if (IsElementPresent(By.XPath("//*[@id='divEtapaError2']/div/ul/li/span")))
            {
                var msg = driver.FindElement(By.XPath("//*[@id='divEtapaError2']/div/ul/li/span")).Text;
                if (msg.ToLowerInvariant() == "ESCC8017: Servidor não permite a Consulta da Margem".ToLowerInvariant())
                {
                    OutLine("\nEste servidor não permite a consulta da margem");
                    FecharAviso();
                    throw new Exception(TipoErro.PrivacidadeServidor);
                }
                else if (IsElementPresent(By.XPath("//*[@id='divEtapaError2']/div/ul/li/span")))
                {
                    if (msg.ToLowerInvariant() == "ESCC0289: Órgão não vinculado ao CPF informado no SCC.".ToLowerInvariant())
                    {
                        OutLine("\nNenhum servidor encontrado");
                        FecharAviso();
                        throw new Exception(TipoErro.CPFNaoEncontrado);
                    }
                }
                else if (IsElementPresent(By.XPath("//*[@id='divEtapaError2']/div/ul/li/span")))
                {
                    if (msg.ToLowerInvariant() == "CPF do Servidor é obrigatório e deve ser informado.".ToLowerInvariant())
                    {
                        OutLine("\nCPF do Servidor é obrigatório e deve ser informado.");
                        FecharAviso();
                        throw new Exception(TipoErro.NaoBuscouCPF);
                    }
                }
                else if (IsElementPresent(By.XPath("//*[@id='divEtapaError2']/div/ul/li/span")))
                {
                    if (msg.ToLowerInvariant() == "ESCC0018: Servidor não localizado".ToLowerInvariant())
                    {
                        OutLine("\nNenhum servidor encontrado");
                        FecharAviso();
                        throw new Exception(TipoErro.CPFNaoEncontrado);
                    }
                }
            }
        }
        private static void FecharAviso()
        {
            driver.FindElement(By.XPath("/html/body/div/div/div[2]/div/form/div[2]/div/div[1]/div/span[2]")).Click();
        }
        private static void FecharAvisoLogin()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/div[2]/form/div[3]/div/div/div/div[1]/div/span[2]")).Click();
        }
        private static void EscolhaMenu()
        {
            if (regraRobo == 1)
            {
                driver.FindElement(By.XPath("/html/body/div/div/form/div[2]/div/div[4]/div[2]/fieldset/span/label/input")).Click();
            }
            if (regraRobo == 2)
            {
                driver.FindElement(By.XPath("/html/body/div/div/form/div[2]/div/div[5]/div[2]/fieldset/span/label[1]/input")).Click();
            }
            if (regraRobo == 3)
            {
                driver.FindElement(By.XPath("/html/body/div/div/form/div[2]/div/div[5]/div[2]/fieldset/span/label[2]/input")).Click();
            }
            if (regraRobo == 4)
            {
                driver.FindElement(By.XPath("/html/body/div/div/form/div[2]/div/div[5]/div[2]/fieldset/span/label[3]/input")).Click();
            }
            if (regraRobo == 5)
            {
                driver.FindElement(By.XPath("/html/body/div/div/form/div[2]/div/div[6]/div[2]/fieldset/span/label/input")).Click();
            }
            Thread.Sleep(1000);
            NavegaMenu();
        }
        private static void NavegaMenu()
        {
            driver.FindElement(By.XPath("//input[@value='Acessar']")).Click();
            Thread.Sleep(3000);
            driver.FindElement(By.XPath("//*[@id='menu']/div[2]/a/span")).Click();
            driver.FindElement(By.XPath("//*[@href='/consignatario/pesquisarMargem']")).Click();
        }
        private static bool DeCaptcha(IWebElement elementCaptcha, IWebElement elementTxt, IWebElement click)
        {
            try
            {
                var arrScreen = ((ITakesScreenshot)driver).GetScreenshot().AsByteArray;
                using (var msScreen = new MemoryStream(arrScreen))
                {
                    using (var bmpScreen = new Bitmap(msScreen))
                    {
                        var cap = elementCaptcha;
                        var rcCrop = new Rectangle(cap.Location, cap.Size);
                        var imgCap = bmpScreen.Clone(rcCrop, bmpScreen.PixelFormat);
                        var captchaTimout = 20;


                        imgCap.Save(imgFileName, ImageFormat.Png);

                        Out("Decoding CAPTCHA...");
                        captcha = client.Decode(imgFileName, 30);

                        if (captcha.Solved && captcha.Correct)
                        {
                            Console.WriteLine("{0} - Saldo: {1}", captcha.Text, client.Balance / 100);
                        }
                        else
                        {
                            if (client.Balance == 0)
                                throw new Exception(TipoErro.SemSaldo);

                            return false;
                        }

                        elementTxt.Clear();
                        elementTxt.SendKeys(captcha.Text);
                        click.Click();
                        Thread.Sleep(5000);

                        //-----Testa Erro no Captcha-----//

                        if (IsElementPresent(By.XPath("/html/body/div/div[2]/form/div[3]/div/div/div/div[1]/div/div/ul/li/span")))
                        {
                            var captchaError = driver.FindElement(By.XPath("/html/body/div/div[2]/form/div[3]/div/div/div/div[1]/div/div/ul/li/span")).Text;
                            if (captchaError.ToLowerInvariant() == "Os caracteres digitados não correspondem à imagem. Por favor tente novamente".ToLowerInvariant())
                                Out("CAPTCHA ERRADA. Reportando...");
                            client.Report(captcha);
                            Console.WriteLine("Reportado");
                            Log.LogThis(captchaError);
                            captchaOk = false;
                            FecharAvisoLogin();
                            Thread.Sleep(1000);
                        }
                        else if (IsElementPresent(By.XPath("/html/body/div/center/div/div/div/div/span/p")))
                        {
                            var msg = driver.FindElement(By.XPath("/html/body/div/center/div/div/div/div/span/p")).Text;
                            if (msg.ToLowerInvariant() == "Grade horária fechada.".ToLowerInvariant())
                            {
                                captchaOk = true;
                                Console.WriteLine("Grade horária fechada.");
                                throw new Exception(TipoErro.ForaDoHorario);
                            }
                        }
                        else if (IsElementPresent(By.XPath("/html/body/div/div[2]/form/div[3]/div/div/div/div[1]/div/div/ul/li/span")))
                        {
                            var msg = driver.FindElement(By.XPath("/html/body/div/div[2]/form/div[3]/div/div/div/div[1]/div/div/ul/li/span")).Text;
                            if (msg.ToLowerInvariant() == "Dados inválidos.".ToLowerInvariant())
                            {
                                Console.WriteLine("Dados inválidos.");
                                FecharAvisoLogin();
                                SairDoRobo(msg = "Login ou Senha Inválido, verifique para não bloquear usuário");
                            }
                        }
                        else if (IsElementPresent(By.XPath("/html/body/div/div/form/div[2]/div/div[1]")))
                        {
                            OutLine("Captcha Decifrado com Sucesso");
                            captchaOk = true;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                if (e.Message == TipoErro.UsuarioBloqueado ||
                    e.Message == TipoErro.SenhaInvalida ||
                    e.Message == TipoErro.DadosInvalidos ||
                    e.Message == TipoErro.ForaDoHorario ||
                    e.Message == TipoErro.SemSaldo)

                    SairDoRobo(e.Message);
                else
                {
                    Log.LogThis(TipoErro.ErroInesperado);
                    throw new Exception(TipoErro.ErroInesperado);
                }
            }
            return true;
        }
        static private void Login(string usuario, string senha)
        {
            IWebElement usr = driver.FindElement(By.Id("username"));
            usr.Clear();
            usr.Click();
            usr.SendKeys(usuario);

            IWebElement pws = driver.FindElement(By.Id("password"));
            pws.Clear();
            pws.Click();
            pws.SendKeys(senha);
            Thread.Sleep(1000);
        }
        private static void AbrirPagina()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.saopauloconsig.org.br/home");
            driver.FindElement(By.XPath("//*[@id='guias']/div[2]/span")).Click();
            Thread.Sleep(3000);
        }
        static private void SairDoRobo(string msg = "")
        {
            if (msg != "")
                OutLine(msg);

            Log.LogThis(msg, true);
            Log.LogThis("### Encerrando Robô ###");
            OutLine("##### Tecle algo para encerrar o robô ####");
            Console.ReadKey();
            finalizador = true;
            driver.Quit();
            Environment.Exit(-1);
        }
        static private void OutLine(string msg)
        {
            Console.WriteLine($"[{DateTime.Now}] {msg}");
        }
        static private void Out(string msg)
        {
            Console.Write($"[{DateTime.Now}] {msg}");
        }
        static private bool IsElementPresent(By by)
        {
            try
            {
                driver.FindElement(by);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
    }
}