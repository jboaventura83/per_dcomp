using NLog;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace PER_DComp.Robo
{
    class Program
    {
        #region Variáveis Globais
        private static string _diretorioPadrao = $@"{Directory.GetCurrentDirectory()}\dados";
        private static string _diretorioPadraoLogs = $@"{Directory.GetCurrentDirectory()}\dados\logs";

        private static Logger logger = LogManager.GetCurrentClassLogger();
        private static IWebDriver driver = null;
        #endregion

        static void Main(string[] args)
        {
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("pt-BR");

            Util.Log(logger, "Iniciando Robo PER/DComp v1.0.0 ...");

            // cria pasta dados                
            if (!Directory.Exists(_diretorioPadrao))
            {
                Util.Log(logger, $"Criando Diretorio '{_diretorioPadrao}'...");
                Directory.CreateDirectory(_diretorioPadrao);
            }

            // cria pasta logs                
            if (!Directory.Exists(_diretorioPadraoLogs))
            {
                Util.Log(logger, $"Criando Diretorio '{_diretorioPadraoLogs}'...");
                Directory.CreateDirectory(_diretorioPadraoLogs);
            }

            Util.Log(logger, $"Lendo planilha dados ");


            Util.Log(logger, $"Leu a planilha de dados");


            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArgument("--disable-blink-features");
            chromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
            // options.AddUserProfilePreference("intl.accept_languages", "nl");
            chromeOptions.AddArgument("no-sandbox");
            chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            chromeOptions.AddExcludedArguments(new List<string>() { "enable-automation" });

            if (driver == null) { driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), chromeOptions, TimeSpan.FromMinutes(3)); }

            driver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(30));

            Util.Log(logger, $"Drivers do ChromeDriver OpenQA.Selenium.Chrome");

            Util.Log(logger, $"Abrindo página portal e-Cac");

            driver.Navigate().GoToUrl("https://cav.receita.fazenda.gov.br/autenticacao/login");

            AguardarCarregamentoTela(driver);

            Console.WriteLine("POR FAVOR, EFETUE O LOGIN NO PORTAL eCAC e pressione ENTER");

            Util.Log(logger, $"Aguardando usuário fazer login...");

            Console.ReadKey();

            Util.Log(logger, $"Abrindo página portal e-Cac");

            driver.Navigate().GoToUrl("https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx?id=10006&origem=pesquisa");
        }

        private static void AguardarCarregamentoTela(IWebDriver driver)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
        }
    }
}
