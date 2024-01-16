using NLog;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using static System.Net.WebRequestMethods;

namespace PER_DComp.Robo
{
    class Program
    {
        #region Variáveis Globais
         private static Logger logger = LogManager.GetCurrentClassLogger();
        private static IWebDriver driver = null;
        #endregion

        static void Main(string[] args)
        {
            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("pt-BR");

            Util.Log(logger, "Iniciando Robo PER/DComp v1.0.0 ...");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Iniciando Robo PER/DComp v1.0.0 ...");
            Console.ResetColor();


            var chromeOptions = new ChromeOptions();
            //chromeOptions.AddArgument("--disable-blink-features");
            //chromeOptions.AddArgument("--disable-blink-features=AutomationControlled");
            // options.AddUserProfilePreference("intl.accept_languages", "nl");
            //chromeOptions.AddArgument("no-sandbox");/html/body/div[1]/header/div/div[2]/div/div/div/nav/div[3]/ul/li/a
            //chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
            //chromeOptions.AddExcludedArguments(new List<string>() { "enable-automation" });

            if (driver == null) { driver = new ChromeDriver(ChromeDriverService.CreateDefaultService(), chromeOptions, TimeSpan.FromMinutes(3)); }

            driver.Manage().Timeouts().PageLoad.Add(System.TimeSpan.FromSeconds(30));

            var url = "http://localhost:60932/";
            url = "https://site-devtest.planicare.pt/";
            url = "https://site-staging.planicare.pt/";

            driver.Navigate().GoToUrl(url);
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);

            driver.Navigate().GoToUrl(url + "network");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);

            driver.Navigate().GoToUrl(url + "contact");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);

            driver.Navigate().GoToUrl(url + "about");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);

            driver.Navigate().GoToUrl(url + "soft");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);


            driver.Navigate().GoToUrl(url + "legal");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);


            driver.Navigate().GoToUrl(url + "easy50plus");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);


            driver.Navigate().GoToUrl(url + "easycare");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);


            driver.Navigate().GoToUrl(url + "protection");
            AguardarCarregamentoTela(driver);
            Thread.Sleep(100);


            Util.Log(logger, $"Drivers do ChromeDriver OpenQA.Selenium.Chrome");
            for (int i = 0; i < 35; i++)
            {
                driver.Navigate().GoToUrl(url);

                AguardarCarregamentoTela(driver);

                Thread.Sleep(1000);

                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Request nº " + i);
            }

            for (int i = 0; i < 100; i++)
            {
                driver.Navigate().GoToUrl(url);

                AguardarCarregamentoTela(driver);

                Thread.Sleep(1000);

                driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[2]/div/div/div/nav/div[3]/ul/li/a")).Click();

                Thread.Sleep(500);

                driver.FindElement(By.XPath("//*[@id=\"sim_ages-selectized\"]")).Click();

                Thread.Sleep(500);

                driver.FindElement(By.XPath("//*[@id=\"sim_ages-selectized\"]")).SendKeys("40");

                driver.FindElement(By.XPath("//*[@id=\"sim_ages-selectized\"]")).SendKeys(Keys.Return);

                Thread.Sleep(500);

                driver.FindElement(By.XPath("//*[@id=\"sim_name\"]")).SendKeys("Joe Banana");

                driver.FindElement(By.XPath("//*[@id=\"sim_name\"]")).SendKeys(Keys.Return);

                Thread.Sleep(500);

                driver.FindElement(By.XPath("//*[@id=\"sim_cellphone\"]")).SendKeys("921077473");

                driver.FindElement(By.XPath("//*[@id=\"sim_cellphone\"]")).SendKeys(Keys.Return);

                //
                ///

                Thread.Sleep(500);

                driver.FindElement(By.XPath("//*[@id=\"simulator_btn_sim\"]")).Click();

                Thread.Sleep(3000);

                if (driver.FindElements(By.XPath("//*[@id=\"simulator_page_2\"]/div[1]/div[1]/h2")).Count > 0)
                {
                    continue;
                }
                else
                {
                    Thread.Sleep(3000);
                    if (driver.FindElements(By.XPath("//*[@id=\"simulator_page_2\"]/div[1]/div[1]/h2")).Count > 0)
                    {
                        continue;
                    }
                    else
                    {
                        Thread.Sleep(3000);

                        if (driver.FindElements(By.XPath("//*[@id=\"simulator_page_2\"]/div[1]/div[1]/h2")).Count > 0)
                        {
                            continue;
                        }
                        else
                        {
                            if (driver.FindElements(By.XPath("//*[@id=\"simulator_page_1\"]/div[2]/div/div[2]/div/div/div[4]/div[1]")).Count > 0)
                            {
                                break;
                            }
                        }
                    }
                    

                }


            }


        }

        private static void AguardarCarregamentoTela(IWebDriver driver)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
        }


        private static decimal StrToDecimal(string valor)
        {
            valor = valor.Trim();

            CultureInfo _provider = new CultureInfo("pt-BR");
            decimal retorno = 0;
            try
            {
                retorno = decimal.Parse(valor, _provider);
                return retorno;
            }
            catch (Exception ex)
            {
                Util.Log(logger, $" Erro conversão decimal, valor string = {valor} - { ex.Message} - { ex.ToString()}");
                return 0;
            }
        }

        private class PlanilhaGuiaCompensacao
        {
            public int Id { get; set; }
            public string NomeArquivo { get; set; }
            public int TotalLinhas { get; set; }
            public List<DadosGuiaCompensacao> DadosCompensacoes { get; set; }
        }

        private class DadosGuiaCompensacao
        {
            public int Id { get; set; }

            /* Seção 01*/
            public string NovoDocumento_01 { get; set; }
            public bool DocumentoRetificador_01 { get; set; }
            public string TipoCredito_01 { get; set; }
            public string ApelidoDocumento_01 { get; set; }
            public string QualificacaoContribuinte_01 { get; set; }
            public string DetalhamentoCredito_01 { get; set; }
            public bool AlegacaoInconstitucional_01 { get; set; }

            /* Seção 02*/
            public string DetentorCredito_02 { get; set; }
            public string CnpjDetentor_02 { get; set; }
            public int AnoCompetencia_02 { get; set; }
            public string MesCompetencia_02 { get; set; }
            public bool RecolhimentoEfetuado_02 { get; set; }
            public string CodigoPagamento_02 { get; set; }

            /* Seção 03*/
            public decimal ValorInss_03 { get; set; }
            public decimal ValorOutrasEntidades_03 { get; set; }
            public decimal ValorAtmMultaJuros_03 { get; set; }
            public string DataArrecadacao_03 { get; set; }

            /* Seção 04*/
            public decimal ValorOriginal_04 { get; set; }
            public decimal SelicAcumulada_04 { get; set; }
            public decimal CreditoAtualizado_04 { get; set; }
            public string TipoDebito_04 { get; set; }

            /* Seção 05*/
            public string Categoria_05 { get; set; }
            public int AnoApuracao_05 { get; set; }
            public string MesApuracao_05 { get; set; }
            public string DataVencimento_05 { get; set; }
            public string CodigoReceita_05 { get; set; }
            public decimal ValorCompensar_05 { get; set; }

            /* Seção 06*/
            public string Cpf_06 { get; set; }


        }


    }

    
}



