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

namespace PER_DComp.Robo
{
    class Program
    {
        #region Variáveis Globais
        private static string _diretorioPadrao = $@"{Directory.GetCurrentDirectory()}\dados";
        private static string _diretorioPadraoLogs = $@"{Directory.GetCurrentDirectory()}\dados\logs";
        private static PlanilhaGuiaCompensacao _guiaCompensacao = new PlanilhaGuiaCompensacao();
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

            _guiaCompensacao.Id = 1;
            _guiaCompensacao.NomeArquivo = "Guia Compensação.xlsx";
            _guiaCompensacao.DadosCompensacoes = new List<DadosGuiaCompensacao>();


            Util.Log(logger, $"Lendo planilha dados ");

            if (LerPlanilha())
            {
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

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("POR FAVOR, EFETUE O LOGIN NO PORTAL eCAC e pressione ENTER");
                Console.ResetColor();

                Util.Log(logger, $"Aguardando usuário fazer login...");

                Console.ReadKey();

                Util.Log(logger, $"Abrindo página portal e-Cac");

                driver.Navigate().GoToUrl("https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx?id=10006&origem=pesquisa");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("A planilha de compensação não possui dados ou não pode ser lida.");
                Console.ResetColor();
                
                Util.Log(logger, $"A planilha de compensação não possui dados ou não pode ser lida.");
            }
            
            
        }

        private static void AguardarCarregamentoTela(IWebDriver driver)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(20)).Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
        }

        private static bool LerPlanilha()
        {
            try
            {
                FileInfo arquivo = new FileInfo($"{_diretorioPadrao}/{_guiaCompensacao.NomeArquivo}");
                
                using (ExcelPackage package = new ExcelPackage(arquivo))
                {
                    ExcelWorksheet ws;
                    //pega dados da aba "Dados" ou primeira aba
                    ws = package.Workbook.Worksheets[0];
                    Console.WriteLine("Iniciando leitura planilha guia de compensação...");
                    Util.Log(logger, $"Iniciando leitura planilha guia de compensação...");
                    var totalLinhas = 0;
                    for (int j = 4; j < ws.Cells.Rows; j++)
                    {
                        try
                        {
                            if (!String.IsNullOrEmpty(ws.Cells[j, 1].Value.ToString()))
                            {
                                totalLinhas++;
                            }
                        }
                        catch (Exception)
                        {

                            break;
                        }

                    }


                    /*
                     CAMPOS da Planilha Guia de Compensação

                        Seção 1: Identificar Documento                     
                        A: Novo Documento
                        B: Documento Retificafor?
                        C: Tipo de Crédito
                        D: Apelido documento
                        E: Qualificação do Contribuinte
                        F: Detalhamento do Crédito
                        G: Alegação de inconstitucionalidade?

                        Seção2: Identificação do Crédito
                        H: Detentor do crédito
                        I: CNPJ do Detentor
                        J: Ano da competência
                        K: Mês da competência
                        L: Recolhimento efetuado?
                        M: Código de pagamento      
                    
                        Seção3: Detalhamento GPS
                        N: Valor INSS
                        O: Valor Outras Entidades
                        P: Valor de ATM, Multa e Juros
                        Q: Data de Arrecadação

                        Seção4: Demonstrativo do Crédito
                        R: Valor Original do Crédito Inicial
                        S: Selic Acumulada
                        T: Crédito Atualizado
                        U: Tipo de Débito

                        Seção5: Informar e Ordenar Débito
                        V: Categoria
                        W: Ano de Apuração
                        X: Mês de Apuração
                        Y: Data de Vencimento
                        Z: Código da Receita
                        AA: Valor a Compensar

                        Seção6: Informar Dados Gerais
                        AB: CPF
                     */

                    if (totalLinhas > 0)
                    {
                        
                        Console.WriteLine($"Total de linhas encontrado = {totalLinhas}");
                        totalLinhas += 4; // compensa o início do cabeçalho 

                        var totalLinhasOK = 0;

                        for (int i = 4; i < totalLinhas; i++)
                        {
                            var dadosLinha = new DadosGuiaCompensacao();

                            try
                            {
                                dadosLinha.NovoDocumento_01 = ws.Cells[i, 1].Value.ToString();
                                dadosLinha.DocumentoRetificador_01 = ws.Cells[i, 2].Value.ToString().Trim() == "Não" ? false : true;
                                dadosLinha.TipoCredito_01 = ws.Cells[i, 3].Value.ToString();
                                dadosLinha.ApelidoDocumento_01 = ws.Cells[i, 4].Value.ToString();
                                dadosLinha.QualificacaoContribuinte_01 = ws.Cells[i, 5].Value.ToString();
                                dadosLinha.DetalhamentoCredito_01 = ws.Cells[i, 6].Value.ToString();
                                dadosLinha.AlegacaoInconstitucional_01 = ws.Cells[i, 7].Value.ToString().Trim() == "Não" ? false : true;

                                dadosLinha.DetentorCredito_02 = ws.Cells[i, 8].Value.ToString();
                                dadosLinha.CnpjDetentor_02 = ws.Cells[i, 9].Value.ToString();
                                dadosLinha.AnoCompetencia_02 = Convert.ToInt32(ws.Cells[i, 10].Value.ToString());
                                dadosLinha.MesCompetencia_02 = ws.Cells[i, 11].Value.ToString();
                                dadosLinha.RecolhimentoEfetuado_02 = ws.Cells[i, 12].Value.ToString().Trim() == "Não" ? false : true;
                                dadosLinha.CodigoPagamento_02 = ws.Cells[i, 13].Value.ToString();

                                dadosLinha.ValorInss_03 = ws.Cells[i, 14].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 14].Value.ToString().Trim());
                                dadosLinha.ValorOutrasEntidades_03 = ws.Cells[i, 15].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 15].Value.ToString().Trim());
                                dadosLinha.ValorAtmMultaJuros_03 = ws.Cells[i, 16].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 16].Value.ToString().Trim());
                                dadosLinha.DataArrecadacao_03 = ws.Cells[i, 17].Value.ToString();

                                dadosLinha.ValorOriginal_04 = ws.Cells[i, 18].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 18].Value.ToString().Trim());
                                dadosLinha.SelicAcumulada_04 = ws.Cells[i, 19].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 19].Value.ToString().Trim());
                                dadosLinha.CreditoAtualizado_04 = ws.Cells[i, 20].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 20].Value.ToString().Trim());
                                dadosLinha.TipoDebito_04 = ws.Cells[i, 21].Value.ToString();

                                dadosLinha.Categoria_05 = ws.Cells[i, 22].Value.ToString();
                                dadosLinha.AnoApuracao_05 = Convert.ToInt32(ws.Cells[i, 23].Value.ToString());
                                dadosLinha.MesApuracao_05 = ws.Cells[i, 24].Value.ToString();
                                dadosLinha.DataVencimento_05 = ws.Cells[i, 25].Value.ToString();
                                dadosLinha.CodigoReceita_05 = ws.Cells[i, 26].Value.ToString();
                                dadosLinha.ValorCompensar_05 = ws.Cells[i, 27].Value.ToString().Trim() == "-" ? 0m : StrToDecimal(ws.Cells[i, 27].Value.ToString().Trim());

                                dadosLinha.Cpf_06 = ws.Cells[i, 28].Value.ToString();

                                _guiaCompensacao.DadosCompensacoes.Add(dadosLinha);

                                totalLinhasOK++;

                            }
                            catch (Exception)
                            {
                                continue;
                                throw;
                            }
                                                       


                        }

                        if(totalLinhasOK > 0)
                        {
                            _guiaCompensacao.TotalLinhas = totalLinhasOK;
                            Console.WriteLine($"Total de linhas lidas OK = {totalLinhasOK}");
                            return true;
                        }

                        
                    }

                    return false;

                }
            }
            catch (Exception ex)
            {
                Util.Log(logger, $" Erro indeterminado - { ex.Message} - { ex.ToString()}");
                return false;
            }

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



