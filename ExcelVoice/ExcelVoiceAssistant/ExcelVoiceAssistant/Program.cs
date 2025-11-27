using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WebSocketSharp;

namespace ExcelVoiceAssistant
{
    class Program
    {
        private static WebSocket _client;
        private static Application _excelApp;
        private static Workbook _workbook;
        private static Worksheet _sheet;

        private static string excelPathBase;
        private static string excelPathFinal;

        static async Task Main(string[] args)
        {
            string host = "localhost";
            string path = "/IM/USER1/APP";
            string uri = $"wss://{host}:8005{path}"; 

            Console.WriteLine(" Conectando ao IM via WebSocket...");

            _client = new WebSocket(uri);

            _client.SslConfiguration.EnabledSslProtocols = System.Security.Authentication.SslProtocols.Tls12;
            _client.SslConfiguration.ServerCertificateValidationCallback = (sender, cert, chain, errors) =>
            {
                Console.WriteLine($" Ignorando certificado inválido: {errors}");
                return true;
            };

            _client.OnOpen += (s, e) => Console.WriteLine(" Conectado ao IM!");
            _client.OnMessage += (s, e) => ProcessMessage(e.Data);
            _client.OnError += (s, e) => Console.WriteLine(" Erro WebSocket: " + e.Message);
            _client.OnClose += (s, e) => Console.WriteLine(" Conexão encerrada.");

            try
            {
                _client.Connect();
            }
            catch (Exception ex)
            {
                Console.WriteLine(" Falha ao conectar: " + ex.Message);
                return;
            }

            InicializarExcel();

            Console.WriteLine("💬 Aguardando mensagens do IM...");
            await Task.Delay(-1); 
        }

        // =========================================================
        // INICIALIZAR EXCEL
        // =========================================================
        private static void InicializarExcel()
        {
            try
            {
                _excelApp = new Application();
                _excelApp.Visible = true;

                excelPathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_Excel\ETP3.xlsx";
                excelPathFinal = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_ExcelS\Relatorio_Final.xlsx";
                // excelPathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_NODEPENDENCIES\ETP.xlsx";
                //excelPathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_NODEPENDENCIES\Relatorio_Final.xlsx";

                if (!File.Exists(excelPathBase))
                {
                    Console.WriteLine("❌ Ficheiro Excel não encontrado!");
                    return;
                }

                _workbook = _excelApp.Workbooks.Open(excelPathBase);
                _sheet = _workbook.Sheets[1];

                // 👉 Ligar o ExcelController ao Excel já aberto
                ExcelController.SetExcel(_excelApp, _workbook, _sheet);

                Console.WriteLine("✅ Excel inicializado com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao abrir Excel: " + ex.Message);
            }
        }



        // =========================================================
        // PROCESSAR MENSAGENS MMI
        // =========================================================
        private static void ProcessMessage(string message)
        {
            if (message == "OK" || message == "RENEW") return;

            try
            {
                var doc = XDocument.Parse(message);
                var com = doc.Descendants("command").FirstOrDefault()?.Value;
                if (string.IsNullOrEmpty(com)) return;

                dynamic json = JsonConvert.DeserializeObject(com);
                if (json.nlu == null) return;

                string intent = json.nlu.intent;
                Console.WriteLine($"🎯 Intent recebido: {intent}");

                string resposta = ExecutarComando(intent, json);
                SendMessage(messageMMI(resposta));
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao processar mensagem: " + ex.Message);
                SendMessage(messageMMI("Ocorreu um erro ao processar o comando."));
            }
        }

        // =========================================================
        // EXECUTAR COMANDOS
        // =========================================================
        private static string ExecutarComando(string intent, dynamic json)
        {
            try
            {
                switch (intent)
                {
                    case "calcular_media":
                        return ExcelController.CalcularMedia(json);
                    case "destacar_aprovados_reprovados":
                        return ExcelController.DestacarAprovados();

                    case "inserir_colunas":
                        return ExcelController.InserirSituacao();

                    case "identificar_melhoria":
                        return ExcelController.IdentificarMelhoria();

                    case "operacoes_matematicas":
                        return ExcelController.OperacoesMatematicas(json);

                    case "gerar_grafico_turma":
                        return ExcelController.GerarGraficoTurma(json);

                    case "gerar_grafico_barras_aluno":
                        return ExcelController.GerarGraficoBarras(json);

                    case "apagar_grafico":
                        return ExcelController.ApagarGrafico(json);

                    case "apagar_todos_graficos":
                        return ExcelController.ApagarTodosGraficos();

                    case "guardar_ficheiro":
                        return ExcelController.GuardarRelatorio();

                    case "atualizar_notas":
                        return ExcelController.AtualizarNotas(json);

                    default:
                        return "Comando não reconhecido.";
                }
            }
            catch (Exception ex)
            {
                return "❌ Erro ao executar comando: " + ex.Message;
            }
        }

        // =========================================================
        // ENVIAR MENSAGEM MMI
        // =========================================================
        private static void SendMessage(string message)
        {
            _client.Send(message);
            Console.WriteLine("📤 Enviada resposta MMI.");
        }

        // =========================================================
        // FORMATA MENSAGEM MMI PARA TTS
        // =========================================================
        public static string messageMMI(string msg)
        {
            return "<mmi:mmi xmlns:mmi=\"http://www.w3.org/2008/04/mmi-arch\" mmi:version=\"1.0\">" +
                    "<mmi:startRequest mmi:context=\"ctx-1\" mmi:requestId=\"text-1\" mmi:source=\"APPSPEECH\" mmi:target=\"IM\">" +
                        "<mmi:data>" + msg +
                            "<emma:emma xmlns:emma=\"http://www.w3.org/2003/04/emma\" emma:version=\"1.0\">" +
                                "<emma:interpretation emma:confidence=\"1\" emma:id=\"text-\" emma:medium=\"text\" emma:mode=\"command\" emma:start=\"0\">" +
                                    "<command>\"&lt;speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='pt-PT'&gt;&lt;p&gt;" + msg + "&lt;/p&gt;&lt;/speak&gt;\"</command>" +
                                "</emma:interpretation>" +
                            "</emma:emma>" +
                        "</mmi:data>" +
                    "</mmi:startRequest>" +
                "</mmi:mmi>";
        }
    }
}
