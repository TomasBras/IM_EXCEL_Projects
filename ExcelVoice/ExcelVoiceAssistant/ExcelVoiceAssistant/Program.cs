using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using WebSocketSharp;


namespace ExcelVoiceAssistant
{
    class Program
    {
        private static WebSocket _client;
        private static ExcelApp _excelApp;
        private static ExcelWorkbook _workbook;
        private static ExcelWorksheet _sheet;

        private static string excelPathBase;
        private static string excelPathFinal;

        private static string _acaoPendente = null;

        private static Form _confirmacaoForm;
        private static Thread _popupThread;

        static async Task Main(string[] args)
        {
            string host = "localhost";
            string path = "/IM/USER1/APP";
            string uri = $"wss://{host}:8005{path}";

            Console.WriteLine(" Conectando ao IM via WebSocket...");
            Console.WriteLine("🔥🔥🔥 VERSÃO NOVA DO CÓDIGO 🔥🔥🔥");


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

        private static string PedirConfirmacao(string acao, string mensagem)
        {
            _acaoPendente = acao;
            Console.WriteLine($" Ação pendente: {_acaoPendente}");

            MostrarConfirmacaoPopup(mensagem);

            return mensagem;
        }

        private static void MostrarConfirmacaoPopup(string mensagem)
        {
            // evita múltiplos popups
            if (_confirmacaoForm != null)
                return;

            _popupThread = new Thread(() =>
            {
                _confirmacaoForm = new Form
                {
                    Text = "Confirmação",
                    Width = 420,
                    Height = 160,
                    StartPosition = FormStartPosition.CenterScreen,
                    TopMost = true
                };

                var label = new Label
                {
                    Text = mensagem + "\n\nDiga 'Confirmar' ou 'Cancelar'.",
                    Dock = DockStyle.Fill,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter
                };

                _confirmacaoForm.Controls.Add(label);
                Application.Run(_confirmacaoForm);
            });

            _popupThread.SetApartmentState(ApartmentState.STA);
            _popupThread.IsBackground = true;
            _popupThread.Start();
        }

        private static void FecharConfirmacaoPopup()
        {
            if (_confirmacaoForm == null)
                return;

            try
            {
                _confirmacaoForm.Invoke(new Action(() =>
                {
                    _confirmacaoForm.Close();
                    _confirmacaoForm = null;
                }));
            }
            catch
            {
                // ignore (thread já terminou)
            }
        }



        // =========================================================
        // INICIALIZAR EXCEL
        // =========================================================
        private static void InicializarExcel()
        {
            try
            {
                _excelApp = new ExcelApp();
                _excelApp.Visible = true;

                //excelPathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_Excel\ETP3.xlsx";
                //excelPathFinal = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_ExcelS\Relatorio_Final.xlsx";
                excelPathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\ETP.xlsx";
                excelPathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\Relatorio_Final.xlsx";

                if (!File.Exists(excelPathBase))
                {
                    Console.WriteLine("❌ Ficheiro Excel não encontrado!");
                    return;
                }

                _workbook = _excelApp.Workbooks.Open(excelPathBase);
                _sheet = _workbook.Sheets[1];

                ExcelController.SetExcel(_excelApp, _workbook, _sheet);

                Console.WriteLine("✅ Excel inicializado com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao abrir Excel: " + ex.Message);
            }
        }

        

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

                string intent = json.nlu.intent.ToString();
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


        private static string ExecutarAcaoConfirmada()
        {
            if (_acaoPendente == null)
                return "Não há nenhuma ação para confirmar.";

            string acao = _acaoPendente;
            _acaoPendente = null;

            switch (acao)
            {
                case "undolastaction":
                    _excelApp.ActiveCell.ClearContents();
                    return "Conteúdo da célula apagado.";

                case "closeexcel":
                    try
                    {
                        _excelApp.DisplayAlerts = false;
                        _workbook?.SaveAs(excelPathFinal);
                        _workbook?.Close(false);
                        _excelApp?.Quit();
                        return "Excel guardado e fechado.";
                    }
                    finally
                    {
                        if (_excelApp != null)
                            _excelApp.DisplayAlerts = true;
                    }

                case "apagar_todos_graficos":
                    return ExcelController.ApagarTodosGraficos();

                default:
                    return "Ação desconhecida.";
            }
        }


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

                    case "melhoria_real":
                        return ExcelController.MelhoriaReal();

                    case "melhoria_possivel":
                        return ExcelController.MelhoriaPossivel();

                    case "operacoes_matematicas":
                        return ExcelController.OperacoesMatematicas(json);

                    case "inserir_perguntas":
                        return ExcelController.InserirPerguntas(json);

                    case "gerar_grafico_turma":
                        return ExcelController.GerarGraficoTurma(json);

                    case "gerar_grafico_barras_aluno":
                        return ExcelController.GerarGraficoBarras(json);

                    case "gerar_grafico_perguntas_t2":
                        return ExcelController.GerarGraficoPerguntasT2();

                    case "apagar_todos_graficos":
                        return PedirConfirmacao(
                            "apagar_todos_graficos",
                            "Tem a certeza que quer apagar TODOS os gráficos da folha?"
                        );


                    case "guardar_ficheiro":
                        return ExcelController.GuardarRelatorio();

                    case "atualizar_notas":
                        return ExcelController.AtualizarNotas(json);
                    case "undolastaction":
                        if (_excelApp?.ActiveCell == null)
                            return "Nenhuma célula ativa.";

                        return PedirConfirmacao(
                            "undolastaction",
                            "Tem a certeza que quer apagar o conteúdo da célula?"
                        );


                    case "closeexcel":
                        return PedirConfirmacao(
                            "closeexcel",
                            "Tem a certeza que quer guardar e fechar o Excel?"
                        );

                    case "confirmar":
                        if (_acaoPendente == null)
                            return "Não há nenhuma ação pendente para confirmar.";

                        FecharConfirmacaoPopup();
                        return ExecutarAcaoConfirmada();

                    case "cancelar":
                        _acaoPendente = null;
                        FecharConfirmacaoPopup();
                        return "Ação cancelada.";

                    case "swipeleft":
                        _excelApp.ActiveCell?.Offset[0, -1].Select();
                        return "Mover para a esquerda.";

                    case "swiperight":
                        _excelApp.ActiveCell?.Offset[0, 1].Select();
                        return "Mover para a direita.";

                    case "swipeup":
                        _excelApp.ActiveCell?.Offset[-1, 0].Select();
                        return "Mover para cima.";

                    case "swipedown":
                        _excelApp.ActiveCell?.Offset[1, 0].Select();
                        return "Mover para baixo.";
                    case "criar_pivot_table":
                        return ExcelController.CriarPivotTable(json);
                    case "greet":
                        return "Olá! Estou pronto para ajudar no Excel.";

                    case "ask_how_are_you":
                        return "Estou ótimo e pronto para trabalhar com dados!";

                    case "respond_how_am_i":
                        return "Ainda bem! O que queres fazer a seguir?";

                    case "helper":
                        return ExcelController.Helper();

                    case "fallback":
                        return "Não percebi o comando. Podes repetir ou pedir ajuda?";

                    default:
                        Console.WriteLine($"Intent não reconhecida: {intent}");
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