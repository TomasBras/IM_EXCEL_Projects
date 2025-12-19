

using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
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

			Console.WriteLine("Aguardando mensagens do IM...");
			await Task.Delay(-1);
		}

		private static void InicializarExcel()
		{
			try
			{
				_excelApp = new Application();
				_excelApp.Visible = true;

				excelPathBase = @"E:\ExcelGestures\IM_Excel\ETP3.xlsx";
				excelPathFinal = @"E:\ExcelGestures\IM_Excel\Relatorio_Final.xlsx";
				//excelPathBase = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\ETP3.xlsx";
				//excelPathFinal = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\Relatorio_Final.xlsx";
				
				//excelPathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\ETP.xlsx";
				//excelPathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\Relatorio_Final.xlsx";

				if (!File.Exists(excelPathBase))
				{
					Console.WriteLine("Ficheiro Excel não encontrado!");
					return;
				}

				_workbook = _excelApp.Workbooks.Open(excelPathBase);
				_sheet = _workbook.Sheets[1];

				ExcelController.SetExcel(_excelApp, _workbook, _sheet);

				Console.WriteLine("Excel inicializado com sucesso!");
			}
			catch (Exception ex)
			{
				Console.WriteLine("Erro ao abrir Excel: " + ex.Message);
			}
		}


		private static void ProcessMessage(string message)
		{
			if (message == "OK" || message == "RENEW") return;

			try
			{
				var doc = XDocument.Parse(message);
				var com = doc.Descendants("command").FirstOrDefault()?.Value;
				if (string.IsNullOrWhiteSpace(com)) return;

				Console.WriteLine("Command recebido: " + com);

				dynamic json = JsonConvert.DeserializeObject(com);

				if (json.recognized != null && json.recognized.Count > 1)
				{
					string gesture = json.recognized[1].ToString().ToLower();
					Console.WriteLine($"GESTO RECEBIDO: {gesture}");

					string resposta = ExecutarGesto(gesture);

					if (!string.IsNullOrEmpty(resposta))
						SendMessage(messageMMI(resposta));

					return;
				}

				Console.WriteLine("⚠ Mensagem ignorada (não contém gestos).");
			}
			catch (Exception ex)
			{
				Console.WriteLine("Erro ao processar mensagem: " + ex.Message);
			}
		}


		private static string ExecutarGesto(string gesture)
		{
			if (string.IsNullOrWhiteSpace(gesture))
				return null;

			gesture = gesture
				.ToLower()
				.Replace(".a", "")
				.Replace("_", "")
				.Replace("-", "")
				.Trim();

			Console.WriteLine("Gesto normalizado: " + gesture);

			switch (gesture)
			{

				case "calculateaverage":
					return ExcelController.CalcularMediaTurma();

				case "insertcolumn":
					return ExcelController.InserirSituacao();

				case "studentsapproved":
					return ExcelController.DestacarApenasAprovados();

				case "studentsfailed":
					return ExcelController.DestacarApenasReprovados();

				case "undolastaction":
					if (_excelApp?.ActiveCell != null)
					{
						_excelApp.ActiveCell.ClearContents();
						return "Valor da célula apagado.";
					}
					return "Nenhuma célula ativa para apagar.";

				case "closeexcel":
					try
					{
						_excelApp.DisplayAlerts = false;   

						_workbook?.SaveAs(excelPathFinal);
						_workbook?.Close(false);           
						_excelApp?.Quit();

						return "Excel fechado.";
					}
					finally
					{
						if (_excelApp != null)
							_excelApp.DisplayAlerts = true; 
					}


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


				case "zoomin":
					_excelApp.ActiveWindow.Zoom += 10;
					return "Zoom aumentado.";

				case "zoomout":
					_excelApp.ActiveWindow.Zoom -= 10;
					return "Zoom reduzido.";

				default:
					Console.WriteLine("⚠ Gesto não reconhecido: " + gesture);
					return null;
			}
		}

		private static void SendMessage(string message)
		{
			_client.Send(message);
			Console.WriteLine("📤 Enviada resposta MMI.");
		}

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