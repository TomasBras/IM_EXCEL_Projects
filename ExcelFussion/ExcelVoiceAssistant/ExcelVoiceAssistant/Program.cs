using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Collections.Generic;
using WebSocketSharp;

namespace ExcelVoiceAssistant
{
	class Program
	{
		private static WebSocket _client;
		private static Excel.Application _excelApp;
		private static Excel.Workbook _workbook;
		private static Excel.Worksheet _sheet;
		private static volatile bool _excelReady;

		// Cooldown defensivo para gestos (evita execução repetida em loop)
		private static readonly object _gestureCooldownLock = new object();
		private static readonly System.Collections.Generic.Dictionary<string, DateTime> _gestureLastUtc =
			new System.Collections.Generic.Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
		private const int GestureCooldownMs = 1500;

		private static string excelPathBase;
		private static string excelPathFinal;

		private static readonly UiDispatcher _ui = new UiDispatcher();
		private static readonly object _undoLock = new object();
		private static bool _undoAwaitingConfirm;
		private static DateTime _undoConfirmExpiresUtc;
		private static UndoConfirmForm _undoConfirmForm;
		private static string _pendingConfirmKey;
		private static Func<string> _pendingConfirmAction;
		private const int UndoConfirmWindowMs = 15000;

		[STAThread]
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

			// Excel COM must run on an STA thread; WebSocket callbacks run on background threads.
			// We use the existing UI STA dispatcher thread as the single-threaded COM home.
			_ui.Invoke(() =>
			{
				// Avoid RPC_E_SERVERCALL_RETRYLATER when Excel is busy (thread-affine).
				OleMessageFilter.Register();
				InicializarExcel();
			});

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
				_excelReady = false;
				_excelApp = new Excel.Application();
				_excelApp.Visible = true;

				excelPathBase = @"E:\ExcelFussion\IM_Excel\ETP3.xlsx";
				excelPathFinal = @"E:\ExcelFussion\IM_Excel\Relatorio_Final.xlsx";
				//excelPathBase = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\ETP3.xlsx";
				//excelPathFinal = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\Relatorio_Final.xlsx";

				//excelPathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\ETP.xlsx";
				//excelPathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\Relatorio_Final.xlsx";

				if (!File.Exists(excelPathBase))
				{
					Console.WriteLine("❌ Ficheiro Excel não encontrado!");
					return;
				}

				_workbook = _excelApp.Workbooks.Open(excelPathBase);
				_sheet = _workbook.Sheets[1];

				try
				{
					_workbook.Activate();
					_sheet.Activate();
					// Ensure the window is ready before selecting A1.
					_excelApp.Goto(_sheet.Range["A1"], true);
					((Excel.Range)_sheet.Range["A1"]).Select();
				}
				catch (Exception ex)
				{
					Console.WriteLine("⚠ Não foi possível selecionar A1: " + ex.Message);
				}

				ExcelController.SetExcel(_excelApp, _workbook, _sheet);
				_excelReady = true;

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
				var commands = doc.Descendants("command")
					.Select(x => x?.Value)
					.Where(x => !string.IsNullOrWhiteSpace(x))
					.ToList();
				if (commands.Count == 0) return;

				Newtonsoft.Json.Linq.JObject fusionJson = null;
				Newtonsoft.Json.Linq.JObject speechJson = null;
				Newtonsoft.Json.Linq.JObject gesturesJson = null;

				List<string> fusionRecognized = null;
				List<string> speechRecognized = null;
				List<string> gesturesRecognized = null;

				foreach (var cmd in commands)
				{
					Newtonsoft.Json.Linq.JObject cmdJson;
					try
					{
						cmdJson = JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(cmd);
					}
					catch
					{
						continue;
					}

					var tokens = GetRecognizedTokens(cmdJson);
					string modality = tokens.Count > 0 ? tokens[0] : null;
					if (!IsModalityToken(modality) && tokens.Count > 1 && IsModalityToken(tokens[1]))
						modality = tokens[1];

					if (!IsModalityToken(modality))
						continue;

					if (fusionJson == null && modality.Equals("FUSION", StringComparison.OrdinalIgnoreCase))
					{
						fusionJson = cmdJson;
						fusionRecognized = tokens;
						continue;
					}
					if (speechJson == null && modality.Equals("SPEECH", StringComparison.OrdinalIgnoreCase))
					{
						speechJson = cmdJson;
						speechRecognized = tokens;
						continue;
					}
					if (gesturesJson == null &&
						(modality.Equals("GESTURES", StringComparison.OrdinalIgnoreCase) || modality.Equals("GESTURE", StringComparison.OrdinalIgnoreCase)))
					{
						gesturesJson = cmdJson;
						gesturesRecognized = tokens;
						continue;
					}
				}

				// Prefer running the FUSION command, but if the same EMMA group also
				// includes a SPEECH command with NLU payload, attach it so parameterized
				// commands (e.g., aluno_numero) are honored.
				if (fusionJson != null)
				{
					if (speechJson != null)
					{
						try
						{
							if (fusionJson["nlu"] == null && speechJson["nlu"] != null)
								fusionJson["nlu"] = speechJson["nlu"];
							if (fusionJson["text"] == null && speechJson["text"] != null)
								fusionJson["text"] = speechJson["text"];
						}
						catch
						{
							// best-effort
						}
					}

					var recognizedTokens = fusionRecognized ?? GetRecognizedTokens(fusionJson);
					Console.WriteLine("?? Command recebido (FUSION): " + fusionJson.ToString(Formatting.None));

					var fusionTokens = new List<string>();
					for (int i = 1; i < recognizedTokens.Count; i++)
						fusionTokens.Add(recognizedTokens[i]);
					HandleFusionMessage(fusionTokens, fusionJson);
					return;
				}

				if (speechJson != null)
				{
					var recognizedTokens = speechRecognized ?? GetRecognizedTokens(speechJson);
					Console.WriteLine("?? Command recebido (SPEECH): " + speechJson.ToString(Formatting.None));
					string intent = ExtractSpeechIntent(recognizedTokens, speechJson);
					HandleSpeechMessage(intent, speechJson);
					return;
				}

				if (gesturesJson != null)
				{
					var recognizedTokens = gesturesRecognized ?? GetRecognizedTokens(gesturesJson);
					Console.WriteLine("?? Command recebido (GESTURES): " + gesturesJson.ToString(Formatting.None));
					string gesture = recognizedTokens.Count > 1 ? recognizedTokens[1] : recognizedTokens.FirstOrDefault();
					HandleGestureMessage(gesture);
					return;
				}

				// Absolute fallback: use the first command node.
				var com = commands[0];
				Console.WriteLine("?? Command recebido: " + com);
				dynamic json = JsonConvert.DeserializeObject(com);
				var fallbackTokens = GetRecognizedTokens(json);
				if (fallbackTokens.Count > 0)
					HandleGestureMessage(fallbackTokens[fallbackTokens.Count - 1]);
				else
					Console.WriteLine("? Mensagem ignorada (n?o cont?m modalidade reconhecida).");
			}
			catch (Exception ex)
			{
				Console.WriteLine("? Erro ao processar mensagem: " + ex.Message);
			}
		}

		private static readonly HashSet<string> _speechNonIntentTokens =
			new HashSet<string>(StringComparer.OrdinalIgnoreCase)
			{
				"SPEECH",
				"SPEECHIN",
				"APP",
				"USER",
				"ASR",
				"NLU"
			};

		private static string ExtractSpeechIntent(List<string> recognizedTokens, dynamic json)
		{
			try
			{
				var nluIntent = json?.nlu?.intent?.ToString();
				if (!string.IsNullOrWhiteSpace(nluIntent))
					return nluIntent;
			}
			catch
			{
				// best-effort
			}

			if (recognizedTokens != null)
			{
				foreach (var token in recognizedTokens)
				{
					if (string.IsNullOrWhiteSpace(token))
						continue;
					if (_speechNonIntentTokens.Contains(token))
						continue;
					return token;
				}
			}

			return null;
		}

		private static List<string> GetRecognizedTokens(dynamic json)
		{
			var tokens = new List<string>();
			try
			{
				if (json?.recognized != null)
				{
					foreach (var item in json.recognized)
					{
						if (item != null)
							tokens.Add(item.ToString());
					}
				}
			}
			catch
			{
				// best-effort
			}

			return tokens;
		}

		private static bool IsModalityToken(string token)
		{
			if (string.IsNullOrWhiteSpace(token)) return false;
			return token.Equals("GESTURES", StringComparison.OrdinalIgnoreCase) ||
			       token.Equals("GESTURE", StringComparison.OrdinalIgnoreCase) ||
			       token.Equals("SPEECH", StringComparison.OrdinalIgnoreCase) ||
			       token.Equals("FUSION", StringComparison.OrdinalIgnoreCase);
		}

		private static void HandleGestureMessage(string gesture)
		{
			if (string.IsNullOrWhiteSpace(gesture))
			{
				Console.WriteLine("? Gesto vazio.");
				return;
			}

			gesture = gesture.ToLower();
			Console.WriteLine($"?? GESTO RECEBIDO: {gesture}");

			// Cooldown por gesto
			lock (_gestureCooldownLock)
			{
				var now = DateTime.UtcNow;
				if (_gestureLastUtc.TryGetValue(gesture, out var last) && (now - last).TotalMilliseconds < GestureCooldownMs)
				{
					Console.WriteLine($"?? Gesto ignorado (cooldown): {gesture}");
					return;
				}
				_gestureLastUtc[gesture] = now;
			}

			string resposta;
			try
			{
				// Ensure all Excel COM operations execute on the STA UI thread.
				resposta = _ui.Invoke(() => ExecutarGesto(gesture));
			}
			catch (Exception ex)
			{
				Console.WriteLine("? Erro ao executar gesto no thread STA: " + ex.Message);
				resposta = "O Excel est? ocupado. Tente novamente.";
			}

			if (!string.IsNullOrEmpty(resposta))
				SendMessage(messageMMI(resposta));
		}

		private static void HandleSpeechMessage(string recognizedIntent, dynamic json)
		{
			string intent = recognizedIntent;
			string nluIntent = null;

			try
			{
				nluIntent = json?.nlu?.intent?.ToString();
				if (!string.IsNullOrWhiteSpace(nluIntent))
					intent = nluIntent;
				else if (!string.IsNullOrWhiteSpace(intent) && _speechNonIntentTokens.Contains(intent))
					intent = null;
			}
			catch
			{
				// best-effort
			}

			Console.WriteLine($"🎙️ VOZ RECEBIDA (recognized): {recognizedIntent} | (nlu): {nluIntent} | (chosen): {intent}");

			if (string.IsNullOrWhiteSpace(intent))
			{
				SendMessage(messageMMI("Não percebi o comando."));
				return;
			}

			string resposta;
			try
			{
				resposta = _ui.Invoke(() => ExecutarFala(intent, json));
			}
			catch (Exception ex)
			{
				Console.WriteLine("Erro STA: " + ex.Message);
				resposta = "O Excel está ocupado.";
			}

			if (!string.IsNullOrEmpty(resposta))
				SendMessage(messageMMI(resposta));
		}

		private static void HandleFusionMessage(List<string> fusionTokens, dynamic json)
		{
			string fusionCommand = fusionTokens.FirstOrDefault();
			Console.WriteLine($"?? FUSION RECEBIDA: {fusionCommand}");

			string resposta;
			try
			{
				resposta = _ui.Invoke(() => ExecutarFusao(fusionCommand, fusionTokens, json));
			}
			catch (Exception ex)
			{
				Console.WriteLine("? Erro ao executar fus?o no thread STA: " + ex.Message);
				resposta = "O Excel est? ocupado. Tente novamente.";
			}

			if (!string.IsNullOrEmpty(resposta))
				SendMessage(messageMMI(resposta));
		}

		private static string ExecutarFala(string intent, dynamic json)
		{
			if (string.IsNullOrWhiteSpace(intent))
				return "Comando n?o reconhecido.";

			if (!_excelReady || _excelApp == null)
				return "Excel n?o est? inicializado.";

			switch (intent.ToLowerInvariant())
			{
				case "confirmar":
					return ConfirmarAcaoPendentePorVoz();
				case "cancelar":
					return CancelarAcaoPendentePorVoz();
				case "calcular_media":
					// Se houver seleção via gesto (handgrab), calcula só nessa seleção.
					if (ExcelController.HasSelection())
						return ExcelController.CalcularMediaSelecao();
					return ExcelController.CalcularMedia(json);
				case "destacar_aprovados_reprovados":
					return ExcelController.DestacarAprovados();
				case "inserir_colunas":
					return BeginPendingConfirmation(
						"Quer mesmo inserir a coluna?\r\nDiga confirmar ou cancelar.",
						"INSERIR_COLUNAS",
						() => ExcelController.InserirColuna(json));
				case "melhoria_real":
					return ExcelController.MelhoriaReal();
				case "melhoria_possivel":
					return ExcelController.MelhoriaPossivel();
				case "operacoes_matematicas":
					return ExcelController.OperacoesMatematicas(json);
				case "inserir_perguntas":
					return ExcelController.InserirPerguntas(json);
				case "gerar_grafico_turma":
					return ExcelController.HasSelection()
						? ExcelController.GerarGraficoTurmaSelecao(json)
						: ExcelController.GerarGraficoTurma(json);
				case "gerar_grafico_barras_aluno":
					return ExcelController.HasSelection()
						? ExcelController.GerarGraficoBarrasSelecao(json)
						: ExcelController.GerarGraficoBarras(json);
				case "gerar_grafico_perguntas_t2":
					return ExcelController.HasSelection()
						? ExcelController.GerarGraficoPerguntasT2Selecao()
						: ExcelController.GerarGraficoPerguntasT2();
				case "apagar_todos_graficos":
					return BeginPendingConfirmation(
						"Quer mesmo apagar todos os gráficos?\r\nDiga confirmar ou cancelar.",
						"APAGAR_TODOS_GRAFICOS",
						() => ExcelController.ApagarTodosGraficos());
				case "guardar_ficheiro":
					return BeginPendingConfirmation(
						"Quer mesmo guardar o ficheiro?\r\nDiga confirmar ou cancelar.",
						"GUARDAR_FICHEIRO",
						() => ExcelController.GuardarRelatorio());
				case "atualizar_notas":
					return BeginPendingConfirmation(
						"Quer mesmo atualizar as notas?\r\nDiga confirmar ou cancelar.",
						"ATUALIZAR_NOTAS",
						() => ExcelController.AtualizarNotas(json));
				case "criar_pivot_table":
					return ExcelController.HasSelection()
						? ExcelController.CriarPivotTableSelecao(json)
						: ExcelController.CriarPivotTable(json);
				case "helper":
					return ExcelController.Helper();
				default:
					Console.WriteLine("Intent não reconhecida: " + intent);
					return "Comando não reconhecido.";
			}
		}

		private static string ConfirmarAcaoPendentePorVoz()
		{
			var nowUtc = DateTime.UtcNow;
			Func<string> actionToRun = null;
			lock (_undoLock)
			{
				if (_undoAwaitingConfirm && nowUtc <= _undoConfirmExpiresUtc)
				{
					actionToRun = _pendingConfirmAction;
				}
				_pendingConfirmAction = null;
				_pendingConfirmKey = null;
				_undoAwaitingConfirm = false;
				_undoConfirmExpiresUtc = DateTime.MinValue;
			}

			CloseUndoConfirmFormAsync();

			if (actionToRun == null)
				return "Não há nenhuma ação pendente para confirmar.";

			try
			{
				return actionToRun();
			}
			catch (Exception ex)
			{
				Console.WriteLine("Erro ao confirmar ação pendente: " + ex.Message);
				return "Ocorreu um erro ao confirmar a ação.";
			}
		}

		private static string CancelarAcaoPendentePorVoz()
		{
			bool wasPending;
			lock (_undoLock)
			{
				wasPending = _undoAwaitingConfirm;
				_pendingConfirmAction = null;
				_pendingConfirmKey = null;
				_undoAwaitingConfirm = false;
				_undoConfirmExpiresUtc = DateTime.MinValue;
			}

			CloseUndoConfirmFormAsync();

			return wasPending ? "Ação cancelada." : "Não há nenhuma ação pendente para cancelar.";
		}

		private static string BeginPendingConfirmation(string prompt, string key, Func<string> action)
		{
			if (action == null)
				return "Não foi possível preparar a confirmação.";

			var nowUtc = DateTime.UtcNow;
			lock (_undoLock)
			{
				_pendingConfirmAction = action;
				_pendingConfirmKey = key;
				_undoAwaitingConfirm = true;
				_undoConfirmExpiresUtc = nowUtc.AddMilliseconds(UndoConfirmWindowMs);
			}

			var text = string.IsNullOrWhiteSpace(prompt)
				? "Confirmação necessária. Diga confirmar ou cancelar."
				: prompt;
			ShowUndoConfirmFormAsync(text);
			return text;
		}

		private static string ExecutarFusao(string fusionCommand, List<string> fusionTokens, dynamic json)
		{
			if (string.IsNullOrWhiteSpace(fusionCommand))
				return null;

			if (!_excelReady || _excelApp == null)
				return "Excel não está inicializado.";

			// After executing a FUSION command, clear any handgrab-based selection so it doesn't linger.
			// (Exception: the selection toggle command itself.)
			bool shouldClearSelectionAfter =
				!fusionCommand.Equals("TOGGLE_SELECTION_AT_ACTIVE_ROW", StringComparison.OrdinalIgnoreCase);

			string result;

			switch (fusionCommand.ToUpperInvariant())
			{
				case "TOGGLE_SELECTION_AT_ACTIVE_ROW":
						result = ExcelController.ToggleSelectionAtActiveRow();
					break;
				case "HIGHLIGHT_RESULTS":
					if (fusionTokens.Any(t => t.Equals("STUDENTSAPPROVED", StringComparison.OrdinalIgnoreCase)))
					{
						result = ExcelController.DestacarApenasAprovados();
						break;
					}
					if (fusionTokens.Any(t => t.Equals("STUDENTSFAILED", StringComparison.OrdinalIgnoreCase)))
					{
						result = ExcelController.DestacarApenasReprovados();
						break;
					}
					result = ExcelController.DestacarAprovados();
					break;
				case "HIGHLIGHT_RESULTS_ON_SELECTION":
					if (!ExcelController.HasSelection())
						ExcelController.ToggleSelectionAtActiveRow();

					// Garante a coluna Situação (se necessário) e atualiza Aprovado/Reprovado na seleção.
					result = ExcelController.DestacarAprovados();
					if (!string.IsNullOrWhiteSpace(result) &&
						result.IndexOf("coluna situa", StringComparison.OrdinalIgnoreCase) >= 0)
					{
						var pre = ExcelController.InserirSituacaoSelecao();
						if (!string.IsNullOrWhiteSpace(pre) &&
							(pre.IndexOf("calcule", StringComparison.OrdinalIgnoreCase) >= 0 ||
							 pre.IndexOf("erro", StringComparison.OrdinalIgnoreCase) >= 0))
						{
							result = pre;
							break;
						}
						result = ExcelController.DestacarAprovados();
					}
					break;
				case "CLOSE_EXCEL":
					result = BeginPendingConfirmation(
						"Quer mesmo fechar o Excel?\r\nDiga confirmar ou cancelar.",
						"CLOSE_EXCEL",
						() => CloseExcelInternal());
					break;
				case "CLOSE_EXCEL_CONFIRMED":
					result = CloseExcelInternal();
					break;
				case "UNDO_LAST_ACTION_CONFIRMED":
					result = ConfirmarAcaoPendentePorVoz();
					break;
					case "GUARDAR_FICHEIRO_CONFIRMED":
						result = ExcelController.GuardarRelatorio();
						break;
					case "APAGAR_TODOS_GRAFICOS_CONFIRMED":
						result = ExcelController.ApagarTodosGraficos();
						break;
					case "ATUALIZAR_NOTAS_CONFIRMED":
						result = ExcelController.AtualizarNotas(json);
						break;
				case "UNDO_LAST_ACTION":
					result = HandleUndoLastActionGesture();
					break;
				case "CALCULATE_AVERAGE":
					if (ExcelController.HasSelection())
					{
						result = ExcelController.CalcularMediaSelecao();
						break;
					}
					result = ExcelController.CalcularMedia(json);
					break;
				case "CALCULATE_AVERAGE_ON_SELECTION":
					if (!ExcelController.HasSelection())
							ExcelController.ToggleSelectionAtActiveRow();
					result = ExcelController.CalcularMediaSelecao();
					break;
				case "CREATE_PIVOT":
					result = ExcelController.CriarPivotTable(json);
					break;
				case "GENERATE_GRAPH_TURMA":
					result = ExcelController.HasSelection()
						? ExcelController.GerarGraficoTurmaSelecao(json)
						: ExcelController.GerarGraficoTurma(json);
					break;
				case "GENERATE_GRAPH_TURMA_ON_SELECTION":
					if (!ExcelController.HasSelection())
							ExcelController.ToggleSelectionAtActiveRow();
					result = ExcelController.GerarGraficoTurmaSelecao(json);
					break;
				case "GENERATE_GRAPH_ALUNO":
					// If the fused NLU includes a specific aluno, honor it even if a gesture selection exists.
					try
					{
						var alunoNumero = json?.nlu?["aluno_numero"]?.ToString();
						var alunoNome = json?.nlu?["aluno_nome"]?.ToString();
						if (!string.IsNullOrWhiteSpace(alunoNumero) || !string.IsNullOrWhiteSpace(alunoNome))
						{
							result = ExcelController.GerarGraficoBarras(json);
							break;
						}
					}
					catch
					{
						// best-effort
					}

					result = ExcelController.HasSelection()
						? ExcelController.GerarGraficoBarrasSelecao(json)
						: ExcelController.GerarGraficoBarras(json);
					break;
				case "GENERATE_GRAPH_ALUNO_ON_SELECTION":
					if (!ExcelController.HasSelection())
							ExcelController.ToggleSelectionAtActiveRow();
					result = ExcelController.GerarGraficoBarrasSelecao(json);
					break;
				case "GENERATE_GRAPH_PERGUNTAS_T2":
					result = ExcelController.GerarGraficoPerguntasT2();
					break;
				case "INSERT_COLUMN":
					result = ExcelController.InserirColuna(json);
					break;
				case "INSERT_COLUMN_THEN_HIGHLIGHT_APPROVED":
					{
						var ins = ExcelController.InserirColuna(json);
						// Se não conseguiu criar/validar pré-condições (ex: média em falta), devolve já.
						if (!string.IsNullOrWhiteSpace(ins) &&
							(ins.IndexOf("calcule", StringComparison.OrdinalIgnoreCase) >= 0 ||
							 ins.IndexOf("erro", StringComparison.OrdinalIgnoreCase) >= 0))
						{
							result = ins;
							break;
						}
						result = ExcelController.DestacarApenasAprovados();
						break;
					}
				case "INSERT_COLUMN_THEN_HIGHLIGHT_FAILED":
					{
						var ins = ExcelController.InserirColuna(json);
						if (!string.IsNullOrWhiteSpace(ins) &&
							(ins.IndexOf("calcule", StringComparison.OrdinalIgnoreCase) >= 0 ||
							 ins.IndexOf("erro", StringComparison.OrdinalIgnoreCase) >= 0))
						{
							result = ins;
							break;
						}
						result = ExcelController.DestacarApenasReprovados();
						break;
					}
				default:
					// Some SCXML rules forward gesture-style commands via FUSION (e.g., ZOOM_IN, SWIPE_RIGHT).
					// Reuse the gesture executor as a best-effort fallback.
					var gestureResult = ExecutarGesto(fusionCommand);
					if (!string.IsNullOrEmpty(gestureResult))
					{
						result = gestureResult;
						break;
					}

					Console.WriteLine("? Comando de fus?o n?o reconhecido: " + fusionCommand);
					result = null;
					break;
			}

			if (shouldClearSelectionAfter && ExcelController.HasSelection())
			{
				try { ExcelController.ClearSelectionAndRestoreFormatting(); }
				catch { /* best-effort */ }
			}

			return result;
		}

		private static string CloseExcelInternal()
		{
			try
			{
				_excelApp.DisplayAlerts = false;   // ? desativa popups

				_workbook?.SaveAs(excelPathFinal);
				_workbook?.Close(false);           // false = n?o perguntar nada
				_excelApp?.Quit();
				_excelReady = false;

				try { OleMessageFilter.Revoke(); } catch { /* best-effort */ }
				try
				{
					if (_sheet != null) Marshal.FinalReleaseComObject(_sheet);
					if (_workbook != null) Marshal.FinalReleaseComObject(_workbook);
					if (_excelApp != null) Marshal.FinalReleaseComObject(_excelApp);
				}
				catch
				{
					// ignore
				}
				finally
				{
					_sheet = null;
					_workbook = null;
					_excelApp = null;
				}

				return "Excel fechado.";
			}
			finally
			{
				if (_excelApp != null)
					_excelApp.DisplayAlerts = true; // (opcional) reativa alertas
			}
		}


		private static string ExecutarGesto(string gesture)
		{
			if (string.IsNullOrWhiteSpace(gesture))
				return null;

			// 🔹 Normalização (remove .a, underscores, etc.)
			gesture = gesture
				.ToLower()
				.Replace(".a", "")
				.Replace("_", "")
				.Replace("-", "")
				.Trim();

			Console.WriteLine("🎯 Gesto normalizado: " + gesture);

			if (!_excelReady || _excelApp == null)
			{
				Console.WriteLine("⚠ Excel ainda não está pronto.");
				return "Excel não está inicializado.";
			}

			switch (gesture)
			{
				// =========================
				// 📊 OPERAÇÕES EXCEL
				// =========================

				case "calculateaverage":
					return ExcelController.HasSelection()
						? ExcelController.CalcularMediaSelecao()
						: ExcelController.CalcularMediaTurma();

				case "insertcolumn":
					return ExcelController.InserirSituacao();

				case "studentsapproved":
					return ExcelController.DestacarApenasAprovados();

				case "studentsfailed":
					return ExcelController.DestacarApenasReprovados();

				case "undolastaction":
					return HandleUndoLastActionGesture();

					case "closeexcel":
						return BeginPendingConfirmation(
							"Quer mesmo fechar o Excel?\r\nDiga confirmar ou cancelar.",
							"CLOSE_EXCEL",
							() => CloseExcelInternal());

					// Kinect semantic (Gestures.xml): handgrab -> toggle seleção da linha atual.
					// Toggle por linha: permite selecionar várias linhas e também desselecionar repetindo o gesto na mesma linha.
					case "handgrab":
						return ExcelController.ToggleSelectionAtActiveRow();

				// =========================
				// 🔍 NAVEGAÇÃO / MOVIMENTO
				// =========================

				case "swipeleft":
					var leftCell = _excelApp.ActiveCell;
					if (leftCell != null && leftCell.Column > 1)
					{
						leftCell.Offset[0, -1].Select();
						return "Mover para a esquerda.";
					}
					return "Já está na primeira coluna.";

				case "swiperight":
					var rightCell = _excelApp.ActiveCell;
					if (rightCell != null && rightCell.Column < 16384) // XFD = column 16384
					{
						rightCell.Offset[0, 1].Select();
						return "Mover para a direita.";
					}
					return "Já está na última coluna.";

				case "swipeup":
					var upCell = _excelApp.ActiveCell;
					if (upCell != null && upCell.Row > 1)
					{
						upCell.Offset[-1, 0].Select();
						return "Mover para cima.";
					}
					return "Já está na primeira linha.";

				case "swipedown":
					var downCell = _excelApp.ActiveCell;
					if (downCell != null && downCell.Row < 1048576)
					{
						downCell.Offset[1, 0].Select();
						return "Mover para baixo.";
					}
					return "Já está na última linha.";

				// =========================
				// 🔎 ZOOM
				// =========================

				case "zoomin":
					_excelApp.ActiveWindow.Zoom += 10;
					return "Zoom aumentado.";

				case "zoomout":
					_excelApp.ActiveWindow.Zoom -= 10;
					return "Zoom reduzido.";

				// =========================
				// ⚠ FALLBACK
				// =========================

				default:
					Console.WriteLine("⚠ Gesto não reconhecido: " + gesture);
					return null;
			}
		}



		private static string HandleUndoLastActionGesture()
		{
			if (_excelApp?.ActiveCell == null)
			{
				ClearPendingUndoConfirmation();
				CloseUndoConfirmFormAsync();
				return "Nenhuma célula ativa para apagar.";
			}

			var nowUtc = DateTime.UtcNow;
			lock (_undoLock)
			{
				_pendingConfirmAction = () =>
				{
					_excelApp.ActiveCell.ClearContents();
					return "Valor da célula apagado.";
				};
				_pendingConfirmKey = "UNDO_LAST_ACTION";
				_undoAwaitingConfirm = true;
				_undoConfirmExpiresUtc = nowUtc.AddMilliseconds(UndoConfirmWindowMs);
			}

			const string msg = "Confirmação necessária. Diga confirmar ou cancelar.";
			ShowUndoConfirmFormAsync(msg);
			return msg;
		}

		private static void ClearPendingUndoConfirmation()
		{
			lock (_undoLock)
			{
				_undoAwaitingConfirm = false;
				_undoConfirmExpiresUtc = DateTime.MinValue;
				_pendingConfirmKey = null;
				_pendingConfirmAction = null;
			}
		}

		private static void ShowUndoConfirmFormAsync(string text)
		{
			_ui.BeginInvoke(() =>
			{
				try
				{
					if (_undoConfirmForm != null && !_undoConfirmForm.IsDisposed)
					{
						_undoConfirmForm.SetText(text);
						_undoConfirmForm.Activate();
						return;
					}

					_undoConfirmForm = new UndoConfirmForm(text);
					_undoConfirmForm.FormClosed += (s, e) =>
					{
						try
						{
							if (_undoConfirmForm != null && _undoConfirmForm.IsDisposed)
								_undoConfirmForm = null;
						}
						catch
						{
							// best-effort
						}
					};

					_undoConfirmForm.Show();
					_undoConfirmForm.Activate();
				}
				catch (Exception ex)
				{
					Console.WriteLine("❌ Erro ao mostrar confirmação: " + ex.Message);
				}
			});
		}

		private static void CloseUndoConfirmFormAsync()
		{
			_ui.BeginInvoke(() =>
			{
				try
				{
					if (_undoConfirmForm != null && !_undoConfirmForm.IsDisposed)
					{
						_undoConfirmForm.Close();
						_undoConfirmForm.Dispose();
						_undoConfirmForm = null;
					}
				}
				catch
				{
					// best-effort
				}
			});
		}

		private sealed class UiDispatcher
		{
			private readonly ManualResetEventSlim _ready = new ManualResetEventSlim(false);
			private readonly Thread _thread;
			private InvokerForm _invoker;

			public UiDispatcher()
			{
				_thread = new Thread(ThreadMain)
				{
					IsBackground = true,
					Name = "UI-Thread"
				};
				_thread.SetApartmentState(ApartmentState.STA);
				_thread.Start();
				_ready.Wait();
			}

			private void ThreadMain()
			{
				Application.EnableVisualStyles();
				Application.SetCompatibleTextRenderingDefault(false);

				_invoker = new InvokerForm();
				_invoker.Shown += (s, e) =>
				{
					_invoker.Hide();
					_ready.Set();
				};
				Application.Run(_invoker);
			}

			public void BeginInvoke(Action action)
			{
				try
				{
					if (_invoker == null || _invoker.IsDisposed) return;
					_invoker.BeginInvoke(action);
				}
				catch
				{
					// best-effort
				}
			}

			public void Invoke(Action action)
			{
				try
				{
					if (_invoker == null || _invoker.IsDisposed) return;
					if (_invoker.InvokeRequired)
						_invoker.Invoke(action);
					else
						action();
				}
				catch
				{
					// best-effort
				}
			}

			public T Invoke<T>(Func<T> func)
			{
				try
				{
					if (_invoker == null || _invoker.IsDisposed) return default(T);
					if (_invoker.InvokeRequired)
						return (T)_invoker.Invoke(func);
					return func();
				}
				catch
				{
					return default(T);
				}
			}

			private sealed class InvokerForm : Form
			{
				public InvokerForm()
				{
					FormBorderStyle = FormBorderStyle.FixedToolWindow;
					ShowInTaskbar = false;
					StartPosition = FormStartPosition.Manual;
					Size = new System.Drawing.Size(1, 1);
					Location = new System.Drawing.Point(-2000, -2000);
					Opacity = 0;
				}
			}
		}

		private sealed class UndoConfirmForm : Form
		{
			private readonly Label _label;

			public UndoConfirmForm(string text)
			{
				Text = "Confirmação";
				FormBorderStyle = FormBorderStyle.FixedToolWindow;
				ShowInTaskbar = false;
				StartPosition = FormStartPosition.CenterScreen;
				TopMost = true;
				MaximizeBox = false;
				MinimizeBox = false;
				Size = new Size(520, 140);

				_label = new Label
				{
					Dock = DockStyle.Fill,
					TextAlign = ContentAlignment.MiddleCenter,
					Font = new Font("Segoe UI", 11f, FontStyle.Regular),
					Padding = new Padding(16),
					Text = text
				};
				Controls.Add(_label);
			}

			public void SetText(string text)
			{
				if (InvokeRequired)
				{
					BeginInvoke(new Action(() => SetText(text)));
					return;
				}
				_label.Text = text ?? string.Empty;
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

	internal sealed class OleMessageFilter : OleMessageFilter.IOleMessageFilter
	{
		public static void Register()
		{
			IOleMessageFilter newFilter = new OleMessageFilter();
			CoRegisterMessageFilter(newFilter, out _);
		}

		public static void Revoke()
		{
			CoRegisterMessageFilter(null, out _);
		}

		int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
		{
			return 0; // SERVERCALL_ISHANDLED
		}

		int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
		{
			if (dwRejectType == 2) // SERVERCALL_RETRYLATER
				return 99; // Retry the thread call after 99 milliseconds
			return -1;
		}

		int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
		{
			return 2; // PENDINGMSG_WAITDEFPROCESS
		}

		[System.Runtime.InteropServices.DllImport("ole32.dll")]
		private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

		[System.Runtime.InteropServices.ComImport, System.Runtime.InteropServices.Guid("00000016-0000-0000-C000-000000000046"), System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)]
		internal interface IOleMessageFilter
		{
			[PreserveSig]
			int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
			[PreserveSig]
			int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
			[PreserveSig]
			int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
		}
	}
}
