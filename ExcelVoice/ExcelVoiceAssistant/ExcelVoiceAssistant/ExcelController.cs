using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVoiceAssistant
{
    public static class ExcelController
    {
        private static Excel.Application app;
        private static Excel.Workbook workbook;
        private static Excel.Worksheet sheet;

        private static string pathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_Excel\ETP3.xlsx";
        private static string pathFinal = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_Excel\Relatorio_Final.xlsx";

        //private static string pathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\ETP.xlsx";
        //private static string pathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\Relatorio_Final.xlsx";

        // =====================================================
        // Ligar Excel já aberto
        // =====================================================
        public static void SetExcel(Excel.Application excelApp, Excel.Workbook wb, Excel.Worksheet ws)
        {
            app = excelApp;
            workbook = wb;
            sheet = ws;
        }

        // =====================================================
        // Normalizar texto
        // =====================================================
        private static bool IgualIgnorandoAcentos(string a, string b)
        {
            if (a == null || b == null) return false;

            string Normalize(string s) =>
                new string(
                    s.Normalize(NormalizationForm.FormD)
                     .Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
                     .ToArray()
                ).ToLower().Trim();

            return Normalize(a) == Normalize(b);
        }

        // =====================================================
        // Converter número → letra
        // =====================================================
        private static string ColunaParaLetra(int coluna)
        {
            string letra = "";
            while (coluna > 0)
            {
                int resto = (coluna - 1) % 26;
                letra = (char)(65 + resto) + letra;
                coluna = (coluna - 1) / 26;
            }
            return letra;
        }

        // =====================================================
        // Encontrar cabeçalho "Aluno"
        // =====================================================
        private static (int headerRow, int headerCol) EncontrarCabecalho()
        {
            Excel.Range used = sheet.UsedRange;

            int firstRow = used.Row;
            int lastRow = firstRow + used.Rows.Count - 1;
            int firstCol = used.Column;
            int lastCol = firstCol + used.Columns.Count - 1;

            for (int r = firstRow; r <= lastRow; r++)
            {
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var valor = sheet.Cells[r, c].Value;
                    if (valor != null && IgualIgnorandoAcentos(valor.ToString(), "nome"))
                        return (r, c);
                }
            }

            throw new Exception("Cabeçalho Nome não encontrado.");
        }


        // =====================================================
        // CALCULAR MÉDIA DINAMICAMENTE

        public static string CalcularMedia(dynamic json)
        {
            if (json.nlu.entities == null)
                return CalcularMediaTurma();

            string nome = null;
            string numero = null;

            foreach (var e in json.nlu.entities)
            {
                if (e.entity == "aluno_nome") nome = e.value.ToString();
                if (e.entity == "aluno_numero") numero = e.value.ToString();
            }

            if (!string.IsNullOrEmpty(nome)) return CalcularMediaAluno(nome);
            if (!string.IsNullOrEmpty(numero)) return CalcularMediaAlunoNumero(numero);

            return CalcularMediaTurma();
        }


        public static string CalcularMediaTurma()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                List<int> testes = new List<int>();
                int colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (titulo.ToLower().StartsWith("teste")) testes.Add(c);
                    if (IgualIgnorandoAcentos(titulo, "média")) colMedia = c;
                }

                if (testes.Count == 0)
                    return "Nenhuma coluna de teste encontrada.";

                testes.Sort();

                if (colMedia == -1)
                {
                    colMedia = testes.Last() + 1;
                    sheet.Cells[headerRow, colMedia].Value2 = "Média";
                }

                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    string formula = "=MÉDIA(" +
                        string.Join(";", testes.Select(col => $"{ColunaParaLetra(col)}{row}")) +
                        ")";

                    sheet.Cells[row, colMedia].FormulaLocal = formula;
                    row++;
                }

                workbook.Save();
                return "Média turma calculada.";
            }
            catch
            {
                return "Erro ao calcular média turma.";
            }
        }


        public static string CalcularMediaAluno(string nomeAluno)
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                List<int> colTestes = new List<int>();
                int colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (titulo.Trim().ToLower().StartsWith("teste")) colTestes.Add(c);
                    if (IgualIgnorandoAcentos(titulo, "média")) colMedia = c;
                }

                if (colTestes.Count == 0)
                    return "Sem testes.";

                if (colMedia == -1)
                {
                    colMedia = colTestes.Last() + 1;
                    sheet.Cells[headerRow, colMedia].Value2 = "Média";
                }

                int rowAluno = -1;
                int row = headerRow + 1;


                var partes = nomeAluno.ToLower()
                    .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    string excelNome = sheet.Cells[row, headerCol].Value.ToString().ToLower();

                    if (partes.All(p => excelNome.Contains(p)))
                    {
                        rowAluno = row;
                        break;
                    }

                    row++;
                }

                if (rowAluno == -1)
                    return $"Aluno {nomeAluno} não encontrado.";

                string formula = "=MÉDIA(" +
                    string.Join(";", colTestes.Select(c => $"{ColunaParaLetra(c)}{rowAluno}")) + ")";

                sheet.Cells[rowAluno, colMedia].FormulaLocal = formula;

                workbook.Save();
                return $"Média calculada para {nomeAluno}.";
            }
            catch
            {
                return "Erro ao calcular média.";
            }
        }


        public static string CalcularMediaAlunoNumero(string numeroMec)
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMec = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string raw = sheet.Cells[headerRow, c].Value?.ToString();
                    if (raw != null && IgualIgnorandoAcentos(raw, "Número mecanográfico"))
                    {
                        colMec = c;
                        break;
                    }
                }

                if (colMec == -1)
                    return "Coluna mec não encontrada.";

                int rowAluno = -1;
                int r = headerRow + 1;

                while (sheet.Cells[r, colMec].Value != null)
                {
                    if (sheet.Cells[r, colMec].Value.ToString() == numeroMec)
                    {
                        rowAluno = r;
                        break;
                    }
                    r++;
                }

                if (rowAluno == -1)
                    return "Aluno não encontrado.";

                List<int> colTestes = new List<int>();
                int colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (titulo.ToLower().StartsWith("teste")) colTestes.Add(c);
                    if (IgualIgnorandoAcentos(titulo, "média")) colMedia = c;
                }

                if (colMedia == -1)
                {
                    colMedia = colTestes.Last() + 1;
                    sheet.Cells[headerRow, colMedia].Value2 = "Média";
                }

                string formula = "=MÉDIA(" +
                    string.Join(";", colTestes.Select(c => $"{ColunaParaLetra(c)}{rowAluno}")) + ")";

                sheet.Cells[rowAluno, colMedia].FormulaLocal = formula;

                workbook.Save();
                return $"Média calculada para {numeroMec}.";
            }
            catch
            {
                return "Erro ao calcular média.";
            }
        }
        public static string InserirSituacao()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "média"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                if (colMedia == -1)
                    return "Calcule a média primeiro.";

                int colSit = colMedia + 1;
                sheet.Cells[headerRow, colSit].Value2 = "Situação";

                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    sheet.Cells[row, colSit].Value2 = "";
                    row++;
                }

                return "Coluna situação criada.";
            }
            catch
            {
                return "Erro ao criar Situação.";
            }
        }
        public static string DestacarAprovados()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMedia = -1;
                int colSit = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v == null) continue;

                    if (IgualIgnorandoAcentos(v.ToString(), "média"))
                        colMedia = c;

                    if (IgualIgnorandoAcentos(v.ToString(), "situação"))
                        colSit = c;
                }

                // 🔥 PRIMEIRA VERIFICAÇÃO: Falta coluna Situação
                if (colSit == -1)
                    return "Criar coluna situação primeiro.";

                // 🔥 SEGUNDA VERIFICAÇÃO: Falta coluna Média
                if (colMedia == -1)
                    return "Calcular média primeiro.";

                // ⭐ Ambas existem → processar normalmente
                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double media = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (media >= 10)
                    {
                        sheet.Cells[row, colSit].Value2 = "Aprovado";
                        sheet.Cells[row, colSit].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                    }
                    else
                    {
                        sheet.Cells[row, colSit].Value2 = "Reprovado";
                        sheet.Cells[row, colSit].Interior.Color = ColorTranslator.ToOle(Color.LightCoral);
                    }

                    row++;
                }

                return "Situação atualizada com sucesso";
            }
            catch
            {
                return "Erro ao destacar.";
            }
        }

        public static string IdentificarMelhoria()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // 1️⃣ Encontrar dinamicamente todas as colunas Teste X
                List<(int col, int num)> testes = new List<(int col, int num)>();

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    var match = Regex.Match(titulo.ToLower().Replace(" ", ""), @"teste(\d+)");
                    if (match.Success)
                    {
                        testes.Add((c, int.Parse(match.Groups[1].Value)));
                    }
                }

                if (testes.Count < 1)
                    return "Nenhum teste encontrado.";

                // Ordenar Teste 1, Teste 2, Teste 3, ...
                testes = testes.OrderBy(t => t.num).ToList();

                int numTestes = testes.Count;
                int colUltimoTeste = testes.Last().col;

                // 2️⃣ Encontrar coluna Média
                int colMedia = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string v = sheet.Cells[headerRow, c].Value?.ToString();
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "média"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                if (colMedia == -1)
                    return "Calcule a média primeiro.";

                // 3️⃣ Criar coluna Melhoria se necessário
                int colMelhoria = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string raw = sheet.Cells[headerRow, c].Value?.ToString();
                    if (raw != null && IgualIgnorandoAcentos(raw, "melhoria"))
                    {
                        colMelhoria = c;
                        break;
                    }
                }

                if (colMelhoria == -1)
                {
                    colMelhoria = lastCol + 1;
                    sheet.Cells[headerRow, colMelhoria].Value2 = "Melhoria";
                }

                // 4️⃣ Calcular MP linha a linha
                int row = headerRow + 1;
                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double mediaAtual = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (mediaAtual >= 10)
                    {
                        sheet.Cells[row, colMelhoria].Value2 = "";
                        row++;
                        continue;
                    }

                    // somar todos os testes exceto o último
                    double somaAnteriores = 0;

                    foreach (var t in testes.Take(testes.Count - 1))
                    {
                        somaAnteriores += Convert.ToDouble(sheet.Cells[row, t.col].Value2 ?? 0);
                    }

                    // 5️⃣ Nota necessária no último teste para atingir média 10
                    double notaNecessaria =
                        10 * numTestes - somaAnteriores;

                    // 6️⃣ Se a nota necessária for possível (<=20) → MP
                    if (notaNecessaria <= 20)
                        sheet.Cells[row, colMelhoria].Value2 = "MP";
                    else
                        sheet.Cells[row, colMelhoria].Value2 = "";

                    row++;
                }

                return $"Melhoria calculada usando todos os {numTestes} testes.";
            }
            catch (Exception ex)
            {
                return "Erro ao identificar melhoria: " + ex.Message;
            }
        }


        public static string GerarGraficoTurma(dynamic json)
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colT1 = -1, colT2 = -1, colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "teste 2")) colT2 = c;
                    if (IgualIgnorandoAcentos(titulo, "média")) colMedia = c;
                }

                if (colT1 == -1 || colT2 == -1 || colMedia == -1)
                    return "Colunas T1, T2 ou média não encontradas.";

                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerCol].Value != null)
                    lastRow++;

                int count = lastRow - headerRow - 1;
                if (count <= 0)
                    return "Sem alunos.";

                double somaT1 = 0, somaT2 = 0, somaM = 0;

                for (int r = headerRow + 1; r < lastRow; r++)
                {
                    somaT1 += Convert.ToDouble(sheet.Cells[r, colT1].Value2 ?? 0);
                    somaT2 += Convert.ToDouble(sheet.Cells[r, colT2].Value2 ?? 0);
                    somaM += Convert.ToDouble(sheet.Cells[r, colMedia].Value2 ?? 0);
                }

                double mT1 = somaT1 / count;
                double mT2 = somaT2 / count;
                double mMF = somaM / count;

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 20
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 30;

                Excel.ChartObject chartObj = charts.Add(40, posY, 650, 360);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Médias da Turma";

                Excel.Series s = chart.SeriesCollection().NewSeries();
                s.Name = "Médias";
                s.Values = new double[] { mT1, mT2, mMF };
                s.XValues = new string[] { "Teste 1", "Teste 2", "Média" };

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                return "Gráfico criado.";
            }
            catch
            {
                return "Erro ao criar gráfico.";
            }
        }
        private static (string nome, string numero) ExtrairAluno(dynamic json)
        {
            string nome = null;
            string numero = null;

            if (json?.nlu?.entities != null)
            {
                foreach (var ent in json.nlu.entities)
                {
                    if (ent.entity == "aluno_nome")
                        nome = ent.value.ToString();

                    if (ent.entity == "aluno_numero")
                        numero = ent.value.ToString();
                }
            }

            return (nome, numero);
        }
        public static string GerarGraficoBarras(dynamic json)
        {
            try
            {
                var aluno = ExtrairAluno(json);
                string numero = aluno.numero;

                if (string.IsNullOrEmpty(numero))
                    return "Número mec não encontrado.";

                var (headerRow, colNome) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMec = -1;
                int colT1 = -1, colT2 = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "número mecanográfico")) colMec = c;
                    if (IgualIgnorandoAcentos(titulo, "teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "teste 2")) colT2 = c;
                }

                if (colMec == -1 || colT1 == -1 || colT2 == -1)
                    return "Colunas do aluno não encontradas.";

                int row = headerRow + 1;
                int rowAluno = -1;

                while (sheet.Cells[row, colNome].Value != null)
                {
                    string val = sheet.Cells[row, colMec].Value?.ToString();
                    if (val == numero)
                    {
                        rowAluno = row;
                        break;
                    }
                    row++;
                }

                if (rowAluno == -1)
                    return "Aluno não encontrado.";

                string nomeReal = sheet.Cells[rowAluno, colNome].Value?.ToString();

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[row].Top + 20
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 30;

                Excel.ChartObject chartObj = charts.Add(40, posY, 700, 360);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = $"Notas de {nomeReal}";

                Excel.SeriesCollection sc = chart.SeriesCollection();

                var s1 = sc.NewSeries();
                s1.Name = "Teste 1";
                s1.Values = sheet.Range[$"{ColunaParaLetra(colT1)}{rowAluno}"];
                s1.XValues = "\"Teste 1\"";

                var s2 = sc.NewSeries();
                s2.Name = "Teste 2";
                s2.Values = sheet.Range[$"{ColunaParaLetra(colT2)}{rowAluno}"];
                s2.XValues = "\"Teste 2\"";

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                return "Gráfico do aluno criado.";
            }
            catch
            {
                return "Erro no gráfico.";
            }
        }


        public static string AtualizarNotas(dynamic json)
        {
            try
            {
                // 1️⃣ ALUNO (número mecanográfico)
                var alunoInfo = ExtrairAluno(json);
                string numeroMec = alunoInfo.numero;
                if (string.IsNullOrEmpty(numeroMec))
                {
                    Console.WriteLine("❌ Número mecanográfico não encontrado.");
                    return "";
                }
                // 2️⃣ TEXTO BASE64 → frase original
                string base64 = json.text.ToString();
                string frase = Encoding.UTF8.GetString(Convert.FromBase64String(base64));
                Console.WriteLine("📥 Texto decodificado: " + frase);

                // -------------------------------------------------------------
                // NORMALIZAÇÃO INTELIGENTE DO INPUT
                // -------------------------------------------------------------

                // 1) Corrigir número mecanográfico dito com pausas ("978 76" → "97876")
                // Declarar antes de qualquer uso
                double[] notas = null;

                if (!string.IsNullOrEmpty(numeroMec))
                {
                    numeroMec = Regex.Replace(numeroMec, @"\s+", "");   // remover espaços
                }


                // 2) Verificar se as notas vieram coladas ("12345")
                bool notasColadas = false;

                // Contamos quantos dígitos existem depois da palavra "para"
                int idxPara2 = frase.ToLower().IndexOf("para");
                if (idxPara2 != -1)
                {
                    string depois = frase.Substring(idxPara2);

                    // Se só há UM MATCH e tem 2 ou mais dígitos → provavelmente são várias notas coladas
                    var matchesPossiveis = Regex.Matches(depois, @"\d");
                    var matchGrande = Regex.Match(depois, @"\d{2,}");

                    if (matchGrande.Success && matchesPossiveis.Count > 1 && matchGrande.Value.Length > 1)
                        notasColadas = true;
                }

                // Se as notas estiverem coladas, expandimos cada dígito individualmente
                if (notasColadas)
                {
                    Console.WriteLine("⚠️ Detetado padrão de notas coladas. A separar dígitos...");

                    string apenasDigitos = Regex.Replace(frase.Substring(frase.ToLower().IndexOf("para")), @"\D", "");

                    // converter cada dígito numa nota
                    List<double> lista = new List<double>();
                    foreach (char ch in apenasDigitos)
                    {
                        if (char.IsDigit(ch))
                            lista.Add(double.Parse(ch.ToString()));
                    }

                    // substituir as notas
                    notas = lista.ToArray();

                    Console.WriteLine("📌 Notas corrigidas: " + string.Join(", ", notas));
                }


                // 3️⃣ INTERPRETAR PERGUNTA / TESTE NATURAL
                string perguntaRaw = json.nlu.pergunta != null ? json.nlu.pergunta.ToString().ToLower() : "";
                string colunaAlvo = "";

                // PERGUNTA 1..5 → T2_P1..T2_P5
                var matchPerg = Regex.Match(frase.ToLower(), @"pergunta\s*(\d)");
                if (matchPerg.Success)
                {
                    int num = int.Parse(matchPerg.Groups[1].Value);
                    colunaAlvo = $"T2_P{num}";
                }

                if (frase.ToLower().Contains("teste 1"))
                    colunaAlvo = "Teste 1";

                if (frase.ToLower().Contains("teste 2"))
                    colunaAlvo = "Teste 2";


                // 4️⃣ EXTRAIR NOTAS (todos os números)
                var matches = System.Text.RegularExpressions.Regex.Matches(frase, @"\d+[.,]?\d*");

                if (matches.Count == 0)
                {
                    Console.WriteLine("❌ Nenhum valor encontrado.");
                    return "";
                }

                notas = matches
                    .Cast<System.Text.RegularExpressions.Match>()
                    .Select(m => double.Parse(m.Value.Replace(",", "."), CultureInfo.InvariantCulture))
                    .ToArray();

                Console.WriteLine("📌 Notas extraídas: " + string.Join(", ", notas));

                // 🛑 REMOVER O NÚMERO MECANOGRÁFICO DA LISTA DE NOTAS
                // 🧹 REMOVER números que não são notas (mec, pergunta, teste)

                List<double> filtradas = new List<double>();

                // A frase sempre tem a estrutura "... para X Y Z"
                int idxPara = frase.ToLower().IndexOf("para");
                if (idxPara != -1)
                {
                    string soDepois = frase.Substring(idxPara); // só texto depois de "para"
                    var matchesAfter = System.Text.RegularExpressions.Regex.Matches(soDepois, @"\d+[.,]?\d*");

                    foreach (Match m in matchesAfter)
                    {
                        double v = double.Parse(m.Value.Replace(",", "."), CultureInfo.InvariantCulture);

                        // remover número mecanográfico
                        if (v.ToString() == numeroMec) continue;

                        // remover "1" de "teste 1"
                        if (perguntaRaw.Contains("teste 1") && v == 1) continue;

                        // remover "2" de "teste 2"
                        if (perguntaRaw.Contains("teste 2") && v == 2) continue;

                        // remover número da pergunta (ex: "pergunta 1")
                        var pm = System.Text.RegularExpressions.Regex.Match(perguntaRaw, @"\d");
                        if (pm.Success && v == int.Parse(pm.Value)) continue;

                        filtradas.Add(v);
                    }
                }

                notas = filtradas.ToArray();

                if (notas.Length == 0)
                {
                    Console.WriteLine("❌ Não foram encontradas notas válidas.");
                    return "";
                }

                Console.WriteLine("📌 Notas finais filtradas: " + string.Join(", ", notas));



                // 5️⃣ LOCALIZAR LINHA DO ALUNO POR NÚMERO MECANOGRÁFICO
                Excel.Range used = sheet.UsedRange;

                var (headerRow, headerColNome) = EncontrarCabecalho();
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;


                // ================================================================
                // 🔥 **CORREÇÃO CRÍTICA: DETETAR COLUNA 'Número mecanográfico'**
                // ================================================================
                int colNumeroMec = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var raw = sheet.Cells[headerRow, c].Value?.ToString();
                    if (raw == null) continue;

                    string v = raw
                        .Trim()
                        .ToLower()
                        .Replace("  ", " ")
                        .Normalize(NormalizationForm.FormD);

                    v = new string(v.Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark).ToArray());

                    if (v.Contains("numero") && v.Contains("mecanograf"))
                    {
                        colNumeroMec = c;
                        break;
                    }
                }

                if (colNumeroMec == -1)
                {
                    Console.WriteLine("❌ Coluna 'Número mecanográfico' não encontrada.");
                    return "";
                }

                // Procurar linha do aluno
                int alunoRow = -1;
                int rPtr = headerRow + 1;

                while (sheet.Cells[rPtr, colNumeroMec].Value != null)
                {
                    string val = sheet.Cells[rPtr, colNumeroMec].Value.ToString().Trim();
                    if (val == numeroMec)
                    {
                        alunoRow = rPtr;
                        break;
                    }
                    rPtr++;
                }

                if (alunoRow == -1)
                {
                    Console.WriteLine($"❌ Número mecanográfico {numeroMec} não encontrado.");
                    return "";
                }

                // 6️⃣ LOCALIZAR TODAS AS COLUNAS DE TESTES E PERGUNTAS
                Dictionary<string, int> mapaColunas = new Dictionary<string, int>();

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    string normal = titulo.Trim();

                    if (IgualIgnorandoAcentos(normal, "Teste 1")) mapaColunas["Teste 1"] = c;
                    if (IgualIgnorandoAcentos(normal, "Teste 2")) mapaColunas["Teste 2"] = c;
                    if (normal.StartsWith("T2_P")) mapaColunas[normal] = c;
                }

                // Verifica coluna alvo
                if (!string.IsNullOrEmpty(colunaAlvo) && !mapaColunas.ContainsKey(colunaAlvo))
                {
                    Console.WriteLine($"❌ Coluna '{colunaAlvo}' não encontrada.");
                    return "";
                }

                // 7️⃣ Atualizar Pergunta específica
                if (!string.IsNullOrEmpty(colunaAlvo) && colunaAlvo.StartsWith("T2_P"))
                {
                    int col = mapaColunas[colunaAlvo];
                    sheet.Cells[alunoRow, col].Value2 = notas[0];

                    Console.WriteLine($"✏️ Atualizada {colunaAlvo} do aluno {numeroMec} para {notas[0]}.");
                }
                else
                {
                    // 8️⃣ Atualizar Teste 1
                    if (colunaAlvo == "Teste 1" && mapaColunas.ContainsKey("Teste 1"))
                    {
                        sheet.Cells[alunoRow, mapaColunas["Teste 1"]].Value2 = notas[0];
                        Console.WriteLine("✏️ Atualizado Teste 1.");
                    }

                    // Atualizar Teste 2
                    else if (colunaAlvo == "Teste 2" && mapaColunas.ContainsKey("Teste 2"))
                    {
                        sheet.Cells[alunoRow, mapaColunas["Teste 2"]].Value2 = notas[0];
                        Console.WriteLine("✏️ Atualizado Teste 2.");
                    }

                    // 9️⃣ Atualizar várias perguntas (ex: "12 14 15 10 9")
                    else if (notas.Length > 1)
                    {
                        int i = 0;
                        foreach (var kv in mapaColunas.Where(k => k.Key.StartsWith("T2_P")).OrderBy(k => k.Key))
                        {
                            if (i >= notas.Length) break;
                            sheet.Cells[alunoRow, kv.Value].Value2 = notas[i];
                            i++;
                        }

                        Console.WriteLine("✏️ Atualizadas várias perguntas do Teste 2.");
                    }
                }

                // Atualizar Teste 2 como soma das perguntas
                if (mapaColunas.ContainsKey("Teste 2"))
                {
                    double soma = 0;
                    foreach (var kv in mapaColunas.Where(k => k.Key.StartsWith("T2_P")))
                    {
                        soma += Convert.ToDouble(sheet.Cells[alunoRow, kv.Value].Value2 ?? 0);
                    }

                    sheet.Cells[alunoRow, mapaColunas["Teste 2"]].Value2 = soma;
                    Console.WriteLine($"🔄 Teste 2 recalculado automaticamente = {soma}");
                }


                // 🔟 RE-CALCULAR MÉDIA SE EXISTIR
                int colMedia = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value?.ToString();
                    if (v != null && IgualIgnorandoAcentos(v, "Média"))
                        colMedia = c;
                }

                if (colMedia != -1 && mapaColunas.ContainsKey("Teste 1") && mapaColunas.ContainsKey("Teste 2"))
                {
                    string cT1 = ColunaParaLetra(mapaColunas["Teste 1"]);
                    string cT2 = ColunaParaLetra(mapaColunas["Teste 2"]);

                    sheet.Cells[alunoRow, colMedia].FormulaLocal =
                        $"=MÉDIA({cT1}{alunoRow};{cT2}{alunoRow})";

                    Console.WriteLine("📊 Média recalculada.");
                }

                workbook.Save();
                Console.WriteLine("✅ Atualização concluída!");
                return "Notas atualizadas com sucesso.";
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao atualizar notas: " + ex.Message);
                return "Erro ao atualizar notas: " + ex.Message;
            }
        }
        public static string ApagarGrafico(dynamic json)
        {
            try
            {
                var info = ExtrairAluno(json);
                string alvo = (info.nome ?? info.numero ?? "").ToLower();

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                for (int i = charts.Count; i >= 1; i--)
                {
                    var chart = charts.Item(i).Chart;
                    string titulo = chart.ChartTitle?.Text?.ToLower() ?? "";

                    if (!string.IsNullOrEmpty(alvo) && titulo.Contains(alvo))
                    {
                        charts.Item(i).Delete();
                        return "Gráfico apagado.";
                    }

                    if (titulo.Contains("médias") || titulo.Contains("media"))
                    {
                        charts.Item(i).Delete();
                        return "Gráfico apagado.";
                    }
                }

                return "Nenhum gráfico encontrado.";
            }
            catch
            {
                return "Erro ao apagar gráfico.";
            }
        }


        public static string ApagarTodosGraficos()
        {
            try
            {
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                for (int i = charts.Count; i >= 1; i--)
                    charts.Item(i).Delete();

                return "Todos os gráficos apagados.";
            }
            catch
            {
                return "Erro ao apagar todos.";
            }
        }

        public static string OperacoesMatematicas(dynamic json)
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // Encontrar coluna média
                int colMedia = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "media"))
                    {
                        colMedia = c;
                        break;
                    }
                }
                if (colMedia == -1)
                {
                    Console.WriteLine("⚠️ É necessário calcular a média primeiro.");
                    return "É necessário calcular a média primeiro.";
                }

                int row = headerRow + 1;
                int aprovados = 0, reprovados = 0, acimaDe16 = 0, total = 0;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double media = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (media >= 10) aprovados++;
                    else reprovados++;

                    if (media >= 16) acimaDe16++;

                    total++;
                    row++;
                }

                string resultado = $"Aprovados: {aprovados}, Reprovados: {reprovados}, Acima de 16: {acimaDe16}, Total: {total}";
                Console.WriteLine($"📊 Estatísticas: {resultado}");
                return resultado;
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro em Operações Matemáticas: " + ex.Message);
                return "Erro em Operações Matemáticas: " + ex.Message;
            }
        }
        public static string GuardarRelatorio()
        {
            try
            {
                workbook.SaveAs(pathFinal);
                Console.WriteLine("💾 Relatório guardado!");
                return "Relatório guardado.";
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao guardar relatório: " + ex.Message);
                return "Erro ao guardar relatório: " + ex.Message;
            }
        }


    }
}
