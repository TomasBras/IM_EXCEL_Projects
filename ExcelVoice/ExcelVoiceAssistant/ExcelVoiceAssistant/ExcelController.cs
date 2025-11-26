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

        //private static string pathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM\IM_EXCEL_NODEPENDENCIES\ETP.xlsx";
        //private static string pathFinal = @"C:\Users\trmbr\OneDrive\Desktop\IM\IM_EXCEL_NODEPENDENCIES\Relatorio_Final.xlsx";

        private static string pathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_NODEPENDENCIES\ETP.xlsx";
        private static string pathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_NODEPENDENCIES\Relatorio_Final.xlsx";

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

            throw new Exception("Cabeçalho 'Nome' não encontrado.");
        }

        // =====================================================
        // CALCULAR MÉDIA DINAMICAMENTE
        public static void CalcularMedia()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colT1 = -1;
                int colT2 = -1;
                int colMedia = -1;

                // 1️⃣ Encontrar colunas pelo nome EXACTO do cabeçalho
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "Teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "Teste 2")) colT2 = c;
                    if (IgualIgnorandoAcentos(titulo, "Média")) colMedia = c;
                }

                if (colT1 == -1 || colT2 == -1)
                {
                    Console.WriteLine("❌ Não encontrei Teste 1 e Teste 2.");
                    return;
                }

                // 2️⃣ Criar coluna Média se não existir
                if (colMedia == -1)
                {
                    colMedia = colT2 + 1;
                    sheet.Cells[headerRow, colMedia].Value2 = "Média";
                }

                // 3️⃣ Preencher média SOMENTE entre Teste1 e Teste2
                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    string letraT1 = ColunaParaLetra(colT1);
                    string letraT2 = ColunaParaLetra(colT2);

                    // ⚠️ AQUI ESTÁ A CORREÇÃO: MÉDIA DE DOIS VALORES, NÃO INTERVALO
                    sheet.Cells[row, colMedia].FormulaLocal =
                        $"=MÉDIA({letraT1}{row};{letraT2}{row})";

                    row++;
                }

                workbook.Save();
                Console.WriteLine("📊 Média corrigida (apenas Teste 1 + Teste 2) calculada com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao calcular média: " + ex.Message);
            }
        }

        // =====================================================
        // INSERIR COLUNA SITUAÇÃO APÓS MÉDIA
        // =====================================================
        public static void InserirSituacao()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMedia = -1;

                // Encontrar coluna "Média"
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var valor = sheet.Cells[headerRow, c].Value;
                    if (valor != null && IgualIgnorandoAcentos(valor.ToString(), "media"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                if (colMedia == -1)
                {
                    Console.WriteLine("⚠ Calcule a média primeiro.");
                    return;
                }

                int colSituacao = colMedia + 1;
                sheet.Cells[headerRow, colSituacao] = "Situação";

                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    sheet.Cells[row, colSituacao].Value2 = "";
                    row++;
                }

                Console.WriteLine("📗 Coluna 'Situação' criada.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao inserir coluna situação: " + ex.Message);
            }
        }

       
        // =====================================================
        // DESTACAR APROVADOS
        // =====================================================
        public static void DestacarAprovados()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMedia = -1;
                int colSituacao = -1;

                // Encontrar colunas
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var valor = sheet.Cells[headerRow, c].Value;
                    if (valor == null) continue;

                    if (IgualIgnorandoAcentos(valor.ToString(), "media"))
                        colMedia = c;

                    if (IgualIgnorandoAcentos(valor.ToString(), "situacao"))
                        colSituacao = c;
                }

                if (colMedia == -1 || colSituacao == -1)
                {
                    Console.WriteLine("⚠ Calcule a média e adicione Situação primeiro.");
                    return;
                }

                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double media = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (media >= 10)
                    {
                        sheet.Cells[row, colSituacao].Value2 = "Aprovado";
                        sheet.Cells[row, colSituacao].Interior.Color =
                            ColorTranslator.ToOle(Color.LightGreen);
                    }
                    else
                    {
                        sheet.Cells[row, colSituacao].Value2 = "Reprovado";
                        sheet.Cells[row, colSituacao].Interior.Color =
                            ColorTranslator.ToOle(Color.LightCoral);
                    }

                    row++;
                }

                Console.WriteLine("🎨 Situação preenchida com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro: " + ex.Message);
            }
        }

        // =====================================================
        // IDENTIFICAR MELHORIA ≥20%
        // =====================================================
        public static void IdentificarMelhoria()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();

                int colT1 = headerCol + 6;   
                int colT2 = headerCol + 13;  

                int row = headerRow + 1;

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double t1 = sheet.Cells[row, colT1].Value ?? 0;
                    double t2 = sheet.Cells[row, colT2].Value ?? 0;

                    if (t1 > 0 && (t2 - t1) / t1 >= 0.2)
                        sheet.Rows[row].Interior.Color = ColorTranslator.ToOle(Color.Yellow);

                    row++;
                }

                Console.WriteLine("📈 Melhorias destacadas!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro: " + ex.Message);
            }
        }

        // =====================================================
        // GERAR GRÁFICO (VERSÃO FINAL — APENAS 2 TIPOS)
        // =====================================================
        public static void GerarGraficoTurma(dynamic json)
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colT1 = -1, colT2 = -1, colMedia = -1;

                // Encontrar colunas corretas
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "Teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "Teste 2")) colT2 = c;
                    if (IgualIgnorandoAcentos(titulo, "Média")) colMedia = c;
                }

                if (colT1 == -1 || colT2 == -1 || colMedia == -1)
                {
                    Console.WriteLine("❌ Não encontrei colunas de teste.");
                    return;
                }

                // Descobrir última linha
                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerCol].Value != null)
                    lastRow++;

                int count = lastRow - headerRow - 1;
                if (count <= 0)
                {
                    Console.WriteLine("❌ Não há alunos.");
                    return;
                }

                // Calcular médias reais
                double somaT1 = 0, somaT2 = 0, somaM = 0;

                for (int r = headerRow + 1; r < lastRow; r++)
                {
                    somaT1 += Convert.ToDouble(sheet.Cells[r, colT1].Value2 ?? 0);
                    somaT2 += Convert.ToDouble(sheet.Cells[r, colT2].Value2 ?? 0);
                    somaM += Convert.ToDouble(sheet.Cells[r, colMedia].Value2 ?? 0);
                }

                double mediaT1 = somaT1 / count;
                double mediaT2 = somaT2 / count;
                double mediaMF = somaM / count;

                // Criar gráfico
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + sheet.Rows[lastRow].Height + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 650, 380);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Médias da Turma — T1, T2 e Final";

                Excel.SeriesCollection sc = chart.SeriesCollection();

                Excel.Series s = sc.NewSeries();
                s.Name = "Média da Turma";
                s.Values = new double[] { mediaT1, mediaT2, mediaMF };
                s.XValues = new string[] { "Teste 1", "Teste 2", "Média Final" };

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                Console.WriteLine("📊 Gráfico de médias criado com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao criar gráfico: " + ex.Message);
            }
        }


        public static void GerarGraficoBarras(dynamic json)
        {
            try
            {
                // Agora o "aluno" vem como número mecanográfico
                string numeroMec = json.nlu.aluno != null ? json.nlu.aluno.ToString() : "";
                if (string.IsNullOrEmpty(numeroMec))
                {
                    Console.WriteLine("❌ Nenhum número mecanográfico encontrado.");
                    return;
                }

                Excel.Range used = sheet.UsedRange;

                // 1️⃣ Encontrar cabeçalho da coluna Nome (já existente)
                var (headerRowNome, colNome) = EncontrarCabecalho();

                // 2️⃣ Encontrar coluna do número mecanográfico
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colNumeroMec = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRowNome, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "Número Mecanográfico"))
                    {
                        colNumeroMec = c;
                        break;
                    }
                }

                if (colNumeroMec == -1)
                {
                    Console.WriteLine("❌ Coluna 'Número Mecanográfico' não encontrada.");
                    return;
                }

                // 3️⃣ Encontrar colunas Teste 1 e Teste 2
                int colT1 = -1;
                int colT2 = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRowNome, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "Teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "Teste 2")) colT2 = c;
                }

                if (colT1 == -1 || colT2 == -1)
                {
                    Console.WriteLine("❌ Não encontrei Teste 1 / Teste 2.");
                    return;
                }

                // 4️⃣ Descobrir última linha
                int lastRow = headerRowNome + 1;
                while (sheet.Cells[lastRow, colNome].Value != null)
                    lastRow++;

                // 5️⃣ Procurar o aluno pelo número mecanográfico
                int rowAluno = -1;
                for (int r = headerRowNome + 1; r < lastRow; r++)
                {
                    var valor = sheet.Cells[r, colNumeroMec].Value?.ToString().Trim();
                    if (valor != null && valor == numeroMec)
                    {
                        rowAluno = r;
                        break;
                    }
                }

                if (rowAluno == -1)
                {
                    Console.WriteLine($"❌ Número mecanográfico {numeroMec} não encontrado.");
                    return;
                }

                // 6️⃣ Obter o nome verdadeiro do aluno
                string nomeAluno = sheet.Cells[rowAluno, colNome].Value?.ToString() ?? "(Sem nome)";

                // 7️⃣ Criar gráfico
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 700, 380);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = $"Notas de {nomeAluno} (NMec {numeroMec})";

                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();

                Excel.Series s1 = sc.NewSeries();
                s1.Name = "Teste 1";
                s1.Values = sheet.Range[$"{ColunaParaLetra(colT1)}{rowAluno}"];
                s1.XValues = $"\"Teste 1\"";

                Excel.Series s2 = sc.NewSeries();
                s2.Name = "Teste 2";
                s2.Values = sheet.Range[$"{ColunaParaLetra(colT2)}{rowAluno}"];
                s2.XValues = $"\"Teste 2\"";

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                Console.WriteLine($"📊 Gráfico de barras criado para o aluno {nomeAluno} (NMec {numeroMec})!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao criar gráfico de barras: " + ex.Message);
            }
        }

        public static void AtualizarNotas(dynamic json)
        {
            try
            {
                // 1️⃣ ALUNO (número mecanográfico)
                string numeroMec = json.nlu.aluno != null ? json.nlu.aluno.ToString() : "";
                if (string.IsNullOrEmpty(numeroMec))
                {
                    Console.WriteLine("❌ Número mecanográfico não encontrado.");
                    return;
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
                    return;
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
                    return;
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
                    return;
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
                    return;
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
                    return;
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
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao atualizar notas: " + ex.Message);
            }
        }




        public static void ApagarGrafico(dynamic json)
        {
            try
            {
                string aluno = json.nlu.aluno != null ? json.nlu.aluno.ToString().ToLower() : "";

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                for (int i = charts.Count; i >= 1; i--)
                {
                    Excel.Chart chart = charts.Item(i).Chart;

                    if (!chart.HasTitle) continue;

                    string caption = (chart.ChartTitle.Caption ?? "").ToLower();
                    string titulo = (chart.ChartTitle.Text ?? "").ToLower();

                    // Apagar gráfico do aluno
                    if (!string.IsNullOrEmpty(aluno) &&
                        (caption.Contains($"#tag#aluno#{aluno}") || titulo.Contains(aluno)))
                    {
                        charts.Item(i).Delete();
                        Console.WriteLine($"🗑 Gráfico do aluno {aluno} apagado!");
                        return;
                    }

                    // Apagar gráfico das médias
                    if (caption.Contains("#tag#medias#") ||
                        titulo.Contains("médias") || titulo.Contains("media"))
                    {
                        charts.Item(i).Delete();
                        Console.WriteLine("🗑 Gráfico das médias apagado!");
                        return;
                    }
                }

                Console.WriteLine("⚠ Nenhum gráfico correspondente encontrado.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao apagar gráfico: " + ex.Message);
            }
        }


        public static void ApagarTodosGraficos()
        {
            try
            {
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                for (int i = charts.Count; i >= 1; i--)
                    charts.Item(i).Delete();

                Console.WriteLine("🧹 Todos os gráficos foram apagados!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao apagar todos os gráficos: " + ex.Message);
            }
        }

        public static void OperacoesMatematicas(dynamic json)
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
                    return;
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

                Console.WriteLine($"📊 Estatísticas: {aprovados} aprovados, {reprovados} reprovados, {acimaDe16} acima de 16 valores.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro em Operações Matemáticas: " + ex.Message);
            }
        }




        // =====================================================
        // GUARDAR RELATÓRIO
        // =====================================================
        public static void GuardarRelatorio()
        {
            try
            {
                workbook.SaveAs(pathFinal);
                Console.WriteLine("💾 Relatório guardado!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao guardar relatório: " + ex.Message);
            }
        }
    }
}
