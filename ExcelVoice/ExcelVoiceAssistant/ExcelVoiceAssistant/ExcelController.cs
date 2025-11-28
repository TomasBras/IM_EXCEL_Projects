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

        private static string pathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM_EXCEL_Projects\ExcelVoice\IM_Excel\ETP.xlsx";
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

            // Procurar folha correta automaticamente
            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                if (sh.Cells[1, 1].Value?.ToString() == "Número mecanográfico")
                {
                    sheet = sh;
                    return;
                }
            }

            // fallback: usa a enviada
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

                    if (valor == null) continue;

                    string texto = valor.ToString();

                    // 🔥 Normalizador universal
                    string clean = new string(
                        texto.Normalize(NormalizationForm.FormD)
                        .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                        .ToArray()
                    )
                    .Replace("\u00A0", " ") // remove non-breaking space
                    .Replace("\t", " ")     // remove tabs
                    .Replace("  ", " ")
                    .Trim()
                    .ToLower();

                    if (clean == "nome")
                        return (r, c);
                }
            }

            throw new Exception("Cabeçalho Nome não encontrado (mesmo após limpeza).");
        }


        public static string CalcularMedia(dynamic json)
        {
            try
            {
                // o intent vem sempre em nlu.intent
                string intent = json.nlu.intent.ToString();

                // ENTIDADES VÊM COMO CAMPOS DIRETOS (não dentro de "entities")
                string nome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : null;
                string numero = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : null;

                // 🎯 Se o nome existir → calcula só para esse aluno
                if (!string.IsNullOrEmpty(nome))
                    return CalcularMediaAluno(nome);

                // 🎯 Se houver número mecanográfico
                if (!string.IsNullOrEmpty(numero))
                    return CalcularMediaAlunoNumero(numero);

                // Caso contrário → média da turma
                return CalcularMediaTurma();
            }
            catch
            {
                return CalcularMediaTurma();
            }
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
            catch (Exception ex)
            {
                Console.WriteLine("❌ ERRO CalcularMediaTurma: " + ex.Message);
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

        public static string MelhoriaReal()
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // Encontrar colunas de testes automaticamente
                List<(int col, int num)> testes = new List<(int col, int num)>();

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    var match = System.Text.RegularExpressions.Regex.Match(
                        titulo.ToLower().Replace(" ", ""),
                        @"teste(\d+)"
                    );

                    if (match.Success)
                        testes.Add((c, int.Parse(match.Groups[1].Value)));
                }

                if (testes.Count < 2)
                    return "São necessários pelo menos dois testes para calcular melhoria.";

                testes = testes.OrderBy(t => t.num).ToList();

                int colPenultimo = testes[testes.Count - 2].col;
                int colUltimo = testes[testes.Count - 1].col;

                // Criar coluna Melhoria Real se não existir
                int colMelhoria = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var val = sheet.Cells[headerRow, c].Value?.ToString();
                    if (val != null && IgualIgnorandoAcentos(val, "melhoria real"))
                    {
                        colMelhoria = c;
                        break;
                    }
                }

                if (colMelhoria == -1)
                {
                    colMelhoria = lastCol + 1;
                    sheet.Cells[headerRow, colMelhoria].Value2 = "Melhoria Real";
                    lastCol++;
                }

                // Processar aluno a aluno
                int row = headerRow + 1;

                while (sheet.Cells[row, headerColNome].Value != null)
                {
                    double penultimo = Convert.ToDouble(sheet.Cells[row, colPenultimo].Value2 ?? 0);
                    double ultimo = Convert.ToDouble(sheet.Cells[row, colUltimo].Value2 ?? 0);

                    double diferenca = ultimo - penultimo;
                    double percent = (penultimo != 0)
                        ? (diferenca / penultimo) * 100
                        : (diferenca > 0 ? 100 : 0);

                    // Mensagem
                    string texto;

                    if (diferenca > 0)
                    {
                        texto = $"Melhorou (+{Math.Round(diferenca, 2)} valores, +{Math.Round(percent, 1)}%)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

                        // Verde
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else if (diferenca < 0)
                    {
                        texto = $"Piorou ({Math.Round(diferenca, 2)} valores, {Math.Round(percent, 1)}%)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

                        // Vermelho
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);
                    }
                    else
                    {
                        texto = $"Igual (0)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

                        // Cinza
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    }

                    row++;
                }

                return "Melhoria Real atualizada com detalhes e cores.";
            }
            catch (Exception ex)
            {
                return "Erro em Melhoria Real: " + ex.Message;
            }
        }

        public static string MelhoriaPossivel()
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // Encontrar colunas dos testes dinamicamente
                List<int> colTestes = new List<int>();

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    var m = System.Text.RegularExpressions.Regex.Match(
                        titulo.ToLower().Replace(" ", ""), @"teste(\d+)"
                    );

                    if (m.Success)
                        colTestes.Add(c);
                }

                if (colTestes.Count < 2)
                    return "São necessários pelo menos dois testes para calcular MP.";

                colTestes = colTestes.OrderBy(c => c).ToList();
                int colUltimoTeste = colTestes.Last();

                // Encontrar coluna média
                int colMedia = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo != null && IgualIgnorandoAcentos(titulo, "média"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                if (colMedia == -1)
                    return "Calcule a média antes de verificar MP.";

                // Criar coluna MP se não existir
                int colMP = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo != null && IgualIgnorandoAcentos(titulo, "mp"))
                    {
                        colMP = c;
                        break;
                    }
                }

                if (colMP == -1)
                {
                    colMP = lastCol + 1;
                    sheet.Cells[headerRow, colMP].Value2 = "MP";
                    lastCol++;
                }

                // Criar coluna Nota Necessária se não existir
                int colNotaNecessaria = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo != null && IgualIgnorandoAcentos(titulo, "nota necessária"))
                    {
                        colNotaNecessaria = c;
                        break;
                    }
                }

                if (colNotaNecessaria == -1)
                {
                    colNotaNecessaria = lastCol + 1;
                    sheet.Cells[headerRow, colNotaNecessaria].Value2 = "Nota Necessária";
                    lastCol++;
                }

                // Calcular MP aluno a aluno
                int row = headerRow + 1;

                while (sheet.Cells[row, headerColNome].Value != null)
                {
                    double mediaAtual = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (mediaAtual >= 10)
                    {
                        sheet.Cells[row, colMP].Value2 = "";
                        sheet.Cells[row, colNotaNecessaria].Value2 = "—";
                        row++;
                        continue;
                    }

                    // Calcular soma dos testes exceto o último
                    double soma = 0;
                    foreach (int col in colTestes.Take(colTestes.Count - 1))
                        soma += Convert.ToDouble(sheet.Cells[row, col].Value2 ?? 0);

                    int n = colTestes.Count;

                    // Nota necessária para média >= 10
                    double notaNecessaria = 10 * n - soma;

                    // Preencher coluna Nota Necessária
                    if (notaNecessaria > 20)
                        sheet.Cells[row, colNotaNecessaria].Value2 = ">20";
                    else
                        sheet.Cells[row, colNotaNecessaria].Value2 = Math.Round(notaNecessaria, 2);

                    // Preencher MP
                    if (notaNecessaria <= 20)
                        sheet.Cells[row, colMP].Value2 = "MP";
                    else
                        sheet.Cells[row, colMP].Value2 = "";

                    row++;
                }

                return "Melhoria possível e nota necessária calculadas dinamicamente.";
            }
            catch (Exception ex)
            {
                return "Erro em Melhoria Possível: " + ex.Message;
            }
        }

        public static string InserirPerguntas(dynamic json)
        {
            try
            {
                // 1) Obter número do teste
                int testeNum = -1;

                if (json.nlu.teste_numero != null)
                {
                    string raw = json.nlu.teste_numero.ToString();
                    var m = Regex.Match(raw, @"(\d+)");
                    if (m.Success)
                        testeNum = int.Parse(m.Groups[1].Value);
                }

                if (testeNum == -1)
                    return "Não percebi qual é o teste.";

                string prefixo = $"T{testeNum}_P";

                // 2) Texto do utilizador
                string texto = json.text != null
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower()
                    : "";

                // 3) Procurar intervalo tipo “1 a 5”
                int pInicio = -1, pFim = -1;
                var intervalo = Regex.Match(texto, @"(\d+)\s*(a|à|até|-)\s*(\d+)");
                if (intervalo.Success)
                {
                    pInicio = int.Parse(intervalo.Groups[1].Value);
                    pFim = int.Parse(intervalo.Groups[3].Value);
                }

                // 4) Procurar pergunta única
                var unico = Regex.Match(texto, @"(p|pergunta|questao|questão|q)\s*(número\s*)?(\d+)");
                if (unico.Success)
                {
                    int p = int.Parse(unico.Groups[3].Value);
                    pInicio = pFim = p;
                }

                // 5) Se nada foi encontrado → ERRO (não default!)
                if (pInicio == -1)
                    return "Não percebi qual pergunta queres adicionar.";

                // 6) Cabeçalho
                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // 7) Encontrar posição da coluna “Teste N”
                int colTeste = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo != null && IgualIgnorandoAcentos(titulo, $"teste {testeNum}"))
                    {
                        colTeste = c;
                        break;
                    }
                }

                if (colTeste == -1)
                    return $"Não encontrei o Teste {testeNum}.";

                // 8) Mapear perguntas existentes ANTES do teste
                Dictionary<int, int> existentes = new Dictionary<int, int>();

                for (int c = firstCol; c < colTeste; c++)
                {
                    string t = sheet.Cells[headerRow, c].Value?.ToString();
                    if (t == null) continue;

                    string norm = t.Replace(" ", "").ToUpper();

                    if (norm.StartsWith(prefixo.ToUpper()))
                    {
                        var mm = Regex.Match(norm, @"P(\d+)");
                        if (mm.Success)
                        {
                            int per = int.Parse(mm.Groups[1].Value);
                            existentes[per] = c;
                        }
                    }
                }

                // 9) Inserir NOVAS perguntas
                int adicionadas = 0;

                for (int p = pInicio; p <= pFim; p++)
                {
                    if (!existentes.ContainsKey(p))
                    {
                        // Inserir nova coluna ANTES do teste
                        sheet.Columns[colTeste].Insert();

                        sheet.Cells[headerRow, colTeste].Value2 = $"{prefixo}{p}";

                        int r = headerRow + 1;
                        while (sheet.Cells[r, headerCol].Value != null)
                        {
                            sheet.Cells[r, colTeste].Value2 = "";
                            r++;
                        }

                        adicionadas++;

                        // mover o teste uma coluna para a direita
                        colTeste++;
                        lastCol++;
                    }
                }

                if (adicionadas == 0)
                    return $"As perguntas pedidas já existiam no Teste {testeNum}.";

                return $"Foram adicionadas {adicionadas} perguntas ao Teste {testeNum}.";
            }
            catch (Exception ex)
            {
                return "Erro ao inserir perguntas: " + ex.Message;
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
                // 📌 1) Obter aluno_nome e aluno_numero do JSON
                string numeroMec = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : "";
                string alunoNome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : "";

                if (string.IsNullOrEmpty(numeroMec) && string.IsNullOrEmpty(alunoNome))
                {
                    Console.WriteLine("❌ Não foi indicado nome nem número do aluno.");
                    return "Não foi indicado nome nem número do aluno.";
                }

                Excel.Range used = sheet.UsedRange;

                // 📌 2) Cabeçalho e coluna Nome
                var (headerRow, colNome) = EncontrarCabecalho();

                // 📌 3) Encontrar coluna "Número Mecanográfico"
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colNumeroMec = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRow, c].Value?.ToString();
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
                    return "Coluna 'Número Mecanográfico' não encontrada.";
                }

                // 📌 4) Encontrar colunas do Teste 1 e Teste 2
                int colT1 = -1, colT2 = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "Teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "Teste 2")) colT2 = c;
                }

                if (colT1 == -1 || colT2 == -1)
                {
                    Console.WriteLine("❌ Não encontrei Teste 1 / Teste 2.");
                    return "Não encontrei Teste 1 / Teste 2.";
                }

                // 📌 5) Encontrar última linha
                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, colNome].Value != null)
                    lastRow++;

                // 📌 6) Procurar aluno
                int rowAluno = -1;

                // 🔍 6A — Procurar pelo número MEC
                if (!string.IsNullOrEmpty(numeroMec))
                {
                    for (int r = headerRow + 1; r < lastRow; r++)
                    {
                        var valor = sheet.Cells[r, colNumeroMec].Value?.ToString().Trim();

                        if (valor != null && valor == numeroMec)
                        {
                            rowAluno = r;
                            break;
                        }
                    }
                }

                // 🔍 6B — Procurar pelo NOME (caso não tenha encontrado pelo número)
                if (rowAluno == -1 && !string.IsNullOrEmpty(alunoNome))
                {
                    string[] partes = alunoNome.ToLower().Split(' ');

                    for (int r = headerRow + 1; r < lastRow; r++)
                    {
                        string excelNome = sheet.Cells[r, colNome].Value?.ToString().ToLower() ?? "";

                        bool match = partes.All(p => excelNome.Contains(p));
                        if (match)
                        {
                            rowAluno = r;
                            break;
                        }
                    }
                }

                // 📌 Falha total
                if (rowAluno == -1)
                {
                    Console.WriteLine($"❌ Aluno não encontrado: {alunoNome} / {numeroMec}");
                    return $"Aluno não encontrado: {alunoNome} / {numeroMec}";
                }

                // 📌 7) Nome verdadeiro para o título
                string nomeFinal = sheet.Cells[rowAluno, colNome].Value?.ToString() ?? "(Sem nome)";
                string textoNumero = string.IsNullOrEmpty(numeroMec) ? "" : $" (NMec {numeroMec})";

                // 📌 8) Criar gráfico
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 700, 380);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = $"Notas de {nomeFinal}{textoNumero}";

                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();

                Excel.Series s1 = sc.NewSeries();
                s1.Name = "Teste 1";
                s1.Values = sheet.Range[$"{ColunaParaLetra(colT1)}{rowAluno}"];
                s1.XValues = "\"Teste 1\"";

                Excel.Series s2 = sc.NewSeries();
                s2.Name = "Teste 2";
                s2.Values = sheet.Range[$"{ColunaParaLetra(colT2)}{rowAluno}"];
                s2.XValues = "\"Teste 2\"";

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                Console.WriteLine($"📊 Gráfico de barras criado para o aluno {nomeFinal}{textoNumero}!");
                return $"Gráfico de barras criado para o aluno {nomeFinal}{textoNumero}!";
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao criar gráfico de barras: " + ex.Message);
                return "Erro ao criar gráfico de barras.";
            }
        }


        public static string GerarGraficoPerguntasT2()
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // Encontrar colunas T2_P1 ... T2_P5
                Dictionary<string, int> perguntas = new Dictionary<string, int>();

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (titulo.Trim().StartsWith("T2_P"))
                    {
                        perguntas[titulo.Trim()] = c;
                    }
                }

                if (perguntas.Count == 0)
                {
                    Console.WriteLine("❌ Nenhuma coluna T2_P encontrada.");
                    return "Nenhuma coluna T2_P encontrada.";
                }

                // Ordenar T2_P1, T2_P2, ...
                var ordenadas = perguntas.OrderBy(k => k.Key).ToList();

                // Descobrir última linha com alunos
                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerColNome].Value != null)
                    lastRow++;

                int totalAlunos = lastRow - headerRow - 1;
                if (totalAlunos <= 0)
                {
                    Console.WriteLine("❌ Nenhum aluno encontrado.");
                    return "Nenhum aluno encontrado.";
                }

                // Calcular média de cada pergunta
                List<double> medias = new List<double>();

                foreach (var kv in ordenadas)
                {
                    double soma = 0;
                    for (int r = headerRow + 1; r < lastRow; r++)
                    {
                        soma += Convert.ToDouble(sheet.Cells[r, kv.Value].Value2 ?? 0);
                    }

                    medias.Add(soma / totalAlunos);
                }

                // Criar gráfico
                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 700, 400);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Médias das Perguntas do Teste 2 (T2_P1 a T2_P5)";

                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                Excel.Series s = sc.NewSeries();

                s.Name = "Média";
                s.Values = medias.ToArray();
                s.XValues = ordenadas.Select(k => k.Key).ToArray();

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                Console.WriteLine("📊 Gráfico das médias das perguntas do Teste 2 criado com sucesso!");
                return "Gráfico das perguntas do teste 2 criado com sucesso.";

            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao gerar gráfico das perguntas: " + ex.Message);
                return "Erro ao gerar gráfico das perguntas.";
            }
        }
        public static string AtualizarNotas(dynamic json)
        {
            try
            {
                // ==========================================================
                // 1) DECODIFICAR TEXTO ORIGINAL
                // ==========================================================
                string textoOriginal = json.text != null
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower()
                    : "";

                // ==========================================================
                // 2) ENTIDADES: aluno (nome/numero), teste, pergunta
                // ==========================================================
                string numeroMec = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : null;
                string alunoNome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : null;

                // TESTE
                int testeNum = -1;
                Match matchTeste = Regex.Match(textoOriginal, @"teste ?([0-9]{1,2})");
                if (matchTeste.Success)
                    testeNum = int.Parse(matchTeste.Groups[1].Value);

                // PERGUNTA
                int perguntaNum = -1;
                Match matchPerg = Regex.Match(textoOriginal, @"(pergunta|quest[aã]o) ?([0-9]{1,2})");
                if (matchPerg.Success)
                    perguntaNum = int.Parse(matchPerg.Groups[2].Value);

                // REGRA: PERGUNTA sem TESTE → erro
                if (perguntaNum != -1 && testeNum == -1)
                    return "Tens de indicar o número do teste. Ex.: 'pergunta 2 do teste 1'.";


                // ==========================================================
                // 2B) EXTRAIR VALORES -> APENAS APÓS "COM" ou "PARA"
                // ==========================================================
                List<double> valores = new List<double>();

                Match matchValores = Regex.Match(textoOriginal, @"(?:com|para)\s+([0-9.,\s]+)");
                if (matchValores.Success)
                {
                    string bloco = matchValores.Groups[1].Value;
                    string[] parts = bloco.Split(new char[] { ' ', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string p in parts)
                    {
                        double v;
                        if (double.TryParse(p.Replace(",", "."), NumberStyles.Any,
                            CultureInfo.InvariantCulture, out v))
                        {
                            valores.Add(v);
                        }
                    }
                }

                // ==========================================================
                // 3) CABEÇALHO E COLUNAS
                // ==========================================================
                var header = EncontrarCabecalho();
                int headerRow = header.Item1;
                int colNome = header.Item2;

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // Coluna número mecanográfico
                int colNMec = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string t = sheet.Cells[headerRow, c].Value?.ToString();
                    if (t != null && IgualIgnorandoAcentos(t, "número mecanográfico"))
                    {
                        colNMec = c;
                        break;
                    }
                }
                if (colNMec == -1)
                    return "Coluna 'Número Mecanográfico' não encontrada.";

                // Última linha
                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, colNome].Value != null)
                    lastRow++;

                // ==========================================================
                // 4) ENCONTRAR ALUNO
                // ==========================================================
                int alunoRow = -1;

                for (int r = headerRow + 1; r < lastRow; r++)
                {
                    object nm = sheet.Cells[r, colNMec].Value;

                    // Por número
                    if (numeroMec != null && nm != null && nm.ToString() == numeroMec)
                    {
                        alunoRow = r;
                        break;
                    }

                    // Por nome
                    if (alunoNome != null)
                    {
                        string excelNome = (sheet.Cells[r, colNome].Value ?? "").ToString().ToLower();
                        string[] partes = alunoNome.ToLower().Split(' ');

                        bool matchAll = true;
                        foreach (string p in partes)
                            if (!excelNome.Contains(p)) matchAll = false;

                        if (matchAll)
                        {
                            alunoRow = r;
                            break;
                        }
                    }
                }

                // Operação turma -> apenas se explicitamente pedido
                bool operacaoTurma =
                    alunoRow == -1 &&
                    (textoOriginal.Contains("toda a turma") || textoOriginal.Contains("todos os alunos"));


                // ==========================================================
                // 5) MAPEAR PERGUNTAS (Tn_Px)
                // ==========================================================
                if (testeNum == -1)
                    return "Tens de indicar o número do teste.";

                string prefixo = "T" + testeNum + "_P";

                Dictionary<int, int> colsPerguntas = new Dictionary<int, int>();
                int colTesteFinal = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;

                    if (IgualIgnorandoAcentos(titulo, "teste " + testeNum))
                        colTesteFinal = c;

                    string norm = titulo.Replace(" ", "").ToUpper();

                    if (norm.StartsWith(prefixo.ToUpper()))
                    {
                        Match m = Regex.Match(norm, @"P(\d+)");
                        if (m.Success)
                        {
                            colsPerguntas[int.Parse(m.Groups[1].Value)] = c;
                        }
                    }
                }

                if (colsPerguntas.Count == 0)
                    return "Nenhuma pergunta encontrada no teste " + testeNum + ".";


                // ==========================================================
                // 6) TIPOS DE OPERAÇÕES
                // ==========================================================
                bool pedirZero = textoOriginal.Contains(" zero");
                bool pedirRandom = textoOriginal.Contains("random") || textoOriginal.Contains("aleat");
                bool pedirCotacaoMax = textoOriginal.Contains("cotação máxima") || textoOriginal.Contains("nota máxima");
                bool apenasVazias = textoOriginal.Contains("vazia");

                Random rnd = new Random();


                // ==========================================================
                // 7) APLICAR OPERAÇÃO A UM ALUNO
                // ==========================================================
                Action<int> AplicarOperacao = delegate (int r)
                {
                    // ZERO
                    if (pedirZero)
                    {
                        foreach (int col in colsPerguntas.Values)
                            sheet.Cells[r, col].Value2 = 0;
                    }

                    // RANDOM
                    else if (pedirRandom)
                    {
                        foreach (int col in colsPerguntas.Values)
                        {
                            if (apenasVazias &&
                                sheet.Cells[r, col].Value2 != null &&
                                sheet.Cells[r, col].Value2.ToString() != "")
                                continue;

                            double randomNota;
                            if (rnd.Next(2) == 0)
                                randomNota = rnd.Next(0, 21);    // inteiro
                            else
                                randomNota = Math.Round(rnd.NextDouble() * 20, 1);

                            sheet.Cells[r, col].Value2 = randomNota;
                        }
                    }

                    // COTAÇÃO MÁXIMA – alterar APENAS uma pergunta
                    else if (pedirCotacaoMax && perguntaNum != -1)
                    {
                        if (colsPerguntas.ContainsKey(perguntaNum))
                            sheet.Cells[r, colsPerguntas[perguntaNum]].Value2 = 20.0;
                    }

                    // PERGUNTA INDIVIDUAL
                    else if (perguntaNum != -1 && valores.Count >= 1)
                    {
                        if (colsPerguntas.ContainsKey(perguntaNum))
                            sheet.Cells[r, colsPerguntas[perguntaNum]].Value2 = valores[0];
                    }

                    // LISTA DE PERGUNTAS (ex: 1 2 3 4 5)
                    else if (valores.Count > 1)
                    {
                        List<KeyValuePair<int, int>> ord =
                            colsPerguntas.OrderBy(k => k.Key).ToList();

                        for (int i = 0; i < valores.Count && i < ord.Count; i++)
                            sheet.Cells[r, ord[i].Value].Value2 = valores[i];
                    }

                    // ======================================================
                    // NORMALIZAÇÃO 0–20 → peso
                    // ======================================================
                    double peso = 20.0 / colsPerguntas.Count;
                    double soma = 0;

                    foreach (int col in colsPerguntas.Values)
                    {
                        double bruto = 0;
                        object valObj = sheet.Cells[r, col].Value2;

                        if (valObj != null)
                            bruto = Convert.ToDouble(valObj);

                        double normalizado = (bruto / 20.0) * peso;

                        sheet.Cells[r, col].Value2 = normalizado;
                        soma += normalizado;
                    }

                    // Teste final
                    if (colTesteFinal != -1)
                        sheet.Cells[r, colTesteFinal].Value2 = soma;
                };


                // ==========================================================
                // 8) EXECUTAR PARA 1 ALUNO OU PARA A TURMA
                // ==========================================================
                if (operacaoTurma)
                {
                    for (int r = headerRow + 1; r < lastRow; r++)
                        AplicarOperacao(r);
                }
                else
                {
                    AplicarOperacao(alunoRow);
                }


                // ==========================================================
                // 9) REFAZER MÉDIAS
                // ==========================================================
                int colMedia = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string t = sheet.Cells[headerRow, c].Value?.ToString();
                    if (t != null && IgualIgnorandoAcentos(t, "média"))
                        colMedia = c;
                }

                if (colMedia != -1)
                {
                    List<int> colTestes = new List<int>();

                    for (int c = firstCol; c <= lastCol; c++)
                    {
                        string t = sheet.Cells[headerRow, c].Value?.ToString();
                        if (t != null && t.ToLower().StartsWith("teste"))
                            colTestes.Add(c);
                    }

                    for (int r = headerRow + 1; r < lastRow; r++)
                    {
                        List<string> refs = new List<string>();
                        foreach (int c in colTestes)
                            refs.Add(ColunaParaLetra(c) + r);

                        string formula = "=MÉDIA(" + string.Join(";", refs.ToArray()) + ")";
                        sheet.Cells[r, colMedia].FormulaLocal = formula;
                    }
                }

                workbook.Save();
                return operacaoTurma ? "Notas atualizadas para toda a turma!" : "Notas atualizadas!";
            }
            catch (Exception ex)
            {
                return "Erro ao atualizar notas: " + ex.Message;
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
                // Texto original decodificado do Base64
                string texto = json.text != null
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower()
                    : "";

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
                    return "É necessário calcular a média primeiro.";

                // Ler todas as médias
                int row = headerRow + 1;
                List<double> medias = new List<double>();

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double media = sheet.Cells[row, colMedia].Value2 ?? 0;
                    medias.Add(media);
                    row++;
                }

                int total = medias.Count;
                if (total == 0) return "Nenhum aluno encontrado.";

                // Estatísticas
                int aprovados = medias.Count(m => m >= 10);
                int reprovados = medias.Count(m => m < 10);
                int acima16 = medias.Count(m => m >= 16);
                int acima18 = medias.Count(m => m >= 18);
                double mediaGeral = medias.Average();
                double melhor = medias.Max();
                double pior = medias.Min();
                double mediana = medias.OrderBy(v => v).ToList()[total / 2];
                double desvio = Math.Sqrt(medias.Sum(v => Math.Pow(v - mediaGeral, 2)) / total);
                double percAprov = (double)aprovados / total * 100;

                // --------------------------------------------------------
                // 🔍 DETEÇÃO: É pedido geral?
                // --------------------------------------------------------
                bool pedidoGeral =
                    texto.Contains("estatistic") ||
                    texto.Contains("resumo") ||
                    texto.Contains("tabela") ||
                    texto.Contains("relatório") ||
                    texto.Contains("estatísticas gerais");

                // --------------------------------------------------------
                // 📌 CASO 1: PEDIDOS ESPECÍFICOS → escrever 1 linha no Excel
                // --------------------------------------------------------
                if (!pedidoGeral)
                {
                    int writeRow = headerRow + total + 3;
                    int col = headerCol;

                    string titulo = "";
                    string valor = "";

                    if (texto.Contains("aprovad"))
                    {
                        titulo = "Aprovados";
                        valor = aprovados.ToString();
                    }
                    else if (texto.Contains("reprovad"))
                    {
                        titulo = "Reprovados";
                        valor = reprovados.ToString();
                    }
                    else if (texto.Contains("acima de 16") || texto.Contains("superior a 16"))
                    {
                        titulo = "Média ≥ 16";
                        valor = acima16.ToString();
                    }
                    else if (texto.Contains("acima de 18") || texto.Contains("superior a 18"))
                    {
                        titulo = "Média ≥ 18";
                        valor = acima18.ToString();
                    }
                    else if (texto.Contains("percentagem") || texto.Contains("aprovação"))
                    {
                        titulo = "Percentagem aprovação";
                        valor = $"{percAprov:0.0}%";
                    }
                    else if (texto.Contains("média geral") || texto.Contains("media geral"))
                    {
                        titulo = "Média geral";
                        valor = $"{mediaGeral:0.00}";
                    }
                    else if (texto.Contains("soma das médias"))
                    {
                        titulo = "Soma das médias";
                        valor = $"{medias.Sum():0.00}";
                    }
                    else
                    {
                        return "Não consegui interpretar a pergunta.";
                    }

                    // Escrever no Excel
                    sheet.Cells[writeRow, col].Value2 = titulo;
                    sheet.Cells[writeRow, col + 1].Value2 = valor;

                    Excel.Range r = sheet.Range[
                        sheet.Cells[writeRow, col],
                        sheet.Cells[writeRow, col + 1]
                    ];
                    r.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    r.Columns.AutoFit();

                    return $"{titulo}: {valor}";
                }


                // --------------------------------------------------------
                // 📌 CASO 2: ESTATÍSTICAS GERAIS → criar tabela completa
                // --------------------------------------------------------
                int startTableRow = headerRow + total + 3;
                int baseCol = headerCol;

                sheet.Cells[startTableRow, baseCol].Value2 = "ESTATÍSTICAS GERAIS DA TURMA";
                sheet.Cells[startTableRow, baseCol].Font.Bold = true;

                int r2 = startTableRow + 1;

                void Linha(string nome, object val)
                {
                    sheet.Cells[r2, baseCol].Value2 = nome;
                    sheet.Cells[r2, baseCol + 1].Value2 = val;
                    r2++;
                }

                Linha("Total de alunos", total);
                Linha("Aprovados", aprovados);
                Linha("Reprovados", reprovados);
                Linha("Percentagem de aprovação", $"{percAprov:0.0}%");
                Linha("Média geral", $"{mediaGeral:0.00}");
                Linha("Melhor nota", $"{melhor:0.00}");
                Linha("Pior nota", $"{pior:0.00}");
                Linha("Mediana", $"{mediana:0.00}");
                Linha("Desvio padrão", $"{desvio:0.00}");
                Linha("Notas ≥ 16", acima16);
                Linha("Notas ≥ 18", acima18);

                Excel.Range range = sheet.Range[
                    sheet.Cells[startTableRow, baseCol],
                    sheet.Cells[r2 - 1, baseCol + 1]
                ];

                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                range.Columns.AutoFit();

                return "Tabela de estatísticas gerais criada no Excel!";
            }
            catch (Exception ex)
            {
                return "Erro em Operações Matemáticas: " + ex.Message;
            }
        }

        public static void ImprimirCabecalhosComUnicode()
        {
            Excel.Range used = sheet.UsedRange;
            int headerRow = used.Row;   // normalmente é 1

            Console.WriteLine("=== DEBUG: A imprimir cabeçalhos ===");
            ImprimirCabecalhosComUnicode();

            for (int c = used.Column; c < used.Column + used.Columns.Count; c++)
            {
                var valor = sheet.Cells[headerRow, c].Value;

                if (valor == null)
                {
                    Console.WriteLine($"{ColunaParaLetra(c)}: (vazio)");
                    continue;
                }

                string texto = valor.ToString();
                Console.WriteLine($"{ColunaParaLetra(c)}: \"{texto}\"  (len={texto.Length})");

                // imprimir cada carácter com o seu código Unicode
                for (int i = 0; i < texto.Length; i++)
                {
                    char ch = texto[i];
                    Console.WriteLine($"   [{i}] '{ch}'  U+{((int)ch).ToString("X4")}");
                }

                Console.WriteLine();
            }

            Console.WriteLine("=================================");
        }

        private static bool ColunaEhNumerica(string colunaNome)
        {
            string[] numericFields =
            {
        "Média", "Teste 1", "Teste 2", "Teste 3",
        "Nota Necessária", "Melhoria Real"
    };

            return numericFields.Contains(colunaNome, StringComparer.OrdinalIgnoreCase);
        }

        public static void DebugCabecalhos()
        {
            Excel.Range used = sheet.UsedRange;
            int headerRow = 1;

            Console.WriteLine("=== CABEÇALHOS ENCONTRADOS ===");

            for (int c = 1; c <= used.Columns.Count; c++)
            {
                var v = sheet.Cells[headerRow, c].Value?.ToString() ?? "(vazio)";

                Console.Write($"{c}: \"{v}\"   |   ");

                // mostrar cada caracter
                foreach (char ch in v)
                    Console.Write($"[{ch} U+{((int)ch).ToString("X4")}] ");

                Console.WriteLine();
            }

            Console.WriteLine("===============================");
        }

        public static string CriarPivotTable(dynamic json)
        {
            try
            {
                Excel.Range used = sheet.UsedRange;

                int firstRow = used.Row;
                int lastRow = used.Row + used.Rows.Count - 1;
                int firstCol = used.Column;
                int lastCol = used.Column + used.Columns.Count - 1;

                Excel.Range dataRange =
                    sheet.Range[sheet.Cells[firstRow, firstCol], sheet.Cells[lastRow, lastCol]];

                // Criar folha para Pivot
                Excel.Worksheet pivotSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                pivotSheet.Name = "Pivot_" + DateTime.Now.Ticks;

                Excel.PivotCache cache = workbook.PivotCaches().Create(
                    Excel.XlPivotTableSourceType.xlDatabase,
                    dataRange
                );

                Excel.PivotTable pivot = cache.CreatePivotTable(
                    pivotSheet.Cells[1, 1],
                    "TabelaDinamica"
                );

                // Ler campos enviados pelo Rasa
                string rowField = json?.nlu?.coluna_excel_row?.ToString();
                string valueField = json?.nlu?.coluna_excel_value?.ToString();
                string filterRegime = json?.nlu?.regime?.ToString();

                bool comandoBasico = (rowField == null && valueField == null);

                // Mapa RASA → cabeçalhos Excel
                Dictionary<string, string> map = new Dictionary<string, string>
        {
            { "regime", "REGIME" },
            { "média", "Média" },
            { "media", "Média" },
            { "teste 1", "Teste 1" },
            { "teste 2", "Teste 2" },
            { "nome", "Nome" },
            { "numero mecanografico", "Número mecanográfico" }
        };

                string Resolve(string key)
                {
                    if (key == null) return null;
                    key = key.ToLower().Trim();
                    return map.ContainsKey(key) ? map[key] : null;
                }

                rowField = Resolve(rowField);
                valueField = Resolve(valueField);

                // Caso básico: pivot geral
                if (comandoBasico)
                {
                    Excel.PivotField pfNome = pivot.PivotFields("Nome");
                    pfNome.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    Excel.PivotField pfRegime = pivot.PivotFields("REGIME");
                    pfRegime.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    // Valores = média por defeito
                    Excel.PivotField pf = pivot.PivotFields("Média");
                    pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    pf.Function = Excel.XlConsolidationFunction.xlAverage;
                    pf.Name = "Média";

                    return "Tabela dinâmica criada com campos padrão.";
                }

                // 1) RowField → OK sempre
                if (rowField != null)
                {
                    Excel.PivotField row = pivot.PivotFields(rowField);
                    row.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }

                // 2) ValueField → tem de ser numérico
                if (valueField != null)
                {
                    if (!ColunaEhNumerica(valueField))
                    {
                        // ⚠️ Campo não-numérico → mover para linhas automaticamente
                        Excel.PivotField pf = pivot.PivotFields(valueField);
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                        return $"O campo '{valueField}' não é numérico e foi movido automaticamente para as linhas.";
                    }
                    else
                    {
                        Excel.PivotField pf = pivot.PivotFields(valueField);
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                        pf.Function = Excel.XlConsolidationFunction.xlAverage;
                        pf.Name = "Média de " + valueField;
                    }
                }

                // 3) Filtrar por regime se pedido
                if (!string.IsNullOrEmpty(filterRegime))
                {
                    Excel.PivotField filtro = pivot.PivotFields("REGIME");
                    filtro.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

                    app.Calculate();

                    foreach (Excel.PivotItem item in filtro.PivotItems())
                    {
                        if (item.Name.Equals(filterRegime, StringComparison.OrdinalIgnoreCase))
                        {
                            filtro.CurrentPage = filterRegime;
                            return "Tabela dinâmica criada com filtro aplicado.";
                        }
                    }

                    filtro.ClearAllFilters();
                }

                return "Tabela dinâmica criada com sucesso!";
            }
            catch (Exception ex)
            {
                return "Erro ao criar tabela dinâmica: " + ex.Message;
            }
        }


        public static string Helper()
        {
            return "Pode pedir para calcular médias, destacar aprovados, inserir colunas, atualizar notas, criar gráficos ou gerar tabelas dinâmicas.";
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