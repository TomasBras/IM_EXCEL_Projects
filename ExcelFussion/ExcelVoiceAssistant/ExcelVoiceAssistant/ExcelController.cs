using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelVoiceAssistant
{
    public static class ExcelController
    {
        private static Excel.Application app;
        private static Excel.Workbook workbook;
        private static Excel.Worksheet sheet;

		private static readonly HashSet<int> _selectedRows = new HashSet<int>();
		private static readonly object _selectionLock = new object();
        private static readonly Dictionary<int, RowBorderSnapshot> _rowBorderSnapshots = new Dictionary<int, RowBorderSnapshot>();

        private sealed class RowBorderSnapshot
        {
            public BorderSnapshot Left;
            public BorderSnapshot Top;
            public BorderSnapshot Right;
            public BorderSnapshot Bottom;
        }

        private sealed class BorderSnapshot
        {
            public object LineStyle;
            public object Color;
            public object Weight;
        }



        private static string pathBase = @"E:\ExcelGestures\IM_Excel\ETP3.xlsx";
		private static string pathFinal = @"E:\ExcelGestures\IM_Excel\Relatorio_Final.xlsx";

		//private static string pathBase = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\ETP.xlsx";
		//private static string pathFinal = @"C:\Users\User\Desktop\ExcelGestures\IM_Excel\Relatorio_Final.xlsx";

		//private static string pathBase = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\ETP.xlsx";
		//private static string pathFinal = @"C:\Users\carol\Desktop\IM\IM_EXCEL_Projects\ExcelVoice\Relatorio_Final.xlsx";

		public static void SetExcel(Excel.Application excelApp, Excel.Workbook wb, Excel.Worksheet ws)
        {
            app = excelApp;
            workbook = wb;

            foreach (Excel.Worksheet sh in workbook.Worksheets)
            {
                if (sh.Cells[1, 1].Value?.ToString() == "Número mecanográfico")
                {
                    sheet = sh;
                    return;
                }
            }

            sheet = ws;
        }

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

        public static bool HasSelection()
        {
            lock (_selectionLock)
                return _selectedRows.Count > 0;
        }

        private static List<int> GetSelectedRowsOrdered()
        {
            lock (_selectionLock)
                return _selectedRows.OrderBy(r => r).ToList();
        }

        private static List<int> GetSelectedRowsBelowHeader(int headerRow)
        {
            return GetSelectedRowsOrdered().Where(r => r > headerRow).ToList();
        }

        private static int? GetSingleSelectedRowBelowHeader(int headerRow)
        {
            var rows = GetSelectedRowsBelowHeader(headerRow);
            return rows.Count > 0 ? rows[0] : (int?)null;
        }

        public static void ClearSelectionAndRestoreFormatting()
        {
            if (sheet == null)
                return;

            List<int> rows;
            Dictionary<int, RowBorderSnapshot> snapshots;
            lock (_selectionLock)
            {
                rows = _selectedRows.OrderBy(r => r).ToList();
                _selectedRows.Clear();

                snapshots = new Dictionary<int, RowBorderSnapshot>(_rowBorderSnapshots);
                _rowBorderSnapshots.Clear();
            }

            if (rows.Count == 0)
                return;

            try
            {
                var (headerRow, _) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                foreach (int row in rows)
                {
                    if (row <= headerRow)
                        continue;

                    try
                    {
                        var range = (Excel.Range)sheet.Range[sheet.Cells[row, firstCol], sheet.Cells[row, lastCol]];
                        if (snapshots.TryGetValue(row, out var snapshot) && snapshot != null)
                            RestoreRowBorderSnapshot(range, snapshot);
                        else
                            ClearRowSelectionBorder(range);
                    }
                    catch
                    {
                        // best-effort por linha (não abortar tudo se uma linha falhar)
                    }
                }
            }
            catch
            {
                // best-effort
            }
        }

        public static string ToggleSelectionAtActiveRow()
        {
            if (app?.ActiveCell == null || sheet == null)
                return "Nenhuma célula ativa.";

            int row = app.ActiveCell.Row;
            var (headerRow, _) = EncontrarCabecalho();
            if (row <= headerRow)
                return "Selecione uma linha de dados (abaixo do cabeçalho).";

            Excel.Range used = sheet.UsedRange;
            int firstCol = used.Column;
            int lastCol = firstCol + used.Columns.Count - 1;

            var range = (Excel.Range)sheet.Range[sheet.Cells[row, firstCol], sheet.Cells[row, lastCol]];

            bool isSelectedNow;
            RowBorderSnapshot snapshotToRestore = null;
            lock (_selectionLock)
            {
                if (_selectedRows.Contains(row))
                {
                    _selectedRows.Remove(row);
                    isSelectedNow = false;
                    if (_rowBorderSnapshots.TryGetValue(row, out var existing))					
                    {
                        snapshotToRestore = existing;
                        _rowBorderSnapshots.Remove(row);
                    }
                }
                else
                {
                    _selectedRows.Add(row);
                    isSelectedNow = true;
                }
            }

            try
            {
                if (isSelectedNow)
                {
                    // Realce via handgrab: borda vermelha (sem mexer no fundo)
                    lock (_selectionLock)
                    {
                        if (!_rowBorderSnapshots.ContainsKey(row))
                            _rowBorderSnapshots[row] = CaptureRowBorderSnapshot(range);
                    }

                    ApplyRowSelectionBorder(range);
                }
                else
                {
                    // remove realce (restaura bordas anteriores quando possível)
                    try
                    {
                        if (snapshotToRestore != null)
                            RestoreRowBorderSnapshot(range, snapshotToRestore);
                        else
                            ClearRowSelectionBorder(range);
                    }
                    catch
                    {
                        // se a restauração falhar, pelo menos remove a borda vermelha
                        try { ClearRowSelectionBorder(range); } catch { /* ignore */ }
                    }
                }
            }
            catch (COMException)
            {
                // best-effort; mesmo que a cor falhe, mantém seleção lógica
            }

            int count;
            lock (_selectionLock)
                count = _selectedRows.Count;

            return isSelectedNow
                ? $"Linha selecionada. Total: {count}."
                : $"Linha removida da seleção. Total: {count}.";
        }

        public static string AddSelectionAtActiveRow()
        {
            if (app?.ActiveCell == null || sheet == null)
                return "Nenhuma célula ativa.";

            int row = app.ActiveCell.Row;
            var (headerRow, _) = EncontrarCabecalho();
            if (row <= headerRow)
                return "Selecione uma linha de dados (abaixo do cabeçalho).";

            Excel.Range used = sheet.UsedRange;
            int firstCol = used.Column;
            int lastCol = firstCol + used.Columns.Count - 1;

            var range = (Excel.Range)sheet.Range[sheet.Cells[row, firstCol], sheet.Cells[row, lastCol]];

            bool addedNow = false;
            lock (_selectionLock)
            {
                if (!_selectedRows.Contains(row))
                {
                    _selectedRows.Add(row);
                    addedNow = true;
                }

                if (!_rowBorderSnapshots.ContainsKey(row))				
                    _rowBorderSnapshots[row] = CaptureRowBorderSnapshot(range);
            }

            try
            {
                ApplyRowSelectionBorder(range);
            }
            catch (COMException)
            {
                // best-effort
            }

            int count;
            lock (_selectionLock)
                count = _selectedRows.Count;

            return addedNow
                ? $"Linha selecionada. Total: {count}."
                : $"Linha já estava selecionada. Total: {count}.";
        }

        private static RowBorderSnapshot CaptureRowBorderSnapshot(Excel.Range range)
        {
            return new RowBorderSnapshot
            {
                Left = CaptureBorder(range, Excel.XlBordersIndex.xlEdgeLeft),
                Top = CaptureBorder(range, Excel.XlBordersIndex.xlEdgeTop),
                Right = CaptureBorder(range, Excel.XlBordersIndex.xlEdgeRight),
                Bottom = CaptureBorder(range, Excel.XlBordersIndex.xlEdgeBottom),
            };
        }

        private static BorderSnapshot CaptureBorder(Excel.Range range, Excel.XlBordersIndex idx)
        {
            try
            {
                var b = range.Borders[idx];
                return new BorderSnapshot
                {
                    LineStyle = b.LineStyle,
                    Color = b.Color,
                    Weight = b.Weight
                };
            }
            catch
            {
                return new BorderSnapshot
                {
                    LineStyle = Excel.XlLineStyle.xlLineStyleNone,
                    Color = ColorTranslator.ToOle(Color.Black),
                    Weight = Excel.XlBorderWeight.xlThin
                };
            }
        }

        private static void ApplyRowSelectionBorder(Excel.Range range)
        {
            SetBorder(range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlLineStyle.xlContinuous, Color.Red, Excel.XlBorderWeight.xlMedium);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlLineStyle.xlContinuous, Color.Red, Excel.XlBorderWeight.xlMedium);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlLineStyle.xlContinuous, Color.Red, Excel.XlBorderWeight.xlMedium);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlLineStyle.xlContinuous, Color.Red, Excel.XlBorderWeight.xlMedium);
        }

        private static void ClearRowSelectionBorder(Excel.Range range)
        {
            SetBorder(range, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlLineStyle.xlLineStyleNone, null, null);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeTop, Excel.XlLineStyle.xlLineStyleNone, null, null);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeRight, Excel.XlLineStyle.xlLineStyleNone, null, null);
            SetBorder(range, Excel.XlBordersIndex.xlEdgeBottom, Excel.XlLineStyle.xlLineStyleNone, null, null);
        }

        private static void RestoreRowBorderSnapshot(Excel.Range range, RowBorderSnapshot snapshot)
        {
            RestoreBorder(range, Excel.XlBordersIndex.xlEdgeLeft, snapshot.Left);
            RestoreBorder(range, Excel.XlBordersIndex.xlEdgeTop, snapshot.Top);
            RestoreBorder(range, Excel.XlBordersIndex.xlEdgeRight, snapshot.Right);
            RestoreBorder(range, Excel.XlBordersIndex.xlEdgeBottom, snapshot.Bottom);
        }

        private static void RestoreBorder(Excel.Range range, Excel.XlBordersIndex idx, BorderSnapshot snapshot)
        {
            try
            {
                var b = range.Borders[idx];
                b.LineStyle = snapshot.LineStyle;
                b.Color = snapshot.Color;
                b.Weight = snapshot.Weight;
            }
            catch
            {
                // ignore
            }
        }

        private static void SetBorder(
            Excel.Range range,
            Excel.XlBordersIndex idx,
            Excel.XlLineStyle lineStyle,
            Color? color,
            Excel.XlBorderWeight? weight)
        {
            try
            {
                var b = range.Borders[idx];
                b.LineStyle = lineStyle;
                if (color.HasValue) b.Color = ColorTranslator.ToOle(color.Value);
                if (weight.HasValue) b.Weight = weight.Value;
            }
            catch
            {
                // ignore
            }
        }

        public static string CalcularMediaSelecao()
        {
            List<int> rows;
            lock (_selectionLock)
                rows = _selectedRows.OrderBy(r => r).ToList();

            if (rows.Count == 0)
                return "Nenhuma linha selecionada.";

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

                var appliedRows = new List<int>();
                foreach (int row in rows)
                {
                    if (row <= headerRow) continue;
                    if (sheet.Cells[row, headerCol].Value == null) continue;

                    var cell = (Excel.Range)sheet.Cells[row, colMedia];
                    string formulaEn = BuildAverageFormula(testes, row);
                    string formulaPt = BuildAverageFormulaPt(testes, row);
                    SetCellFormulaWithFallback(cell, formulaEn, formulaPt);
					appliedRows.Add(row);
                }

                workbook.Save();

				if (appliedRows.Count == 0)
					return "Nenhuma linha válida selecionada.";

				string colLetter = ColunaParaLetra(colMedia);
				var appliedCells = appliedRows.Select(r => $"{colLetter}{r}").ToList();
				string coords = FormatCoordinateList(appliedCells, 8);

				return appliedRows.Count == 1
					? $"Média calculada para 1 linha selecionada ({coords})."
					: $"Média calculada para as {appliedRows.Count} linhas selecionadas ({coords}).";
            }
            catch (COMException ex)
            {
                Console.WriteLine($"❌ ERRO CalcularMediaSelecao: HRESULT=0x{ex.HResult:X8} MSG={ex.Message}");
                return "Erro ao calcular média na seleção.";
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ ERRO CalcularMediaSelecao: " + ex.Message);
                return "Erro ao calcular média na seleção.";
            }
        }

        private static string FormatCoordinateList(List<string> coordinates, int maxItems)
        {
            if (coordinates == null || coordinates.Count == 0)
                return "sem coordenadas";

            if (coordinates.Count <= maxItems)
                return "células " + string.Join(", ", coordinates);

            return "células " + string.Join(", ", coordinates.Take(maxItems)) + $", … (+{coordinates.Count - maxItems})";
        }


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

                    string clean = new string(
                        texto.Normalize(NormalizationForm.FormD)
                        .Where(ch => CharUnicodeInfo.GetUnicodeCategory(ch) != UnicodeCategory.NonSpacingMark)
                        .ToArray()
                    )
                    .Replace("\u00A0", " ") 
                    .Replace("\t", " ")     
                    .Replace("  ", " ")
                    .Trim()
                    .ToLower();

                    if (clean == "nome")
                        return (r, c);
                }
            }

            throw new Exception("Cabeçalho Nome não encontrado (mesmo após limpeza).");
        }

        private static string BuildAverageFormula(IEnumerable<int> cols, int row)
        {
            // IMPORTANT:
            // Range.Formula expects an invariant (US-English) formula.
            // That means English function names + comma as argument separator.
            // Using the localized list separator (often ';' in pt-PT) together with
            // English functions can lead to formulas that parse but evaluate as #NAME?.
            const string sep = ",";
            var refs = cols.Select(col => $"{ColunaParaLetra(col)}{row}");
            return $"=AVERAGE({string.Join(sep, refs)})";
        }

        private static string BuildAverageFormulaPt(IEnumerable<int> cols, int row)
        {
            // Range.FormulaLocal expects localized function names and localized separators.
            string sep = app?.International[Excel.XlApplicationInternational.xlListSeparator]?.ToString() ?? ";";
            var refs = cols.Select(col => $"{ColunaParaLetra(col)}{row}");
            return $"=MÉDIA({string.Join(sep, refs)})";
        }

        private static void SetCellFormulaWithFallback(Excel.Range cell, string englishFormula, string portugueseFormula)
        {
            try
            {
                cell.Formula = englishFormula;
            }
            catch (COMException)
            {
                // Fallback to .FormulaLocal using Portuguese function name + localized separators.
                cell.FormulaLocal = portugueseFormula;
            }
        }


        public static string CalcularMedia(dynamic json)
        {
            try
            {
                string intent = json.nlu.intent.ToString();

                string nome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : null;
                string numero = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : null;

                if (!string.IsNullOrEmpty(nome))
                    return CalcularMediaAluno(nome);

                if (!string.IsNullOrEmpty(numero))
                    return CalcularMediaAlunoNumero(numero);

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
                    var cell = (Excel.Range)sheet.Cells[row, colMedia];
                    string formulaEn = BuildAverageFormula(testes, row);
                    string formulaPt = BuildAverageFormulaPt(testes, row);
                    SetCellFormulaWithFallback(cell, formulaEn, formulaPt);
                    row++;
                }

                workbook.Save();
                return "Média turma calculada.";
            }
            catch (COMException ex)
            {
                Console.WriteLine($"❌ ERRO CalcularMediaTurma: HRESULT=0x{ex.HResult:X8} MSG={ex.Message}");
                return "Erro ao calcular média turma.";
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

                var cell = (Excel.Range)sheet.Cells[rowAluno, colMedia];
                string formulaEn = BuildAverageFormula(colTestes, rowAluno);
                string formulaPt = BuildAverageFormulaPt(colTestes, rowAluno);
                SetCellFormulaWithFallback(cell, formulaEn, formulaPt);

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
                    return "Coluna Número mecanográfico não encontrada.";

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

                var cell = (Excel.Range)sheet.Cells[rowAluno, colMedia];
                string formulaEn = BuildAverageFormula(colTestes, rowAluno);
                string formulaPt = BuildAverageFormulaPt(colTestes, rowAluno);
                SetCellFormulaWithFallback(cell, formulaEn, formulaPt);

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
                int colSit = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "média"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "situação"))
                    {
                        colSit = c;
                        break;
                    }
                }

                if (colMedia == -1)
                    return "Calcule a média primeiro.";

                if (colSit != -1)
                    return "Coluna situação já existe.";

                colSit = colMedia + 1;
                // insere para não sobrescrever a coluna seguinte
                sheet.Columns[colSit].Insert();
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

        public static string InserirSituacaoSelecao()
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                var selectedRows = GetSelectedRowsBelowHeader(headerRow);
                if (selectedRows.Count == 0)
                    return "Nenhuma linha selecionada.";

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colMedia = -1;
                int colSit = -1;

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "média"))
                    {
                        colMedia = c;
                        break;
                    }
                }

                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v != null && IgualIgnorandoAcentos(v.ToString(), "situação"))
                    {
                        colSit = c;
                        break;
                    }
                }

                if (colMedia == -1)
                    return "Calcule a média primeiro.";

                if (colSit == -1)
                {
                    colSit = colMedia + 1;
                    // insere para não sobrescrever a coluna seguinte
                    sheet.Columns[colSit].Insert();
                    sheet.Cells[headerRow, colSit].Value2 = "Situação";
                }

                foreach (int row in selectedRows)
                {
                    if (sheet.Cells[row, headerCol].Value == null)
                        continue;
                    sheet.Cells[row, colSit].Value2 = "";
                }

                return selectedRows.Count == 1
                    ? "Coluna situação aplicada à linha selecionada."
                    : $"Coluna situação aplicada a {selectedRows.Count} linhas selecionadas.";
            }
            catch
            {
                return "Erro ao criar Situação na seleção.";
            }
        }

        public static string InserirColuna(dynamic json)
        {
            try
            {
                string nomeColuna = null;

                try
                {
                    if (json?.nlu?.coluna_nome != null)
                        nomeColuna = json.nlu.coluna_nome.ToString();
                }
                catch
                {
                    // ignore
                }

                if (string.IsNullOrWhiteSpace(nomeColuna))
                {
                    try
                    {
                        string texto = json?.text != null
                            ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString()))
                            : "";

                        if (!string.IsNullOrWhiteSpace(texto))
                        {
                            // exemplos:
                            // - "cria coluna chamada feedback"
                            // - "cria uma nova coluna chamada Situação"
                            // - "insere a coluna melhoria"
                            // - "adiciona coluna feedback"
                            texto = Regex.Replace(texto, @"\s+", " ").Trim();

                            Match m;

                            // "... coluna ... chamada/nome ... <X>"
                            m = Regex.Match(
                                texto,
                                @"\bcoluna\b.*?\b(chamada|com\s+nome|nome)\b\s+(?<nome>.+)$",
                                RegexOptions.IgnoreCase
                            );
                            if (m.Success)
                            {
                                nomeColuna = m.Groups["nome"].Value;
                            }
                            else
                            {
                                // "inserir/adicionar/criar coluna <X>" (sem 'chamada')
                                m = Regex.Match(
                                    texto,
                                    @"\b(inserir|insere|adicionar|adiciona|criar|cria)\b.*?\bcoluna\b\s+(?<nome>.+)$",
                                    RegexOptions.IgnoreCase
                                );
                                if (m.Success)
                                {
                                    nomeColuna = m.Groups["nome"].Value;
                                }
                                else
                                {
                                    // fallback: se terminar em "coluna <X>" tenta extrair <X>
                                    m = Regex.Match(
                                        texto,
                                        @"\bcoluna\b\s+(?<nome>.+)$",
                                        RegexOptions.IgnoreCase
                                    );
                                    nomeColuna = m.Success ? m.Groups["nome"].Value : texto;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // ignore
                    }
                }

                nomeColuna = (nomeColuna ?? "").Trim().Trim('"', '\'', '.', '!', '?');

                // remove artigos/adjetivos comuns no início (ex.: "a Situação", "uma Melhoria")
                nomeColuna = Regex.Replace(
                    nomeColuna,
                    @"^(a|o|um|uma|novo|nova|novos|novas)\s+",
                    "",
                    RegexOptions.IgnoreCase
                ).Trim();

                if (string.IsNullOrWhiteSpace(nomeColuna))
                    nomeColuna = "Situação";

                // normaliza capitalização (ex.: "melhoria" -> "Melhoria")
                try
                {
                    var ti = CultureInfo.GetCultureInfo("pt-PT").TextInfo;
                    nomeColuna = Regex.Replace(nomeColuna, @"\s+", " ").Trim();
                    nomeColuna = ti.ToTitleCase(nomeColuna.ToLower());
                }
                catch
                {
                    // fallback simples
                    nomeColuna = Regex.Replace(nomeColuna, @"\s+", " ").Trim();
                    if (nomeColuna.Length > 0)
                        nomeColuna = char.ToUpper(nomeColuna[0]) + (nomeColuna.Length > 1 ? nomeColuna.Substring(1) : "");
                }

                // pedido específico: quando for Situação, usar InserirSituacao()
                if (IgualIgnorandoAcentos(nomeColuna, "situação"))
                    return InserirSituacao();

                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                // evita duplicados
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var v = sheet.Cells[headerRow, c].Value;
                    if (v == null) continue;

                    if (IgualIgnorandoAcentos(v.ToString(), nomeColuna))
                        return $"A coluna {nomeColuna} já existe.";
                }

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

                int colNova = colMedia + 1;

                // insere uma coluna para não sobrescrever a existente
                sheet.Columns[colNova].Insert();
                sheet.Cells[headerRow, colNova].Value2 = nomeColuna;

                int row = headerRow + 1;
                while (sheet.Cells[row, headerCol].Value != null)
                {
                    sheet.Cells[row, colNova].Value2 = "";
                    row++;
                }

                return $"Coluna {nomeColuna} criada.";
            }
            catch (Exception ex)
            {
                return "Erro ao criar coluna: " + ex.Message;
            }
        }

        public static string GerarGraficoTurmaSelecao(dynamic json)
        {
            try
            {
                var (headerRow, headerCol) = EncontrarCabecalho();
                var selectedRows = GetSelectedRowsBelowHeader(headerRow);
                if (selectedRows.Count == 0)
                    return "Nenhuma linha selecionada.";

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

                double somaT1 = 0, somaT2 = 0, somaM = 0;
                int count = 0;

                foreach (int r in selectedRows)
                {
                    if (sheet.Cells[r, headerCol].Value == null)
                        continue;
                    somaT1 += Convert.ToDouble(sheet.Cells[r, colT1].Value2 ?? 0);
                    somaT2 += Convert.ToDouble(sheet.Cells[r, colT2].Value2 ?? 0);
                    somaM += Convert.ToDouble(sheet.Cells[r, colMedia].Value2 ?? 0);
                    count++;
                }

                if (count <= 0)
                    return "Seleção vazia.";

                double mT1 = somaT1 / count;
                double mT2 = somaT2 / count;
                double mMF = somaM / count;

                // Posiciona o gráfico abaixo do último aluno (usa End(xlUp) na coluna "Média" para ignorar formatação solta).
                int maxRow = sheet.Rows.Count;
                int lastDataRow = ((Excel.Range)sheet.Cells[maxRow, colMedia]).End[Excel.XlDirection.xlUp].Row;
                if (lastDataRow < headerRow + 1)
                    lastDataRow = headerRow + 1;

                // Coloca o gráfico ao lado da tabela principal (mais visível do que muito abaixo).
                Excel.Range tableRange = sheet.Range[sheet.Cells[headerRow, firstCol], sheet.Cells[lastDataRow, lastCol]];
                double posX = tableRange.Left + tableRange.Width + 30;
                double baseY = tableRange.Top;

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? baseY
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 20;

                Excel.ChartObject chartObj = charts.Add(posX, posY, 650, 360);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                if (count == 1)
                {
                    int rowAluno = selectedRows[0];

                    int colNumeroMec = -1;
                    for (int c = firstCol; c <= lastCol; c++)
                    {
                        var titulo = sheet.Cells[headerRow, c].Value?.ToString();
                        if (titulo == null) continue;
                        if (IgualIgnorandoAcentos(titulo, "Número mecanográfico") || IgualIgnorandoAcentos(titulo, "Número Mecanográfico"))
                        {
                            colNumeroMec = c;
                            break;
                        }
                    }

                    string nome = sheet.Cells[rowAluno, headerCol].Value?.ToString() ?? "(Sem nome)";
                    string nmec = colNumeroMec > 0 ? (sheet.Cells[rowAluno, colNumeroMec].Value?.ToString() ?? "") : "";
                    string textoNmec = string.IsNullOrWhiteSpace(nmec) ? "" : $" (nº mec {nmec})";
                    chart.ChartTitle.Text = $"Média do(a) aluno(a) {nome}{textoNmec}";
                }
                else
                {
                    chart.ChartTitle.Text = $"Médias (seleção: {count} linha(s))";
                }

                Excel.Series s = chart.SeriesCollection().NewSeries();
                s.Name = "Médias";
                s.Values = new double[] { mT1, mT2, mMF };
                s.XValues = new string[] { "Teste 1", "Teste 2", "Média" };

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                // Leva a vista para o gráfico recém-criado (evita parecer que "não apareceu").
                try
                {
                    chartObj.Activate();
                    chartObj.Select();
                    app.Goto(chartObj.TopLeftCell, true);
                }
                catch
                {
                    // best-effort
                }

                return "Gráfico criado (apenas seleção).";
            }
            catch (Exception ex)
            {
                return "Erro ao criar gráfico (seleção): " + ex.Message;
            }
        }

        public static string GerarGraficoBarrasSelecao(dynamic json)
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                var selectedRows = GetSelectedRowsBelowHeader(headerRow);
                if (selectedRows.Count == 0)
                    return "Nenhuma linha selecionada.";
                if (selectedRows.Count > 1)
                    return "Selecione só 1 aluno para gerar o gráfico de barras.";

                int rowAluno = selectedRows[0];

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                int colNumeroMec = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;
                    if (IgualIgnorandoAcentos(titulo, "Número mecanográfico") || IgualIgnorandoAcentos(titulo, "Número Mecanográfico"))
                    {
                        colNumeroMec = c;
                        break;
                    }
                }

                int colT1 = -1, colT2 = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    var titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;
                    if (IgualIgnorandoAcentos(titulo, "Teste 1")) colT1 = c;
                    if (IgualIgnorandoAcentos(titulo, "Teste 2")) colT2 = c;
                }

                if (colT1 == -1 || colT2 == -1)
                    return "Não encontrei Teste 1 / Teste 2.";

                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerColNome].Value != null)
                    lastRow++;

                string nomeFinal = sheet.Cells[rowAluno, headerColNome].Value?.ToString() ?? "(Sem nome)";

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 700, 380);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                {
                    string nmec = colNumeroMec > 0 ? (sheet.Cells[rowAluno, colNumeroMec].Value?.ToString() ?? "") : "";
                    string textoNmec = string.IsNullOrWhiteSpace(nmec) ? "" : $" (nº mec {nmec})";
                    chart.ChartTitle.Text = $"Média do(a) aluno(a) {nomeFinal}{textoNmec}";
                }

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

                return "Gráfico criado (apenas seleção).";
            }
            catch
            {
                return "Erro ao criar gráfico de barras (seleção).";
            }
        }

        public static string GerarGraficoPerguntasT2Selecao()
        {
            try
            {
                var (headerRow, headerColNome) = EncontrarCabecalho();
                var selectedRows = GetSelectedRowsBelowHeader(headerRow);
                if (selectedRows.Count == 0)
                    return "Nenhuma linha selecionada.";

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

                Dictionary<string, int> perguntas = new Dictionary<string, int>();
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo == null) continue;
                    if (titulo.Trim().StartsWith("T2_P"))
                        perguntas[titulo.Trim()] = c;
                }

                if (perguntas.Count == 0)
                    return "Nenhuma coluna T2_P encontrada.";

                var ordenadas = perguntas.OrderBy(k => k.Key).ToList();

                // posicionamento dos gráficos usa o último aluno da folha (não da seleção)
                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerColNome].Value != null)
                    lastRow++;

                var rowsValidas = selectedRows.Where(r => sheet.Cells[r, headerColNome].Value != null).ToList();
                if (rowsValidas.Count == 0)
                    return "Seleção vazia.";

                List<double> medias = new List<double>();
                foreach (var kv in ordenadas)
                {
                    double soma = 0;
                    foreach (int r in rowsValidas)
                        soma += Convert.ToDouble(sheet.Cells[r, kv.Value].Value2 ?? 0);
                    medias.Add(soma / rowsValidas.Count);
                }

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
                double posY = charts.Count == 0
                    ? sheet.Rows[lastRow].Top + 30
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 40;

                Excel.ChartObject chartObj = charts.Add(50, posY, 700, 400);
                Excel.Chart chart = chartObj.Chart;

                chart.ChartType = Excel.XlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = $"Médias T2 (seleção: {rowsValidas.Count} linha(s))";

                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                Excel.Series s = sc.NewSeries();
                s.Name = "Média";
                s.Values = medias.ToArray();
                s.XValues = ordenadas.Select(k => k.Key).ToArray();

                chart.Axes(Excel.XlAxisType.xlValue).MinimumScale = 0;
                chart.Axes(Excel.XlAxisType.xlValue).MaximumScale = 20;

                return "Gráfico das perguntas T2 criado (apenas seleção).";
            }
            catch
            {
                return "Erro ao gerar gráfico das perguntas (seleção).";
            }
        }

        public static string CriarPivotTableSelecao(dynamic json)
        {
            try
            {
                var (headerRow, _) = EncontrarCabecalho();
                var selectedRows = GetSelectedRowsBelowHeader(headerRow);
                if (selectedRows.Count == 0)
                    return "Nenhuma linha selecionada.";

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = used.Column + used.Columns.Count - 1;

                // Copia cabeçalho + linhas selecionadas para uma folha temporária (range contíguo).
                Excel.Worksheet dataSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                dataSheet.Name = "DataSel_" + DateTime.Now.Ticks;

                // Header
                Excel.Range headerRange = (Excel.Range)sheet.Range[sheet.Cells[headerRow, firstCol], sheet.Cells[headerRow, lastCol]];
                headerRange.Copy(dataSheet.Cells[1, 1]);

                int outRow = 2;
                foreach (int r in selectedRows)
                {
                    Excel.Range rowRange = (Excel.Range)sheet.Range[sheet.Cells[r, firstCol], sheet.Cells[r, lastCol]];
                    rowRange.Copy(dataSheet.Cells[outRow, 1]);
                    outRow++;
                }

                int outLastRow = outRow - 1;
                Excel.Range dataRange = (Excel.Range)dataSheet.Range[dataSheet.Cells[1, 1], dataSheet.Cells[outLastRow, lastCol - firstCol + 1]];

                Excel.Worksheet pivotSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                pivotSheet.Name = "PivotSel_" + DateTime.Now.Ticks;

                Excel.PivotCache cache = workbook.PivotCaches().Create(
                    Excel.XlPivotTableSourceType.xlDatabase,
                    dataRange
                );

                Excel.PivotTable pivot = cache.CreatePivotTable(
                    pivotSheet.Cells[1, 1],
                    "TabelaDinamica"
                );

                string rowField = json?.nlu?.coluna_excel_row?.ToString();
                string valueField = json?.nlu?.coluna_excel_value?.ToString();

                bool comandoBasico = (rowField == null && valueField == null);

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

                if (comandoBasico)
                {
                    Excel.PivotField pfNome = pivot.PivotFields("Nome");
                    pfNome.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    Excel.PivotField pfRegime = pivot.PivotFields("REGIME");
                    pfRegime.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    Excel.PivotField pf = pivot.PivotFields("Média");
                    pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    pf.Function = Excel.XlConsolidationFunction.xlAverage;
                    pf.Name = "Média";

                    return "Tabela dinâmica criada (apenas seleção).";
                }

                if (rowField != null)
                {
                    Excel.PivotField row = pivot.PivotFields(rowField);
                    row.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }

                if (valueField != null)
                {
                    Excel.PivotField pf = pivot.PivotFields(valueField);
                    if (!ColunaEhNumerica(valueField))
                    {
                        pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                        return $"O campo '{valueField}' não é numérico e foi movido automaticamente para as linhas.";
                    }
                    pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    pf.Function = Excel.XlConsolidationFunction.xlAverage;
                    pf.Name = "Média de " + valueField;
                }

                return "Tabela dinâmica criada (apenas seleção).";
            }
            catch (Exception ex)
            {
                return "Erro ao criar tabela dinâmica (seleção): " + ex.Message;
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

                if (colSit == -1)
                    return "Criar coluna situação primeiro.";

                if (colMedia == -1)
                    return "Calcular média primeiro.";

                List<int> selectedRows;
                lock (_selectionLock)
                    selectedRows = _selectedRows.OrderBy(r => r).ToList();

                if (selectedRows.Count > 0)
                {
                    int applied = 0;
                    foreach (int row in selectedRows)
                    {
                        if (row <= headerRow) continue;
                        if (sheet.Cells[row, headerCol].Value == null) continue;

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
                        applied++;
                    }

                    return selectedRows.Count == 1
                        ? "Situação atualizada na linha selecionada."
                        : $"Situação atualizada nas {selectedRows.Count} linhas selecionadas.";
                }

                int scanRow = headerRow + 1;

                while (sheet.Cells[scanRow, headerCol].Value != null)
                {
                    double media = sheet.Cells[scanRow, colMedia].Value2 ?? 0;

                    if (media >= 10)
                    {
                        sheet.Cells[scanRow, colSit].Value2 = "Aprovado";
                        sheet.Cells[scanRow, colSit].Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                    }
                    else
                    {
                        sheet.Cells[scanRow, colSit].Value2 = "Reprovado";
                        sheet.Cells[scanRow, colSit].Interior.Color = ColorTranslator.ToOle(Color.LightCoral);
                    }

                    scanRow++;
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

                int row = headerRow + 1;

                while (sheet.Cells[row, headerColNome].Value != null)
                {
                    double penultimo = Convert.ToDouble(sheet.Cells[row, colPenultimo].Value2 ?? 0);
                    double ultimo = Convert.ToDouble(sheet.Cells[row, colUltimo].Value2 ?? 0);

                    double diferenca = ultimo - penultimo;
                    double percent = (penultimo != 0)
                        ? (diferenca / penultimo) * 100
                        : (diferenca > 0 ? 100 : 0);

                    string texto;

                    if (diferenca > 0)
                    {
                        texto = $"Melhorou (+{Math.Round(diferenca, 2)} valores, +{Math.Round(percent, 1)}%)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else if (diferenca < 0)
                    {
                        texto = $"Piorou ({Math.Round(diferenca, 2)} valores, {Math.Round(percent, 1)}%)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightCoral);
                    }
                    else
                    {
                        texto = $"Igual (0)";

                        var cell = sheet.Cells[row, colMelhoria];
                        cell.Value2 = texto;

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

                // Prefer filling a column explicitly named "Melhoria".
                // For backward compatibility, if an "MP" column already exists, we also update it,
                // but we no longer auto-create it.
                int colMelhoria = -1;
                for (int c = firstCol; c <= lastCol; c++)
                {
                    string titulo = sheet.Cells[headerRow, c].Value?.ToString();
                    if (titulo != null && IgualIgnorandoAcentos(titulo, "melhoria"))
                    {
                        colMelhoria = c;
                        break;
                    }
                }

                if (colMelhoria == -1)
                {
                    colMelhoria = lastCol + 1;
                    sheet.Cells[headerRow, colMelhoria].Value2 = "Melhoria";
                    lastCol++;
                }

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

                int row = headerRow + 1;

                while (sheet.Cells[row, headerColNome].Value != null)
                {
                    double mediaAtual = sheet.Cells[row, colMedia].Value2 ?? 0;

                    if (mediaAtual >= 10)
                    {
                        sheet.Cells[row, colMelhoria].Value2 = "";
                        if (colMP != -1) sheet.Cells[row, colMP].Value2 = "";
                        sheet.Cells[row, colNotaNecessaria].Value2 = "—";
                        row++;
                        continue;
                    }

                    double soma = 0;
                    foreach (int col in colTestes.Take(colTestes.Count - 1))
                        soma += Convert.ToDouble(sheet.Cells[row, col].Value2 ?? 0);

                    int n = colTestes.Count;

                    double notaNecessaria = 10 * n - soma;

                    if (notaNecessaria > 20)
                        sheet.Cells[row, colNotaNecessaria].Value2 = ">20";
                    else
                        sheet.Cells[row, colNotaNecessaria].Value2 = Math.Round(notaNecessaria, 2);

                    if (notaNecessaria <= 20)
                    {
                        sheet.Cells[row, colMelhoria].Value2 = "MP";
                        if (colMP != -1) sheet.Cells[row, colMP].Value2 = "MP";
                    }
                    else
                    {
                        sheet.Cells[row, colMelhoria].Value2 = "";
                        if (colMP != -1) sheet.Cells[row, colMP].Value2 = "";
                    }

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

                string texto = json.text != null
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower()
                    : "";

                int pInicio = -1, pFim = -1;
                var intervalo = Regex.Match(texto, @"(\d+)\s*(a|à|até|-)\s*(\d+)");
                if (intervalo.Success)
                {
                    pInicio = int.Parse(intervalo.Groups[1].Value);
                    pFim = int.Parse(intervalo.Groups[3].Value);
                }

                var unico = Regex.Match(texto, @"(p|pergunta|questao|questão|q)\s*(número\s*)?(\d+)");
                if (unico.Success)
                {
                    int p = int.Parse(unico.Groups[3].Value);
                    pInicio = pFim = p;
                }

                if (pInicio == -1)
                    return "Não percebi qual pergunta queres adicionar.";

                var (headerRow, headerCol) = EncontrarCabecalho();
                Excel.Range used = sheet.UsedRange;

                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

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

                int adicionadas = 0;

                for (int p = pInicio; p <= pFim; p++)
                {
                    if (!existentes.ContainsKey(p))
                    {
                        sheet.Columns[colTeste].Insert();

                        sheet.Cells[headerRow, colTeste].Value2 = $"{prefixo}{p}";

                        int r = headerRow + 1;
                        while (sheet.Cells[r, headerCol].Value != null)
                        {
                            sheet.Cells[r, colTeste].Value2 = "";
                            r++;
                        }

                        adicionadas++;

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

                // Determina a última linha de alunos com End(xlUp) na coluna "Média" (ignora UsedRange gigante por formatação).
                int maxRow = sheet.Rows.Count;
                int lastDataRow = ((Excel.Range)sheet.Cells[maxRow, colMedia]).End[Excel.XlDirection.xlUp].Row;
                if (lastDataRow <= headerRow)
                    return "Sem alunos.";

                int count = 0;
                double somaT1 = 0, somaT2 = 0, somaM = 0;
                for (int r = headerRow + 1; r <= lastDataRow; r++)
                {
                    // Conta apenas linhas que parecem ter aluno (evita linhas vazias no meio).
                    if (sheet.Cells[r, headerCol].Value == null)
                        continue;
                    somaT1 += Convert.ToDouble(sheet.Cells[r, colT1].Value2 ?? 0);
                    somaT2 += Convert.ToDouble(sheet.Cells[r, colT2].Value2 ?? 0);
                    somaM += Convert.ToDouble(sheet.Cells[r, colMedia].Value2 ?? 0);
                    count++;
                }

                if (count <= 0)
                    return "Sem alunos.";

                double mT1 = somaT1 / count;
                double mT2 = somaT2 / count;
                double mMF = somaM / count;

                // Coloca o gráfico ao lado da tabela principal para ser imediatamente visível.
                Excel.Range tableRange = sheet.Range[sheet.Cells[headerRow, firstCol], sheet.Cells[lastDataRow, lastCol]];
                double posX = tableRange.Left + tableRange.Width + 30;
                double baseY = tableRange.Top;

                Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();

                double posY = charts.Count == 0
                    ? baseY
                    : charts.Item(charts.Count).Top + charts.Item(charts.Count).Height + 30;

                Excel.ChartObject chartObj = charts.Add(posX, posY, 650, 360);
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

                try
                {
                    chartObj.Activate();
                    chartObj.Select();
                    app.Goto(chartObj.TopLeftCell, true);
                }
                catch
                {
                    // best-effort
                }

                return "Gráfico criado.";
            }
            catch (Exception ex)
            {
                return "Erro ao criar gráfico: " + ex.Message;
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
                string numeroMec = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : "";
                string alunoNome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : "";

                if (string.IsNullOrEmpty(numeroMec) && string.IsNullOrEmpty(alunoNome))
                {
                    Console.WriteLine("❌ Não foi indicado nome nem número do aluno.");
                    return "Não foi indicado nome nem número do aluno.";
                }

                Excel.Range used = sheet.UsedRange;

                var (headerRow, colNome) = EncontrarCabecalho();

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

                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, colNome].Value != null)
                    lastRow++;

                int rowAluno = -1;

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

                if (rowAluno == -1)
                {
                    Console.WriteLine($"❌ Aluno não encontrado: {alunoNome} / {numeroMec}");
                    return $"Aluno não encontrado: {alunoNome} / {numeroMec}";
                }

                string nomeFinal = sheet.Cells[rowAluno, colNome].Value?.ToString() ?? "(Sem nome)";
                string textoNumero = string.IsNullOrEmpty(numeroMec) ? "" : $" (NMec {numeroMec})";

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

                var ordenadas = perguntas.OrderBy(k => k.Key).ToList();

                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, headerColNome].Value != null)
                    lastRow++;

                int totalAlunos = lastRow - headerRow - 1;
                if (totalAlunos <= 0)
                {
                    Console.WriteLine("❌ Nenhum aluno encontrado.");
                    return "Nenhum aluno encontrado.";
                }

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
                string textoOriginal = json.text != null
                    ? Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower()
                    : "";

                string numeroMec = json.nlu.aluno_numero != null ? json.nlu.aluno_numero.ToString() : null;
                string alunoNome = json.nlu.aluno_nome != null ? json.nlu.aluno_nome.ToString() : null;

                int testeNum = -1;
                Match matchTeste = Regex.Match(textoOriginal, @"teste ?([0-9]{1,2})");
                if (matchTeste.Success)
                    testeNum = int.Parse(matchTeste.Groups[1].Value);

                int perguntaNum = -1;
                Match matchPerg = Regex.Match(textoOriginal, @"(pergunta|quest[aã]o) ?([0-9]{1,2})");
                if (matchPerg.Success)
                    perguntaNum = int.Parse(matchPerg.Groups[2].Value);

                if (perguntaNum != -1 && testeNum == -1)
                    return "Tens de indicar o número do teste. Ex.: 'pergunta 2 do teste 1'.";

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


                var header = EncontrarCabecalho();
                int headerRow = header.Item1;
                int colNome = header.Item2;

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

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

                int lastRow = headerRow + 1;
                while (sheet.Cells[lastRow, colNome].Value != null)
                    lastRow++;

               
                int alunoRow = -1;

                for (int r = headerRow + 1; r < lastRow; r++)
                {
                    object nm = sheet.Cells[r, colNMec].Value;

                    if (numeroMec != null && nm != null && nm.ToString() == numeroMec)
                    {
                        alunoRow = r;
                        break;
                    }

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

                bool operacaoTurma =
                    alunoRow == -1 &&
                    (textoOriginal.Contains("toda a turma") || textoOriginal.Contains("todos os alunos"));


               
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


                
                bool pedirZero = textoOriginal.Contains(" zero");
                bool pedirRandom = textoOriginal.Contains("random") || textoOriginal.Contains("aleat");
                bool pedirCotacaoMax = textoOriginal.Contains("cotação máxima") || textoOriginal.Contains("nota máxima");
                bool apenasVazias = textoOriginal.Contains("vazia");

                Random rnd = new Random();

                Action<int> AplicarOperacao = delegate (int r)
                {
                    if (pedirZero)
                    {
                        foreach (int col in colsPerguntas.Values)
                            sheet.Cells[r, col].Value2 = 0;
                    }

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
                                randomNota = rnd.Next(0, 21);  
                            else
                                randomNota = Math.Round(rnd.NextDouble() * 20, 1);

                            sheet.Cells[r, col].Value2 = randomNota;
                        }
                    }

                    else if (pedirCotacaoMax && perguntaNum != -1)
                    {
                        if (colsPerguntas.ContainsKey(perguntaNum))
                            sheet.Cells[r, colsPerguntas[perguntaNum]].Value2 = 20.0;
                    }

                    else if (perguntaNum != -1 && valores.Count >= 1)
                    {
                        if (colsPerguntas.ContainsKey(perguntaNum))
                            sheet.Cells[r, colsPerguntas[perguntaNum]].Value2 = valores[0];
                    }

                    else if (valores.Count > 1)
                    {
                        List<KeyValuePair<int, int>> ord =
                            colsPerguntas.OrderBy(k => k.Key).ToList();

                        for (int i = 0; i < valores.Count && i < ord.Count; i++)
                            sheet.Cells[r, ord[i].Value].Value2 = valores[i];
                    }

                   
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

                    if (colTesteFinal != -1)
                        sheet.Cells[r, colTesteFinal].Value2 = soma;
                };


                
                if (operacaoTurma)
                {
                    for (int r = headerRow + 1; r < lastRow; r++)
                        AplicarOperacao(r);
                }
                else
                {
                    AplicarOperacao(alunoRow);
                }


                
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
                        var cell = (Excel.Range)sheet.Cells[r, colMedia];
                        string formulaEn = BuildAverageFormula(colTestes, r);
                        string formulaPt = BuildAverageFormulaPt(colTestes, r);
                        SetCellFormulaWithFallback(cell, formulaEn, formulaPt);
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
                string texto = "";

                if (json != null && json.text != null)
                {
                    texto = Encoding.UTF8.GetString(Convert.FromBase64String(json.text.ToString())).ToLower();
                }

                Console.WriteLine($"[OperacoesMatematicas] start | texto.len={texto?.Length ?? 0}");

                var (headerRow, headerCol) = EncontrarCabecalho();

                Excel.Range used = sheet.UsedRange;
                int firstCol = used.Column;
                int lastCol = firstCol + used.Columns.Count - 1;

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

                Console.WriteLine($"[OperacoesMatematicas] colMedia={colMedia}");

                if (colMedia == -1)
                    return "É necessário calcular a média primeiro.";

                int row = headerRow + 1;
                List<double> medias = new List<double>();

                while (sheet.Cells[row, headerCol].Value != null)
                {
                    double media = sheet.Cells[row, colMedia].Value2 ?? 0;
                    medias.Add(media);
                    row++;
                }

                Console.WriteLine($"[OperacoesMatematicas] total_alunos={medias.Count}");

                int total = medias.Count;
                if (total == 0) return "Nenhum aluno encontrado.";

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

                
                bool pedidoGeral =
                    texto.Contains("estatistic") ||
                    texto.Contains("resumo") ||
                    texto.Contains("tabela") ||
                    texto.Contains("relatório") ||
                    texto.Contains("estatísticas gerais");

               
                            Console.WriteLine($"[OperacoesMatematicas] pedidoGeral={pedidoGeral}");

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
            int headerRow = used.Row;   

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

                string rowField = json?.nlu?.coluna_excel_row?.ToString();
                string valueField = json?.nlu?.coluna_excel_value?.ToString();
                string filterRegime = json?.nlu?.regime?.ToString();

                bool comandoBasico = (rowField == null && valueField == null);

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

                if (comandoBasico)
                {
                    Excel.PivotField pfNome = pivot.PivotFields("Nome");
                    pfNome.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    Excel.PivotField pfRegime = pivot.PivotFields("REGIME");
                    pfRegime.Orientation = Excel.XlPivotFieldOrientation.xlRowField;

                    Excel.PivotField pf = pivot.PivotFields("Média");
                    pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField;
                    pf.Function = Excel.XlConsolidationFunction.xlAverage;
                    pf.Name = "Média";

                    return "Tabela dinâmica criada com campos padrão.";
                }

                if (rowField != null)
                {
                    Excel.PivotField row = pivot.PivotFields(rowField);
                    row.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }

                if (valueField != null)
                {
                    if (!ColunaEhNumerica(valueField))
                    {
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

        public static string DestacarApenasAprovados()
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
                    var v = sheet.Cells[headerRow, c].Value?.ToString();
                    if (v == null) continue;

                    if (IgualIgnorandoAcentos(v, "média")) colMedia = c;
                    if (IgualIgnorandoAcentos(v, "situação")) colSit = c;
                }

                if (colMedia == -1 || colSit == -1)
                    return "É necessário calcular a média e criar a coluna situação.";

                List<int> selectedRows;
                lock (_selectionLock)
                    selectedRows = _selectedRows.OrderBy(r => r).ToList();

                int applied = 0;
                if (selectedRows.Count > 0)
                {
                    foreach (int row in selectedRows)
                    {
                        if (row <= headerRow) continue;
                        if (sheet.Cells[row, headerCol].Value == null) continue;

                        // Mantém o estado existente; apenas marca/aplica quando corresponde.
                        var sitCell = (Excel.Range)sheet.Cells[row, colSit];

                        double media = sheet.Cells[row, colMedia].Value2 ?? 0;
                        if (media >= 10)
                        {
                            sitCell.Value2 = "Aprovado";
                            sitCell.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                            applied++;
                        }
                    }

                    return selectedRows.Count == 1
                        ? $"Aprovados destacados na linha selecionada."
                        : $"Aprovados destacados nas {selectedRows.Count} linhas selecionadas.";
                }

                int scanRow = headerRow + 1;
                while (sheet.Cells[scanRow, headerCol].Value != null)
                {
                    // Mantém o estado existente; apenas marca/aplica quando corresponde.
                    var sitCell = (Excel.Range)sheet.Cells[scanRow, colSit];

                    double media = sheet.Cells[scanRow, colMedia].Value2 ?? 0;
                    if (media >= 10)
                    {
                        sitCell.Value2 = "Aprovado";
                        sitCell.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                        applied++;
                    }
                    scanRow++;
                }

                return "Aprovados destacados (sem ocultar restantes).";
            }
            catch
            {
                return "Erro ao destacar aprovados.";
            }
        }

        public static string DestacarApenasReprovados()
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
                    var v = sheet.Cells[headerRow, c].Value?.ToString();
                    if (v == null) continue;

                    if (IgualIgnorandoAcentos(v, "média")) colMedia = c;
                    if (IgualIgnorandoAcentos(v, "situação")) colSit = c;
                }

                if (colMedia == -1 || colSit == -1)
                    return "É necessário calcular a média e criar a coluna situação.";

                List<int> selectedRows;
                lock (_selectionLock)
                    selectedRows = _selectedRows.OrderBy(r => r).ToList();

                int applied = 0;
                if (selectedRows.Count > 0)
                {
                    foreach (int row in selectedRows)
                    {
                        if (row <= headerRow) continue;
                        if (sheet.Cells[row, headerCol].Value == null) continue;

                        // Mantém o estado existente; apenas marca/aplica quando corresponde.
                        var sitCell = (Excel.Range)sheet.Cells[row, colSit];

                        double media = sheet.Cells[row, colMedia].Value2 ?? 0;
                        if (media < 10)
                        {
                            sitCell.Value2 = "Reprovado";
                            sitCell.Interior.Color = ColorTranslator.ToOle(Color.LightCoral);
                            applied++;
                        }
                    }

                    return selectedRows.Count == 1
                        ? $"Reprovados destacados na linha selecionada."
                        : $"Reprovados destacados nas {selectedRows.Count} linhas selecionadas.";
                }

                int scanRow = headerRow + 1;
                while (sheet.Cells[scanRow, headerCol].Value != null)
                {
                    // Mantém o estado existente; apenas marca/aplica quando corresponde.
                    var sitCell = (Excel.Range)sheet.Cells[scanRow, colSit];

                    double media = sheet.Cells[scanRow, colMedia].Value2 ?? 0;
                    if (media < 10)
                    {
                        sitCell.Value2 = "Reprovado";
                        sitCell.Interior.Color = ColorTranslator.ToOle(Color.LightCoral);
                        applied++;
                    }
                    scanRow++;
                }

                return "Reprovados destacados (sem ocultar restantes).";
            }
            catch
            {
                return "Erro ao destacar reprovados.";
            }
        }

private static string CloseExcel()
{
    try
    {
        app.DisplayAlerts = false;
        workbook.SaveAs(pathFinal);
        workbook.Close(false);
        app.Quit();

        // _excelReady = false; // Remova ou ajuste conforme necessário, pois não existe no contexto apresentado
        return "Excel fechado com sucesso.";
    }
    catch
    {
        return "Erro ao fechar o Excel.";
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