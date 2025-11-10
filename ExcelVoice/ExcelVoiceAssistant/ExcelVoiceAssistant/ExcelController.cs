using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq; 


namespace ExcelVoiceAssistant
{
    public static class ExcelController
    {
        private static Excel.Application app;
        private static Excel.Workbook workbook;
        private static Excel.Worksheet sheet;

        private static string pathBase = @"C:\Users\trmbr\OneDrive\Desktop\IM\IM_EXCEL_NODEPENDENCIES\dados_turma.xlsx";
        private static string pathFinal = @"C:\Users\trmbr\OneDrive\Desktop\IM\IM_EXCEL_NODEPENDENCIES\Relatorio_Final.xlsx";
        public static void SetExcel(Excel.Application excelApp, Excel.Workbook wb, Excel.Worksheet ws)
        {
            app = excelApp;
            workbook = wb;
            sheet = ws;
        }


        // ===== ABRIR EXCEL =====
        public static void Inicializar()
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true; // Mostra o Excel
                workbook = app.Workbooks.Open(pathBase);
                sheet = workbook.Sheets[1];
                Console.WriteLine("✅ Excel aberto com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao abrir Excel: " + ex.Message);
            }
        }

        // ===== EXECUTAR COMANDO =====
        public static void Executar(string comando)
        {
            if (app == null)
                Inicializar();

            switch (comando.ToLower())
            {
                case "calcular média":
                    CalcularMedia();
                    break;

                case "destacar aprovados":
                    DestacarAprovados();
                    break;

                case "identificar melhoria":
                    IdentificarMelhoria();
                    break;

                case "inserir coluna situação":
                    InserirSituacao();
                    break;

                case "gerar gráfico":
                    GerarGrafico();
                    break;

                case "guardar relatório":
                    GuardarRelatorio();
                    break;

                default:
                    Console.WriteLine("⚠️ Comando não reconhecido: " + comando);
                    break;
            }
        }
        public static void CalcularMedia()
        {
            try
            {
                if (sheet == null)
                {
                    Console.WriteLine("⚠️ Folha não inicializada.");
                    return;
                }

                app.Visible = true;
                sheet.Activate();

                // 👉 Cria o cabeçalho da coluna Média
                sheet.Cells[1, "E"] = "Média";

                // 👉 Calcula médias usando a função local (MÉDIA em PT-PT)
                for (int i = 2; i <= 30; i++)
                {
                    if (sheet.Cells[i, "A"].Value != null)
                        sheet.Cells[i, "E"].FormulaLocal = $"=MÉDIA(B{i}:D{i})";
                }

                workbook.Save();
                app.CalculateFullRebuild();

                Console.WriteLine("📊 Coluna 'Média' criada e médias calculadas!");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao calcular médias: " + ex.Message);
            }
        }


        // ===== DESTACAR APROVADOS =====
        public static void DestacarAprovados()
        {
            Excel.Range range = sheet.Range["E2:E30"];
            range.FormatConditions.Delete();

            var condAprov = (Excel.FormatCondition)range.FormatConditions.Add(
                Excel.XlFormatConditionType.xlCellValue,
                Excel.XlFormatConditionOperator.xlGreaterEqual, "10");
            condAprov.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);

            var condReprov = (Excel.FormatCondition)range.FormatConditions.Add(
                Excel.XlFormatConditionType.xlCellValue,
                Excel.XlFormatConditionOperator.xlLess, "10");
            condReprov.Interior.Color = ColorTranslator.ToOle(Color.LightCoral);

            Console.WriteLine("✅ Aprovados e reprovados destacados.");
        }

        // ===== IDENTIFICAR MELHORIA =====
        public static void IdentificarMelhoria()
        {
            for (int i = 2; i <= 30; i++)
            {
                double t1 = sheet.Cells[i, "B"].Value ?? 0;
                double t2 = sheet.Cells[i, "C"].Value ?? 0;

                if (t1 > 0 && (t2 - t1) / t1 > 0.2)
                    sheet.Rows[i].Interior.Color = ColorTranslator.ToOle(Color.Yellow);
            }
            Console.WriteLine("📈 Alunos com melhoria identificados.");
        }

        // ===== INSERIR COLUNA “SITUAÇÃO” =====
        public static void InserirSituacao()
        {
            sheet.Cells[1, "F"] = "Situação";
            for (int i = 2; i <= 30; i++)
            {
                if (sheet.Cells[i, "E"].Value != null)
                    sheet.Cells[i, "F"].Formula = $"=IF(E{i}>=10,\"Aprovado\",\"Reprovado\")";
            }
            Console.WriteLine("📘 Coluna 'Situação' adicionada.");
        }

        // ===== GERAR GRÁFICO =====
        public static void GerarGrafico()
        {
            Excel.Range range = sheet.Range["A1:E30"];
            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            Excel.ChartObject chartObj = charts.Add(300, 20, 400, 300);
            Excel.Chart chart = chartObj.Chart;
            chart.ChartType = Excel.XlChartType.xlColumnClustered;
            chart.SetSourceData(range);
            chart.ChartTitle.Text = "Desempenho da Turma";
            Console.WriteLine("📊 Gráfico criado.");
        }

        // ===== GUARDAR RELATÓRIO FINAL =====
        public static void GuardarRelatorio()
        {
            try
            {
                workbook.SaveAs(pathFinal);
                Console.WriteLine("💾 Relatório guardado em: " + pathFinal);
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Erro ao guardar relatório: " + ex.Message);
            }
        }
    }
}
