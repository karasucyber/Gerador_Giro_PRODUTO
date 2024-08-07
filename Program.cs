using System;
using System.Data;
using System.IO;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

public class ExcelExporter
{
    public void GetExcel(DateTime startDate, DateTime endDate)
    {
        // String de conexão com o banco de dados
        string connectionString = "Server=192.168.252.2;Database=SANKHYA_PRODUCAO;User Id=Sankhya;Password=tecsis;TrustServerCertificate=True;";
        
        // Consulta SQL com parâmetros para datas
        string query = @"
            SELECT 
                est.codemp,
                pro.referencia,
                pro.descrprod,
                CAST(est.RESERVADO AS INT) AS RESERVADO,
                CAST(est.ESTOQUE AS INT) AS ESTOQUE,
                CAST(SUM(CASE WHEN cab.codemp = '1' THEN ite.qtdneg ELSE 0 END) AS INT) AS VendaSP,
                CAST(SUM(CASE WHEN cab.codemp = '2' THEN ite.qtdneg ELSE 0 END) AS INT) AS VendaPB,
                CAST(SUM(CASE WHEN cab.codemp = '3' THEN ite.qtdneg ELSE 0 END) AS INT) AS VendaMG,
                CAST(SUM(CASE WHEN cab.codemp = '4' THEN ite.qtdneg ELSE 0 END) AS INT) AS VendaRN,
                CAST((SUM(ite.qtdneg) / 3) AS INT) AS MVD_30Dias,
                CAST(SUM(ite.qtdneg) AS INT) AS VTT_90Dias,
                CAST(est.ESTMIN AS INT) AS ESTMIN,
                CAST(est.ESTMAX AS INT) AS ESTMAX,
                CAST(ite.VLRUNIT AS INT) AS VLRUNIT, 
                cab.DTNEG
            FROM 
                TGFCAB cab
                JOIN TGFITE ite ON ite.nunota = cab.nunota
                JOIN TGFPRO pro ON pro.codprod = ite.codprod
                JOIN TGFEST est ON pro.codprod = est.codprod
            WHERE 
                cab.CODTIPOPER IN ('1100', '1118', '1119') 
                AND pro.CODGRUPOPROD NOT IN ('999000000', '998000000', '991000000', '990000000')
                AND cab.DTNEG >= @StartDate
                AND cab.DTNEG < @EndDate
                AND est.CODLOCAL NOT IN (9000000)
            GROUP BY 
                est.codemp,
                pro.referencia,
                pro.descrprod,
                est.RESERVADO,
                est.ESTOQUE,
                est.ESTMIN,
                est.ESTMAX,
                ite.VLRUNIT, 
                cab.DTNEG
            ORDER BY 
                pro.referencia;
        ";

        // Executa a consulta e processa os resultados
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(query, connection);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@StartDate", startDate);
            dataAdapter.SelectCommand.Parameters.AddWithValue("@EndDate", endDate);
            DataTable dataTable = new DataTable();
            try
            {
                connection.Open();
                dataAdapter.Fill(dataTable);
                // Criar e salvar o arquivo Excel
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Resultados");

                    // Adicionar os dados do DataTable à planilha
                    worksheet.Cell(2, 1).InsertTable(dataTable);

                    // Adicionar cabeçalhos personalizados
                    worksheet.Cell(1, 1).Value = "Código da Empresa";
                    worksheet.Cell(1, 2).Value = "Referência";
                    worksheet.Cell(1, 3).Value = "Descrição do Produto";
                    worksheet.Cell(1, 4).Value = "Reservado";
                    worksheet.Cell(1, 5).Value = "Estoque";
                    worksheet.Cell(1, 6).Value = "Venda SP";
                    worksheet.Cell(1, 7).Value = "Venda PB";
                    worksheet.Cell(1, 8).Value = "Venda MG";
                    worksheet.Cell(1, 9).Value = "Venda RN";
                    worksheet.Cell(1, 10).Value = "Média Vendas 30 Dias";
                    worksheet.Cell(1, 11).Value = "Vendas Totais 90 Dias";
                    worksheet.Cell(1, 12).Value = "Estoque Mínimo";
                    worksheet.Cell(1, 13).Value = "Estoque Máximo";
                    worksheet.Cell(1, 14).Value = "Valor Unitário";
                    worksheet.Cell(1, 15).Value = "Data de Negociação";

                    // Estilizar cabeçalhos
                    var headerRange = worksheet.Range("A1:O1");
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Ajustar a largura das colunas
                    worksheet.Columns().Width = 15;

                    // Caminho para a pasta Downloads do usuário
                    string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                    string filePath = Path.Combine(downloadsPath, "GiroProdutos.xlsx");

                    // Salvar o arquivo Excel
                    workbook.SaveAs(filePath);

                    Console.WriteLine($"Planilha salva em: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        DateTime startDate;
        DateTime endDate;

        Console.WriteLine("Escolha uma opção:");
        Console.WriteLine("1. Inserir datas manualmente");
        Console.WriteLine("2. Usar datas padrão (2024-01-01 a 2024-07-01)");
        
        int choice = int.Parse(Console.ReadLine());

        switch (choice)
        {
            case 1:
                Console.Write("Insira a data de início (aaaa-mm-dd): ");
                startDate = DateTime.Parse(Console.ReadLine());

                Console.Write("Insira a data de término (aaaa-mm-dd): ");
                endDate = DateTime.Parse(Console.ReadLine());
                break;

            case 2:
                startDate = new DateTime(2024, 1, 1);
                endDate = new DateTime(2024, 7, 1);
                break;

            default:
                Console.WriteLine("Opção inválida. Usando datas padrão.");
                startDate = new DateTime(2024, 1, 1);
                endDate = new DateTime(2024, 7, 1);
                break;
        }

        var exporter = new ExcelExporter();
        exporter.GetExcel(startDate, endDate);
    }
}
