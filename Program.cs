using System;
using System.Data;
using System.IO;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;

public class ExcelExporter
{
    public void GetExcel()
    {
        // String de conexão com o banco de dados
        string connectionString = "Server=192.168.252.2;Database=SANKHYA_PRODUCAO;User Id=Sankhya;Password=tecsis;TrustServerCertificate=True;";

        // Consulta SQL
        string query = @"
            SELECT 
                cab.DTNEG,
                est.codemp,
                pro.codprod,
                est.CODLOCAL,
                est.RESERVADO,
                pro.descrprod,
                pro.referencia,
                ite.VLRUNIT,
                SUM(ite.qtdneg) AS total_qtdneg,
                SUM(CASE WHEN cab.codemp = '1' THEN ite.qtdneg ELSE 0 END) AS total_qtdneg_Emp1,
                SUM(CASE WHEN cab.codemp = '2' THEN ite.qtdneg ELSE 0 END) AS total_qtdneg_Emp2,
                SUM(CASE WHEN cab.codemp = '3' THEN ite.qtdneg ELSE 0 END) AS total_qtdneg_Emp3,
                SUM(CASE WHEN cab.codemp = '4' THEN ite.qtdneg ELSE 0 END) AS total_qtdneg_Emp4,
                (SUM(ite.qtdneg) / 3) AS media_3_mes,
                est.ESTMIN,
                est.ESTOQUE,
                est.ESTMAX
            FROM 
                TGFCAB cab
                JOIN TGFITE ite ON ite.nunota = cab.nunota
                JOIN TGFPRO pro ON pro.codprod = ite.codprod
                JOIN TGFEST est ON pro.codprod = est.codprod
            WHERE 
                cab.CODTIPOPER IN ('1100', '1118', '1119') 
                AND pro.CODGRUPOPROD NOT IN ('999000000', '998000000', '991000000', '990000000')
                AND cab.DTNEG > '2024-01-01'
                AND cab.DTNEG < '2024-07-01'
                AND est.CODLOCAL NOT IN (9000000)
            GROUP BY 
                cab.DTNEG,
                est.codemp,
                pro.codprod, 
                pro.descrprod, 
                pro.referencia, 
                pro.CODGRUPOPROD,
                est.ESTMIN, 
                est.ESTOQUE, 
                est.ESTMAX, 
                est.CODLOCAL, 
                est.RESERVADO,
                ite.VLRUNIT
            ORDER BY 
                pro.codprod;
        ";

        // Executa a consulta e processa os resultados
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(query, connection);
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
                    worksheet.Cell(1, 1).Value = "Data de Negociação";
                    worksheet.Cell(1, 2).Value = "Código da Empresa";
                    worksheet.Cell(1, 3).Value = "Código do Produto";
                    worksheet.Cell(1, 4).Value = "Código do Local";
                    worksheet.Cell(1, 5).Value = "Reservado";
                    worksheet.Cell(1, 6).Value = "Descrição do Produto";
                    worksheet.Cell(1, 7).Value = "Referência";
                    worksheet.Cell(1, 8).Value = "Valor Unitário";
                    worksheet.Cell(1, 9).Value = "Total Quantidade Negociada";
                    worksheet.Cell(1, 10).Value = "Total Quantidade Negociada Empresa 1";
                    worksheet.Cell(1, 11).Value = "Total Quantidade Negociada Empresa 2";
                    worksheet.Cell(1, 12).Value = "Total Quantidade Negociada Empresa 3";
                    worksheet.Cell(1, 13).Value = "Total Quantidade Negociada Empresa 4";
                    worksheet.Cell(1, 14).Value = "Média 3 Meses";
                    worksheet.Cell(1, 15).Value = "Estoque Mínimo";
                    worksheet.Cell(1, 16).Value = "Estoque";
                    worksheet.Cell(1, 17).Value = "Estoque Máximo";

                    // Estilizar cabeçalhos
                    var headerRange = worksheet.Range("A1:Q1");
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Ajustar a largura das colunas
                    worksheet.Columns().AdjustToContents();

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
        var exporter = new ExcelExporter();
        exporter.GetExcel();
    }
}

