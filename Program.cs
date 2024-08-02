using System;
using System.Data;
using Microsoft.Data.SqlClient; // Atualize aqui
using ClosedXML.Excel; // Para manipulação de arquivos Excel

namespace SqlQueryToExcel
{
    class Program
    {
        static void Main(string[] args)
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
            using (SqlConnection connection = new SqlConnection(connectionString)) // Atualize aqui
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
                        worksheet.Cell(1, 1).InsertTable(dataTable);

                        // Salvar o arquivo Excel
                        string filePath = @"C:\Users\Kopermax\Desktop\Teste\Resultados.xlsx"; // Altere o caminho conforme necessário
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
}
