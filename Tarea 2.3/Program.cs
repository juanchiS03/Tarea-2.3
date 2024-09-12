using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.ComponentModel;

class Program
{
    static void Main(string[] args)
    {
        // Ruta del archivo CSV de entrada
        string filePath = @"C:\Users\juansanchez\Downloads\atp_matches_2015.csv";

        // Configuración del lector CSV
        var csvConfig = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ",",
            HasHeaderRecord = true,
        };

        // Se lee el archivo CSV mediante un reader.
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, csvConfig))
        {
            // Variable que almacene todos los partidos leídos anteriormente.
            var partidos = csv.GetRecords<PartidoTenis>().ToList();

            // Mostrar el contenido por consola con tabulaciones como delimitador
            foreach (var partido in partidos)
            {
                // Capitalizar los nombres de los jugadores
                partido.winner_name = Capitalize(partido.winner_name);
                partido.loser_name = Capitalize(partido.loser_name);

                // Convertir fecha a formato "yyyy/MM/dd"
                partido.tourney_date = DateTime.ParseExact(partido.tourney_date, "yyyyMMdd", CultureInfo.InvariantCulture).ToString("yyyy/MM/dd");
            }

            // Cargar los datos en un archivo Excel
            string excelFilePath = @"C:\Users\juansanchez\Downloads\PartidosTransformados.xlsx";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets.Add("Partidos Transformados");

                // Escribir encabezados
                worksheet.Cells[1, 1].Value = "Tourney Name";
                worksheet.Cells[1, 2].Value = "Winner Name";
                worksheet.Cells[1, 3].Value = "Loser Name";
                worksheet.Cells[1, 4].Value = "Tourney Date";

                // Escribir datos transformados
                for (int i = 0; i < partidos.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = partidos[i].tourney_name;
                    worksheet.Cells[i + 2, 2].Value = partidos[i].winner_name;
                    worksheet.Cells[i + 2, 3].Value = partidos[i].loser_name;
                    worksheet.Cells[i + 2, 4].Value = partidos[i].tourney_date;
                }

                package.Save();
                Console.WriteLine("Datos cargados correctamente en el archivo Excel.");
            }
        }

        // Función para capitalizar el texto (primera letra mayúscula)
        static string Capitalize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            text = text.ToLower();
            return char.ToUpper(text[0]) + text.Substring(1);
        }
    }

    // Clase para mapear los datos del CSV
    public class PartidoTenis
    {
        // Torneo
        public string tourney_id { get; set; }
        public string tourney_name { get; set; }
        public string surface { get; set; }
        public int? draw_size { get; set; }
        public string tourney_level { get; set; }
        public string tourney_date { get; set; }

        // Partido
        public int match_num { get; set; }

        // Ganador
        public int? winner_id { get; set; }
        public int? winner_seed { get; set; }
        public string winner_entry { get; set; }
        public string winner_name { get; set; }
        public string winner_hand { get; set; }
        public int? winner_ht { get; set; }
        public string winner_ioc { get; set; }
        public double winner_age { get; set; }

        // Perdedor
        public int? loser_id { get; set; }
        public int? loser_seed { get; set; }
        public string loser_entry { get; set; }
        public string loser_name { get; set; }
        public string loser_hand { get; set; }
        public int? loser_ht { get; set; }
        public string loser_ioc { get; set; }
        public double loser_age { get; set; }

        // Otros detalles del partido
        public string score { get; set; }
        public int? best_of { get; set; }
        public string round { get; set; }
        public int? minutes { get; set; }

        // Estadísticas del ganador
        public int? w_ace { get; set; }
        public int? w_df { get; set; }
        public int? w_svpt { get; set; }
        public int? w_1stIn { get; set; }
        public int? w_1stWon { get; set; }
        public int? w_2ndWon { get; set; }
        public int? w_SvGms { get; set; }
        public int? w_bpSaved { get; set; }
        public int? w_bpFaced { get; set; }

        // Estadísticas del perdedor
        public int? l_ace { get; set; }
        public int? l_df { get; set; }
        public int? l_svpt { get; set; }
        public int? l_1stIn { get; set; }
        public int? l_1stWon { get; set; }
        public int? l_2ndWon { get; set; }
        public int? l_SvGms { get; set; }
        public int? l_bpSaved { get; set; }
        public int? l_bpFaced { get; set; }

        // Ranking
        public int? winner_rank { get; set; }
        public int? winner_rank_points { get; set; }
        public int? loser_rank { get; set; }
        public int? loser_rank_points { get; set; }
    }
}
