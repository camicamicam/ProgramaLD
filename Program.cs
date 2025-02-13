using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ProgramaLD
{
    public class Program
    {

        static void Main(string[] args)
        {
            string Path = "DATOS 2025.xlsx";
            if (!File.Exists(Path))
            {
                Console.WriteLine("El archivo no existe.");
                return;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            List<Rangos> listaRangos = new List<Rangos>();
            List<float> listaValores = new List<float>();

            using (var package = new ExcelPackage(new FileInfo(Path)))
            {
                var hoja = package.Workbook.Worksheets[0];
                int rowCount = hoja.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string nombre = hoja.Cells[row, 1].Text;
                    string minimoTexto = hoja.Cells[row, 2].Text;
                    string maximoTexto = hoja.Cells[row, 3].Text;
                    string valorTexto = hoja.Cells[row, 5].Text;

                    if (string.IsNullOrWhiteSpace(nombre) ||
                        string.IsNullOrWhiteSpace(minimoTexto) ||
                        string.IsNullOrWhiteSpace(maximoTexto) ||
                        string.IsNullOrWhiteSpace(valorTexto))
                    {
                        listaValores.Add(int.Parse(valorTexto));
                        continue;
                    }

                    if (!int.TryParse(minimoTexto, out int minimo) ||
                        !int.TryParse(maximoTexto, out int maximo) ||
                        !float.TryParse(valorTexto, out float valor))
                    {
                        continue;
                    }
                    listaRangos.Add(new Rangos(nombre, minimo, maximo));
                    listaValores.Add(valor);
                }
                //Imprimirlistas(listaRangos, listaValores);

                if (listaRangos.Count < 2 || listaRangos.Count > 4)
                {
                    
                    throw new Exception("Debes tener entre minimo 2 y maximo 4 rangos");
                }

                if (!verificarRangos(listaRangos))
                {
                    throw new Exception("No existe traslape en sus rangos, por favor vuelva a escribirlos");
                }
                Evaluar(listaRangos, listaValores);

            }
        }

        static void Imprimirlistas(List<Rangos> listaRangos, List<float> listaValores)
        {
            Console.WriteLine("Rangos:");
            foreach (var rango in listaRangos)
            {
                Console.WriteLine($"{rango.getNombre()}: [{rango.getMinimo()}, {rango.getMaximo()}]");
            }
            Console.WriteLine("\nValores:");
            foreach (var valor in listaValores)
            {
                Console.WriteLine(valor);
            }
        }

        static bool verificarRangos(List<Rangos> listaRangos)
        {
            for (int i = 0; i < listaRangos.Count - 1; i++)
            {
                if (listaRangos[i].getMaximo() <= listaRangos[i+1].getMinimo())
                {
                    //Console.WriteLine(listaRangos[i].getMaximo() + "y" + listaRangos[i+1].getMinimo());
                    return false;
                }

            }
            return true;
        }
        static void Evaluar(List<Rangos> listaRangos, List<float> listaValores)
        {
            Console.WriteLine("Sucess!");
        }
    }
}