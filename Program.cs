//Camila Patricia Mata Gallegos
//Cristian Ledesma Ortiz
//Funcionalidad: Programa que maneja lógica difusa, a través del acceso de un archivo de
//excel, el usuario inserta los datos de los rangos que quiere usar junto con los valores
//a evaluar, como resultado el programa da los grados de exactitud de cada uno y el rango 
//seleccionado, también se imprimen las medias en la consola

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
            string Path = "Entrada.xlsx";
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
                        listaValores.Add(float.Parse(valorTexto));
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
                if (listaRangos[i].getMaximo() <= listaRangos[i + 1].getMinimo())
                {
                    //Console.WriteLine(listaRangos[i].getMaximo() + "y" + listaRangos[i+1].getMinimo());
                    return false;
                }

            }
            return true;
        }
        static void Evaluar(List<Rangos> listaRangos, List<float> listaValores)
        {
            string Pathresultado = "Resultado.xlsx";
            using (ExcelPackage rpackage = new ExcelPackage())
            {
                ExcelWorksheet hojar = rpackage.Workbook.Worksheets.Add("Resultados");
                List<object[]> salida = new List<object[]>();

                hojar.Cells[1, 1].Value = "T";
                for (int i = 0; i < listaRangos.Count; i++)
                {
                    hojar.Cells[1, i + 2].Value = listaRangos[i].getNombre();
                }
                hojar.Cells[1, listaRangos.Count + 2].Value = "δ";


                Calculo(salida, listaValores, listaRangos);
                for (int f = 0; f < salida.Count; f++)
                {
                    for (int c = 0; c < salida[f].Length; c++)
                    {
                        hojar.Cells[f + 2, c + 1].Value = salida[f][c];
                    }
                }

                hojar.Cells.AutoFitColumns();
                File.WriteAllBytes(Pathresultado, rpackage.GetAsByteArray());
                salida.Clear();
            }
            Console.WriteLine("Exito! Su archivo con los resultados se encuentran en el archivo Resultado.xlsx");
            
        }

        static void Calculo(List<object[]> listaS, List<float> listaV, List<Rangos> listaR)
        {
            float[] medias = new float[listaR.Count];
            Console.WriteLine("Medias");
            for (int j = 0; j < listaR.Count; j++)
            {
                medias[j] = (listaR[j].getMinimo() + listaR[j].getMaximo()) / 2;
                Console.WriteLine(medias[j]);
            }

            foreach (float valor in listaV)
            {
                //Console.WriteLine(valor);
                float[] grados = new float[listaR.Count];

                for (int i = 0; i < listaR.Count; i++)
                {
                    var rango = listaR[i];
                    if (valor < listaR[0].getMinimo() || valor > listaR.Last().getMaximo())
                    {
                        grados[i] = FuncionTrapezoidal(valor, rango.getMinimo(), medias[i], medias[i], rango.getMaximo());
                        //Console.WriteLine(grados[i]);
                    }
                    else
                    {
                        grados[i] = FuncionTriangulo(valor, rango.getMinimo(), medias[i], rango.getMaximo());
                        //Console.WriteLine(grados[i]);
                    }
                }

                float gradoMax = grados.Max();
                int rangop = Array.IndexOf(grados, gradoMax);

                string rangoe = listaR[rangop].getNombre();
                object [] fila = new object[listaR.Count+2];
                fila[0] = valor;
                for (int i = 0; i < listaR.Count; i++)
                {
                    fila[i + 1] = grados[i];
                }
                fila[listaR.Count+1] = rangoe;
                listaS.Add(fila);
            }
        }
        static float FuncionTrapezoidal(float x, float a, float b, float c, float d)
        {
            if (x <= a || x >= d) return 0;
            if (x > a && x < b) return (x - a) / (b - a);
            if (x >= b && x <= c) return 1;
            if (x > c && x < d) return (d - x) / (d - c);
            return 0;
        }

        static float FuncionTriangulo(float x, float a, float b, float c)
        {
            if (x<a) return 0;
            if (x >= a && x <= b) return (x - a) / (b - a);
            if (x >= b && x <= c) return (c - x) / (c - b);
            return 0;
        }

    }
}
