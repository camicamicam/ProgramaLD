//Camila Patricia Mata Gallegos
//Cristian Ledesma Ortiz
//Funcionalidad: Programa que maneja lógica difusa, a través del acceso de un archivo de
//excel, el usuario inserta los datos de los rangos que quiere usar junto con los valores
//a evaluar, como resultado el programa da los grados de exactitud de cada uno y el rango 
//seleccionado, también se incluyen las medias en el archivo

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
                List<float> Medias = new List<float>();

                Calculo(salida, listaValores, listaRangos, Medias);
                hojar.Cells[1, 1].Value = "Medias";
                hojar.Cells[2, 1].Value = "T";
                for (int i = 0; i < listaRangos.Count; i++)
                {
                    hojar.Cells[1, i + 2].Value = Medias[i];
                    hojar.Cells[2, i + 2].Value = listaRangos[i].getNombre();
                }
                hojar.Cells[2, listaRangos.Count + 2].Value = "δ";



                for (int f = 0; f < salida.Count; f++)
                {
                    for (int c = 0; c < salida[f].Length; c++)
                    {
                        hojar.Cells[f + 3, c + 1].Value = salida[f][c];
                    }
                }

                hojar.Cells.AutoFitColumns();
                File.WriteAllBytes(Pathresultado, rpackage.GetAsByteArray());
                salida.Clear();
            }
            Console.WriteLine("Exito! Su archivo con los resultados se encuentran en el archivo Resultado.xlsx");

        }

        static void Calculo(List<object[]> listaS, List<float> listaV, List<Rangos> listaR, List<float> listaM)
        {
            for (int j = 0; j < listaR.Count; j++)
            {
                listaM.Add(((float)listaR[j].getMinimo() + (float)listaR[j].getMaximo()) / 2);
            }

            foreach (float valor in listaV)
            {
                //Console.WriteLine(valor);
                float[] grados = new float[listaR.Count];

                for (int i = 0; i < listaR.Count; i++)
                {
                    var rango = listaR[i];
                    // Determinamos si el rango es extremo:
                    bool extremoIzquierdo = false;
                    bool extremoDerecho = false;
                    string nombre = rango.getNombre();
                    if (nombre == listaR[0].getNombre())
                        extremoIzquierdo = true;
                    else if (nombre == listaR.Last().getNombre())
                        extremoDerecho = true;

                    grados[i] = CalcularMembresia(valor, rango.getMinimo(), rango.getMaximo(), extremoIzquierdo, extremoDerecho, listaM, i);
                }

                float gradoMax = grados.Max();
                int rangop = Array.IndexOf(grados, gradoMax);

                string rangoe = listaR[rangop].getNombre();
                object[] fila = new object[listaR.Count + 2];
                fila[0] = valor;
                for (int i = 0; i < listaR.Count; i++)
                {
                    fila[i + 1] = grados[i];
                }
                fila[listaR.Count + 1] = rangoe;
                listaS.Add(fila);
            }
        }
        static float CalcularMembresia(float x, float min, float max, bool extremoIzquierdo, bool extremoDerecho, List<float> medias, int i)
        {
            if (extremoIzquierdo && x <= min)
                return 1;
            if (extremoDerecho && x >= max)
                return 1;

            if (x <= min || x >= max)
                return 0;

            if (x <= medias[i])
                return (x - min) / (medias[i] - min);
            else
                return (max - x) / (max - medias[i]);
        }

    }
}
