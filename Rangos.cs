using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ProgramaLD
{
    public class Rangos
    {
        private string nombre;
        public int minimo;
        public int maximo;
        public Rangos(string nombre, int minimo, int maximo)
        {
            this.nombre = nombre;
            this.minimo = minimo;
            this.maximo = maximo;
        }
        public string getNombre()
        {
            return this.nombre;
        }
        public int getMinimo()
        {
            return this.minimo;
        }
        public int getMaximo()
        {
            return this.maximo;
        }
    }
}