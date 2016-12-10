using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace miLibreria
{
    public class Menu
    {
        private List<String> _opciones;
        public Menu(params string[] opciones)
        {
            this._opciones = new List<string>();
            for (int i = 0; i < opciones.Length; i++)
            {
                this._opciones.Add(opciones[i]);
            }
        }
        

        public int MostrarMenu()
        {
            int nroOpcion = 1;
            int resultado;
            foreach (String opcion in this._opciones)
            {
                Console.WriteLine("{0}.{1}", nroOpcion, opcion);
                nroOpcion++;
            }
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("{0}. Salir", nroOpcion);
            Console.ForegroundColor = ConsoleColor.White;
            int.TryParse(Console.ReadLine(), out resultado);
            return resultado;
        }
    }
}
