using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using miLibreria;
namespace conversorExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Menu miMenu = new Menu("opcion 1", "opcion 2");
            int opcion;
            do
            {
                opcion = miMenu.MostrarMenu();
                switch (opcion)
                {
                    case 1:
                        
                        break;
                    case 2:
                        
                        break;

                    default:
                        
                        break;
                }
            } while (opcion != 3);
        }
    }
}
