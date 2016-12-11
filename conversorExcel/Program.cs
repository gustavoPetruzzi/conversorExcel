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
            String path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            String original;
            String nuevo;

            Menu miMenu = new Menu("opcion 1", "opcion 2");
            int opcion;
            excelConversor convertidor;
            do
            {
                opcion = miMenu.MostrarMenu();
                switch (opcion)
                {
                    case 1:
                        Console.Clear();
                        Console.WriteLine("Ingrese el nombre del archivo original");
                        original = Console.ReadLine();
                        Console.Clear();
                        Console.WriteLine("Ingrese el nombre del template");
                        nuevo = Console.ReadLine();
                        Console.Clear();
                        convertidor = new excelConversor(original, nuevo);
                        if (convertidor.copiarDatos(EBancos.cmf))
                        {
                            Console.WriteLine("Archivo Creado");
                            Console.ReadKey();
                        }
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
