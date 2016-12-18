using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace miLibreria
{
    public enum EBancos
    {
        cmf,
        finansur,
        otro
    }
    public class excelConversor
    {
        private String _path;
        private String _archivoOriginal;
        private String _archivoNuevo;

        public excelConversor(String archivoOriginal, String archivoNuevo)
        {
            this._path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            this._archivoNuevo = archivoNuevo;
            this._archivoOriginal = archivoOriginal;
            
        }
        public Boolean copiarDatos(EBancos banco)
        {
            Excel.Application excelApp = new Excel.Application();

            String fechaHoy = String.Format("{0:dd-MM-yy}", DateTime.Now.Date);
            try
            {
                Excel.Workbook original = excelApp.Workbooks.Open(this._path + "\\" + this._archivoOriginal);
                Excel.Workbook nuevo = excelApp.Workbooks.Open(this._path + "\\" + this._archivoNuevo);
                switch (banco)
                {
                    case EBancos.cmf:
                        original = this.formatColumn(original, 7);
                        this.copiarColumna(original, nuevo, 2, 5, 2);
                        this.copiarColumna(original, nuevo, 3, 6, 2);
                        this.copiarColumna(original, nuevo, 10, 7, 2);
                        this.copiarColumna(original, nuevo, 11, 8, 2);
                        this.copiarColumna(original, nuevo, 6, 9, 2);
                        this.copiarColumna(original, nuevo, 7, 10, 2);
                        this.copiarColumna(original, nuevo, 8, 11, 2);
                        this.copiarColumna(original, nuevo, 9, 12, 2);
                        break;
                    case EBancos.finansur:
                        this.copiarColumna(original, nuevo, 2, 1, 7);
                        this.copiarColumna(original, nuevo, 1, 2, 7);
                        this.copiarColumna(original, nuevo, 7, 3, 7);
                        this.copiarColumna(original, nuevo, 8, 4, 7);
                        this.copiarColumna(original, nuevo, 5, 5, 7);
                        this.copiarColumna(original, nuevo, 4, 6, 7);
                        break;
                    case EBancos.otro:
                        break;
                    default:
                        break;
                }
                original.Save();
                nuevo.SaveAs(this._path + "\\" + fechaHoy+ this._archivoOriginal );
                return true;
                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
                
            }
            finally
            {
                excelApp.Quit();
            }
            

        }
        private Excel.Workbook copiarColumna(Excel.Workbook original, Excel.Workbook nuevo, int columnaOriginal, int columnaNueva, int desdeFila)
        {
            Excel._Worksheet hojaOriginal = original.Sheets[1];
            Excel._Worksheet hojaNuevo = nuevo.Sheets[1];
            Excel.Range rangoOriginal = hojaOriginal.UsedRange;
            int cantidadFilas = rangoOriginal.Rows.Count;
            for (int i = 2; i < cantidadFilas; i++)
            {
                hojaNuevo.Cells[desdeFila, columnaNueva] = hojaOriginal.Cells[i, columnaOriginal];
                desdeFila++;
            }
            return nuevo;
        }
        private Excel.Workbook formatColumn(Excel.Workbook original, int columna)
        {
            Excel._Worksheet hojaOriginal = original.Sheets[1];
            hojaOriginal.Cells[1, 4].EntireColumn.NumberFormat = "####-##-##";
            Excel.Range rangoOriginal = hojaOriginal.UsedRange;
            int cantidadFilas = rangoOriginal.Rows.Count;
            for (int i = 2; i < cantidadFilas; i++)
            {
                hojaOriginal.Cells[i, columna].value2 = int.Parse(hojaOriginal.Cells[i, columna].value2);
            }

            return original;
        }
    }
}
