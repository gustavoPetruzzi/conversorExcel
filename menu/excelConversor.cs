using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace menu
{
    public enum EBancos
    {
        cmf,
        otro
    }
    public class excelConversor
    {
        private String _path;
        private EBancos _banco;
        
        
        public excelConversor(String archivo, EBancos banco)
        {
            this._path = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + archivo;
            
            this._banco = banco;
        }
        public void copiarDatos()
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook original = excelApp.Workbooks.Open(this._path);
            Excel.Workbook nuevo = excelApp.Workbooks.Add();
            switch (this._banco)
            {
                case EBancos.cmf:

                    break;
                case EBancos.otro:
                    break;
                default:
                    break;
            }

        }
        private Excel.Workbook copiarColumna(Excel.Workbook original, Excel.Workbook nuevo, int columna, int desdeFila)
        {
            Excel._Worksheet hojaOriginal = original.Sheets[1];
            Excel._Worksheet hojaNuevo = nuevo.Sheets[1];
            Excel.Range rangoOriginal = hojaOriginal.UsedRange;
            int cantidadFilas = rangoOriginal.Rows.Count;
            for (int i = 2; i < cantidadFilas; i++)
            {

            }
            return nuevo;
        }
    }
}
