using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using _Excel = Microsoft.Office.Interop.Excel;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using System.Reflection;

namespace AU.MySoftware
{
    public class Initialization : IExtensionApplication
    {
        static double areaC, perimC, areaP, perimP;

        #region Commands
        [CommandMethod("Inovar")]
        public static void Inovar()
        {

            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();

                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;

                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        if (acSSObj != null)
                        {
 
                            if (acSSObj.ObjectId.ObjectClass.DxfName == "CIRCLE")
                            {
                                Circle acCircle = (Circle) acTrans.GetObject(acSSObj.ObjectId,
                                                                OpenMode.ForRead);

                                areaC = acCircle.Area;
                                perimC = acCircle.Circumference;

                            }
                            if (acSSObj.ObjectId.ObjectClass.DxfName == "LWPOLYLINE")
                            {
                                Polyline acPolyline = (Polyline)acTrans.GetObject(acSSObj.ObjectId,
                                                                OpenMode.ForRead);

                                areaP = acPolyline.Area;
                                perimP = acPolyline.Length;
                            }
                        }
                    }
                }
            }
            SalvarExcel(areaC, perimC, areaP, perimP);
        }

        public static void SalvarExcel(double areaC, double perimC, double areaP, double perimP)
        {
            string[,] valores = new string[2, 3];

            valores[0, 0] = "Circunferência";
            valores[0, 1] = areaC.ToString();
            valores[0, 2] = perimC.ToString();
            valores[1, 0] = "Polígono";
            valores[1, 1] = areaP.ToString();
            valores[1, 2] = perimP.ToString();

            ExportarExcel(valores);
        }

        public static void ExportarExcel(string[,] valores)
        {
            _Excel.Application oExcel;
            _Excel._Workbook oWB;
            _Excel._Worksheet oWS;
            _Excel.Range oRng;


            //Start excel
            oExcel = new _Excel.Application();
            oExcel.Visible = true;

            //Get a new workbook.
            oWB = (_Excel._Workbook)(oExcel.Workbooks.Add(Missing.Value));
            oWS = (_Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oWS.Cells[1, 1] = "Objeto";
            oWS.Cells[1, 2] = "Área";
            oWS.Cells[1, 3] = "Perímetro";

            //Format A1:C1 as bold, vertical alignment = center.
            oWS.get_Range("A1", "C1").Font.Bold = true;
            oWS.get_Range("A1", "C1").VerticalAlignment =
            _Excel.XlVAlign.xlVAlignCenter;


            //Fill A2:C6 with an array of values.
            oWS.get_Range("A2", "C3").Value2 = valores;

            //AutoFit columns A:D.
            oRng = oWS.get_Range("A1", "C1");
            oRng.EntireColumn.AutoFit();

            //Manipulate a variable number of columns for Quarterly Sales Data.
            //DisplayQuarterlySales(oWS);

            //Make sure Excel is visible and give the user control
            //of Microsoft Excel's lifetime.
            oExcel.Visible = true;
            oExcel.UserControl = true;
        }

        #endregion

        #region Initialization

        void IExtensionApplication.Initialize()
        {

        }

        void IExtensionApplication.Terminate()
        {

        }
        #endregion


    }
}