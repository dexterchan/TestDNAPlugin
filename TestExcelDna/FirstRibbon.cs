using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;
using System.Threading;
using TestWinForm;

namespace TestExcelDna
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        DataResponse cr = new DataResponse();
        TestWinForm.TestInputForm ctrl = new TestWinForm.TestInputForm();
        object[,] result = null;
        //bool waitSet = false;
        ExcelReference refCell = null;

         [ExcelFunction(IsMacroType = true)] 
        public void OnButtonPressed(IRibbonControl control)
        {
            //MessageBox.Show("Hello from control " + control.Id);
            

            string str=FirstAddIn.MyGetHostname();

            var UIHandler = new Action<object>((o) =>
            {
                ctrl.ShowDialog();
            });

            
             

             //if (waitSet) //avoid double submission
             //{

             //    ThreadPool.QueueUserWorkItem(new WaitCallback(UIHandler));
             //    return;
             //}
             //else
             //{
             //    waitSet = true;
             //}

            
            
            //var wait = new ManualResetEvent(false);

            var handler = new EventHandler((o, e) =>
            {
                cr = (DataResponse)o;
                result = cr.data;
                //waitSet = false;
                //wait.Set();

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    //ExcelReference cell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);
                    ExcelReference cell = refCell;
                    int testRowSize = cr.row;
                    int testColSize = cr.col;

                    var activeCell = new ExcelReference(cell.RowFirst, testRowSize + cell.RowFirst - 1, cell.ColumnLast, cell.ColumnLast + testColSize - 1);

                    activeCell.SetValue(result);
                    XlCall.Excel(XlCall.xlcSelect, activeCell);

                });

            });

            if (this.ctrl.getRunState() == RunState.IDLE)
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    refCell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);
                });
                ctrl.registerCallback(handler);
            }
            


            ThreadPool.QueueUserWorkItem(new WaitCallback(UIHandler));

            ////For simplicity, we implement the wait here
            //wait.WaitOne();
            //if (result == null)
            //{
            //    return;
            //}

            
            //ExcelAsyncUtil.QueueAsMacro(() =>
            //{
            //    //ExcelReference cell = (ExcelReference)XlCall.Excel(XlCall.xlfActiveCell);
            //    ExcelReference cell = refCell;
            //    int testRowSize = cr.row;
            //    int testColSize = cr.col;

            //    var activeCell = new ExcelReference(cell.RowFirst,testRowSize+cell.RowFirst-1, cell.ColumnLast ,cell.ColumnLast + testColSize-1);
               


            //    activeCell.SetValue(result);
            //    XlCall.Excel(XlCall.xlcSelect, activeCell);
                
            //});
        }

        public static void ShowHelloMessage()
        {
            MessageBox.Show("Hello from 'ShowHelloMessage'.");
        }

        
    }


    public static class MyFunctions
    {
        public static string TestFunction()
        {
            return "Testing...OK";
        }

    }
}
