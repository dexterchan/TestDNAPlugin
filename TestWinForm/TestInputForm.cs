using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace TestWinForm
{
    public partial class TestInputForm : Form
    {
        private EventHandler sendRequest;
        //public GetRowColTmpl getRowCol;
        //private bool IsSubmittedResult = false;
        DataResponse r = null;

        

        RunState myState = RunState.IDLE;

        public TestInputForm()
        {
            InitializeComponent();
        }

        public RunState getRunState()
        {
            return this.myState;
        }
        
        //public delegate void GetRowColTmpl(ref int row, ref int col, object[,] obj);

        
        public void registerCallback(EventHandler req)
        {
            myState = RunState.IDLE;
            sendRequest = req;
        }

        private void btnGo_Click(object sender, EventArgs e)
        {
            myState = RunState.RUN;
            r = new DataResponse();
            r.row = Int32.Parse( this.txtRow.Text);
            r.col = Int32.Parse(this.txtCol.Text);

            ThreadPool.QueueUserWorkItem(this.runCalculation, r);
            //this.Close();
        }

        public void runCalculation(Object o)
        {
            DataResponse r = (DataResponse)o;
            Thread.Sleep(1000 * 20);
            r.data = MakeArrayetest(r.row, r.col);
            sendRequest(r, null);

            myState = RunState.IDLE;
        }

        public static object[,] MakeArrayetest(int rows, int columns)
        {
            object[,] result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
            {
                result[i, 0] = i;
                for (int j = 1; j < columns - 1; j++)
                {
                    result[i, j] = string.Format("({0},{1})", i, j);
                }
                result[i, columns - 1] = DateTime.Now;
            }

            return result;
        }

        private void UserControl1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //if (!IsSubmittedResult)
            //{
            //    sendRequest(null, null);
            //}
            //else
            //{
            //    sendRequest(r, null);
            //}
        }
    }
}
