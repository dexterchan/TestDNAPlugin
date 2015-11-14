using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TestWinForm
{
    public class DataResponse
    {
        public int row;
        public int col;
        public object[,] data;
    }
    public enum RunState { IDLE = 0, RUN };
}
