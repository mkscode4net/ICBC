using System;
using System.Collections.Generic;
using System.Text;
using ICBC.Utility;

namespace ICBC.BL.Base
{
    public class BaseBL
    {
        protected readonly ILoggerManager Logger;
        public BaseBL(ILoggerManager logger)
        {
            this.Logger = logger;
        }
    }
}
