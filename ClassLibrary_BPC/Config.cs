﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassLibrary_BPC
{
    public class Config
    {

        static public string Database = "HRM";
        static public string Userid = "sa";

        static public string Server = ".\\SQLEXPRESS";
        static public string Password = "Sql2019*";

        //static public string Server = ".\\SQL2019E";
        //static public string Password = "2019";

        static public string FormatDateSQL = "MM/dd/yyyy";

        static public string PathFileImport = "D:\\Temp\\HR365";
        static public string PathFileExport = "D:\\Temp\\HR365\\Export";

    }
}
