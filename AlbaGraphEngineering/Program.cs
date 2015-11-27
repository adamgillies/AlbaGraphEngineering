using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Reflection;
using System.IO;

namespace AlbaGraphEngineering
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            EmbeddedAssembly.Load("AlbaGraphEngineering.Resources.AlbaLibrary.dll", "AlbaLibrary.dll");
            EmbeddedAssembly.Load("AlbaGraphEngineering.Resources.ZedGraph.dll", "ZedGraph.dll");
            EmbeddedAssembly.Load("AlbaGraphEngineering.Resources.ClosedXML.dll", "ClosedXML.dll");
            //EmbeddedAssembly.Load("AlbaGraphEngineering.DocumentFormat.OpenXml.dll", "DocumentFormat.OpenXml.dll");
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new GraphForm());
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name);
        }
    }
}
