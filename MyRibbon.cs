using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using static System.Runtime.CompilerServices.RuntimeHelpers;


namespace RibbonStartCopy
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {

        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);

        public void AutoOpen()
        {
            
        }
        public override string GetCustomUI(string RibbonID)
        {
            
            AppDomain.CurrentDomain.FirstChanceException += FirstChanceException;
            return RibbonResources.Ribbon;
        }

        public override object LoadImage(string imageId)
        {
            // This will return the image resource with the name specified in the image='xxxx' tag
            return RibbonResources.ResourceManager.GetObject(imageId);
        }

        public void OnButton2Pressed(IRibbonControl control)
        {
            MessageBox.Show("This is a test");
            for(int i = 0; i < 200; i++)
            {
                Console.WriteLine(i.ToString());
            }
            MessageBox.Show("Test finished");
        }

        public void OnButtonPressed(IRibbonControl control)
        {
            string tempDirPath = System.IO.Path.GetTempPath();
            string fileExt = ".png";
            string tempxlsFile = Path.Combine(tempDirPath, Path.ChangeExtension(Guid.NewGuid().ToString(), fileExt));
            string tempxlsFile2 = Path.Combine(tempDirPath, Path.ChangeExtension(Guid.NewGuid().ToString(), ".xlsm"));
            //RunHelper("c82e0385-ef5b-4e10-a872-a68e36de5f9e", tempxlsFile);
            //AzureControl.GetAzureDataAsync("ConnectionTest.xlsm", tempxlsFile);
            AzureControl.GetAzureDataAsyncOmni("ConnectionTest.xlsm", tempxlsFile2);
            //ExcelAsyncUtil.QueueAsMacro(() => RunOperationSafely());
            //await _lock.WaitAsync();
            //try
            //{

            //}
            //catch (Exception ex) { _lock.Release(); }
        }

        [SecurityCritical]
        private static void FirstChanceException(object sender, FirstChanceExceptionEventArgs e)
        {
            if (e != null)
            {
                //System.Windows.Forms.MessageBox.Show("Error!");
            }
        }
    }
}


