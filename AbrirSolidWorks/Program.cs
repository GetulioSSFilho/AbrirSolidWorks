using System;
using System.Runtime.InteropServices;

namespace AbrirSolidWorks
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //add solidworks reference to project
            SldWorks.SldWorks swApp;

            
            try//If there is no solidworks instance open, 
            {
                //This function takes an open solidworks instance
                swApp = Marshal.GetActiveObject("SldWorks.Application") as SldWorks.SldWorks;//To use, a solidworks instance must be open.

            }
            catch (Exception ex)
            {
                
                //we will catch the error and create a new instance
                if (ex.HResult == -2147221021)
                {
                    swApp = Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")) as SldWorks.SldWorks;//Its opened in background
                    swApp.Visible = true;//make the instance visible.
                }
            }
        }
    }
}
