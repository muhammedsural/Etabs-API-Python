using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SAP2000v1;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //set the following flag to true to attach to an existing instance of the program

            //otherwise a new instance of the program will be started

            bool AttachToInstance;

            AttachToInstance = false;



            //set the following flag to true to manually specify the path to SAP2000.exe

            //this allows for a connection to a version of SAP2000 other than the latest installation

            //otherwise the latest installed version of SAP2000 will be launched

            bool SpecifyPath;

            SpecifyPath = false;



            //if the above flag is set to true, specify the path to SAP2000 below

            string ProgramPath;

            ProgramPath = @"C:\Program Files\Computers and Structures\SAP2000 22\SAP2000.exe";



            //full path to the model

            //set it to the desired path of your model

            string ModelDirectory = @"C:\CSiAPIexample";

            try

            {

                System.IO.Directory.CreateDirectory(ModelDirectory);

            }

            catch (Exception ex)

            {

                Console.WriteLine("Could not create directory: " + ModelDirectory);

            }

            string ModelName = "API_1-001.sdb";

            string ModelPath = ModelDirectory + System.IO.Path.DirectorySeparatorChar + ModelName;



            //dimension the SapObject as cOAPI type

            cOAPI mySapObject = null;



            //Use ret to check if functions return successfully (ret = 0) or fail (ret = nonzero)

            int ret = 0;



            //create API helper object

            cHelper myHelper;

            try

            {

                myHelper = new Helper();

            }

            catch (Exception ex)

            {

                Console.WriteLine("Cannot create an instance of the Helper object");

                return;

            }





            if (AttachToInstance)

            {

                //attach to a running instance of SAP2000

                try

                {

                    //get the active SapObject

                    mySapObject = myHelper.GetObject("CSI.SAP2000.API.SapObject");

                }

                catch (Exception ex)

                {

                    Console.WriteLine("No running instance of the program found or failed to attach.");

                    return;

                }

            }

            else

            {





                if (SpecifyPath)

                {

                    //'create an instance of the SapObject from the specified path

                    try

                    {

                        //create SapObject

                        mySapObject = myHelper.CreateObject(ProgramPath);

                    }

                    catch (Exception ex)

                    {

                        Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);

                        return;

                    }

                }

                else

                {

                    //'create an instance of the SapObject from the latest installed SAP2000

                    try

                    {

                        //create SapObject

                        mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");

                    }

                    catch (Exception ex)

                    {

                        Console.WriteLine("Cannot start a new instance of the program.");

                        return;

                    }

                }

                //start SAP2000 application

                ret = mySapObject.ApplicationStart();

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
