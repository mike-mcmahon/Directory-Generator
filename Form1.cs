using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace PM_Folder_Generator
{
    public partial class Form1 : Form
    {
        //Variables for storing the text entered by the user as strings.  
        //These variables are for the project name and Q2C order number.
        String projectNameString = ""; //This is the project name entered by the user in the GUI textbox
        String Q2CNumberAsString = ""; //This is the new Q2C order number entered by the user in the GUI textbox
        String directoryQ2CSeries = ""; // REFER TO METHOD doesSeriesFolderExist(String Q2CNumber).  This generates the directory
                               // 'xx000000_Series' where xx = the first two digits of the Q2C order number.
                               // This is used in the above method for testing if this directory exists or not, and subsequently for
                               // creating this directory if it doesn't already exist.
        String basePath = @"M:\Application Centre\Orders"; //This is the base path (absolute/fully-qualified) 
                                                                                        //for the generation of the prject directory.
        String secondaryPath = ""; // basePath\directoryQ2CSeries....secondaryPath - this is the point where the project specific folders
                                    //will be generated.
        String projectPath = ""; //project specific path

        //Partial class constructor.
        public Form1()
        {
            InitializeComponent();
        }


        //Event handler method for assigning the Q2C number entered by the user to the variable Q2CNumberAsString.
        private void EnterQ2COrderNumber(object sender, EventArgs e)
        {
            //Create the Q2C textbox object 
            TextBox Q2COrderNumberTextBox = sender as TextBox;
            //Assign the textbox's contents to Q2CNumberAsString
            Q2CNumberAsString = Q2COrderNumberTextBox.Text;
        }


        //Event handler method for assigning the project name entered by the user to the variable projectNameString.
        private void EnterProjectName(object sender, EventArgs e)
        {
            //Create the project name textbox object
            TextBox projectNameTextBox = sender as TextBox;
            //Assign the textbox's contents to projectNameString
            projectNameString = projectNameTextBox.Text;
        }


        //Event handler for when the CREATE button is selected/clicked.
        //Calls method validQ2CNumber to validate the Q2C number.
        //Calls method TBA that verifies if series folder exists, and if not creates one.
        //Calls method doesProjectDirectoryExist to verify if the project directory exists.
        //Calls method createProjectDirectory if no project directory exists for valid Q2C number.
        //@param - EVENT HANDLER for createFoldersButton_CLICK sender
        //@return - void
        private void createFoldersButton_Click(object sender, EventArgs e)
        {
            if(validQ2CNumber(Q2CNumberAsString)) //Test to make sure the Q2C number entered in the textbox is a valid number.
            {
                //Check to see if series directory exists, and if not create it.
                seriesFolderCheckAndCreate(Q2CNumberAsString);

                //If the Q2C number is a valid number, test to see if the project folder exists.
                if (doesProjectDirectoryExist(Q2CNumberAsString))
                {
                    MessageBox.Show("Project Folder Already Exists");
                    Q2COrderNumberTextBox.Text = ""; //Blank out the Q2C number textbox
                    projectNameTextBox.Text = ""; //Blank out the project name textbox
                }
                else 
                {
                    //call a method that generates the project folders at secondaryPath ..\Orders\xx000000_Series\
                    createProjectDirectory();
                    //call method to copy templates, resource files into new directory
                    copyFiles();
                    MessageBox.Show("New Project Folders Now Generated");
                    Q2COrderNumberTextBox.Text = ""; //Blank out the Q2C number textbox
                    projectNameTextBox.Text = ""; //Blank out the project name textbox
                    
                    //Added February 19, 2015
                    //Open windows explorer
                    String command = "explorer.exe";
                    String argument = projectPath;
                    Process.Start(command, argument);
                    //Close this application
                    Application.Exit();
                }    
            }
            else {
                    MessageBox.Show("Please enter a valid 8 Digit Q2C number."); //Message user that the Q2C number entered in the textbox is not valid.
                    Q2COrderNumberTextBox.Text = ""; //Blank out the Q2C number textbox...forces user to re-enter the number.
                  }  
        }


        //This method verifies that the Q2C number entered is valid.
        //The valid Q2C number is eight (8) characters long and all numeric digits.
        //The method returns true if the Q2C number is valid, else false.
        //@param - String Q2CNumber; this is Q2CNumberAsString.
        //@return - isValidQ2CNumber; returns true or false if Q2C number is valid.
        private bool validQ2CNumber(String Q2CNumber)
        {
            bool isValidQ2CNumber = false; //Variable for use on method return

            //Test to make sure the Q2C number contains all numeric digits
            //We assume the Q2C number enter is all digits, and thus set isAllDigits true.
            //isAllDigits is set false is any character in the string is not a digit.
            bool isAllDigits = true;
            for(int i = 0; i < Q2CNumber.Length; i++)
            {
                if(!Char.IsDigit(Q2CNumber, i))
                {
                    isAllDigits = false;
                }
            }

            //Test to make sure the Q2C number is 8 digits in length
            bool isEightDigits = false;
            if(Q2CNumber.Length == 8)
            {
                isEightDigits = true;
            }

            //If Q2C Number is all digits and is 8 characters in length
            //then set isValidQ2CNumber true.
            if(isAllDigits && isEightDigits)
            {
                isValidQ2CNumber = true;
            }
            return isValidQ2CNumber;
        }


        //Method checks for the existance of the series directory folder and 
        //if it doesn't exist it creates one.
        //@param - Q2CNumber; this is Q2CNumberAsString.
        //@return - void
        private void seriesFolderCheckAndCreate(String Q2CNumber)
        {
            //Parse first two digits of new Q2C number and concatinate them to "000000_Series"
            directoryQ2CSeries = Q2CNumber.Substring(0, 2);
            directoryQ2CSeries = directoryQ2CSeries + "000000_Series"; //The variable directoryQ2CSeries is now ready to be used as a part of the project directory path.                                                   
            secondaryPath = Path.Combine(basePath, directoryQ2CSeries); //secondaryPath - this is the point where the project specific folders will be generated.

            //Test to see if the series directory path exists at
            //drive:\Project Management Folder Project\Test Folders\Orders
            String[] seriesPath = Directory.GetDirectories(basePath + @"\"); //creates an array of all the directories in ..\Orders\
            bool doesSeriesExist = false; //assume series directory doesn't exist
            for (int i = 0; i < seriesPath.Length; i++)
            {
                if (secondaryPath.Equals(seriesPath[i]))
                {
                    doesSeriesExist = true; //if series directory exists, set doesSeriesExist true
                }
            }
            if (!doesSeriesExist) //If the series directory doesn't exist, create the new directory
            {
                Directory.CreateDirectory(secondaryPath); //creates a new series directory based on the Q2C order number.
            }
        }

        //This method checks to see if the project folder exists.
        //@param - String Q2CNumber which represents Q2CNumberAsString
        //@return - bool doesProjectExist which indicates whether the project directory existed or not.
        private bool doesProjectDirectoryExist(String Q2CNumber)
        {
            //Test to see if the project directory path exists at
            //drive:\Project Management Folder Project\Test Folders\Orders\xx000000_Series\
            String[] projectDirectory = Directory.GetDirectories(secondaryPath); //Get all the subDirectories in the secondary path

            //Add Q2C# to secondary path. 
            String projectPathToCompare = secondaryPath + @"\" + Q2CNumber;
            bool doesProjectExist = false; //Assume the project directory doesn't exist
            for(int i = 0; i < projectDirectory.Length; i++)
            {
                if(projectPathToCompare.Equals(projectDirectory[i].Substring(0, projectPathToCompare.Length))) //if projectPathToCompare exists in projectDirectory
                {                                                    //this means the project folder exists.  Set doesProjectExist to true.
                    doesProjectExist = true;
                }
            }
            return doesProjectExist; //returns to method createFoldersButton_Click
        }

        //Method creates the project directory specifically for the new user supplied Q2C order 
        //number and project name.  
        //@param - n/a
        //@return - void
        private void createProjectDirectory()
        {
            //Create the parent project directory after the secondaryPath
            projectPath = Path.Combine(secondaryPath, Q2CNumberAsString + " - " + projectNameString);

            //Create the first level of project specific directory folders
            String[] projectFirstLevelPath = {"1 Bid", "2 Order Conversion", "3 Drawings", "4 Release From Customer", "5 Change Orders",
                                        "6 Communications", "7 Schedule", "8 Post Shipment Close-Out", "9 Issue Log - Risk Register",
                                        "10 Engineering Studies"};

            for (int i = 0; i < projectFirstLevelPath.Length; i++)
            {
                Directory.CreateDirectory(Path.Combine(projectPath, projectFirstLevelPath[i]));
            }

            //create sub folders for specific first level project folders
            String[] subPathOrderConversion = {"Customer PO", "Final Documents for Construction", "Hand-Off Meeting", "Order Entry Coversheet",
            "Kick-Off Meeting"};
            for (int i = 0; i < subPathOrderConversion.Length; i++)
            {
                Directory.CreateDirectory(Path.Combine(projectPath, @"2 Order Conversion\" + subPathOrderConversion[i]));
            }

            String[] subPathDrawings = { "Approval Gate 2", "Approval Gate 3", "Record" };
            for (int i = 0; i < subPathDrawings.Length; i++)
            {
                Directory.CreateDirectory(Path.Combine(projectPath, @"3 Drawings\" + subPathDrawings[i]));
            }

            String[] subPathCommunications = { "Customer and Supplier", "Financial Project Review", "Schneider Electric" };
            for (int i = 0; i < subPathCommunications.Length; i++)
            {
                Directory.CreateDirectory(Path.Combine(projectPath, @"6 Communications\" + subPathCommunications[i]));
            }

            String[] subPathPostShip = { "ETO Warranty Claim Credit", "Financial Close Out", "Manuals", "Post Mortem" };
            for (int i = 0; i < subPathPostShip.Length; i++)
            {
                Directory.CreateDirectory(Path.Combine(projectPath, @"8 Post Shipment Close-Out\" + subPathPostShip[i]));
            }
        }


        //Method copies resource files/templates to the newly created directory.
        //@param - n/a
        //@return - void
        private void copyFiles()
        {
            //***********************************************************************************************************
            //CUSTOMER SAT TOOL REMOVED FROM THIS ROUTINE TEMPORARILY AS THE COPY FUNCTION IS VERY SLOW OVER THE NETWORK
            //***********************************************************************************************************

            //Source file locations for resource files
            String resourcePath = @"S:\Application Centre\Resource Templates";
            String ETOWarrantyTemplatePath = Path.Combine(resourcePath, "ETO_Warranty_Form_CSMP-580A_rev37.xls"); //Path to ETO Warranty Form
            String serviceToSiteTemplatePath = Path.Combine(resourcePath, "Send_Service_CC_Form_CSMP-580A_rev36.xls"); //Path to Send Service to Site Form
            //String cusSatToolTemplatePath = Path.Combine(resourcePath, "CusSAT_Tool_rev53.xls"); //Path to CustSat Tool form
            String hydroTransmittalTemplatePath = Path.Combine(resourcePath, "hydro transmittal.doc"); //Path to hydro transmittal template
            String drawingTransmittalTemplatePath = Path.Combine(resourcePath, "Drg Transmittal.doc"); //Path to drawing transmittal template

            //The following strings are the paths to various Excel/Word templatess used for customer satisfaction, warranty claims
            //and communication transmittals.
            String copyOfETOWarrantyPath = Path.Combine(projectPath, @"8 Post Shipment Close-Out\ETO Warranty Claim Credit\ETO_Warranty_Form_CSMP-580A_rev37.xls"); // Path to ETO Warranty form
            String copyOfSendServiceToSitePath = Path.Combine(projectPath, @"8 Post Shipment Close-Out\ETO Warranty Claim Credit\Send_Service_CC_Form_CSMP-580A_rev36.xls"); //Path to Send Service to Site form
            //String copyOfCusSatToolPath = Path.Combine(projectPath, @"8 Post Shipment Close-Out\CusSAT_Tool_rev53.xls"); //Customer Sat Tool
            String copyOfHydroTransmittalPath = Path.Combine(projectPath, @"3 Drawings\hydro transmittal.doc");//hydro tranmittal
            String copyOfDrawingTransmittalPath = Path.Combine(projectPath, @"3 Drawings\Drg Transmittal.doc");//drawing transmittal

            //The following copies the resource files to the appropriate directories in the new project directory.
            File.Copy(ETOWarrantyTemplatePath, copyOfETOWarrantyPath);
            File.Copy(serviceToSiteTemplatePath, copyOfSendServiceToSitePath);
            //File.Copy(cusSatToolTemplatePath, copyOfCusSatToolPath);
            File.Copy(hydroTransmittalTemplatePath, copyOfHydroTransmittalPath);
            File.Copy(drawingTransmittalTemplatePath, copyOfDrawingTransmittalPath);
        }
        

        
    }
        
}

