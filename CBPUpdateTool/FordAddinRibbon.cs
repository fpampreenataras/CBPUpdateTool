using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using MyExcel = Microsoft.Office.Interop.Excel;
using Core = Microsoft.Office.Core;
using Aras.IOM;
using ArasOC = OfficeConnector;
using ArasOCUtilities = OfficeConnector.Utils;
using Forms = System.Windows.Forms;
using Settings = CBPUpdateTool.Properties.Settings;
using System.Xml;


//This add-in is for a SALT excel file.
//It will do the following.
//  1. verify the range for CBP names is in the excel file.  If not end and tell user.
//  2. Get all names of x4_CBP ItemTypes in a List
//  3. Get current row of CPB names in a List
//  4. Update the row of CBP Names.
//  3. Get all the Parameters of the associate document and populate options for each CBP Name.
//      Don't know how to do that yet.
namespace CBPUpdateTool
{

    public partial class FordAddinRibbon
    {
        Innovator inn;
        Core.DocumentProperties properties = null; 
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //login to innovator somehow.  Using the excel file cells on a hidden sheet for now.
          

           //add the event to catch loading a new file form Aras so that we can get the files workflow and activity names
            Globals.FordAddin.Application.WorkbookOpen += new MyExcel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            Globals.FordAddin.Application.WorkbookBeforeClose += new MyExcel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookClose);
        }

        private void Application_WorkbookClose(MyExcel.Workbook Wb, ref bool Cancel)
        {
            if (!Wb.FullName.Contains("mso_viablecopy"))  // hack to get around office connector opening and closing document when doing a Save to Aras
            {
                WorkflowNameRibbonLabel.Label = "";
                ActivityRibbonLabel.Label = "";
            }
        }

        //populate workflow and activity
        void Application_WorkbookOpen(MyExcel.Workbook Wb)
        {
            
            if ((inn = ConnecttoAras) != null)
            {
                properties = Globals.FordAddin.Application.ActiveWorkbook.CustomDocumentProperties;

                //update workflow ribbon label
                UpdateWorkFlowRibbonName(properties[Settings.Default.arasPrimaryLinkItemId].Value);

                //update activity ribbon label
                UpdateActivityRibbonName(properties[Settings.Default.arasPrimaryLinkItemId].Value);
            }
        }


       

        private void UpdateFromAras_Click(object sender, RibbonControlEventArgs e)
        {
            
            MyExcel.Workbook ThisWorkbook = Globals.FordAddin.Application.ActiveWorkbook; ;
            MyExcel.Sheets ThisWorkSheets = ThisWorkbook.Worksheets;
            //get inn from office connector would be great
            //0  Login into Aras
            if ((inn = ConnecttoAras) != null)
            {
                //  1. verify the range for CBP names is in the excel file.  If not end and tell user.
                String RangeNamestr = Settings.Default.CPB_Name_Range;
                MyExcel.Worksheet ThisWorkSheet = Globals.FordAddin.Application.ActiveSheet;
                MyExcel.Range NameRange2 = VerifyNameRange(ThisWorkbook, RangeNamestr);
                if (NameRange2 == null)
                {
                    Forms.MessageBox.Show("No name range for " + RangeNamestr, "Missing Name Range");
                    return;
                }


                //  2. Get all names of x4_CBP ItemTypes and populate the CBP Name range.
                Item CBPs = inn.applyAML(Settings.Default.AML_to_get_all_CBPs);
                int count = CBPs.getItemCount();
                List<MyCBPItem> CBPItems = new List<MyCBPItem>();
                MyCBPItem CBP_Item;
                for (int i = 0; i < count; i++)
                {
                    Item item = CBPs.getItemByIndex(i);
                    CBP_Item = new MyCBPItem()
                    {
                        name = item.getProperty(Settings.Default.CBP_Name_Property_Title),
                        ID = item.getProperty(Settings.Default.CBP_DOCID_Property_Title)
                    };
                    if (!string.IsNullOrEmpty(CBP_Item.name) && !string.IsNullOrEmpty(CBP_Item.ID)) { CBPItems.Add(CBP_Item); }
                }
                //3 Go through CBPItems and see if they already exist in Excel
                //3.1 update current items
                int CBPNamerangestart;
                long CBPNamerangelastRow;
                foreach (MyCBPItem CBP_item in CBPItems) //head lamps
                {

                    CBPNamerangestart = NameRange2.Row + 1;
                    CBPNamerangelastRow = ThisWorkSheet.Cells[ThisWorkSheet.Rows.Count, NameRange2.Column].End(MyExcel.XlDirection.xlUp).Row + 1;
                    for (int i = CBPNamerangestart; i < CBPNamerangelastRow; i++)
                    {
                        string CBPNameinExcel = ThisWorkSheet.Cells[i, NameRange2.Column].Value.ToString(); //seats
                        List<MyCBPItem> foundinCBPitems = CBPItems.FindAll(x => x.name == CBPNameinExcel);

                        if (foundinCBPitems.Count != 0 && !foundinCBPitems[0].Found)
                        {
                            //update the validtion for the cell one column over.
                            UpdateColumnValidation(ThisWorkSheet.Cells[CBPNamerangelastRow, NameRange2.Column + 1], CBP_item.ID, Properties.Settings.Default.numberofArchitecturs);

                            // Put descriptions in column corresponding to row of CBPName
                            UpdateDescriptionfields(ThisWorkbook.Worksheets[Settings.Default.ArchetictureWorksheetName], CBP_item.ID, Properties.Settings.Default.numberofArchitecturs, CBPNamerangelastRow);

                            CBP_item.Found = true;
                            break;
                        }
                        else if (foundinCBPitems.Count == 0)
                        {
                            HighlightCell(ThisWorkSheet.Cells[i, NameRange2.Column]);
                        }
                        else if (foundinCBPitems[0].Found)
                        {
                            continue;

                        }
                    }
                }
                //3.2 Add new items
                foreach (MyCBPItem CBP_item in CBPItems) //head lamps
                {
                    if (!CBP_item.Found)
                    {
                        //update date the CBP name column
                        CBPNamerangestart = NameRange2.Row + 1;
                        CBPNamerangelastRow = ThisWorkSheet.Cells[ThisWorkSheet.Rows.Count, NameRange2.Column].End(MyExcel.XlDirection.xlUp).Row + 1;
                        ThisWorkSheet.Cells[CBPNamerangelastRow, NameRange2.Column].Value = CBP_item.name;
                        //update the validtion for the cell one column over.
                        UpdateColumnValidation(ThisWorkSheet.Cells[CBPNamerangelastRow, NameRange2.Column + 1], CBP_item.ID, Settings.Default.numberofArchitecturs);


                        // Put descriptions in column corresponding to row of CBPName
                        UpdateDescriptionfields(ThisWorkbook.Worksheets[Settings.Default.ArchetictureWorksheetName], CBP_item.ID, Settings.Default.numberofArchitecturs, CBPNamerangelastRow);

                        CBP_item.Found = true;
                        //6 add the vlookup to the description column
                        ThisWorkSheet.Cells[CBPNamerangelastRow, NameRange2.Column + 2].Value = "=VLOOKUP(RC[-1],Architectures!R[-" + (CBPNamerangelastRow - 1).ToString() + "]C[" + ((CBPNamerangelastRow * 2) - (NameRange2.Column + 2)).ToString() + "]:R[" + (NameRange2.Column - (CBPNamerangelastRow - CBPNamerangestart)).ToString() + "]C[" + ((CBPNamerangelastRow * 2) - (NameRange2.Column + 2) + 1).ToString() + "],2,0)";
                    }


                }

                //3.3 Hightlight out of date items
                CBPNamerangestart = NameRange2.Row + 1;
                CBPNamerangelastRow = ThisWorkSheet.Cells[ThisWorkSheet.Rows.Count, NameRange2.Column].End(MyExcel.XlDirection.xlUp).Row + 1;
                for (int i = CBPNamerangestart; i < CBPNamerangelastRow; i++)
                {
                    string CBPNameinExcel = ThisWorkSheet.Cells[i, NameRange2.Column].Value.ToString(); //seats
                    List<MyCBPItem> foundinCBPitems = CBPItems.FindAll(x => x.name == CBPNameinExcel);
                    if (foundinCBPitems.Count == 0)
                    {
                        HighlightCell(ThisWorkSheet.Cells[i, NameRange2.Column]);
                    }
                }
            }

        }

        private void UpdateDescriptionfields(MyExcel.Worksheet ArchSheet, string CBP_ID, int numberofarchitectures, long CBPNamerangelastRow)
        {
            Item CBPInstance = inn.getItemById("Document", CBP_ID);
            for (int iii = 1; iii <= numberofarchitectures; iii++)
            {
                ArchSheet.Cells[iii, CBPNamerangelastRow * 2] = CBPInstance.getProperty(Settings.Default.CBPArchname + iii);
                ArchSheet.Cells[iii, (CBPNamerangelastRow * 2) + 1] = CBPInstance.getProperty(Settings.Default.CBPArchDesc + iii);
            }
        }

        private void UpdateColumnValidation(MyExcel.Range cell, string CBP_ID, int numberofarchitectures)
        {
            cell.Validation.Delete();

            Item CBPInstance = inn.getItemById("Document", CBP_ID);
            //3 Get the architectures from the document Itemtype
            string ArchNames = "";
           
            for (int ii = 1; ii <= numberofarchitectures; ii++)
            {
                ArchNames += CBPInstance.getProperty(Settings.Default.CBPArchname + ii) + ",";
            }
            ArchNames = ArchNames.Remove(ArchNames.Length - 1);
            //4 Add the validation, that only allowes selection of provided values.
            cell.Validation.Add(MyExcel.XlDVType.xlValidateList, MyExcel.XlDVAlertStyle.xlValidAlertStop, MyExcel.XlFormatConditionOperator.xlBetween, ArchNames, Type.Missing);

            cell.Validation.IgnoreBlank = true;
            cell.Validation.InputTitle = "";
            cell.Validation.InputMessage = "";
        }

        private void HighlightCell(MyExcel.Range cell)
        {
            cell.Interior.Pattern = MyExcel.XlPattern.xlPatternGray50;
            cell.Interior.PatternColor = 192;
            cell.Interior.Color = 8420607;
            cell.Interior.PatternTintAndShade = 0;
        }

        private MyExcel.Range VerifyNameRange(MyExcel.Workbook wb, string RangeName)
        {
            MyExcel.Names names = wb.Names;
            foreach (MyExcel.Name name in names)
            {
                if (name.Name == RangeName)
                {
                    return wb.ActiveSheet.Range[RangeName];
                }
            }
            return null;
        }

        private Innovator ConnecttoAras
        {
            get
            {
                string url = "";
                string db = "";
                string windowauth = "";
                //string user = "";
                XmlDocument config = new XmlDocument();
                string path = Environment.GetEnvironmentVariable("Appdata") + "\\OfficeConnector\\config.xml";
                if (System.IO.File.Exists(path))
                {
                    config.Load(path);
                    url = config.ChildNodes[0].FirstChild.Attributes[0].Value;
                    db = config.ChildNodes[0].FirstChild.Attributes[1].Value;
                    windowauth = config.ChildNodes[0].FirstChild.Attributes[3].Value;
                    if (windowauth == "True")
                    {
                        WinAuthHttpServerConnection winnconn = IomFactory.CreateWinAuthHttpServerConnection(url, db);
                        Item login_result = winnconn.Login();
                        if (login_result.isError())
                        {
                            Forms.MessageBox.Show("Login failed.  Check configfile for login information ", "Login Failed");
                            return null;
                        }
                        return IomFactory.CreateInnovator(winnconn);
                    }
                    else  //Office connector set up but windowauth not set up so prompt user for login.
                    {
                        //prompt with OC login
                        // still needs work
                        //OfficeConnector.Dialogs.DialogFactory dialogfactory = new ArasOC.Dialogs.DialogFactory();
                        //dialogfactory.GetLoginDialog();

                        //Dictionary<string, string> logindict = new Dictionary<string, string>();

                        //OfficeConnector.Dialogs.LoginDialog loginDialog = new ArasOC.Dialogs.LoginDialog(logindict, new ArasOC.OfficeApps.OfficeApp(),ArasOC.Configurations.IConfigurationStorage()) ;
                        //loginDialog.ShowDialog();
                        ////ArasOC.OfficeApps.OfficeApp loginofficeapp;
                        ////ArasOC.Configurations.IConfigurationStorage iconfig = new 

                        //////OfficeConnector.Dialogs.LoginDialog OCLoginDialog = new ArasOC.Dialogs.LoginDialog(logindict, loginofficeapp,);
                        //////OCLoginDialog.ShowDialog();
                        //////url = OCLoginDialog.Url;
                        ////HttpServerConnection conn = IomFactory.CreateHttpServerConnection(OCLoginDialog.Url, OCLoginDialog.DB, OCLoginDialog.UserName, OCLoginDialog.Password);
                        ////return IomFactory.CreateInnovator(conn);
                        Forms.MessageBox.Show("Login failed.  This Add-in only works with Windows Authentication active", "Login Failed");
                        return null;
                    }
                }
                else  //Office Connector not set up so prompt for Aras Login
                {

                    //HttpServerConnection conn = IomFactory.CreateHttpServerConnection(url, db, user, password);
                    //Item login_result = conn.Login();
                    //if (login_result.isError())
                    //{
                    //    //prompt with OC login
                    //    OfficeConnector.Dialogs.LoginDialog OCLoginDialog = new ArasOC.Dialogs.LoginDialog();
                    //    OCLoginDialog.ShowDialog();
                    //    url = OCLoginDialog.Url;

                    //}

                    Forms.MessageBox.Show("Login failed.  This Add-in only works with Windows Authentication active and Office Connector enabled", "Login Failed");
                    //return IomFactory.CreateInnovator(conn);
                    return null;
                }
            }
        }

        private void Complete_Activity_click(object sender, RibbonControlEventArgs e)
        {
            
            if ((inn = ConnecttoAras)  != null)
            {
                
                ActivityForm AF = new ActivityForm()
                {
                    docid = properties[Settings.Default.ArasDocumentIDCPName].Value,
                    inn = inn,
                    primarylinkedid = properties[Settings.Default.arasPrimaryLinkItemId].Value
                };

                AF.ShowDialog();

                //update workflow ribbon label
                UpdateWorkFlowRibbonName(properties[Settings.Default.arasPrimaryLinkItemId].Value);

                //update activity ribbon label
              UpdateActivityRibbonName(properties[Settings.Default.arasPrimaryLinkItemId].Value);
            }
        }

        private void UpdateWorkFlowRibbonName(string source_id)
        {
            //update workflow ribbon label
            Item Workflow = inn.applyAML(Settings.Default.getWorkFlowAML.Replace("<source_id>", "<source_id>" + source_id));
            if (Workflow.node != null)
            {
                string Workflowprocessid = Workflow.getProperty("related_id");
                Item WorkFlowProcess = inn.applyAML(Settings.Default.getWorkFlowProcessAML.Replace("<id></id>", "<id>" + Workflowprocessid + "</id>"));
                string WorkflowName = inn.applyAML(Settings.Default.getWorkflowMapNameAML.Replace("<id></id>", "<id>" + WorkFlowProcess.getProperty("copied_from_string") + "</id>")).getProperty("name");
                WorkflowNameRibbonLabel.Label = WorkflowName;
            }
          
        }

        private void UpdateActivityRibbonName(string source_id)
        {
            Item Workflow = inn.applyAML(Settings.Default.getWorkFlowAML.Replace("<source_id>", "<source_id>" + source_id));
            if (Workflow.node != null)
            {
                string Workflowprocessid = Workflow.getProperty("related_id");
                Item WorkFlowProcess = inn.applyAML(Settings.Default.getWorkFlowProcessAML.Replace("<id></id>", "<id>" + Workflowprocessid + "</id>"));
                string WorkflowName = inn.applyAML(Settings.Default.getWorkflowMapNameAML.Replace("<id></id>", "<id>" + WorkFlowProcess.getProperty("copied_from_string") + "</id>")).getProperty("name");

                Item WorkFlowProcessActivities = inn.applyAML(Settings.Default.getActivitiesAML.Replace("<source_id>", "<source_id>" + Workflowprocessid));
                string endactivity = "";
                for (int i = 0; i < WorkFlowProcessActivities.getItemCount(); i++)
                {
                    Item WorkFlowProcessActivity = WorkFlowProcessActivities.getItemByIndex(i);
                    Item Activity = WorkFlowProcessActivity.getPropertyItem("related_id");
                    string currentstate = Activity.getPropertyAttribute("current_state", "keyed_name");
                    endactivity = Activity.getProperty("is_end") == "1" && string.IsNullOrEmpty(endactivity) ? Activity.getID() : endactivity;
                    if (currentstate == "Active")
                    {
                        string currentactivity = Activity.getPropertyAttribute("config_id", "keyed_name");
                        ActivityRibbonLabel.Label = currentactivity;
                        CompleteTaskButton.Enabled = true;
                        return;
                    }
                }
                //if code gets here and didn't break out then end of workflow
                ActivityRibbonLabel.Label = inn.getItemById("Activity", endactivity).getPropertyAttribute("config_id", "keyed_name");
                CompleteTaskButton.Enabled = false;
            }
        }

        public class MyCBPItem
        {
            public string name = "";
            public string ID = "";
            private Boolean found = false;

            public bool Found { get => found; set => found = value; }
        }
    }

}
