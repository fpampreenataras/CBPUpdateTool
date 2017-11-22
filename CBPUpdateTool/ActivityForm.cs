using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aras.IOM;
using Settings = CBPUpdateTool.Properties.Settings;
using System.Xml;
using System.IO;

namespace CBPUpdateTool
{
    public partial class ActivityForm : Form
    {
        public string docid;
        public string primarylinkedid;
        public Innovator inn;
    
        private string currentActivityID = "";
        private string loggedinuserID = "";
        public ActivityForm()
        {
            InitializeComponent();
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ActivityForm_Load(object sender, EventArgs e)
        {
            //1 Get the WorkFlow Name based on the docid is related to primarylinkedid (CBP) populate form
            Item Workflow = inn.applyAML(Settings.Default.getWorkFlowAML.Replace("<source_id>", "<source_id>" + primarylinkedid));
            string Workflowprocessid = Workflow.getProperty("related_id");
            Item WorkFlowProcess = inn.applyAML(Settings.Default.getWorkFlowProcessAML.Replace("<id></id>", "<id>" + Workflowprocessid + "</id>"));
            string WorkflowName = inn.applyAML(Settings.Default.getWorkflowMapNameAML.Replace("<id></id>", "<id>" + WorkFlowProcess.getProperty("copied_from_string") + "</id>")).getProperty("name");
            WorkflowNameLbl.Text = WorkflowName;

            //Get current Activity.  Get list of all activities and see which one is active
            Item WorkFlowProcessActivities = inn.applyAML(Settings.Default.getActivitiesAML.Replace("<source_id>", "<source_id>" + Workflowprocessid));
            for (int i = 0; i < WorkFlowProcessActivities.getItemCount(); i++)
            {
                Item WorkFlowProcessActivity = WorkFlowProcessActivities.getItemByIndex(i);
                Item Activity = WorkFlowProcessActivity.getPropertyItem("related_id");
                string currentstate = Activity.getPropertyAttribute("current_state", "keyed_name");
                if (currentstate == "Active")
                {
                    string currentactivity = Activity.getPropertyAttribute("config_id", "keyed_name");
                    ActivityNameLabel.Text = currentactivity;
                    currentActivityID = Activity.getID();
                    break;
                }
            }

            //Get the tasks to complete based on the currentActivityID
            Item tasks = inn.applyAML(Settings.Default.getTasksAML.Replace("<source_id>", "<source_id>" + currentActivityID));
            for (int i = 0; i < tasks.getItemCount(); i++)
            {

                Item task = tasks.getItemByIndex(i);

                string sequence = task.getProperty("sequence", "");
                string is_required = task.getProperty("is_required");
                string description = task.getProperty("description");
                //add to taskdatagrid
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(TaskDataGridView);
                row.Cells[0].Value = sequence;
                row.Cells[1].Value = is_required == "1" ? 1 : 0;
                row.Cells[2].Value = description;
                row.Cells[3].Value = 0;
                row.Cells[4].Value = task.getID();
                TaskDataGridView.Rows.Add(row);
            }

            //get list of voting options.
            VoteComboBox.Items.Clear();
            Item votes = inn.applyAML(Settings.Default.getVotePaths.Replace("<source_id>", "<source_id>" + currentActivityID));
            for (int i = 0; i < votes.getItemCount(); i++)
            {
                Item vote = votes.getItemByIndex(i);
                string votedescription = vote.getProperty("name");
                VoteComboBox.Items.Add(votedescription);
                Item activity = vote.getPropertyItem("related_id");
                if (activity.getProperty("can_delegate") == "1")
                {
                    VoteComboBox.Items.Add("Delegate");
                }
                if (activity.getProperty("can_refuse") == "1")
                {
                    VoteComboBox.Items.Add("Refuse");
                }
            }

            //get user information this is used when completing the activity
            string UserID = inn.getUserID();
            loggedinuserID = inn.applyAML(Settings.Default.getLoggedInIdentityID.Replace("<id>", "<id>" + inn.getUserID())).getProperty("owned_by_id");


        }

        private void Complete_Click(object sender, EventArgs e)
        {  
            //get activity assignment
            Item ActivityAssignments = inn.applyAML(Settings.Default.getActivityAssignement.Replace("<source_id>", "<source_id>" + currentActivityID));

          
            string activityassignmentid = "";
            for (int i = 0; i < ActivityAssignments.getItemCount(); i++)
            {
                //get the related_id (idendity) and see if current user is a member.  Propogate down if member is not an alias
                Item ActivityAssignment = ActivityAssignments.getItemByIndex(i);
                string name = ActivityAssignment.getProperty("keyed_name");
                if (is_assigned(ActivityAssignment))
                {
                    activityassignmentid = ActivityAssignment.getID();
                    break;
                }

            }
            if (string.IsNullOrEmpty(activityassignmentid))
            {
                System.Windows.Forms.MessageBox.Show("User is not assigned to this task.","Activity Completion warning");
                this.Close();

            }

            string xmlString = null;
            using (StringWriter sw = new StringWriter())
            {
                XmlTextWriter writer = new XmlTextWriter(sw);
                writer.Formatting = Formatting.Indented; // if you want it indented

                //writer.WriteStartDocument(); // <?xml version="1.0" encoding="utf-16"?>

                writer.WriteStartElement("AML"); //<AML>
                //< Item type = "Activity" action = "EvaluateActivity" >
                writer.WriteStartElement("Item"); //<Item>
                writer.WriteStartAttribute("type");
                writer.WriteString("Activity");  
                writer.WriteEndAttribute();
                writer.WriteStartAttribute("action");
                writer.WriteString("EvaluateActivity");
                writer.WriteEndAttribute();

                // <Activity>value</Activity>
                writer.WriteStartElement("Activity");
                writer.WriteString(currentActivityID);
                writer.WriteEndElement();

                // <ActivityAssignment>value</ActivityAssignment>
                writer.WriteStartElement("ActivityAssignment");
                writer.WriteString(activityassignmentid);
                writer.WriteEndElement();

                // <Path>value</Path>>
                writer.WriteStartElement("Paths");
                writer.WriteStartElement("Path");
                writer.WriteString(VoteComboBox.SelectedItem.ToString());
                writer.WriteEndElement();
                writer.WriteEndElement();

                //Task id = '' completed = "1" ></ Task >
                // <SUBTAG attr="hello">world</SUBTAG>
                writer.WriteStartElement("Tasks");
                DataGridViewRowCollection rows = TaskDataGridView.Rows;
                foreach (DataGridViewRow row in rows)
                {
                    writer.WriteStartElement("Task");
                    writer.WriteStartAttribute("id");
                    writer.WriteString(row.Cells[4].Value.ToString());
                    writer.WriteEndAttribute();
                    writer.WriteStartAttribute("completed");
                    writer.WriteString("1");
                    writer.WriteEndAttribute();
                    writer.WriteString(TaskDataGridView.Rows[0].Cells[4].Value.ToString());
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();

                // <comments>value</comments>>
                writer.WriteStartElement("Comment");
                writer.WriteString(Comments.Text);
                writer.WriteEndElement();

                 // <Authentication>value</Authentication>>
                writer.WriteStartElement("Authentication");
                writer.WriteStartAttribute("mode");
                writer.WriteString("");
                writer.WriteEndAttribute();
                writer.WriteEndElement();

                // <Complete>value</Complete >>  
                writer.WriteStartElement("Complete");
                writer.WriteString("1");
                writer.WriteEndElement();

                // <Variables value</Variable >>
                writer.WriteStartElement("Variables");
                writer.WriteString("1");
                writer.WriteEndElement();

                writer.WriteEndElement(); //<ITEM>
                writer.WriteEndElement();//</AML>

                xmlString = sw.ToString();

            }
            Item completedItem = inn.applyAML(xmlString);
            this.Close();
        }

        private bool is_assigned(Item activity)
        {
            Item identity = activity.getPropertyItem("related_id");
            string name = identity.getProperty("name");
            Boolean found = false;
            if (is_member(identity, ref found)){
                return true;

            }
            return false;
        }

        private bool is_member(Item identity,  ref bool found)
        {
            identity.fetchRelationships("Member");
            Item Members = identity.getRelationships();
            for (int i = 0; i < Members.getItemCount(); i++)
            {
                Item member = Members.getItemByIndex(i);
                Item memberidentity = member.getPropertyItem("related_id");
                if (String.IsNullOrEmpty(memberidentity.getProperty("is_alias"))) // a group and not a user so recursively keep looking
                {
                    found = is_member(member, ref found);
                    if (found) return true;
                }
                else  // a user so check it.
                {
                    string identityid = memberidentity.getID();
                    found = (identityid == loggedinuserID);
                    if (found) return true;
                };

            }
            return false;
        }

        private void CheckBox_Click(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewCell cell = ((DataGridView)sender).Rows[e.RowIndex].Cells[e.ColumnIndex];
            cell.Value = cell.Value.ToString() == "0" ? "1" : "0";

        }
    }
}
