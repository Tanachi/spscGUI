using SC.API.ComInterop;
using SC.API.ComInterop.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Threading;
using System.Security;
using Microsoft.SharePoint.Client;

namespace Sharepointbuton
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window
    {
        bool isSharp = true;
        static BackgroundWorker backgroundWorker1 = new BackgroundWorker();
        SharpCloudApi sc = new SharpCloudApi();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void update_Click(object sender, RoutedEventArgs e)
        {
                backgroundWorker1.RunWorkerAsync();
                var progBar = sharpProgress;
                update.IsEnabled = false;
                var teamText = Team.Text;
                var storyText = Story.Text;
                var userText = Username.Text;
                sc = new SharpCloudApi(userText, Password.Password.ToString());
                try
                {
                    Console.WriteLine(sc.Stories());
                }
                catch
                {
                    MessageBox.Show("Invalid Username or Password");
                    update.IsEnabled = true;
                    return;
                }
                StoryLite[] teamBook = new StoryLite[30];
                try
                {
                    teamBook = sc.StoriesTeam(teamText);
                }
                catch
                {
                    MessageBox.Show("Invalid Team");
                    update.IsEnabled = true;
                    return;
                }
                Story dashStory = null;
                try
                {
                    dashStory = sc.LoadStory(storyText);

                }
                catch
                {
                    MessageBox.Show("Invalid Story");
                    update.IsEnabled = true;
                    return;
                }
                sharpProgress.Maximum = teamBook.Length - 1;
                sharpProgress.Value = 0;
                sharpProgress.Minimum = 1;

                // Adds new attributes if story does not have it
                if (dashStory.Attribute_FindByName("Appropriated Budget") == null)
                    dashStory.Attribute_Add("Appropriated Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                if (dashStory.Attribute_FindByName("RAG Status") == null)
                    dashStory.Attribute_Add("RAG Status", SC.API.ComInterop.Models.Attribute.AttributeType.List);
                if (dashStory.Attribute_FindByName("New Requested Budget") == null)
                    dashStory.Attribute_Add("New Requested Budget", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                if (dashStory.Attribute_FindByName("Project Business Value") == null)
                    dashStory.Attribute_Add("Project Business Value", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
                if (dashStory.Attribute_FindByName("Project Dependencies/Assumptions/Risks") == null)
                    dashStory.Attribute_Add("Project Dependencies/Assumptions/Risks", SC.API.ComInterop.Models.Attribute.AttributeType.Text);
                if (dashStory.Attribute_FindByName("Total Spent to Date") == null)
                    dashStory.Attribute_Add("Total Spent to Date", SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                Story tagStory = sc.LoadStory(teamBook[0].Id);

                // Add tags to new story
                foreach (var tag in tagStory.ItemTags)
                {
                    if (dashStory.ItemTag_FindByName(tag.Name) == null)
                    {
                        dashStory.ItemTag_AddNew(tag.Name, tag.Description, tag.Group);
                    }
                }
                MessageBox.Show("Updating");
                sharpProgress.Value++;
                foreach(StoryLite storyTeam in teamBook)
                {
                    Story story = sc.LoadStory(storyTeam.Id);
                    highCost(dashStory, story);
                    sharpProgress.Value++;
                }
                dashStory.Save();
                MessageBox.Show("Done");
                update.IsEnabled = true;
        }
        static void highCost(Story newStory, Story story)
        {
                String catName = "";
                String newCat = "empty";
                // Finds the category that contains the project
                foreach (var cat in story.Categories)
                {
                    var catLine = cat.Name.Split(' ');
                    // Checks to see if last word is project, if so, adds the other words to be a category for the dashboard
                    if (catLine[catLine.Length - 1] == "Projects")
                    {
                        catName = cat.Name;
                        String[] noProject = new String[catLine.Length - 1];
                        Array.Copy(catLine, noProject, catLine.Length - 1);
                        newCat = String.Join(" ", noProject);
                        if (newStory.Category_FindByName(newCat) == null)
                            newStory.Category_AddNew(newCat);
                    }
                }
                // Copies attribute data from team story to dashboard story
                foreach (var item in story.Items)
                {
                    // Checks to see if there's a bad item in the story
                    MatchCollection matchUrl = Regex.Matches(item.Name, @"Item \d+|(DELETE)");
                    // checks for category if there's no category with projects
                    if (newCat == "empty" && item.GetAllAttributeValues().Count > 8)
                    {
                        newCat = item.Category.Name;
                        catName = item.Category.Name;
                        if (newStory.Category_FindByName(newCat) == null)
                        {
                            newStory.Category_AddNew(newCat);
                        }
                    }
                    // Final check to see if item has attribute values for category
                    else if (newCat == "empty")
                    {
                        if (item.GetAttributeValueAsText(story.Attribute_FindByName("Appropriated Budget")) != null
                            || item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value")) != null)
                        {
                            newCat = item.Category.Name;
                            catName = item.Category.Name;
                            if (newStory.Category_FindByName(newCat) == null)
                            {
                                newStory.Category_AddNew(newCat);
                            }
                        }
                    }
                    double checkBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                    // Inserts item into dashboard
                    if (item.Category.Name == catName)
                    {
                        Item scItem = null;
                        if (newStory.Item_FindByName(item.Name) == null && item.Name != "" && matchUrl.Count == 0)
                        {
                            scItem = newStory.Item_AddNew(item.Name);
                            scItem.Category = newStory.Category_FindByName(newCat);
                            scItem.StartDate = Convert.ToDateTime(item.StartDate.ToString());
                            scItem.DurationInDays = item.DurationInDays;
                            scItem.Description = item.Description;
                            // Get values from story
                            if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget")) != null)
                            {
                                double appBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Appropriated Budget"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("Appropriated Budget"), appBudget);
                            }
                            if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget")) != null)
                            {
                                double newBudget = item.GetAttributeValueAsDouble(story.Attribute_FindByName("New Requested Budget"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("New Requested Budget"), newBudget);
                            }
                            if (item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status")) != null)
                            {
                                string RAG = item.GetAttributeValueAsText(story.Attribute_FindByName("RAG Status"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("RAG Status"), RAG);
                            }
                            if (item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value")) != null)
                            {
                                string value = item.GetAttributeValueAsText(story.Attribute_FindByName("Project Business Value"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("Project Business Value"), value);
                            }
                            if (item.GetAttributeValueAsText(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks")) != null)
                            {
                                string risks = item.GetAttributeValueAsText(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("Project Dependencies/Assumptions/Risks"), risks);
                            }
                            if (item.GetAttributeValueAsDouble(story.Attribute_FindByName("Project Dependencies/Assumptions/Risks")) != null)
                            {
                                double total = item.GetAttributeValueAsDouble(story.Attribute_FindByName("Total Spent to Date"));
                                scItem.SetAttributeValue(newStory.Attribute_FindByName("Total Spent to Date"), total);
                            }
                    }
 
  
                }
            }            
        }
        private void share_Click(object sender, RoutedEventArgs e)
        {
            string[] attr = { "Project Lead|Text", "Project Team|Text", "External ID|Numeric","Priority|Numberic",
            "Value Stream/LOB|List","RAG Status|List","Percent Complete|Numeric", "Due Date|Date","New Requested Budget|Numeric",
            "Appropriated Budget|Numeric","Project Business Value|Text","Project Dependencies/Assumptions/Risks|Text",
            "Status Comments|Text","Total Spent to Date|Numberic","Financial Comments|Text"};
            // Load from App.Config
            string sharpCloudUsername = shareSharpUsername.Text;
            string sharpCloudPassword = shareSharpPassword.Password.ToString();
            string sharePointUsername = shareUsername.Text;
            string sharePointPassword = sharePassword.Password.ToString();
            string sharePointSite = "https://hawaiioimt.sharepoint.com/sites/ets";

            // Login into sharpcloud
            var sc = new SharpCloudApi();
            try
            {
                sc = new SharpCloudApi(sharpCloudUsername, sharpCloudPassword);
            }
            catch
            {
                MessageBox.Show("Incorrect Sharpcloud info");
                return;
            }

            Story story = sc.LoadStory(shareStory.Text);
            // Login into Sharepoint
            var securePassword = new SecureString();
            foreach (var character in sharePointPassword)
            {
                securePassword.AppendChar(character);
            }
            securePassword.MakeReadOnly();
            var sharePointCredentials = new SharePointOnlineCredentials(sharePointUsername, securePassword);
            ClientContext context = new ClientContext(sharePointSite);
            context.Credentials = sharePointCredentials;
            Microsoft.SharePoint.Client.List list = context.Web.Lists.GetByTitle("Projects");
            // Setup query
            CamlQuery query = new CamlQuery();
            query.ViewXml = "<View/>";
            Microsoft.SharePoint.Client.ListItemCollection items = list.GetItems(query);
            // Loads list
            context.Load(list);
            context.Load(items);
            context.ExecuteQuery();
            // Goes through List
            // Adds Attribute to story if none exist
            addAttribute(story, attr);
            // Goes through List Items
            foreach (var item in items)
            {
            addItem(story, item, attr);
            }
            story.Save();
        }
        static void addAttribute(Story story, string[] attr)
        {
            foreach (var att in attr)
            {
                string[] split = att.Split('|');
                var name = split[0];
                var type = split[1];
                if (story.Attribute_FindByName(name) == null)
                {
                    if (type == "Text")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Text);
                    }
                    else if (type == "Numeric")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Numeric);
                    }
                    else if (type == "List")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.List);
                    }
                    else if (type == "Date")
                    {
                        story.Attribute_Add(name, SC.API.ComInterop.Models.Attribute.AttributeType.Date);
                    }

                }
            }
        }
        // Adds Attribute data to item
        static void addItem(Story story, Microsoft.SharePoint.Client.ListItem item, string[] attr)
        {
            Item storyItem = story.Item_AddNew(item["Title"].ToString());
            string[] shareAtt = { "Project_x0020_Lead", "Project_x0020_Team", "External_x0020_ID", "Priority","Category","Status",
                "Percent_x0020_Complete", "Due_x0020_Date","New_x0020_Requested_x0020_Budget","Appropriated_x0020_Budget",
                "Project_x0020_Business_x0020_Val","Project_x0020_Dependencies_x002f","Status_x0020_Comments","Total_x0020_Spent_x0020_to_x0020",
            "Financial_x0020_Comments" };

            for (var i = 0; i < attr.Length - 1; i++)
            {
                string[] split = attr[i].Split('|');
                var name = split[0];
                var type = split[1];
                if (item[shareAtt[i]] != null)
                {
                    if (type == "Text" || type == "List")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), item[shareAtt[i]].ToString());
                    }
                    else if (type == "Numeric")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), double.Parse(item[shareAtt[i]].ToString()));
                    }

                    else if (type == "Date")
                    {
                        storyItem.SetAttributeValue(story.Attribute_FindByName(name), DateTime.Parse(item[shareAtt[i]].ToString()));
                    }
                }
            }
            if (item["Notes"] != null)
                storyItem.Description = item["Notes"].ToString();
            if (item["Start_x0020_Date"] != null)
                storyItem.StartDate = DateTime.Parse(item["Start_x0020_Date"].ToString());
            if (item["Tags_x002e_ETS_x0020_Initiatives"] != null)
            {
                string[] tagSplit = item["Tags_x002e_ETS_x0020_Initiatives"].ToString().Split(',');
                foreach (var tag in tagSplit)
                {
                    if (story.ItemTag_FindByNameAndGroup(tag, "ETS Initiatives") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "ETS Initiatives");
                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
            if (item["Tags_x002e_Governor_x0020_Priori"] != null)
            {
                string[] tagSplit = item["Tags_x002e_Governor_x0020_Priori"].ToString().Split(',');
                foreach (var tag in tagSplit)
                {
                    if (story.ItemTag_FindByNameAndGroup(tag, "Governor Priorities") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "Governor Priorities");
                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
            if (item["Tags_x002e_ETS_x0020_Priorities"] != null)
            {
                string[] tagSplit = item["Tags_x002e_ETS_x0020_Priorities"].ToString().Split(',');
                foreach (var tag in tagSplit)
                {
                    if (story.ItemTag_FindByNameAndGroup(tag, "ETS Priorities") == null)
                    {
                        story.ItemTag_AddNew(tag, "", "ETS Priorities");

                        storyItem.Tag_AddNew(story.ItemTag_FindByName(tag));
                    }
                }
            }
        }
    }
}
