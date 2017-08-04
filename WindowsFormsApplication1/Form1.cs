using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace WindowsFormsApplication1
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label6_Click(object sender, EventArgs e)
        {
            MessageBox.Show("DO NOT PROCEED ! wait for site collection creation. Click on ok AFTER you can opened the site collection in browser.");
        }

        private void button1_Click(object sender, EventArgs e)
        {

            CreatDestinationSiteCollection(getSourceSiteCollectionSetting(textBox1.Text,textBox5.Text,textBox4.Text),textBox3.Text,comboBox1.Text,textBox2.Text,textBox7.Text, textBox8.Text );

        }



        private void CreatDestinationSiteCollection(SourceSiteCollectionSettings SourceSiteColl, string tenantname, string managedpath, string newurl,string username, string password)
        {
            string DestinationSiteCollectionURL = $"https://{tenantname}.sharepoint.com/{managedpath}/{newurl}";
            string tenantURL = $"https://{tenantname}-admin.sharepoint.com";

            //create root site collection at destination
            AuthenticationManager AM = new AuthenticationManager();
            using (ClientContext ctx = AM.GetSharePointOnlineAuthenticatedContextTenant(tenantURL, username, password))
            {
                try
                {
                    Tenant objTenant = new Tenant(ctx);

                    SiteCreationProperties newSite = new SiteCreationProperties();

                    newSite.Owner = username;
                    newSite.Title = SourceSiteColl.Title;
                    newSite.Url = DestinationSiteCollectionURL;
                    newSite.CompatibilityLevel = 15;
                    
                    if (SourceSiteColl.Template.ToUpper().Trim() == "BLANKINTERNET#0")
                    {
                        newSite.Template = "BLANKINTERNETCONTAINER#0";
                    }
                    else {
                        newSite.Template = SourceSiteColl.Template;
                    }
                    
                    newSite.Lcid = SourceSiteColl.LanguageID;
                    newSite.UserCodeMaximumLevel = 0;

                    SpoOperation oSpoOps = objTenant.CreateSite(newSite);
                    ctx.Load(oSpoOps, spo => spo.IsComplete);
                    ctx.ExecuteQuery();
                    MessageBox.Show("DO NOT PROCEED ! wait for site collection creation. Click on ok AFTER you can opened the site collection in browser.");

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //create libraies
                createSiteCollectionLists(SourceSiteColl, DestinationSiteCollectionURL, username, password);
                //create subsites
                createSubsites(SourceSiteColl, DestinationSiteCollectionURL, DestinationSiteCollectionURL,username,password );

            }
        }

        private void createSubsites(SourceSiteCollectionSettings sourceSiteColl, string destinationSiteCollectionURL, string SiteCollectionURL, string username, string password )
        {
            AuthenticationManager am = new AuthenticationManager();

            using (ClientContext ctx = am.GetSharePointOnlineAuthenticatedContextTenant(SiteCollectionURL,username,password))
            {
                foreach (SubsiteDef s in sourceSiteColl.Subsites)
                {
                    try
                    {

                        WebCreationInformation WCI = new WebCreationInformation();
                        WCI.Url = s.Url;
                        WCI.Title = s.Title;
                        WCI.WebTemplate = s.WebTemplate;
                        WCI.Language = int.Parse(s.Langauage.ToString());

                        Web w = ctx.Web.Webs.Add(WCI);

                        ctx.ExecuteQuery();
                        System.Threading.Thread.Sleep(5000);

                        //create the library

                        foreach (ListDef l in sourceSiteColl.Lists)
                        {
                            if (!ctx.Web.ListExists(l.Name))
                            {

                                try
                                {
                                    ListCreationInformation lci = new ListCreationInformation();
                                    lci.TemplateType = l.TemplateType;
                                    lci.Title = l.Name;
                                    lci.Description = l.Description;
                                    ctx.Site.RootWeb.Lists.Add(lci);
                                    ctx.ExecuteQuery();

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"error on moving {l.Name}, error: {ex.Message}");
                                }
                            }
                            else
                            {
                                //MessageBox.Show("library exist");
                            }
                        }

                        // end of library creation


                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }

                }
            }
        }

        private void createSiteCollectionLists(SourceSiteCollectionSettings sourceSiteColl, string SiteCollectionURL, string username, string password)
        {
            AuthenticationManager AM = new AuthenticationManager();

            using (ClientContext ctx = AM.GetSharePointOnlineAuthenticatedContextTenant(SiteCollectionURL, username, password))
            {
                foreach (ListDef l in sourceSiteColl.Lists)
                {
                    if (!ctx.Site.RootWeb.ListExists(l.Name))
                    {

                        try
                        {
                            ListCreationInformation lci = new ListCreationInformation();
                            lci.TemplateType = l.TemplateType;
                            lci.Title = l.Name;
                            lci.Description = l.Description;
                            ctx.Site.RootWeb.Lists.Add(lci);
                            ctx.ExecuteQuery();

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"error on moving {l.Name}, error: {ex.Message}");
                        }
                    }
                    else
                    {
                       // MessageBox.Show("library exist");
                    }
                }
            }

        }

        private SourceSiteCollectionSettings getSourceSiteCollectionSetting(string SiteCollURL, string username, string password)
        {

            List<ListDef> AllList = new List<ListDef>();

            SourceSiteCollectionSettings settings = new SourceSiteCollectionSettings();
            
            AuthenticationManager AM = new AuthenticationManager();

            using (ClientContext ctx = AM.GetSharePointOnlineAuthenticatedContextTenant(SiteCollURL, username, password))
            {
                
                //getting title, description, and base template
                ctx.Load(ctx.Site.RootWeb, rw => rw.Title, rw=>rw.Description,rw=>rw.Language, rw=>rw.Lists.Include(l=>l.Title, l=>l.Description, l=>l.BaseTemplate).Where(l=>l.Hidden != true));
                ctx.ExecuteQuery();
                foreach (List l in ctx.Site.RootWeb.Lists)
                {
                    AllList.Add(new ListDef() { Name = l.Title, TemplateType = l.BaseTemplate, Description = l.Description });
                }
                string s = ctx.Site.RootWeb.GetBaseTemplateId();
                ctx.ExecuteQuery();

                settings.Title = ctx.Site.RootWeb.Title;
                settings.Descrition = ctx.Site.RootWeb.Description;
                settings.Template = s;
                settings.Lists = AllList;
                settings.LanguageID = ctx.Site.RootWeb.Language;

                settings.Subsites = convertToSiteDef(ctx.Site.GetAllWebUrls().ToList<string>(),SiteCollURL,username,password);
                
                ctx.ExecuteQuery();

            }

            return settings;

    }

        private List<SubsiteDef> convertToSiteDef(List<string> Subsitelist, string siteCollURL, string username, string password)
        {
            List<SubsiteDef> AllSites = new List<SubsiteDef>();
            List<ListDef> AllList = new List<ListDef>();

            foreach (string s in Subsitelist)
            {
                if (!(s == siteCollURL))
                {
                    string webpath = s.Replace(siteCollURL + "/", "");
                    AuthenticationManager am = new AuthenticationManager();
                    using (ClientContext ctx = am.GetSharePointOnlineAuthenticatedContextTenant(s, username, password))
                    {
                        ctx.Load(ctx.Web, w=>w.Title, w=>w.WebTemplate,w=>w.Language, w=>w.Description, w=>w.ServerRelativeUrl,w=>w.Lists.Include(l => l.Title, l => l.Description, l => l.BaseTemplate).Where(l => l.Hidden != true));
                        ctx.ExecuteQuery();
                        foreach (List l in ctx.Web.Lists)
                        {
                            

                            AllList.Add(new ListDef() {  Name = l.Title, Description = l.Description, TemplateType = l.BaseTemplate});
                        }
                        

                        AllSites.Add(new SubsiteDef() { Title = ctx.Web.Title, Url = webpath, Langauage = ctx.Web.Language, WebTemplate = ctx.Web.WebTemplate, Lists = AllList });
                    }
                }
                
            }
            return AllSites;
        }
        
        class SourceSiteCollectionSettings
        {
            public string Title { get; set; }
            public string Descrition { get; set; }
            public uint LanguageID { get; set; }
            public string Template { get; set; }
            public List<ListDef> Lists { get; set; }
            public List<SubsiteDef> Subsites { get; set; }
        }

        class ListDef {

            public string Name { get; set; }
            public string Description { get; set; }
            public int TemplateType { get; set; }
            

        }
        class SubsiteDef {
            public string Url { get; set; }
            public string Title { get; set; }
            public string WebTemplate { get; set; }
            public uint Langauage { get; set; }
            public List<ListDef> Lists { get; set; }
        }

    }

}
