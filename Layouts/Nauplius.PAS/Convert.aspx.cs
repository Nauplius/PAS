using Microsoft.Office.Server.PowerPoint.Conversion;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Data;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Nauplius.PAS.Layouts.Nauplius.PAS
{
    public partial class Convert : LayoutsPageBase
    {
        protected void Page_Init(object sender, EventArgs e)
        {
            string[] items = Request["Items"] != null
                ? Request["Items"].Split(new string[] { "|" }, StringSplitOptions.RemoveEmptyEntries)
                : new string[] { };

            SPList list = null;

            try
            {
                using (SPSite site = new SPSite(SPContext.Current.Site.ID))
                {
                    using (SPWeb web = site.OpenWeb())
                    {
                        list = web.Lists[new Guid(Request["ListId"])];
                        if (list == null) return;
                        CreateTable(items, list);
                    }
                }
            }
            catch (SPException)
            {
                //ToDO: add exception logic
            }

            gvItems.EmptyDataText = "No files selected. Please select a supported PowerPoint file format.";
            btnCancel.Visible = true;
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack);
        }

        private void CreateTable(string[] items, SPList list)
        {
            var bSPFile = new BoundField {DataField = "File", HeaderText = "File"};
            bSPFile.ItemStyle.CssClass = "hide-file";
            bSPFile.HeaderStyle.CssClass = "hide-file";

            var bSPType = new BoundField {DataField = "Type", HeaderText = "Type", HtmlEncode = false};
            bSPType.ItemStyle.HorizontalAlign = HorizontalAlign.Center;
            bSPType.ItemStyle.Width = Unit.Pixel(35);
            bSPType.ItemStyle.Height = Unit.Pixel(35);

            var bSPFileName = new BoundField {DataField = "FileName", HeaderText = "File Name"};
            var bSPFileType = new BoundField {DataField = "FileType", HeaderText = "File Type"};
            var bSPNewName = new BoundField {DataField = "NewName", HeaderText = "New Name"};
            var bSPDestination = new BoundField {DataField = "Destination", HeaderText = "Destiniation"};
            var bSPButton = new BoundField {DataField = "Browse", HeaderText = "Browse"};

            var bSPWeb = new BoundField {DataField = "Web", HeaderText = "Web"};
            bSPWeb.ItemStyle.CssClass = "hide-web";
            bSPWeb.HeaderStyle.CssClass = "hide-web";

            var bSPSite = new BoundField {DataField = "Site", HeaderText = "Site"};
            bSPSite.ItemStyle.CssClass = "hide-site";
            bSPSite.HeaderStyle.CssClass = "hide-site";

            var bSettings = new BoundField {DataField = "Settings", HeaderText = "Settings"};
            
            var bSettingsOut = new BoundField {DataField = "SettingsOut", HeaderText = "SettingsOut"};
            bSettingsOut.ItemStyle.CssClass = "hide-settings";
            bSettingsOut.HeaderStyle.CssClass = "hide-settings";
            
            gvItems.Columns.Add(bSPFile);
            gvItems.Columns.Add(bSPType);
            gvItems.Columns.Add(bSPFileName);
            gvItems.Columns.Add(bSPFileType);
            gvItems.Columns.Add(bSPNewName);
            gvItems.Columns.Add(bSPDestination);
            gvItems.Columns.Add(bSPButton);
            gvItems.Columns.Add(bSPWeb);
            gvItems.Columns.Add(bSPSite);
            gvItems.Columns.Add(bSettings);
            gvItems.Columns.Add(bSettingsOut);

            var dt = new DataTable("PPTFiles");
            
            dt.Columns.Add("File", typeof (string));
            dt.Columns.Add("Type", typeof (string));
            dt.Columns.Add("FileName", typeof (string));
            dt.Columns.Add("FileType", typeof (DropDownList));
            dt.Columns.Add("NewName", typeof (string));
            dt.Columns.Add("Destination", typeof (Uri));
            dt.Columns.Add("Browse", typeof (Button));
            dt.Columns.Add("Web", typeof (string));
            dt.Columns.Add("Site", typeof (string));
            dt.Columns.Add("Settings", typeof (LinkButton));
            dt.Columns.Add("SettingsOut", typeof (TextBox));

            foreach (var id in items)
            {
                var listItem = list.GetItemById(Int32.Parse(id));

                if (listItem.FileSystemObjectType != SPFileSystemObjectType.File) continue;

                if (!ValidateFileFormat(listItem)) continue;

                var row = dt.NewRow();
                if(Site.ServerRelativeUrl == "/")
                    row["File"] = "/" + listItem.Url;
                else
                    row["File"] = Site.ServerRelativeUrl + "/" + listItem.Url;

                string docicon = SPUtility.ConcatUrls("/_layouts/images",
                    SPUtility.MapToIcon(listItem.Web, SPUtility.ConcatUrls(listItem.Web.Url, listItem.Url), "",
                        IconSize.Size16));
                row["Type"] = string.Format("<img src='{0}' />", docicon);
                row["FileName"] = listItem.DisplayName;
                row["FileType"] = null;
                row["NewName"] = string.Empty;
                row["Destination"] = null;
                row["Browse"] = null;
                row["Web"] = listItem.Web.Url;
                row["Site"] = listItem.Web.Site.Url;
                row["Settings"] = new LinkButton{ ID = "lBtn1" };
                row["SettingsOut"] = null;
                
                dt.Rows.Add(row);
            }

            gvItems.DataSource = dt;
            gvItems.DataBind();

            if (gvItems.Rows.Count < 1) return;
            btnConvert.Visible = true;
            btnCancel.Visible = true;
            chkWait.Visible = false;
            lblWait.Visible = false;
            p1.Visible = true;
        }

        protected void gvItems_OnRowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType != DataControlRowType.DataRow) return;
            var dataRow = (DataRowView) e.Row.DataItem;
            var itemUrl = dataRow["File"].ToString();

            using (SPSite site = new SPSite(SPContext.Current.Site.ID))
            {
                using (SPWeb web = Site.OpenWeb(itemUrl, false))
                {
                    var listItem = web.GetListItem(itemUrl); //try catch
                    e.Row.Cells[3].Controls.Add(OutputFileFormats(listItem));

                    var ddl1 = (DropDownList)e.Row.Cells[3].FindControl("dvddl1");
                    var txtNewName = new TextBox {ID = "txt1"};
                    var txtDest = new TextBox {ID = "txt2"};
                    var button = new Button();

                    e.Row.Cells[4].Controls.Add(txtNewName);

                    txtNewName.Attributes.Add("onBlur", "RewriteOutput(" + txtNewName.ClientID + "," + e.Row.Cells[2].ClientID + 
                        "," + e.Row.Cells[3].Controls[0].ClientID + "); return false;");
                    e.Row.Cells[5].Controls.Add(txtDest);
                    txtDest.ReadOnly = false;
                    button.Attributes.Add("onclick", "ShowLocationTree(" + txtDest.ClientID + "); return false;");
                    button.Text = "Browse";
                    e.Row.Cells[6].Controls.Add(button);

                    var lBtn = new LinkButton
                    {
                        CausesValidation = false,
                        CommandName = "Settings",
                        Text = "Settings"
                    };

                    var settingsOut = new TextBox {ID = "txtSO"};
                    e.Row.Cells[10].Controls.Add(settingsOut);

                    lBtn.Attributes.Add("onclick", "ShowSettings('" + e.Row.ClientID + "','" + ddl1.ClientID + "','" + e.Row.Cells[2].Text + "','" + settingsOut.ClientID + "'); return false;");
                    e.Row.Cells[9].Controls.Add(lBtn);
                }
            }
        }

        public static bool ValidateFileFormat(SPListItem item)
        {
            string fExt = item[SPBuiltInFieldId.DocIcon].ToString().ToLower();
            var fileFormats = new SupportedFileFormats();

            if (fileFormats.ReadFormats.Contains(fExt))
            {
                return true;
            }

            return false;
        }

        protected DropDownList OutputFileFormats(SPListItem listItem)
        {
            var dvddl1 = new DropDownList { ID = "dvddl1", Width = 60 };
            var fileFormats = new SupportedFileFormats();
            var fExt = listItem[SPBuiltInFieldId.DocIcon].ToString().ToLower();

            foreach (var fileFormat in fileFormats.WriteFormats.Where(fileFormat => fileFormat != fExt))
            {
                dvddl1.Items.Add(fileFormat);
            }

            dvddl1.SelectedIndex = 0;
            return dvddl1;
        }

        protected void InitializeConversion(object sender, EventArgs e)
        {
            btnConvert.Enabled = false;
            btnCancel.Enabled = false;

            var result = false;

            foreach (SPGridViewRow dataRow in gvItems.Rows)
            {
                if (dataRow.RowType != DataControlRowType.DataRow) continue;

                var file = dataRow.Cells[0].Text;
                var fileName = dataRow.Cells[2].Text;
                var ddl1 = (DropDownList) dataRow.Cells[3].FindControl("dvddl1");
                var fileNewType = ddl1.SelectedValue;
                var txtNewName = (TextBox) dataRow.Cells[4].FindControl("txt1");
                var fileNewName = txtNewName.Text;
                var txtDest = (TextBox) dataRow.Cells[5].FindControl("txt2");
                var fileDest = txtDest.Text;
                var fileWeb = dataRow.Cells[7].Text;
                var fileSite = dataRow.Cells[8].Text;
                var settingsOut = (TextBox)dataRow.Cells[10].FindControl("txtSO");
                var settings = settingsOut.Text;

                var slideType = PublishOption.Default;
                var fixedFormatSettings = new FixedFormatSettings(slideType);

                if (settings != "")
                {
                    var splitSettings = settings.Split(';');

                    if (fileNewType == splitSettings[0].Remove(0,2))
                    {
                        switch (fileNewType)
                        {
                            case "xps":
                            case "pdf":
                            {
                                PublishOption pubOption;
                                slideType =
                                    (PublishOption) Enum.Parse(typeof (PublishOption), splitSettings[1].Remove(0, 2));

                                var ffSettings = splitSettings.ToDictionary(val => val, val => true);

                                var properties = fixedFormatSettings.GetType().GetProperties().Where(p => p.CanWrite);

                                foreach (var property in properties)
                                {
                                    bool ffSetting;
                                    if (ffSettings.TryGetValue(property.Name, out ffSetting))
                                    {
                                        property.SetValue(fixedFormatSettings, ffSetting);
                                    }
                                }
                            }
                                break;
                                /*
                            case "jpg":
                            case "png":
                            {
                                var picSettings = new PictureSettings(pictureFormat, 
                                    System.Convert.ToUInt32(splitSettings[1].Remove(0, 2)), System.Convert.ToUInt32(splitSettings[2].Remove(0, 2)));
                            }
                                break;
                                 */
                        }
                    }
                }


                SPListItem listItem = null;
                try
                {
                    using (SPSite site = new SPSite(fileSite))
                    {
                        using (SPWeb web = site.OpenWeb(fileWeb))
                        {
                            listItem = web.GetListItem(file);
                        }
                    }
                }
                catch
                {
                    //ToDo: exception handling
                }

                var zipFile = "zipFile.zip";

                //Get the SPFolder of the List Item
                var folder = listItem.Folder == null ? Web.Folders[listItem.File.Web.Url + "/" + listItem.ParentList.RootFolder.Url]
                    : Web.Folders[Web.Url + "/" + listItem.Folder.Url];

                /*If the New Name TextBox is empty, identify if the new file type is a picture.  If it is a picture, save the file as zipFile, else
                 * take the original fileName, append the new file type, and save the variable as fileNewName.
                 * Else, if there is a valid new file name, but the format is a picture, save the zip file as the file's new name with a zip extension.
                */
                if (string.IsNullOrEmpty(fileNewName))
                {
                    if (fileNewType == "png" || fileNewType == "jpg")
                    {
                        zipFile = string.Format("{0}.{1}", fileName, "zip");
                    }
                    else
                    {
                        fileNewName = string.Format("{0}.{1}", fileName, fileNewType);
                    }
                }
                else if (!string.IsNullOrEmpty(fileNewName) && (fileNewType == "jpg" || fileNewType == "png"))
                {
                    zipFile = string.Format("{0}.{1}", fileNewName.Substring(0, fileNewName.LastIndexOf('.')), "zip");
                }

                //Get the folder object based on the URL passed into the file destiniation textbox.
                Uri uri;
                Uri.TryCreate(fileDest.Trim(), UriKind.Absolute, out uri);

                if (uri != null)
                {
                    using (SPSite site = new SPSite(uri.ToString()))
                    {
                        using (SPWeb web = site.OpenWeb())
                        {
                            folder = web.GetFolder(uri.ToString());
                        }
                    }
                }

                var wait = chkWait.Checked;

                switch (fileNewType)
                {
                    case "pptxS":
                        result = Conversion.ConvertToFormat(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(),
                                                    PresentationType.PptxStrict, folder, fileNewName, wait);
                        continue;
                    case "potx":
                        result = Conversion.ConvertToFormat(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(),
                                                    PresentationType.Potx, folder, fileNewName, wait);
                        continue;
                    case "ppsx":
                        result = Conversion.ConvertToFormat(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(),
                                                    PresentationType.Ppsx, folder, fileNewName, wait);
                        continue;
                    case "pptx":
                        result = Conversion.ConvertToFormat(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(),
                                                    PresentationType.Pptx, folder, fileNewName, wait);
                        continue;
                    case "ppt":
                        result = Conversion.ConvertToFormat(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(),
                                                    PresentationType.Default, folder, fileNewName, wait);
                        continue;
                    case "pdf":
                        result = Conversion.ConvertToPdf(listItem.Web.Site.Url, listItem.File, null, listItem[SPBuiltInFieldId.DocIcon].ToString(), folder, fixedFormatSettings, fileNewName);
                        continue;
                    case "xps":
                        result = Conversion.ConverToXps(listItem.Web.Site.Url, listItem.File, null,
                                                    listItem[SPBuiltInFieldId.DocIcon].ToString(), folder, fixedFormatSettings, fileNewName);
                        continue;
                    case "png":
                        result = Conversion.ConvertToPicture(listItem.Web.Site.Url, listItem.File, null,
                                                    listItem[SPBuiltInFieldId.DocIcon].ToString(), PictureFormat.Png, folder,
                                                        zipFile);
                        continue;
                    case "jpg":
                        result = Conversion.ConvertToPicture(listItem.Web.Site.Url, listItem.File, null,
                                                    listItem[SPBuiltInFieldId.DocIcon].ToString(), PictureFormat.Jpg, folder,
                                                        zipFile);
                        continue;
                }

            }

            Page.Response.Clear();
            Page.Response.Write(result
                                    ? "<script type=\"text/javascript\">window.frameElement.commonModalDialogClose(1);</script>"
                                    : "<script type=\"text/javascript\">window.frameElement.commonModalDialogClose(2);</script>");
            Page.Response.End();
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            Page.Response.Clear();
            Page.Response.Write("<script type=\"text/javascript\">window.frameElement.commonModalDialogClose(3);</script>");
            Page.Response.End();
        }
    }

    class SupportedFileFormats
    {
        private string[] _readFormats;
        private string[] _writeFormats;

        public string[] ReadFormats
        {
            get
            {
                return _readFormats = new[] {"pptx", "ppt"};
            }
        }

        public string[] WriteFormats
        {
            get
            {
                return _writeFormats = new[] {"ppt", "pptx", "potx", "ppsx", "pptxS", "jpg", "png", "pdf", "xps"};
            }
        }
    }
}