using Microsoft.Office.Server.PowerPoint.Conversion;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Nauplius.PAS.Layouts.Nauplius.PAS
{
    public partial class ConversionSettings : LayoutsPageBase
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack) return;
            var fileType = Request.QueryString["fileType"];

            if (fileType == "pdf" || fileType == "xps")
            {
                lblPdfOps.Text = fileType == "pdf" ? "PDF Options" : "XPS Options";

                ListItem[] list =
                {
                    new ListItem("Bitmap Unembeddable Fonts", "BitmapUnembeddableFonts"), new ListItem("Frame Slides", "FrameSlides"),
                    new ListItem("Include Document Properties", "IncludeDocumentProperties"), 
                    new ListItem("Include Document Structure Tags", "IncludeDocumentStructureTags"), 
                    new ListItem("Include Hidden Slides", "IncludeHiddenSlides"),
                    new ListItem("Optimize for Minimum Size", "OptimizeForMinimumSize"), new ListItem("Use PDF/A", "UsePdfA"), 
                    new ListItem("Use Vertical Order", "UseVerticalOrder"), 
                };

                foreach (var listItem in list)
                {
                    listItem.Attributes.Add("onmouseover", "PdfHelp('" + listItem.Value + "')");
                    listItem.Attributes.Add("onmouseout", "RemoveHelp()");
                }

                cBoxList.Items.AddRange(list);
                ph1.Visible = true;
                lblPdfOps.Visible = true;
                cBoxList.Visible = true;
                PublishOptions();
            }
                /*
            else if (fileType == "jpg" || fileType == "png")
            {
                rVWidth.MinimumValue = System.Convert.ToString(1);
                rVWidth.MaximumValue = System.Convert.ToString(UInt32.MaxValue);
                rVWidth.ErrorMessage = string.Format("{0} must be between {1} and {2} pixels.", "Width",
                    1, UInt32.MaxValue);

                rvHeight.MinimumValue = System.Convert.ToString(1);
                rvHeight.MaximumValue = System.Convert.ToString(UInt32.MaxValue);
                rvHeight.ErrorMessage = string.Format("{0} must be between {1} and {2} pixels.", "Height",
                    1, UInt32.MaxValue);

                lblPicOpts.Visible = true;
                lblWidth.Visible = true;
                lblHeight.Visible = true;
                txtHeight.Visible = true;
                txtWidth.Visible = true;
                ph2.Visible = true;
            }*/
            else
            {
                var lblNoOpts = new Label {Text = "There are no options for this file type."};
                var lcBR = new LiteralControl("<br />");
                ph3.Controls.Add(lblNoOpts);
                ph3.Controls.Add(lcBR);
                ph3.Visible = true;
                btnSave.Enabled = false;
            }
            
        }

        protected void btnSave_OnClick(object sender, EventArgs e)
        {
            var fileType = Request.QueryString["fileType"];
            var element = Request.QueryString["ParentElement"];
            var fileName = Request.QueryString["fileName"];
            var fileSettings = Request.QueryString["settings"];

            if (fileType == "pdf" || fileType == "xps")
            {

                if (!string.IsNullOrEmpty(element))
                {
                    var pdfOptsOut = new List<string> {"x:" + fileType + ";s:" + dvddl1.SelectedValue};

                    pdfOptsOut.AddRange(from ListItem li in cBoxList.Items where li.Selected select li.Value);

                    var pdfOptsRtn = string.Join(";", pdfOptsOut);
                    var response =
                        string.Format(
                            "<script type='text/javascript'>var retArray = new Array; retArray.push(\'{0}\',\'{1}\',\'{2}\',\'{3}\');" +
                            "window.frameElement.commitPopup(retArray);</script>", pdfOptsRtn, element, fileName, fileSettings);
                    Context.Response.Write(response);
                    Context.Response.Flush();
                    Context.Response.End();
                }
            }
            else if (fileType == "jpg" || fileType == "png")
            {
                //ToDo: Check dimentions prior to return.
                var picWidth = txtWidth.Text;
                var picHeight = txtHeight.Text;
                string[] picDimentions = {picWidth, picHeight};
                var picOptsOut = "x:" + fileType + ";" + string.Join(";", picDimentions);

                var response =
                     string.Format(
                         "<script type='text/javascript'>var retArray = new Array; retArray.push(\'{0}\',\'{1}\',\'{2}\',\'{3}\');" +
                         "window.frameElement.commitPopup(retArray);</script>", picOptsOut, element, fileName, fileSettings);
                Context.Response.Write(response);
                Context.Response.Flush();
                Context.Response.End();
            }
            else
            {
                //do nothing
            }
        }

        protected void btnCancel_Click(object sender, EventArgs e)
        {
            Page.Response.Clear();
            Page.Response.Write("<script type=\"text/javascript\">window.frameElement.commonModalDialogClose(0);</script>");
            Page.Response.End();
        }

        internal void PublishOptions()
        {
            dvddl1.DataSource = Enum.GetNames(typeof (PublishOption));
            dvddl1.Attributes.Add("onChange", "PublishHelp('" + dvddl1.ClientID + "')");
            dvddl1.DataBind();
            lblPubOpts.Visible = true;
            dvddl1.Visible = true;
        }
        internal void PdfOptions(string sV)
        {
            PublishOption publishOption;
            Enum.TryParse(sV, out publishOption);
            var ffs = new FixedFormatSettings(publishOption); 
        }
    }
}