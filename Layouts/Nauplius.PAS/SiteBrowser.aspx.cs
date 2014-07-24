using System;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;

namespace Nauplius.PAS.Layouts.Nauplius.PAS
{
    public partial class SiteBrowser : LayoutsPageBase
    {
        readonly Guid _doclib = new Guid("00bfea71-e717-4e80-aa17-d0c71b360101");

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Page.IsPostBack) return;

            var webs = Web.GetSubwebsForCurrentUser();

            var rootNodeWebIcon = (string)SPUtility.MapWebToIcon(Web).First;
            var rootIcon = "/_layouts/15/images/" + rootNodeWebIcon;
            var rootWebNode = new TreeNode(Web.Title, Web.Url, rootIcon) { SelectAction = TreeNodeSelectAction.None };
            treeView1.Nodes.Add(rootWebNode);

            foreach (SPWeb web in webs)
            {
                var rootNode = treeView1.Nodes[0];
                var first = (string)SPUtility.MapWebToIcon(web).First;
                var icon = "/_layouts/15/images/" + first;
                var webNode = new TreeNode(web.Title, web.Url, icon) { SelectAction = TreeNodeSelectAction.None };
                rootNode.ChildNodes.Add(webNode);
                GetWebs(webNode, web);
                GetLibraries(webNode, web);
                web.Dispose();
            }

            GetLibraries(treeView1.Nodes[0], Web);

            treeView1.ExpandDepth = 1;
        }

        protected void treeView1_SelectedNodeChanged(object sender, EventArgs e)
        {
            var element = Request.QueryString["ParentElement"];
            var hidden = Request.QueryString["HiddenElement"];

            if (!string.IsNullOrEmpty(element))
            {
                var response =
                    string.Format(
                        "<script type='text/javascript'>var retArray = new Array; retArray.push(\'{0}\',\'{1}\',\'{2}\');" +
                        "window.frameElement.commitPopup(retArray);</script>", treeView1.SelectedNode.Value, element, hidden);
                Context.Response.Write(response);
                Context.Response.Flush();
                Context.Response.End();
            }
        }

        private void GetWebs(TreeNode topNode, SPWeb rootWeb)
        {
            var webs = rootWeb.GetSubwebsForCurrentUser();

            foreach (SPWeb web in webs)
            {
                var first = (string)SPUtility.MapWebToIcon(web).First;
                var icon = "/_layouts/15/images/" + first;
                var webNode = new TreeNode(web.Title, web.Url, icon) { SelectAction = TreeNodeSelectAction.None };
                topNode.ChildNodes.Add(webNode);
                GetLibraries(webNode, web);
                GetWebs(topNode, web);
                web.Dispose();
            }
        }

        private void GetLibraries(TreeNode topNode, SPWeb web)
        {
            foreach (SPList list in web.Lists)
            {
                if (list.TemplateFeatureId != _doclib || list.Hidden) continue;
                var libraryTreeNode = new TreeNode(list.Title, list.ParentWeb.Url + "/" + list.RootFolder.Url, list.ImageUrl)
                {
                    SelectAction = TreeNodeSelectAction.Select
                };
                topNode.ChildNodes.Add(libraryTreeNode);
                GetFolders(libraryTreeNode, list.RootFolder);
            }
        }

        private void GetFolders(TreeNode topNode, SPFolder rootFolder)
        {
            var query = new SPQuery { Folder = rootFolder };
            var web = rootFolder.ParentWeb;
            var listColl = web.Lists[rootFolder.ParentListId].GetItems(query);

            foreach (SPListItem listItem in listColl)
            {
                if (listItem.Folder == null) continue;
                var folderTreeNode = new TreeNode(listItem.Folder.Name, listItem.Folder.ParentWeb.Url + "/" + listItem.Folder.Url, "/_layouts/images/folder.gif")
                {
                    SelectAction = TreeNodeSelectAction.Select
                };
                topNode.ChildNodes.Add(folderTreeNode);
                GetFolders(folderTreeNode, listItem.Folder);
            }
        }
    }
}