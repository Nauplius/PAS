<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Convert.aspx.cs" Inherits="Nauplius.PAS.Layouts.Nauplius.PAS.Convert" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <p id="p1" runat="server" Visible="False">
        For each file, select the file type to convert to, and optionally enter a new file name with no extension.
    </p>
    <br/>
    <SharePoint:SPGridView runat="server" ID="gvItems" AutoGenerateColumns="False" Enabled="True" OnRowDataBound="gvItems_OnRowDataBound" EnableViewState="True">
    </SharePoint:SPGridView>
    <br/>
    <div id="spinwait" class="wait">
        <br />
        <br />
        <br />
        <br />
        <br />
    </div>
    <div id="textwait">
        Please wait...
    </div>
    <asp:Button runat="server" ID="btnConvert" Text="Convert" Visible="False" OnClick="InitializeConversion" UseSubmitBehavior="False" OnClientClick=" runSpinner() "/> <asp:Button runat="server" ID="btnCancel" Text="Cancel" Visible="False" OnClick="btnCancel_Click"/>
    <asp:CheckBox runat="server" ID="chkWait" Visible="False"/><asp:Label runat="server" ID="lblWait" Text="Wait for conversion process" Visible="False" />
   </asp:Content>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="js/spin.min.js"></script>
    <script type="text/javascript" src="js/jquery-1.9.1.min.js"></script>
    <style type="text/css">
        .hide-file, .hide-site, .hide-web, .hide-settings   
        {
            display: none;
        }
        #textwait, .wait { text-align: center; }
    </style>
    <script type="text/javascript">
        function hideWaitDiv() {
            var spinDiv = $('#spinwait');
            var textDiv = $('#textwait');
            spinDiv.hide();
            textDiv.hide();
        }

        _spBodyOnLoadFunctionNames.push("hideWaitDiv");

        function RewriteOutput(elementId, inputFile, dropDownList) {
            var text = document.getElementById(elementId.id);
            var ddl = document.getElementById(dropDownList.id).value;

            if (text.value.indexOf('.') !== -1) {
                text.value = text.value.substr(0, text.value.lastIndexOf('.'));
            }

            text.value = text.value + "." + ddl;

            if (text.value == "." + ddl) {
                text.value = "";
            }
        }

        function ShowLocationTree(elementId) {
            var tBox = document.getElementById(elementId.id);
            var siteBrowserUrl = "";
            
            if (_spPageContextInfo.siteServerRelativeUrl == "/") {
                siteBrowserUrl = "/_layouts/15/Nauplius.PAS/SiteBrowser.aspx?ParentElement=" + tBox.id + "&IsDlg=1";
            } else {
                siteBrowserUrl = _spPageContextInfo.siteServerRelativeUrl + "/_layouts/15/Nauplius.PAS/SiteBrowser.aspx?ParentElement=" + tBox.id + "&IsDlg=1";
            }
            
            var options = {
                url: siteBrowserUrl,
                args: null,
                title: 'Save Location',
                dialogReturnValueCallback: dialogCallback,
            };
            SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);

            function dialogCallback(dialogResult, returnValue) {
                if (returnValue != null) {
                    var tBox1 = document.getElementById(returnValue[1]);
                    if (document.all) {
                        tBox1.innerText = returnValue[0]; //IE8 and below support
                    } else {
                        tBox1.textContent = returnValue[0]; //Everything else
                    }
                }
            }
        }

        function ShowSettings(rowId, fileTypeDropDownList, fileName, fileSettings) {

            var itemSettingsUrl = "";
            var fileTypeDdl = document.getElementById(fileTypeDropDownList);
            var fileType = fileTypeDdl.options[fileTypeDdl.selectedIndex].value;
            
            if (_spPageContextInfo.siteServerRelativeUrl == "/") {
                itemSettingsUrl = "/_layouts/15/Nauplius.PAS/ConversionSettings.aspx?ParentElement=" + rowId + "&FileType=" + fileType +
                    "&FileName=" + fileName + "&Settings=" + fileSettings + "&IsDlg=1";
            } else {
                itemSettingsUrl = _spPageContextInfo.siteServerRelativeUrl + "/_layouts/15/Nauplius.PAS/ConversionSettings.aspx?ParentElement=" + rowId + "&FileType=" + fileType +
                    "&FileName=" + fileName + "&Settings=" + fileSettings + "&IsDlg=1";
            }

            var options = {
                url: itemSettingsUrl,
                args: null,
                title: 'Conversion Settings for ' + fileName,
                dialogReturnValueCallback: dialogCallback,
            };
            SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
            
            function dialogCallback(dialogResult, returnValue) {
                if (dialogResult == SP.UI.DialogResult.OK) {
                    var settings = returnValue[0];

                    document.getElementById(fileSettings).innerText = settings;
                    
                    var tBox1 = document.getElementById(returnValue[3]);
                    if (document.all) {
                        tBox1.innerText = settings; //IE8 and below support
                    } else {
                        tBox1.textContent = settings; //Everything else
                    }
                }
            }
        }

        var opts = {
            lines: 11,
            length: 13,
            width: 4,
            radius: 15,
            corners: 0,
            rotate: 0,
            direction: 1,
            color: '#000',
            speed: 1.1,
            trail: 47,
            shadow: true,
            hwaccel: false,
            className: 'wait',
            zIndex: 2e9
        };
        var spinner;

        function runSpinner() {
            var target = document.getElementById('spinwait');
            var ph1 = $('#ctl00_PlaceHolderMain_p1');
            var table = $('#ctl00_PlaceHolderMain_gvItems');
            var spinDiv = $('#spinwait');
            var textDiv = $('#textwait');
            var btnOk = $('#ctl00_PlaceHolderMain_btnConvert');
            var btnCan = $('#ctl00_PlaceHolderMain_btnCancel');
            btnOk.attr("disabled", "disabled");
            btnCan.attr("disabled", "disabled");
            table.hide();
            ph1.hide();
            spinDiv.show();
            if (typeof (spinner) == 'undefined') {
                spinner = new Spinner(opts).spin(target);
            }
            textDiv.show();
        }
    </script>
</asp:Content>


<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
Nauplius.PAS [PowerPoint Automation Services]
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
</asp:Content>