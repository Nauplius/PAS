<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ConversionSettings.aspx.cs" Inherits="Nauplius.PAS.Layouts.Nauplius.PAS.ConversionSettings" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript">
        var notifyId = '';
        
        function PdfHelp(value) {
            var helpText = "";
            switch(value) {
            case "BitmapUnembeddableFonts":
                helpText = "Allow unembeddable fonts to be bitmaped.";
                break;
            case "FrameSlides":
                helpText = "Sets slides in frames.";
                break;
            case "IncludeDocumentProperties":
                helpText = "Include document properties.";
                break;
            case "IncludeDocumentStructureTags":
                helpText = "Include document structure tags.";
                break;
            case "IncludeHiddenSlides":
                helpText = "Include all hidden slides.";
                break;
            case "OptimizeForMinimumSize":
                helpText = "Optimize the output for minimim size.";
                break;
            case "UsePDFA":
                helpText = "Use PDF/A, an ISO standard for<br /> long-term document archival.";
                break;
            case "UseVerticalOrder":
                helpText = "Use vertical ordering.";
                break;
            }
            notifyId = SP.UI.Notify.addNotification(helpText, false);   
        }
        
        function PublishHelp(ddl) {
            var helpText = "";

            var slideTypeDdl = document.getElementById(ddl);
            var value = slideTypeDdl.options[slideTypeDdl.selectedIndex].value;

            if (notifyId != '') {
                RemoveHelp();
            }

            switch (value) {
                case "Slides":
                    helpText = "Specifies the Slides option, <br/ >which outputs a slide per page.";
                    break;
                case "Outline":
                    helpText = "Specifies the Outline option, <br/ >which outputs a slide outline per page.";
                    break;
                case "Handout1":
                    helpText = "Specifies the Handout1 option, <br/ >which outputs a slide handout per page.";
                    break;
                case "Handout2":
                    helpText = "Specifies the Handout2 option, <br/ >which outputs two slide handouts per page.";
                    break;
                case "Handout3":
                    helpText = "Specifies the Handout3 option, <br/ >which outputs three slide handouts per page.";
                    break;
                case "Handout4":
                    helpText = "Specifies the Handout4 option, <br/ >which outputs four slide handouts per page.";
                    break;
                case "Handout6":
                    helpText = "Specifies the Handout6 option, <br/ >which outputs six slide handouts per page.";
                    break;
                case "Handout9":
                    helpText = "Specifies the Handout9 option, <br/ >which outputs nine slide handouts per page.";
                    break;
                case "Default":
                    helpText = "Specifies the default option.";
                    break;
            }
            notifyId = SP.UI.Notify.addNotification(helpText, false);
        }
        
        function RemoveHelp() {
            SP.UI.Notify.removeNotification(notifyId);
            notifyId = '';
        }
        
        function Cancel() {
            window.frameElement.commonModalDialogClose(0);
        }
        
    </script>
</asp:Content>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <asp:PlaceHolder runat="server" ID="ph1" Visible="False" >
        <asp:Label ID="lblPubOpts" runat="server" Visible="False" Text="Slide Publish Options" />&nbsp;&nbsp;&nbsp;<SharePoint:DVDropDownList runat="server" ID="dvddl1" Visible="False"/><br/>
        <asp:Label runat="server" ID="lblPdfOps" Visible="False" /><br />
        <SharePoint:InputFormCheckBoxList runat="server" ID="cBoxList" Visible="False"/>
    </asp:PlaceHolder>
    <asp:PlaceHolder runat="server" ID="ph2" Visible="False">
        <asp:Label runat="server" ID="lblPicOpts" Text="Picture Options" Visible="False" />
        <br />
        <asp:Label runat="server" ID="lblWidth" Text="Slide Width (px)" Visible="False"/>
        <asp:TextBox runat="server" ID="txtWidth" Visible="False" MaxLength="10" TextMode="SingleLine" CausesValidation="true"/>
        <br />
        <asp:RangeValidator runat="server" ID="rVWidth" ControlToValidate="txtWidth" ForeColor="Red" />
        <br />
        <asp:Label runat="server" ID="lblHeight" Text="Slide Height (px)" Visible="False"/> 
        <asp:TextBox runat="server" ID="txtHeight" Visible="False" MaxLength="10" TextMode="SingleLine" CausesValidation="true"/>
        <br />
        <asp:RangeValidator runat="server" ID="rvHeight" ControlToValidate="txtHeight" ForeColor="Red" />
        <br />
    </asp:PlaceHolder>
    <asp:PlaceHolder runat="server" ID="ph3" />
    <br />
    <asp:Button runat="server" ID="btnSave" Text="Save" OnClick="btnSave_OnClick"/> <asp:Button runat="server" ID="btnCancel" Text="Cancel" OnClick="btnCancel_Click"/>
</asp:Content>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
Application Page
</asp:Content>

<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
My Application Page
</asp:Content>
