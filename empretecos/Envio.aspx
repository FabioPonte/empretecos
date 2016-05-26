<%@ Page Title="Envio" Language="C#" MasterPageFile="~/Default.Master" AutoEventWireup="true" CodeBehind="Envio.aspx.cs" Inherits="empretecos.Envio" %>
<asp:Content ID="Content_envio_1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content_envio_2" ContentPlaceHolderID="erContentPlaceHolder1" runat="server">
    
<div align="center">
</asp:Label><asp:FileUpload ID="FileUploadControl" runat="server"    /></div>

  <div align="center">  <asp:Button ID="Button_upload" runat="server" Text="Upload" 
          onclick="Button_upload_Click" /> </div>

          <div align="center">    <asp:GridView ID="GridView_arquivos" runat="server" 
                  BackColor="White" BorderColor="#DEDFDE" BorderStyle="None" BorderWidth="1px" 
                  CellPadding="4" ForeColor="Black" GridLines="Vertical">
              <AlternatingRowStyle BackColor="White" />
              <FooterStyle BackColor="#CCCC99" />
              <HeaderStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
              <PagerStyle BackColor="#F7F7DE" ForeColor="Black" HorizontalAlign="Right" />
              <RowStyle BackColor="#F7F7DE" />
              <SelectedRowStyle BackColor="#CE5D5A" Font-Bold="True" ForeColor="White" />
              <SortedAscendingCellStyle BackColor="#FBFBF2" />
              <SortedAscendingHeaderStyle BackColor="#848384" />
              <SortedDescendingCellStyle BackColor="#EAEAD3" />
              <SortedDescendingHeaderStyle BackColor="#575357" />
                                    </asp:GridView>
    </div>

<div align="center"><asp:Label ID="Label_result_upload" runat="server" Text=""></asp:Label></div>
</asp:Content>
